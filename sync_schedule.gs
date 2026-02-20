/**
 * 設定項目
 */
const CONFIG = {
  // ■ スプレッドシート設定
  DEST_SPREADSHEET_ID: '',
  DEST_SHEET_NAME: '【進行入力用】レタッチスケジュール', // 同期先の集約用シート名

  // ■ Notion設定
  NOTION_API_KEY: '',
  NOTION_DB_ID: ''
};

/**
 * メニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('同期メニュー')
    .addItem('【表示中のシート】を集約シート＆Notionへ同期', 'syncActiveSheet')
    .addToUi();
}

/**
 * アクティブなシートのデータを同期先に統合し、Notionへも連携する関数
 * 修正点: 確認出し・納品日が複数ある場合、日付が古い方（最初に見つかった方）を優先
 */
function syncActiveSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. 同期元の特定
  const srcSs = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = srcSs.getActiveSheet();
  const srcSheetName = srcSheet.getName();

  // 2. 同期先の特定
  let destSs;
  try {
    destSs = SpreadsheetApp.openById(CONFIG.DEST_SPREADSHEET_ID);
  } catch (e) {
    ui.alert('同期先スプレッドシートが見つかりません。IDを確認してください。');
    return;
  }
  
  const destSheet = destSs.getSheetByName(CONFIG.DEST_SHEET_NAME);
  if (!destSheet) {
    ui.alert(`同期先シート「${CONFIG.DEST_SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 確認ダイアログ
  if (ui.alert('同期実行', `シート「${srcSheetName}」のデータを同期しますか？\n（Notionへの反映も行います）`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // --- 3. 同期元データの取得 ---
  const srcLastCol = srcSheet.getLastColumn();
  const srcStartCol = 7; // G列
  if (srcLastCol < srcStartCol) { ui.alert('データ範囲不正: G列以降にデータがありません'); return; }

  const srcPerson = srcSheet.getRange("D6").getValue(); 
  const srcProof = srcSheet.getRange("D8").getValue();

  // ヘッダー情報 (Row7:Cut, Row8:Price, Row11:Title)
  const srcHeaderValues = srcSheet.getRange(7, srcStartCol, 5, srcLastCol - srcStartCol + 1).getValues();

  // 予定データの取得
  const srcLastRow = srcSheet.getLastRow();
  const dateStartRow = 12;
  const dateRowCount = srcLastRow - dateStartRow + 1;
  if (dateRowCount < 1) { ui.alert('予定データがありません'); return; }

  const srcDates = srcSheet.getRange(dateStartRow, 1, dateRowCount, 1).getValues().flat();
  const srcSchedules = srcSheet.getRange(dateStartRow, srcStartCol, dateRowCount, srcLastCol - srcStartCol + 1).getValues();

  // --- 4. 同期先の準備 ---
  const destLastRow = destSheet.getLastRow();
  let destDateMap = new Map();
  if (destLastRow >= 12) {
    const destDates = destSheet.getRange(12, 1, destLastRow - 12 + 1, 1).getValues().flat();
    destDates.forEach((d, idx) => {
      if (d instanceof Date) destDateMap.set(d.getTime(), idx + 12);
    });
  } else {
    ui.alert('同期先の日付設定(A12以降)がありません。');
    return;
  }

  const destLastCol = destSheet.getLastColumn();
  const destStartCol = 3; // C列
  let destProjectNames = [];
  if (destLastCol >= destStartCol) {
    destProjectNames = destSheet.getRange(3, destStartCol, 1, destLastCol - destStartCol + 1).getValues()[0];
  }

  // --- 5. 同期ループ ---
  for (let i = 0; i < srcHeaderValues[0].length; i++) {
    const projectName = srcHeaderValues[4][i]; // 案件名
    const cutCount = srcHeaderValues[0][i];    // カット数
    const price = srcHeaderValues[1][i];       // 売価

    if (!projectName) continue;

    let targetColIndex = -1;
    const existingIndex = destProjectNames.indexOf(projectName);

    // ▼ スプレッドシート列確保・書式コピー処理
    if (existingIndex !== -1) {
      targetColIndex = destStartCol + existingIndex;
    } else {
      const currentLastCol = destSheet.getLastColumn();
      destSheet.insertColumnAfter(currentLastCol);
      targetColIndex = currentLastCol + 1;
      
      // 書式コピー（1行目～9行目）
      if (currentLastCol >= destStartCol) {
        // --- 1. 全体（1〜9行目）の「書式（色・枠線）」をコピー ---
        const sourceRange = destSheet.getRange(1, currentLastCol, 9, 1);
        const targetRange = destSheet.getRange(1, targetColIndex);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT);

        // --- 2. 2行目だけ「数式」をコピー ---
        const sourceRow2 = destSheet.getRange(2, currentLastCol); // 左隣の2行目
        const targetRow2 = destSheet.getRange(2, targetColIndex); // 新しい列の2行目
        sourceRow2.copyTo(targetRow2, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
      }
      destProjectNames.push(projectName);
    }

    // ▼ スプレッドシート：ヘッダー書き込み
    destSheet.getRange(3, targetColIndex).setValue(projectName);
    destSheet.getRange(4, targetColIndex).setValue(cutCount);
    destSheet.getRange(5, targetColIndex).setValue(price);
    destSheet.getRange(6, targetColIndex).setValue(srcPerson);
    destSheet.getRange(7, targetColIndex).setValue(srcProof);

    // ▼ スプレッドシート：予定書き込み & 日付抽出
    let deliveryDate = null;     // 納品日
    let confirmationDate = null; // 確認出し日

    for (let d = 0; d < srcDates.length; d++) {
      const srcDate = srcDates[d];
      const val = srcSchedules[d][i];
      
      if (srcDate instanceof Date) {
        const destRow = destDateMap.get(srcDate.getTime());
        if (destRow) {
          destSheet.getRange(destRow, targetColIndex).setValue(val);

          // ★修正箇所：日付抽出ロジック（古い方を優先）
          const valStr = String(val);
          
          // 「納品」が含まれていて、かつ まだ日付が入っていない場合のみセット
          if (val && valStr.indexOf('納品') !== -1) {
            if (deliveryDate === null) {
              deliveryDate = srcDate;
            }
          }
          
          // 「確認出し」が含まれていて、かつ まだ日付が入っていない場合のみセット
          if (val && valStr.indexOf('確認出し') !== -1) {
            if (confirmationDate === null) {
              confirmationDate = srcDate;
            }
          }
        }
      }
    }

    // 抽出日付の書き込み
    if (confirmationDate) destSheet.getRange(9, targetColIndex).setValue(confirmationDate);
    if (deliveryDate) destSheet.getRange(10, targetColIndex).setValue(deliveryDate);

    // ▼ Notion同期処理の実行
    try {
      syncToNotion({
        title: projectName,
        proof: srcProof,
        cut: cutCount,
        confirmationDate: confirmationDate,
        deliveryDate: deliveryDate
      });
    } catch (e) {
      console.error(`Notion sync failed for ${projectName}: ${e.message}`);
    }
  }

  ui.alert('同期が完了しました。');
}

/**
 * Notionへのデータ登録・更新を行う関数
 * 修正点: 新規登録時のみ Priority (Select) を Middle に設定
 */
function syncToNotion(data) {
  if (!CONFIG.NOTION_API_KEY || !CONFIG.NOTION_DB_ID) return;

  // --- 1. 既存ページの検索 ---
  
  const searchUrl = 'https://api.notion.com/v1/databases/' + CONFIG.NOTION_DB_ID + '/query';
  const searchPayload = {
    filter: {
      property: 'Name',
      title: {
        equals: data.title
      }
    }
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.NOTION_API_KEY,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(searchPayload),
    muteHttpExceptions: true
  };

  let pageId = null;
  try {
    const searchRes = UrlFetchApp.fetch(searchUrl, options);
    if (searchRes.getResponseCode() !== 200) {
      console.error('Notion Search Error:', searchRes.getContentText());
      return; 
    }
    const searchJson = JSON.parse(searchRes.getContentText());
    if (searchJson.results && searchJson.results.length > 0) {
      pageId = searchJson.results[0].id;
    }
  } catch (e) {
    console.error('Notion Search Exception:', e.message);
    return;
  }

  // --- 2. プロパティの構築 ---
  
  const formatDate = (dateObj) => {
    if (!dateObj) return null;
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  };

  const proofVal = String(data.proof);
  const isProofChecked = (proofVal.indexOf('有') !== -1 || proofVal.indexOf('あり') !== -1);

  // 基本プロパティ（更新時・新規時共通）
  const properties = {
    'Name': {
      title: [{ text: { content: data.title } }]
    },
    'Cut': {
      number: Number(data.cut) || 0
    },
    'Proof': {
      checkbox: isProofChecked
    },
    '確認出し': {
      date: data.confirmationDate ? { start: formatDate(data.confirmationDate) } : null
    },
    '納品日': {
      date: data.deliveryDate ? { start: formatDate(data.deliveryDate) } : null
    }
  };

  // --- 3. 送信処理 (Create or Update) ---

  let targetUrl, method, payloadObj;

  if (pageId) {
    // ■ 更新 (UPDATE)
    // 既存ページがある場合はPriorityを変更せず、他の情報だけ更新します
    targetUrl = 'https://api.notion.com/v1/pages/' + pageId;
    method = 'patch';
    payloadObj = { properties: properties };
  } else {
    // ■ 新規作成 (CREATE)
    targetUrl = 'https://api.notion.com/v1/pages';
    method = 'post';
    
    // ★追加: 新規作成時のみ Priority: Middle を追加
    properties['Priority'] = {
      select: { name: 'Middle' }
    };

    payloadObj = {
      parent: { database_id: CONFIG.NOTION_DB_ID },
      properties: properties
    };
  }

  const actionOptions = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + CONFIG.NOTION_API_KEY,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payloadObj),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(targetUrl, actionOptions);
  
  if (response.getResponseCode() !== 200) {
    console.error('Notion Sync Error:', response.getContentText());
  }
}
