const ss = SpreadsheetApp.getActiveSpreadsheet();
const masterSheet = ss.getSheetByName("マスタ"); 
const historySheet = ss.getSheetByName("履歴");
const configSheet = ss.getSheetByName("設定");

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('QR在庫管理アプリPro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 初期設定（倉庫リストと担当者リスト）をアプリに渡す
function getAppConfig() {
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return { locations: [], users: [] };
  
  const locations = configSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  const users = configSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(String);
  
  return { locations: locations, users: users };
}

// 商品情報を探す
function getItemInfo(code, locationName) {
  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  
  const stockColIndex = headers.indexOf(locationName);
  if (stockColIndex === -1) return { error: "倉庫が見つかりません。マスタの列名と設定シートの名前が合っているか確認してください。" };

  const timeColIndex = headers.indexOf(locationName + " 更新");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(code)) {
      let lastUpdate = "--";
      if (timeColIndex !== -1 && data[i][timeColIndex]) {
        lastUpdate = Utilities.formatDate(new Date(data[i][timeColIndex]), "JST", "MM/dd HH:mm");
      }

      return {
        row: i + 1,
        col: stockColIndex + 1,
        timeCol: (timeColIndex !== -1) ? timeColIndex + 1 : null, 
        name: data[i][1],
        stock: data[i][stockColIndex] || 0,
        lastUpdate: lastUpdate
      };
    }
  }
  return null;
}

// 商品を検索する（コードまたは名前で部分一致）
function searchItems(query, locationName) {
  if (!query || query.length < 2) return [];
  
  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  const stockColIndex = headers.indexOf(locationName);
  
  if (stockColIndex === -1) return [];
  
  const results = [];
  const queryLower = query.toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0]);
    const name = String(data[i][1]);
    
    if (code.toLowerCase().includes(queryLower) || name.toLowerCase().includes(queryLower)) {
      results.push({
        code: code,
        name: name,
        stock: data[i][stockColIndex] || 0
      });
      
      if (results.length >= 10) break;
    }
  }
  
  return results;
}

// 在庫更新（排他制御付き）
function updateStock(code, change, locationName, userName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return "Error: アクセスが混み合っています。もう一度押してください。";
  }

  try {
    const item = getItemInfo(code, locationName);
    if (!item || item.error) return "Error";

    const newStock = parseInt(item.stock) + parseInt(change);
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd HH:mm:ss");
    const shortTimeStr = Utilities.formatDate(timestamp, "JST", "MM/dd HH:mm");

    masterSheet.getRange(item.row, item.col).setValue(newStock);
    if (item.timeCol) {
      masterSheet.getRange(item.row, item.timeCol).setValue(timeStr);
    }

    if (historySheet) {
      historySheet.appendRow([timeStr, item.name, change, newStock, locationName, userName]);
    }

    return {
      newStock: newStock,
      lastUpdate: shortTimeStr
    };

  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}
