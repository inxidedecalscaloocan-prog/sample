// ============================================================================
// INXIDE MOTORSHOP INVENTORY SYSTEM — GOOGLE APPS SCRIPT BACKEND (Code.gs)
// Linked Google Sheet ID: 1qxj8OhQ6dMbw0dZCiMpgmJ2Kr8MnJv_nmMVuBDR-4S0
// ============================================================================

const SPREADSHEET_ID = "1qxj8OhQ6dMbw0dZCiMpgmJ2Kr8MnJv_nmMVuBDR-4S0";

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Inxide Motorshop Database API Active" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureSheets(ss);

    if (action === "getAll") {
      const data = {
        inventory: getSheetRecords(ss, "Inventory"),
        stockIn: getSheetRecords(ss, "StockIn"),
        stockOut: getSheetRecords(ss, "StockOut"),
        orderList: getSheetRecords(ss, "OrderList"),
        quotes: getSheetRecords(ss, "SavedQuotes"),
        quoteBorders: getSheetRecords(ss, "QuoteBorders"),
        permits: getSheetRecords(ss, "Permits"),
        plasticCatalog: getSheetRecords(ss, "PlasticCatalog"),
        priceLists: getSheetRecords(ss, "PriceLists"),
        suppliers: getSheetRecords(ss, "Suppliers"),
        settings: getSettingsMap(ss)
      };
      return jsonResponse({ status: "success", data: data });
    }

    if (action === "syncTable") {
      const sheetName = payload.sheetName;
      const rows = payload.rows || [];
      saveSheetRecords(ss, sheetName, rows);
      return jsonResponse({ status: "success", sheet: sheetName });
    }

    if (action === "syncAll") {
      const d = payload.data || {};
      const tableMap = {
        inventory: "Inventory",
        stockIn: "StockIn",
        stockOut: "StockOut",
        orderList: "OrderList",
        quotes: "SavedQuotes",
        quoteBorders: "QuoteBorders",
        permits: "Permits",
        plasticCatalog: "PlasticCatalog",
        priceLists: "PriceLists",
        suppliers: "Suppliers"
      };
      Object.keys(tableMap).forEach(function(k) {
        if (Array.isArray(d[k])) {
          saveSheetRecords(ss, tableMap[k], d[k]);
        }
      });
      if (d.settings) saveSettingsMap(ss, d.settings);
      return jsonResponse({ status: "success" });
    }

    if (action === "saveSettings") {
      saveSettingsMap(ss, payload.settings || {});
      return jsonResponse({ status: "success" });
    }

    return jsonResponse({ status: "error", message: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

const SHEET_SCHEMAS = {
  "Inventory": ["id", "partNumber", "productName", "cost", "price", "initialQty", "currentQty", "warningMode", "manualWarning", "brand", "totalSold", "photosJson"],
  "StockIn": ["id", "partNumber", "productName", "qty", "date", "supplier"],
  "StockOut": ["id", "partNumber", "productName", "qty", "platform", "date", "unitPrice", "unitCost"],
  "OrderList": ["id", "partNumber", "productName", "suggestedQty", "orderQty"],
  "SavedQuotes": ["id", "name", "date", "itemsJson", "discountType", "discountValue", "photosJson", "borderId"],
  "QuoteBorders": ["id", "name", "imageUrl"],
  "Permits": ["id", "name", "fileData", "fileName", "fileType"],
  "PlasticCatalog": ["id", "name", "fileData", "fileName", "fileType"],
  "PriceLists": ["id", "name", "driveLink"],
  "Suppliers": ["id", "name", "contactPerson", "address", "contactNo", "notes", "pic1", "pic2"],
  "Settings": ["key", "value"]
};

function ensureSheets(ss) {
  Object.keys(SHEET_SCHEMAS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(SHEET_SCHEMAS[name]);
    }
  });
}

function getSheetRecords(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = SHEET_SCHEMAS[name];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.map(function(row) {
    const obj = {};
    headers.forEach(function(h, idx) {
      obj[h] = row[idx];
    });
    return obj;
  });
}

function saveSheetRecords(ss, name, records) {
  const sheet = ss.getSheetByName(name);
  const headers = SHEET_SCHEMAS[name];
  if (!sheet || !headers) return;
  sheet.clearContents();
  sheet.appendRow(headers);
  if (records && records.length > 0) {
    const rows = records.map(function(rec) {
      return headers.map(function(h) {
        const val = rec[h];
        return (val === undefined || val === null) ? "" : val;
      });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function getSettingsMap(ss) {
  const rows = getSheetRecords(ss, "Settings");
  const map = {};
  rows.forEach(function(r) {
    if (r.key) map[r.key] = r.value;
  });
  return map;
}

function saveSettingsMap(ss, obj) {
  const rows = Object.keys(obj).map(function(k) {
    return { key: k, value: typeof obj[k] === "string" ? obj[k] : JSON.stringify(obj[k]) };
  });
  saveSheetRecords(ss, "Settings", rows);
}
