/**
 * Kinagin Farm Forms – Google Apps Script
 * Separate sheet per form type + safe handling of missing fields
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const type = (data.type || 'Unknown').toLowerCase();
    
    const sheetNames = {
      'spray': 'Spray Records',
      'insect': 'Insect Reports',
      'irrigation': 'Irrigation',
      'fertigation': 'Fertigation',
      'harvest': 'Harvesting',
      'fuel': 'Diesel & Petrol'
    };
    
    const sheetName = sheetNames[type] || 'Other';
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    if (sheet.getLastRow() === 0) {
      const headers = getHeadersForType(type);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    const row = buildRow(data, type);
    sheet.appendRow(row);
    
    // Full JSON backup
    let fullSheet = ss.getSheetByName('FullRecords');
    if (!fullSheet) {
      fullSheet = ss.insertSheet('FullRecords');
      fullSheet.appendRow(['Timestamp', 'Type', 'Full JSON']);
    }
    fullSheet.appendRow([data.timestamp || new Date().toISOString(), type, JSON.stringify(data)]);
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function safe(value) {
  return (value === undefined || value === null) ? '' : value;
}

function getHeadersForType(type) {
  switch(type) {
    case 'spray':
      return ['Timestamp', 'Date', 'Start Time', 'Finish Time', 'Block', 'Crop', 'Product', 'Batch', 'Target', 'Rate', 'Quantity', 'Equipment', 'Wind Speed', 'Wind Direction', 'WHP/EHD', 'Applicator', 'Notes'];
    case 'insect':
      return ['Timestamp', 'Date', 'Block', 'Trap ID', 'Pest', 'Count', 'Recorded By', 'Notes'];
    case 'irrigation':
      return ['Timestamp', 'Date', 'Block', 'Duration (hrs)', 'Volume', 'Method', 'Operator', 'Notes'];
    case 'fertigation':
      return ['Timestamp', 'Date', 'Block', 'Product', 'Rate', 'Method', 'Wind Speed', 'Wind Direction', 'Operator', 'Notes'];
    case 'harvest':
      return ['Timestamp', 'Date', 'Block', 'Variety', 'Quantity', 'Destination', 'EHD Check', 'Recorded By', 'Notes'];
    case 'fuel':
      return ['Timestamp', 'Date', 'Fuel Type', 'Equipment', 'Litres', 'Hours/Odo', 'Filled By', 'Notes'];
    default:
      return ['Timestamp', 'Type', 'Data'];
  }
}

function buildRow(data, type) {
  switch(type) {
    case 'spray':
      return [
        safe(data.timestamp), safe(data.date), safe(data.startTime), safe(data.finishTime), safe(data.block),
        safe(data.crop), safe(data.product), safe(data.batch), safe(data.target), safe(data.rate),
        safe(data.quantity), safe(data.equipment), safe(data.windSpeed), safe(data.windDirection), safe(data.whp), safe(data.applicator), safe(data.notes)
      ];
    case 'insect':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.trapId), safe(data.pest),
        safe(data.count), safe(data.recorder), safe(data.notes)
      ];
    case 'irrigation':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.duration), safe(data.volume),
        safe(data.method), safe(data.operator), safe(data.notes)
      ];
    case 'fertigation':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.product), safe(data.rate),
        safe(data.method), safe(data.windSpeed), safe(data.windDirection), safe(data.operator), safe(data.notes)
      ];
    case 'harvest':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.variety), safe(data.quantity),
        safe(data.destination), safe(data.ehd), safe(data.recorder), safe(data.notes)
      ];
    case 'fuel':
      return [
        safe(data.timestamp), safe(data.date), safe(data.fuelType), safe(data.equipment), safe(data.litres),
        safe(data.hours), safe(data.operator), safe(data.notes)
      ];
    default:
      return [safe(data.timestamp), type, JSON.stringify(data)];
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const type = (params.type || '').toLowerCase();
    const callback = params.callback || '';
    const sheetNames = {
      'spray': 'Spray Records',
      'insect': 'Insect Reports',
      'irrigation': 'Irrigation',
      'fertigation': 'Fertigation',
      'harvest': 'Harvesting',
      'fuel': 'Diesel & Petrol'
    };

    let result;
    if (!type || !sheetNames[type]) {
      result = { status: 'ok', message: 'Kinagin Farm Forms endpoint is running' };
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetNames[type]);
      if (!sheet || sheet.getLastRow() < 2) {
        result = { status: 'ok', type: type, sheet: sheetNames[type], headers: [], rows: [] };
      } else {
        const values = sheet.getDataRange().getValues();
        const headers = values[0].map(function (h) { return String(h); });
        const tz = Session.getScriptTimeZone();
        const rows = values.slice(1).filter(function (r) {
          return r.some(function (cell) { return cell !== '' && cell !== null; });
        }).map(function (r) {
          const obj = {};
          headers.forEach(function (h, i) {
            const cell = r[i];
            if (cell instanceof Date) {
              obj[h] = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
            } else {
              obj[h] = cell === null || cell === undefined ? '' : cell;
            }
          });
          return obj;
        });
        result = { status: 'ok', type: type, sheet: sheetNames[type], headers: headers, rows: rows };
      }
    }

    const text = JSON.stringify(result);
    if (callback && /^[A-Za-z0-9_]+$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + text + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(text)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const text = JSON.stringify({ status: 'error', message: err.toString() });
    return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
  }
}
