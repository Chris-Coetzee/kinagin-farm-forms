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
    
    const wanted = getHeadersForType(type);
    let headers;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(wanted);
      sheet.getRange(1, 1, 1, wanted.length).setFontWeight('bold');
      headers = wanted;
    } else {
      headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function (h) { return String(h); });
      const missing = wanted.filter(function (h) { return headers.indexOf(h) === -1; });
      if (missing.length) {
        sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
        sheet.getRange(1, headers.length + 1, 1, missing.length).setFontWeight('bold');
        headers = headers.concat(missing);
      }
    }

    const valuesByHeader = buildValueMap(data, type);
    const row = headers.map(function (h) {
      return valuesByHeader[h] !== undefined ? valuesByHeader[h] : '';
    });
    sheet.appendRow(row);
    
    // Full JSON backup
    let fullSheet = ss.getSheetByName('FullRecords');
    if (!fullSheet) {
      fullSheet = ss.insertSheet('FullRecords');
      fullSheet.appendRow(['Timestamp', 'Type', 'Full JSON']);
    }
    const backup = {};
    Object.keys(data).forEach(function (key) { backup[key] = data[key]; });
    if (backup.photo && String(backup.photo).length > 200) backup.photo = '[photo stored on the form sheet]';
    if (backup.signature && String(backup.signature).length > 200) backup.signature = '[signature stored on the form sheet]';
    fullSheet.appendRow([data.timestamp || new Date().toISOString(), type, JSON.stringify(backup)]);
    
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
      return ['Timestamp', 'Date', 'Start Time', 'Block', 'Crop', 'Products', 'Batches', 'Rates', 'Withholding Periods', 'Target Pests', 'Sprayer', 'Wind Speed', 'Wind Direction', 'Applicator', 'Notes'];
    case 'insect':
      return ['Timestamp', 'Date', 'Block', 'Trap ID', 'Insects', 'Counts', 'Photo', 'Recorded By', 'Notes'];
    case 'irrigation':
      return ['Timestamp', 'Date', 'Block', 'Duration (hrs)', 'Method', 'Operator', 'Notes'];
    case 'fertigation':
      return ['Timestamp', 'Date', 'Block', 'Product', 'Rate', 'Method', 'Wind Speed', 'Wind Direction', 'Operator', 'Notes'];
    case 'harvest':
      return ['Timestamp', 'Date', 'Block', 'Bin Owner', 'Crop', 'Variety', 'Bin Numbers', 'Quantity', 'Destination', 'Recorded By', 'Notes'];
    case 'fuel':
      return ['Timestamp', 'Date', 'Fuel Type', 'Vehicle', 'Opening Total', 'Closing Total', 'Litres', 'Hours/Odo', 'Filled By', 'Notes', 'Signature'];
    default:
      return ['Timestamp', 'Type', 'Data'];
  }
}

function buildRow(data, type) {
  switch(type) {
    case 'spray':
      return [
        safe(data.timestamp), safe(data.date), safe(data.startTime), safe(data.block), safe(data.crop),
        safe(data.product), safe(data.batch), safe(data.rate), safe(data.whp), safe(data.target),
        safe(data.equipment), safe(data.windSpeed), safe(data.windDirection), safe(data.applicator), safe(data.notes)
      ];
    case 'insect':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.trapId),
        safe(data.pest), safe(data.count), safe(data.photo), safe(data.recorder), safe(data.notes)
      ];
    case 'irrigation':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.duration),
        safe(data.method), safe(data.operator), safe(data.notes)
      ];
    case 'fertigation':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.product), safe(data.rate),
        safe(data.method), safe(data.windSpeed), safe(data.windDirection), safe(data.operator), safe(data.notes)
      ];
    case 'harvest':
      return [
        safe(data.timestamp), safe(data.date), safe(data.block), safe(data.binOwner), safe(data.crop),
        safe(data.variety), safe(data.binNumbers), safe(data.quantity), safe(data.destination),
        safe(data.recorder), safe(data.notes)
      ];
    case 'fuel':
      return [
        safe(data.timestamp), safe(data.date), safe(data.fuelType), safe(data.equipment),
        safe(data.openingTotal), safe(data.closingTotal), safe(data.litres),
        safe(data.hours), safe(data.operator), safe(data.notes), safe(data.signature)
      ];
    default:
      return [safe(data.timestamp), type, JSON.stringify(data)];
  }
}

function buildValueMap(data, type) {
  const headers = getHeadersForType(type);
  const values = buildRow(data, type);
  const map = {};
  headers.forEach(function (h, i) { map[h] = values[i]; });
  if (type === 'fuel') {
    map['Equipment'] = safe(data.equipment);
    map['Vehicle'] = safe(data.equipment);
  }
  if (type === 'spray') {
    map['Product'] = safe(data.product);
    map['Batch'] = safe(data.batch);
    map['Target'] = safe(data.target);
    map['Rate'] = safe(data.rate);
    map['Equipment'] = safe(data.equipment);
    map['WHP/EHD'] = safe(data.whp);
    map['Sprayer'] = safe(data.equipment);
    map['Products'] = safe(data.product);
    map['Batches'] = safe(data.batch);
    map['Rates'] = safe(data.rate);
    map['Withholding Periods'] = safe(data.whp);
    map['Target Pests'] = safe(data.target);
  }
  if (type === 'insect') {
    map['Pest'] = safe(data.pest);
    map['Count'] = safe(data.count);
    map['Insects'] = safe(data.pest);
    map['Counts'] = safe(data.count);
    map['Photo'] = safe(data.photo);
  }
  if (type === 'harvest') {
    map['Bin Owner'] = safe(data.binOwner);
    map['Crop'] = safe(data.crop);
    map['Bin Numbers'] = safe(data.binNumbers);
  }
  return map;
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
