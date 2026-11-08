// 본인의 구글 시트 ID (URL의 /d/ 와 /edit 사이 문자열)
var SPREADSHEET_ID = '11te3Q7IPAX7q7t8LhEur1fqU2qXOrU6TTq_XfbTz7cY';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('미래채움 강사료 정산');
}

function getRecordSheet() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // 특정 시트 ID로 정확히 연결
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '강사료기록';
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    if (sheet.getLastRow() === 0) {
      var headers = [
        '고유ID', '활동구분', '활동구분명', '활동일자', '학교·교육·모임명',
        '수업·프로그램명', '역할', '활동시간', '시간단가', '수업료',
        '원거리여비', '교구대여비', '세전총액', '메모'
      ];
      sheet.appendRow(headers);
    } else {
      var firstCell = sheet.getRange(1, 1).getValue();
      if (!firstCell || String(firstCell).trim() === '') {
        var headers = [
          '고유ID', '활동구분', '활동구분명', '활동일자', '학교·교육·모임명',
          '수업·프로그램명', '역할', '활동시간', '시간단가', '수업료',
          '원거리여비', '교구대여비', '세전총액', '메모'
        ];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }
    return sheet;
  } finally {
    lock.releaseLock();
  }
}

function formatDateValue(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var yyyy = val.getFullYear();
    var mm = String(val.getMonth() + 1).padStart(2, '0');
    var dd = String(val.getDate()).padStart(2, '0');
    return yyyy + '-' + mm + '-' + dd;
  }
  var str = String(val).trim();
  if (str.length >= 10 && str.indexOf('-') !== -1) {
    return str.substring(0, 10);
  }
  return str;
}

function getAllRecords() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var records = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = row[0] ? String(row[0]).trim() : '';
      var dateStr = formatDateValue(row[3]);
      
      if (!id && !dateStr && !row[4]) {
        continue;
      }
      
      records.push({
        id: id,
        category: row[1] ? String(row[1]) : '',
        categoryName: row[2] ? String(row[2]) : '',
        date: dateStr,
        name: row[4] ? String(row[4]) : '',
        course: row[5] ? String(row[5]) : '',
        role: row[6] ? String(row[6]) : '',
        hours: Number(row[7]) || 0,
        unitPrice: Number(row[8]) || 0,
        teachingFee: Number(row[9]) || 0,
        travel: Number(row[10]) || 0,
        equipment: Number(row[11]) || 0,
        totalFee: Number(row[12]) || 0,
        note: row[13] ? String(row[13]) : ''
      });
    }
    
    return records;
  } finally {
    lock.releaseLock();
  }
}

function addRecord(record) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordSheet();
    var recordId = record.id;
    if (!recordId) {
      recordId = 'rec_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existingIds.length; i++) {
        if (String(existingIds[i][0]).trim() === String(recordId).trim()) {
          return true;
        }
      }
    }
    
    var rowData = [
      recordId,
      record.category || '',
      record.categoryName || '',
      record.date || '',
      record.name || '',
      record.course || '',
      record.role || '',
      Number(record.hours) || 0,
      Number(record.unitPrice) || 0,
      Number(record.teachingFee) || 0,
      Number(record.travel) || 0,
      Number(record.equipment) || 0,
      Number(record.totalFee) || 0,
      record.note || ''
    ];
    
    sheet.appendRow(rowData);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function updateRecord(record) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return false;
    }
    
    var existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRow = -1;
    
    for (var i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0]).trim() === String(record.id).trim()) {
        targetRow = i + 2;
        break;
      }
    }
    
    if (targetRow !== -1) {
      var rowData = [
        record.id,
        record.category || '',
        record.categoryName || '',
        record.date || '',
        record.name || '',
        record.course || '',
        record.role || '',
        Number(record.hours) || 0,
        Number(record.unitPrice) || 0,
        Number(record.teachingFee) || 0,
        Number(record.travel) || 0,
        Number(record.equipment) || 0,
        Number(record.totalFee) || 0,
        record.note || ''
      ];
      sheet.getRange(targetRow, 1, 1, 14).setValues([rowData]);
      return true;
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return false;
    }
    
    var existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 2);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function clearAllRecords() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}