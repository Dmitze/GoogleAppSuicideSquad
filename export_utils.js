// === export_utils.js ===
// Универсальные функции для экспорта диапазонов и отправки файлов на email

/** Получить имена всех листов */
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/** Получить данные диапазона по имени листа и A1-диапазону */
function getPreviewData(sheetName, rangeA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(rangeA1).getValues();
}

/** Получить выделенный пользователем диапазон (A1-нотация) */
function getActiveRangeA1() {
  const range = SpreadsheetApp.getActiveRange();
  return range ? range.getA1Notation() : '';
}

/** Экспорт диапазона в Word (docx) и возвращает ссылку на файл */
function exportRangeToWord(sheetName, rangeA1, fileName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Лист "${sheetName}" не знайдено!`);
  const values = sheet.getRange(rangeA1).getValues();
  if (!values || !values.length) throw new Error("Діапазон порожній або невірний!");
  const doc = DocumentApp.create(fileName.replace(/\.docx$/i, ''));
  doc.getBody().appendTable(values);
  doc.saveAndClose();
  const token = ScriptApp.getOAuthToken();
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob = response.getBlob().setName(fileName.endsWith('.docx') ? fileName : fileName + ".docx");
  const file = DriveApp.createFile(blob);
  return file.getUrl();
}

/** Экспорт диапазона в PDF и возвращает ссылку на файл */
function exportRangeToPdf(sheetName, rangeA1, fileName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Лист "${sheetName}" не знайдено!`);
  const values = sheet.getRange(rangeA1).getValues();
  if (!values || !values.length) throw new Error("Діапазон порожній або невірний!");
  const doc = DocumentApp.create(fileName.replace(/\.pdf$/i, ''));
  doc.getBody().appendTable(values);
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf').setName(fileName.endsWith('.pdf') ? fileName : fileName + ".pdf");
  const exported = DriveApp.createFile(pdfBlob);
  file.setTrashed(true);
  return exported.getUrl();
}

/** Экспорт диапазона в Excel (xlsx) и возвращает ссылку на файл */
function exportRangeToExcel(sheetName, rangeA1, fileName) {
  const ss = SpreadsheetApp.create(fileName.replace(/\.xlsx$/i, ''));
  const srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!srcSheet) throw new Error(`Лист "${sheetName}" не знайдено!`);
  const values = srcSheet.getRange(rangeA1).getValues();
  let newSheet = ss.getSheets()[0];
  newSheet.setName(sheetName);
  newSheet.getRange(1,1,values.length,values[0].length).setValues(values);
  const file = DriveApp.getFileById(ss.getId());
  const blob = file.getBlob().setName(fileName.endsWith('.xlsx') ? fileName : fileName + ".xlsx");
  const exported = DriveApp.createFile(blob);
  file.setTrashed(true);
  return exported.getUrl();
}

/** Отправить файл по email */
function sendFileByEmail(email, url, fileName, message) {
  // Получаем файл по ссылке
  const response = UrlFetchApp.fetch(url);
  const blob = response.getBlob().setName(fileName);
  MailApp.sendEmail({
    to: email,
    subject: `Експортований файл: ${fileName}`,
    body: message || 'Дивіться вкладення.',
    attachments: [blob]
  });
} 