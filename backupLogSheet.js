function backupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не найден!");

  const folder = DriveApp.getFolderById(TMP_FOLDER_ID); // Убедись, что TMP_FOLDER_ID задан в глобальной области
  const data = logSheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error("Нет данных для экспорта!");

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `log_backup_${timestamp}`;

  // Генерация CSV с BOM
  let csv = '\uFEFF' + data.map(row => 
    row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();

  return fileUrl;
}

function backupLogAsCSV() {
  try {
    const url = backupLogSheet();
    SpreadsheetApp.getUi().alert("Резервная копия CSV создана!", url, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Ошибка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function backupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не найден!");

  const folder = DriveApp.getFolderById(TMP_FOLDER_ID);
  const data = logSheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error("Нет данных для экспорта!");

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `log_backup_${timestamp}`;

  // Создаем CSV с BOM для Excel
  let csv = '\uFEFF' + data.map(row =>
    row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();

  Logger.log(`Бэкап создан: ${fileUrl}`);
}

// === ТРИГГЕР ДЛЯ ЕЖЕДНЕВНОГО БЭКАПА ===
function createDailyTrigger() {
  ScriptApp.newTrigger("backupLogSheet")
    .timeBased()
    .everyDays(1)
    .atHour(3) // Выполнять в 3:00 утра
    .create();
}

function showBackupForm() {
  const html = HtmlService.createHtmlOutputFromFile("backupLogSheetForm")
    .setTitle("Импорт и просмотр CSV бэкапа");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === ОБРАБОТЧИК CSV ===
function importCSV(csvData) {
  try {
    // Парсим CSV данные
    const rows = Utilities.parseCsv(csvData);

    if (!rows || rows.length === 0) {
      throw new Error("Файл пустой или не соответствует формату CSV.");
    }

    return rows;
  } catch (e) {
    return { error: e.toString() };
  }
}
