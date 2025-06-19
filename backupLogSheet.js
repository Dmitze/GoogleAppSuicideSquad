/**
 * const TMP_FOLDER_ID = '14EX4nx7NACIv0qCnJL0soVivEmI4SG9G';
 * ID папки Google Drive для збереження резервних копій, кешу і тимчасових файлів
 */

/**
 * Отримати посилання на папку для тимчасових/архівних файлів
 */
function getTmpFolderOrThrow() {
  try {
    return DriveApp.getFolderById(TMP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для тимчасових файлів не знайдено! Перевірте TMP_FOLDER_ID.");
  }
}

/**
 * Експорт листа "Лог змін" як файла CSV або Excel у папку TMP_FOLDER_ID
 * @param {string} format - "csv" або "xlsx"
 * @returns {string} URL створеного файлу
 */
function backupLogSheetToDrive(format = "xlsx") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(typeof LOG_SHEET_NAME !== 'undefined' ? LOG_SHEET_NAME : "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const folder = getTmpFolderOrThrow();
  const data = logSheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error("Немає даних для експорту!");

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `backup_log_${timestamp}`;
  let fileUrl = "";

  const fmt = format.toLowerCase();

  if (fmt === "csv") {
    // Додаємо BOM для коректного відкриття в Excel
    let csv = '\uFEFF' + data.map(row => row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  } else if (fmt === "xlsx") {
    const tempSS = SpreadsheetApp.create(fileName);
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // Видаляємо зайві листи
    tempSS.getSheets().forEach(s => {
      if (s.getName() !== tempSheet.getName()) tempSS.deleteSheet(s);
    });

    const blob = tempSS.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
    const file = folder.createFile(blob.setName(`${fileName}.xlsx`));
    // Важливо: затримка на видалення, щоб уникнути race condition
    Utilities.sleep(500);
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    fileUrl = file.getUrl();
  } else {
    throw new Error("Непідтримуваний формат: " + format);
  }

  return fileUrl;
}

/**
 * Архівує лог у файл CSV і зберігає у TMP_FOLDER_ID, очищає лог
 */
function archiveLogHistory() {
  const folder = getTmpFolderOrThrow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(typeof LOG_SHEET_NAME !== 'undefined' ? LOG_SHEET_NAME : "Лог змін");
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return; // Тільки заголовки

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  let csv = '\uFEFF' + data.map(row =>
    row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `log_archive_${timestamp}.csv`);
  folder.createFile(blob);

  // Очищаємо лог
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow()-1, logSheet.getLastColumn()).clearContent();
  }
  Logger.log(`Лог за ${timestamp} архівовано.`);
}

/**
 * Створює тригер на щоденну архівацію лога (23:00)
 */
function createDailyArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === "archiveLogHistory");
  if (!exists) {
    ScriptApp.newTrigger("archiveLogHistory")
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();
    Logger.log("Тригер на щоденну архівацію створено.");
  }
}

/**
 * Видаляє архіви/тимчасові файли старше N днів з TMP_FOLDER_ID
 * @param {number} daysToKeep - Скільки днів зберігати
 */
function cleanupOldBackups(daysToKeep = 30) {
  const folder = getTmpFolderOrThrow();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const files = folder.getFiles();
  let deleted = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getLastUpdated() < cutoffDate) {
      file.setTrashed(true);
      deleted++;
    }
  }
  Logger.log(`Старі бекапи (${deleted} шт.) видалені.`);
}

/**
 * Виводить діалогове вікно для експорту логу у Excel
 */
function exportLogSheetAsExcel() {
  try {
    const url = backupLogSheetToDrive("xlsx");
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Файл Excel створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Файл Excel створено: " + url);
    }
  } catch (e) {
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Помилка: " + e.message);
    }
  }
}

/**
 * Виводить діалогове вікно для експорту логу у CSV
 */
function exportLogSheetAsCSV() {
  try {
    const url = backupLogSheetToDrive("csv");
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Файл CSV створено", `Завантажити: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Файл CSV створено: " + url);
    }
  } catch (e) {
    if (isUiAvailable()) {
      SpreadsheetApp.getUi().alert("Помилка", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Помилка: " + e.message);
    }
  }
}

/**
 * Перевірка чи доступний SpreadsheetApp.getUi()
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch {
    return false;
  }
}

/**
 * Отримати список всіх лог/архів файлів у TMP_FOLDER_ID (csv/xlsx)
 * @returns {Array<{id:string, name:string, date:Date}>}
 */
function getLogFilesList() {
  const folder = getTmpFolderOrThrow();
  const files = folder.getFiles();
  let list = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!/\.csv$|\.xlsx$/i.test(name)) continue;
    list.push({
      id: file.getId(),
      name: name,
      date: file.getLastUpdated()
    });
  }
  list.sort((a, b) => b.date - a.date);
  return list;
}
