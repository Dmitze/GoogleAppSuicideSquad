/**
 * Всі кеші, тимчасові таблиці, експортовані логи, історії та архіви зберігаються у спеціальній папці Google Drive
 */

const SHEET_NAMES = [
  "2 Бат Загальна", "Ударні БпЛА", "Розвідувальні БпЛА", "НРК", "ППО",
  "НСО БТ", "АТ", "Засоби ураження", "ЗББ та Р", "РЕБ", "Оптика", "РЛС"
];
const LOG_SHEET_NAME = "Лог змін";
const COLOR_GREEN = "#b6d7a8";
const IMPORTANT_RANGES = {
  "2 Бат Загальна": ["A1:C5"],
  "АТ": ["B2:D6"]
};

/**
 * Отримати папку для тимчасових файлів (логів, кешу, експортів, архівів)
 */
function getTmpFolderOrThrow() {
  try {
    return DriveApp.getFolderById(TMP_FOLDER_ID);
  } catch (e) {
    throw new Error("Папку для тимчасових файлів не знайдено! Перевірте TMP_FOLDER_ID.");
  }
}

function showLogsRestoreDialog() {
  const html = HtmlService.createHtmlOutputFromFile('logs_restore.html')
    .setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Архіви логів');
}

// === Меню при відкритті файлу ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // === 1. Меню "🛠 Дії з таблицею" ===
  const mainMenu = ui.createMenu("🛠 Дії з таблицею")
    .addItem("📱 Згенерувати QR-коди з усіма даними", "generateFullInfoQRCodesForSheet")
    .addItem("🔍 Гнучка генерація ключів", "generateKeysWithCustomColumns")
    .addItem("📱 Згенерувати QR-коди за Постійним ID", "generateQRCodesForSheet")
    .addItem("📁 Експортувати товари з ключами (CSV)", "exportProductsWithKeysToCSV");

  // === 2. Меню "📏 Форматування" ===
  const formattingMenu = ui.createMenu("📏 Форматування")
    .addItem("📐 Вирівняти висоту рядків", "showSidebar");

  // === 3. Меню "📄 Експорт до Word" ===
  const wordExportMenu = ui.createMenu("📄 Експорт до Word")
    .addItem("📑 Скласти звіт у Word...", "showWordExportFullForm")
    .addItem("📘 Експортувати до Word...", "showExportToWordDialog")
    .addItem("📖 Експортувати виділений діапазон до Word", "exportSheetRangeToWord");

  // === 4. Меню "🕵️‍♂️ Перевірка та логи" ===
  const validationMenu = ui.createMenu("🕵️‍♂️ Перевірка та логи")
    .addItem("🔧 Ручна перевірка змін", "checkChanges")
    .addItem("📊 Звіт по діям користувачів", "showUsersActionReport")
    .addItem("🔍 Перевірити орфографію/формати", "runValidation");

  // === 5. Меню "📦 Логи та архівація" ===
  const logMenu = ui.createMenu("📦 Логи та архівація")
    .addItem("📥 Експорт логу у Excel", "exportLogSheetAsExcel")
    .addItem("📤 Експорт логу у CSV", "exportLogSheetAsCSV")
    .addItem("📜 Експорт історії у CSV", "exportHistoryToCSV")
    .addItem("🗃 Архівація логів", "archiveLogHistory")
    .addItem("⏰ Створити тригер на архівацію", "createDailyArchiveTrigger")
    .addItem("🗑 Видалити старі бекапи", "cleanupOldBackups");

  // === 6. Додавання підменю "🔍 Пошук і історія" ===
  const searchMenu = ui.createMenu("🔍 Пошук і історія")
    .addItem("🔎 Додати пункт 'Пошук' в головне меню", "addHistorySearchMenu");

  // === Додавання всіх меню до головного UI ===
  ui.createMenu("📋 Головне меню")
    .addSubMenu(mainMenu)
    .addSubMenu(formattingMenu)
    .addSubMenu(wordExportMenu)
    .addSubMenu(validationMenu)
    .addSubMenu(logMenu)
    .addSubMenu(searchMenu)
    .addToUi();

  setupLogSheet();
}


// Відкриває бокову панель
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("Вирівняти строки");
  SpreadsheetApp.getUi().showSidebar(html);
}

function checkChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (!Array.isArray(values)) return;

    const currentHash = JSON.stringify(values);
    const storedHashKey = `prevDataHash_${sheetName}`;
    const storedValuesKey = `prevValues_${sheetName}`;
    const storedHash = props.getProperty(storedHashKey);

    let oldValues = [];
    if (storedHash) {
      const old = props.getProperty(storedValuesKey);
      oldValues = old ? JSON.parse(old) : values.map(row => row.map(() => null));
    } else {
      oldValues = values.map(row => row.map(() => null));
    }

    if (
      storedHash &&
      storedHash !== currentHash &&
      Array.isArray(values) &&
      Array.isArray(oldValues) &&
      values.length > 0 &&
      oldValues.length > 0
    ) {
      highlightChanges(sheet, oldValues, values);
      logChanges(sheet, oldValues, values);
    }

    // Логування додавання/видалення рядків/стовпців
    if (
      Array.isArray(oldValues) &&
      Array.isArray(values) &&
      oldValues.length !== values.length
    ) {
      const type = oldValues.length < values.length ? "Додано рядок" : "Видалено рядок";
      logRowOrColumnAction(sheet, type, oldValues.length, values.length);
    }
    if (
      Array.isArray(oldValues) && Array.isArray(values) &&
      oldValues.length > 0 && values.length > 0 &&
      oldValues[0].length !== values[0].length
    ) {
      const type = oldValues[0].length < values[0].length ? "Додано стовпець" : "Видалено стовпець";
      logRowOrColumnAction(sheet, type, oldValues[0].length, values[0].length);
    }

    props.setProperty(storedHashKey, currentHash);
    props.setProperty(storedValuesKey, JSON.stringify(values));
  });
}

// === Оптимізоване збереження/експорт логів у TMP_FOLDER_ID ===

/**
 * Експортує LOG_SHEET_NAME у Excel/CSV у TMP_FOLDER_ID
 * @param {string} format - "csv" або "xlsx"
 * @returns {string} URL створеного файлу
 */
function backupLogSheetToTmpFolder(format = "xlsx") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const folder = getTmpFolderOrThrow();
  const data = logSheet.getDataRange().getValues();
  if (!data || data.length === 0) throw new Error("Немає даних для експорту!");

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `backup_log_${timestamp}`;
  let fileUrl = "";
  const fmt = format.toLowerCase();

  if (fmt === "csv") {
    let csv = '\uFEFF' + data.map(row => row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = Utilities.newBlob(csv, MimeType.CSV, `${fileName}.csv`);
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  } else if (fmt === "xlsx") {
    const tempSS = SpreadsheetApp.create(fileName);
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    tempSS.getSheets().forEach(s => {
      if (s.getName() !== tempSheet.getName()) tempSS.deleteSheet(s);
    });

    const blob = tempSS.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
    const file = folder.createFile(blob.setName(`${fileName}.xlsx`));
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
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error("Лист 'Лог змін' не знайдено!");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return;

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  let csv = '\uFEFF' + data.map(row =>
    row.map(cell => `"${(cell ?? "").toString().replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const blob = Utilities.newBlob(csv, MimeType.CSV, `log_archive_${timestamp}.csv`);
  folder.createFile(blob);

  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow()-1, logSheet.getLastColumn()).clearContent();
  }
  Logger.log(`Лог за ${timestamp} архівовано.`);
}

/**
 * Створює тригер на щоденну архівацію лога
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
    const url = backupLogSheetToTmpFolder("xlsx");
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
    const url = backupLogSheetToTmpFolder("csv");
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


// === Логування додавання/видалення рядків/стовпців ===
function logRowOrColumnAction(sheet, type, oldLen, newLen) {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const user = Session.getActiveUser().getEmail();
    const time = new Date();
    let actionDesc = "";
    let sheetName = (sheet && typeof sheet.getName === "function") ? sheet.getName() : "[невідомий лист]";
  
    if (type === "Додано рядок") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Видалено рядок") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Додано стовпець") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else if (type === "Видалено стовпець") {
      actionDesc = `Було ${oldLen}, стало ${newLen}`;
    } else {
      actionDesc = `Невідомий тип зміни: ${type}`;
    }
    logSheet.appendRow([
      time,
      user,
      sheetName,
      "",
      type,
      actionDesc,
      "",
      "",
      ""
    ]);
  }


  // === Логування змін значень з типом дії ===
function logChanges(sheet, oldValues, newValues) {
    if (!sheet || !Array.isArray(newValues) || !Array.isArray(oldValues)) return;
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const user = Session.getActiveUser().getEmail();
    const time = new Date();
    let changes = [];
    for (let row = 0; row < newValues.length; row++) {
      const newRow = newValues[row] || [];
      const oldRow = oldValues[row] || [];
      for (let col = 0; col < newRow.length; col++) {
        const oldValue = (oldRow[col] !== undefined ? oldRow[col] : "");
        const newValue = (newRow[col] !== undefined ? newRow[col] : "");
        if (oldValue !== newValue) {
          const cell = sheet.getRange(row + 1, col + 1);
          let formula = cell.getFormula();
          if (formula) formula = "=" + formula;
          const important = isImportantCell(sheet.getName(), row + 1, col + 1) ? "Так" : "Ні";
          // Тип дії для кожної зміни:
          let changeType = "";
          if ((oldValue === "" || oldValue === null) && newValue !== "") {
            changeType = "Додано значення";
          } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
            changeType = "Видалено значення";
          } else {
            changeType = "Змінено";
          }
          // Генерируем link на ячейку для просмотра истории изменений
          const cellLink = `=HYPERLINK("#gid=${sheet.getSheetId()}&range=${cell.getA1Notation()}"; "${cell.getA1Notation()}")`;
          changes.push([
            time,
            user,
            sheet.getName(),
            cellLink, // Ссылка на ячейку в формате HYPERLINK
            changeType,
            oldValue,
            newValue,
            formula || "",
            important
          ]);
        }
      }
    }
    if (changes.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, changes.length, 9).setValues(changes);
    }
  }

  
  // === Підсвітка змінених комірок ===
function highlightChanges(sheet, oldValues, newValues) {
    if (!Array.isArray(newValues) || !Array.isArray(oldValues)) return;
    for (let row = 0; row < newValues.length; row++) {
      for (let col = 0; col < newValues[row].length; col++) {
        const oldValue = (oldValues[row] || [])[col];
        const newValue = newValues[row][col];
        if (oldValue !== newValue) {
          const cell = sheet.getRange(row + 1, col + 1);
          // Додаємо кольори відповідно до типу зміни
          if ((oldValue === "" || oldValue === null) && newValue !== "") {
            // Було пусто -> стало щось (додавання)
            cell.setBackground("#b6d7a8"); // Зелений
          } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
            // Було щось -> стало пусто (видалення)
            cell.setBackground("#ea9999"); // Червоний
          } else {
            // Будь-яка інша зміна (оновлення)
            cell.setBackground("#ffe599"); // Жовтий
          }
        }
      }
    }
  }




  // === Визначення важливих комірок ===
function isImportantCell(sheetName, row, col) {
    if (!IMPORTANT_RANGES[sheetName]) return false;
    for (const rangeStr of IMPORTANT_RANGES[sheetName]) {
      const range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(rangeStr);
      if (
        row >= range.getRow() &&
        row < range.getRow() + range.getNumRows() &&
        col >= range.getColumn() &&
        col < range.getColumn() + range.getNumColumns()
      ) {
        return true;
      }
    }
    return false;
  }


  // === Створення аркуша для логів ===
function setupLogSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      const headers = [[
        "Час зміни",
        "Користувач",
        "Аркуш",
        "Посилання на комірку", // Новый заголовок для ссылки
        "Тип дії",
        "Було",
        "Стало",
        "Формула",
        "Важлива зміна"
      ]];
      logSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      logSheet.autoResizeColumns(1, headers[0].length);
    }
  }

  function getAllHistoryLogs() {
    const logs = google.script.run.withSuccessHandler(function(logs){
      if (!logs || !logs.length) {
        showStatus('Немає записів для пошуку', 'error');
        return [];
      }
      return logs;
    }).getAllHistoryLogs();
  }
  
  function exportHistoryToCSV() {
    const logs = getAllHistoryLogs();
    if (!logs.length) {
      showStatus('Немає даних для експорту!', 'error');
      return;
    }
  
    const headers = ['Дата/час', 'Аркуш', 'Користувач', 'Дія', 'Адреса', 'Було', 'Стало'];
    const rows = [headers].concat(
      logs.map(r => [
        r.dateTime || r.date, r.sheet, r.user, r.action, r.address, r.oldValue, r.newValue
      ])
    );
    const csv = rows.map(row => row.map(cell =>
      `"${(cell||'').toString().replace(/"/g,'""')}"`
    ).join(',')).join('\r\n');
  
    const blob = new Blob([csv], {type:'text/csv'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'history_search_export.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},600);
    showStatus('CSV-файл сформовано. Завантаження розпочато.', 'success');
  }
  
  function getHistoryAnalytics() {
    const logs = getAllHistoryLogs();
    const users = {};
    const sheets = {};
    const days = {};
    logs.forEach(log => {
      if (log.user) users[log.user] = (users[log.user] || 0) + 1;
      if (log.sheet) sheets[log.sheet] = (sheets[log.sheet] || 0) + 1;
      if (log.date) days[log.date] = (days[log.date] || 0) + 1;
    });
    return { users, sheets, days };
  }
