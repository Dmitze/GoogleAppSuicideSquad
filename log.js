// === Налаштування ===
const LOG_SHEET_NAME = "Лог змін";
const COLOR_GREEN = "#b6d7a8";
// Список колонок, які не потрібно логувати (службові/технічні)
const IGNORED_HEADERS = [
  'Постійний ID',
  'Ідентифікатор',
  'QR-код',
  'QR',
  'link',
  'Посилання на QR'
];

/**
 * Основний тригер для логування змін (автоматичний аудит)
 * Записує логи та підсвічує зміни для всіх листів, крім LOG_SHEET_NAME.
 * Логуються тільки зміни в окремих клітинках (без фіксації додавання/видалення рядків/стовпців).
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === LOG_SHEET_NAME) return;

  // Підсвічуємо змінену клітинку
  highlightCell(e);

  // Логуємо зміну (якщо не службова колонка)
  logCellEdit(e);
}

/**
 * Підсвічує змінену клітинку
 */
function highlightCell(e) {
  const cell = e.range;
  const oldValue = e.oldValue !== undefined ? e.oldValue : "";
  const newValue = cell.getValue();

  if (oldValue === newValue) return;

  if ((oldValue === "" || oldValue === null) && newValue !== "") {
    cell.setBackground("#b6d7a8"); // Зелений
  } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
    cell.setBackground("#ea9999"); // Червоний
  } else {
    cell.setBackground("#ffe599"); // Жовтий
  }
}

/**
 * Логує зміну значення клітинки (тільки для несервісних колонок)
 */
function logCellEdit(e) {
  const sheet = e.range.getSheet();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Визначаємо назву колонки
  let header = '';
  try {
    header = sheet.getRange(1, col).getValue();
  } catch (err) {
    header = '';
  }
  if (IGNORED_HEADERS.some(name =>
    String(header).trim().toLowerCase() === name.trim().toLowerCase()
  )) {
    return;
  }

  // Пропускаємо якщо це зміна в заголовку
  if (row === 1) return;

  const user = Session.getActiveUser().getEmail();
  const time = new Date();
  const oldValue = e.oldValue !== undefined ? e.oldValue : "";
  const newValue = e.value !== undefined ? e.value : "";

  let changeType = "";
  if ((oldValue === "" || oldValue === null) && newValue !== "") {
    changeType = "Додано значення";
  } else if (oldValue !== "" && (newValue === "" || newValue === null)) {
    changeType = "Видалено значення";
  } else {
    changeType = "Змінено";
  }

  const cellLink = `=HYPERLINK("#gid=${sheet.getSheetId()}&range=${e.range.getA1Notation()}"; "${e.range.getA1Notation()}")`;

  logSheet.appendRow([
    time,
    user,
    sheet.getName(),
    cellLink,
    changeType,
    oldValue,
    newValue,
    "", // Формула (не використовується тут)
    ""  // Важлива зміна (не використовується тут)
  ]);
}

/**
 * Створення аркуша для логів (якщо не існує)
 */
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    const headers = [[
      "Час зміни",
      "Користувач",
      "Аркуш",
      "Посилання на комірку",
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
