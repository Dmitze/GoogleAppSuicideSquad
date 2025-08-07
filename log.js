
const LOG_SHEET_NAME = "Лог змін";
const COLOR_GREEN = "#b6d7a8";
const IGNORED_HEADERS = [
  'Постійний ID',
  'Ідентифікатор',
  'QR-код',
  'QR',
  'link',
  'Посилання на QR'
];

function promptForUsername() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("Введіть своє ім’я або позивний для логів:");
  if (result.getSelectedButton() === ui.Button.OK) {
    const name = result.getResponseText().trim();
    if (name) {
      PropertiesService.getUserProperties().setProperty("username", name);
      ui.alert("Ім’я збережено як: " + name);
    } else {
      ui.alert("Ім’я не може бути порожнім");
    }
  }
}

function logCellEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() === LOG_SHEET_NAME) return;

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  const header = sheet.getRange(1, col).getValue();

  if (row === 1 || IGNORED_HEADERS.includes(header)) return;

  const user = PropertiesService.getUserProperties().getProperty("username") || "Анонім";
  const time = new Date();
  const oldValue = e.oldValue ?? "";
  const newValue = e.value ?? "";
  const changeType =
    (oldValue === "" && newValue !== "") ? "Додано" :
    (oldValue !== "" && newValue === "") ? "Видалено" : "Змінено";

  const cellLink = `=HYPERLINK("#gid=${sheet.getSheetId()}&range=${e.range.getA1Notation()}"; "${e.range.getA1Notation()}")`;

  logSheet.appendRow([
    time,
    user,
    sheet.getName(),
    cellLink,
    changeType,
    oldValue,
    newValue,
    "", ""
  ]);
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === LOG_SHEET_NAME) return;

  // Підсвічуємо змінену клітинку
  highlightCell(e);

  // Логуємо зміну (якщо не службова колонка)
  logCellEdit(e);
}

function createInstallableTrigger() {
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}


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
  if (sheet.getName() === LOG_SHEET_NAME) return;

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  let header = '';
  try {
    header = sheet.getRange(1, col).getValue();
  } catch (err) {
    return;
  }

  // Проверка на игнорируемые заголовки (регистронезависимо)
  if (IGNORED_HEADERS.some(name =>
    String(header).trim().toLowerCase() === name.trim().toLowerCase()
  )) return;

  if (row === 1) return; // Заголовки не логируем

  // Получаем имя пользователя
  const username = PropertiesService.getUserProperties().getProperty("username");
  const email = Session.getActiveUser().getEmail();
  const user = username || (email ? email.split('@')[0] : "Анонім");

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
    "",
    ""
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

/**
 * Видаляє фрагменти SUMPRODUCT з формул у зазначених аркушах і діапазонах
 * Налаштування: змінюйте startSheetNum, endSheetNum, colStart, colEnd, targetRow, replacementValue, dryRun
 */
function removeSumproductFragments() {
  const config = {
    startSheetNum: 19,
    endSheetNum: 31,
    colStart: "G",
    colEnd: "W",
    targetRow: "105",
    replacementValue: "0", // Що ставити замість видаленого: "0", "", або "0"
    dryRun: false,         // true — лише показати, що буде змінено
  };

  // Генерація імен аркушів
  const sheetNames = [];
  for (let i = config.startSheetNum; i <= config.endSheetNum; i++) {
    sheetNames.push(`Sheet1 (${i})`);
  }

  // Екрануємо імена та формуємо фрагменти
  const escapedNames = sheetNames.map(name =>
    name.replace(/\(/g, '\\(').replace(/\)/g, '\\)')
  );

  const fragments = escapedNames.map(escapedName => {
    return `\\+?SUMPRODUCT\\(\\(MOD\\(COLUMN\\('${escapedName}'!${config.colStart}${config.targetRow}:${escapedName}'!${config.colEnd}${config.targetRow}\\)-COLUMN\\('${escapedName}'!${config.colStart}${config.targetRow}\\);2\\)=0\\)\\*\\('${escapedName}'!${config.colStart}${config.targetRow}:${escapedName}'!${config.colEnd}${config.targetRow}\\)\\)`;
  });

  // Об'єднуємо в один регулярний вираз
  const pattern = `(${fragments.join('|')})`;
  const regex = new RegExp(pattern, 'g');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  let totalChanges = 0;
  const changeLog = [];

  SpreadsheetApp.flush();

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const range = sheet.getDataRange();
    const formulas = range.getFormulas();

    let hasChangesInSheet = false;
    const updatedFormulas = formulas.map((row, rowIndex) =>
      row.map((formula, colIndex) => {
        if (!formula || !formula.includes("SUMPRODUCT")) return formula;

        const originalFormula = formula;
        const startsWithEqual = formula.startsWith('=');
        if (startsWithEqual) formula = formula.slice(1);

        // Застосовуємо заміну
        let newFormula = formula.replace(regex, "");

        // Очищаємо зайві плюси та пробіли
        newFormula = newFormula
          .replace(/\+\+/g, "+")
          .replace(/^\+/, "")
          .replace(/\+$/, "")
          .trim();

        // Якщо порожньо — замінюємо на replacementValue
        if (newFormula === "") {
          newFormula = config.replacementValue;
        }

        // Повертаємо '=' якщо було
        newFormula = startsWithEqual ? `=${newFormula}` : newFormula;

        if (newFormula !== originalFormula) {
          hasChangesInSheet = true;
          totalChanges++;
          changeLog.push({
            sheet: sheetName,
            cell: sheet.getRange(rowIndex + 1, colIndex + 1).getA1Notation(),
            from: originalFormula,
            to: newFormula
          });
          return newFormula;
        }
        return originalFormula;
      })
    );

    // Оновлюємо лише якщо були зміни
    if (hasChangesInSheet && !config.dryRun) {
      range.setFormulas(updatedFormulas);
    }
  });

  SpreadsheetApp.flush();

  // Вивід результату
  const ui = SpreadsheetApp.getUi();
  let msg = `✅ Готово! Змінено формул: ${totalChanges}`;

  if (config.dryRun) {
    msg += "\n\nРежим перегляду (dry run). Формули не змінені.";
  }

  if (totalChanges > 0 && !config.dryRun) {
    msg += "\n\nДеталі в логах (див. журнал виконання).";
  }

  if (totalChanges === 0) {
    msg = "❌ Нічого не знайдено для видалення.";
  }

  ui.alert(msg);

  // Друк у лог для деталей
  if (changeLog.length > 0) {
    console.log("Змінені формули:");
    changeLog.forEach(log => {
      console.log(`${log.sheet}!${log.cell}: \"${log.from}\" → \"${log.to}\"`);
    });
  }
}
