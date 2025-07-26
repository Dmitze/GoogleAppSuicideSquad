function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. Меню "🛠 Дії з таблицею"
  const mainMenu = ui.createMenu("🛠 Дії з таблицею")
    .addItem("📱 Згенерувати QR-коди з усіма даними", "generateFullInfoQRCodesForSheet")
    .addItem("🔍 Гнучка генерація ключів", "generateKeysWithCustomColumns")
    .addItem("📱 Згенерувати QR-коди за Постійним ID", "generateQRCodesForSheet")
    .addItem("📁 Експортувати товари з ключами (CSV)", "exportProductsWithKeysToCSV");

  const formattingMenu = ui.createMenu("📏 Форматування")
    .addItem("💡 Підсвітити збіги ID БпЛА", "highlightMatchingValues")
    .addItem("💡 Очистити кольори ID БпЛА", "clearHighlights")
    .addItem("🧩 Інструменти", "showSidebar");

  // 3. Меню "📄 Експорт до Word"
  const wordExportMenu = ui.createMenu("📄 Експорт до Word")
    .addItem("📑 Скласти звіт у Word...", "showWordExportFullForm")
    .addItem("📘 Експортувати до Word...", "showExportToWordDialog")
    .addItem("📖 Експортувати виділений діапазон до Word", "exportSheetRangeToWord");

  // 4. Меню "🕵️‍♂️ Перевірка та логи"
  const validationMenu = ui.createMenu("🕵️‍♂️ Перевірка та логи")
    .addItem("📊 Звіт по діям користувачів", "showUsersActionReport")
    .addItem("Ввести ім’я", "promptForUsername");

  // 5. Меню "📦 Логи та архівація"
  const logMenu = ui.createMenu("📦 Логи та архівація")
    .addItem("Сделать бекап в CSV", "backupLogAsCSV")
    .addItem("Открыть форму импорта", "showBackupForm");

  // 6. Меню "🔍 Пошук"
  const searchMenu = ui.createMenu("🔍 Пошук")
    .addItem('Гнучкий пошук по всіх листах', 'showGlobalFuzzySearchDialog');

  // 7. Меню "Синхронизация" (ручная синхронизация копий)
  const syncMenu = ui.createMenu('Синхронизация')
    .addItem('1РБпАК 2ББпАК', 'syncToCopy1')
    .addItem('2РБпАК 2ББпАК', 'syncToCopy2')
    .addItem('3РБпАК 2ББпАК', 'syncToCopy3');

  // Главное меню
  ui.createMenu("📋 Головне меню")
    .addSubMenu(mainMenu)
    .addSubMenu(formattingMenu)
    .addSubMenu(wordExportMenu)
    .addSubMenu(validationMenu)
    .addSubMenu(logMenu)
    .addSubMenu(searchMenu)
    .addSubMenu(syncMenu)
    .addToUi();

  if (typeof setupLogSheet === 'function') {
    setupLogSheet();
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle("Форматування та текстові перетворення");
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Синхронизация копий (оставляем из предыдущей версии) ---
const syncToCopies = () => {
  const copyIds = [
    "1j1GmrtdiDnK221kem2MGWQKZFX-8K-306PQIkAN7Xdo",
    "1DXym9zD5kaVj6dKaku8vO7Sl0IhuYqLoxKSoCiV0URk",
    "ID_третьей_копии"
  ];
  const sheetNames = ["Розвідувальні БпЛА", "Ударні БпЛА"];
  const col = 11; 

  const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  copyIds.forEach(copyId => {
    const copySpreadsheet = SpreadsheetApp.openById(copyId);
    sheetNames.forEach(sheetName => {
      const mainSheet = mainSpreadsheet.getSheetByName(sheetName);
      const copySheet = copySpreadsheet.getSheetByName(sheetName);
      if (mainSheet && copySheet) {
        const mainValues = mainSheet.getRange(1, col, mainSheet.getLastRow(), 1).getValues();
        copySheet.getRange(1, col, mainValues.length, 1).setValues(mainValues);
      }
    });
  });
};

const syncToCopy1 = () => {
  syncToSingleCopy("1oYU_XIVq0FAniR4Z0CvwN1iSbcpckWlzDLrO52Ae1Gc");
};

const syncToCopy2 = () => {
  syncToSingleCopy("1j1GmrtdiDnK221kem2MGWQKZFX-8K-306PQIkAN7Xdo");
};

const syncToCopy3 = () => {
  syncToSingleCopy("1DXym9zD5kaVj6dKaku8vO7Sl0IhuYqLoxKSoCiV0URk");
};

const syncToSingleCopy = copyId => {
  const sheetNames = ["Розвідувальні БпЛА", "Ударні БпЛА"];
  const col = 4; 

  const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const copySpreadsheet = SpreadsheetApp.openById(copyId);

  sheetNames.forEach(sheetName => {
    const mainSheet = mainSpreadsheet.getSheetByName(sheetName);
    const copySheet = copySpreadsheet.getSheetByName(sheetName);
    if (mainSheet && copySheet) {
      const mainValues = mainSheet.getRange(1, col, mainSheet.getLastRow(), 1).getValues();
      const mainRows = mainValues.length;
      let copyRows = copySheet.getMaxRows();
 
      if (copyRows < mainRows) {
        copySheet.insertRowsAfter(copyRows, mainRows - copyRows);
        copyRows = copySheet.getMaxRows();
      }

      const allowDelete = mainSpreadsheet.getRange("A1").getValue() === "OK";
      if (copyRows > mainRows && allowDelete) {
        backupRows(copySheet, mainRows + 1, copyRows - mainRows);
        copySheet.deleteRows(mainRows + 1, copyRows - mainRows);
      }
      copySheet.getRange(1, col, mainRows, 1).setValues(mainValues);
    }
  });
};

const backupRows = (copySheet, startRow, numRows) => {
  const backupSheetName = "Резервна копія";
  const ss = copySheet.getParent();
  let backupSheet = ss.getSheetByName(backupSheetName);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(backupSheetName);
  }
  const lastBackupRow = backupSheet.getLastRow();
  const dataToBackup = copySheet.getRange(startRow, 1, numRows, copySheet.getLastColumn()).getValues();
  backupSheet.getRange(lastBackupRow + 1, 1, numRows, dataToBackup[0].length).setValues(dataToBackup);
};
