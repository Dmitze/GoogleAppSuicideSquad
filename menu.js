function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. Меню "🛠 Дії з таблицею"
  const mainMenu = ui.createMenu("🛠 Дії з таблицею")
    .addItem("📱 Згенерувати QR-коди з усіма даними", "generateFullInfoQRCodesForSheet")
    .addItem("🔍 Гнучка генерація ключів", "generateKeysWithCustomColumns")
    .addItem("📱 Згенерувати QR-коди за Постійним ID", "generateQRCodesForSheet")
    .addItem("📁 Експортувати товари з ключами (CSV)", "exportProductsWithKeysToCSV");

  const formattingMenu = ui.createMenu("📏 Форматування")
    .addItem("📐 Вирівняти висоту рядків", "showSidebar");

  // 3. Меню "📄 Експорт до Word"
  const wordExportMenu = ui.createMenu("📄 Експорт до Word")
    .addItem("📑 Скласти звіт у Word...", "showWordExportFullForm")
    .addItem("📘 Експортувати до Word...", "showExportToWordDialog")
    .addItem("📖 Експортувати виділений діапазон до Word", "exportSheetRangeToWord");

  // 4. Меню "🕵️‍♂️ Перевірка та логи"
  const validationMenu = ui.createMenu("🕵️‍♂️ Перевірка та логи")
    .addItem("📊 Звіт по діям користувачів", "showUsersActionReport")
    .addItem("Ввести ім’я", "promptForUsername")

  // 5. Меню "📦 Логи та архівація"
  const logMenu = ui.createMenu("📦 Логи та архівація")
    .addItem("Сделать бекап в CSV", "backupLogAsCSV")
    .addItem("Открыть форму импорта", "showBackupForm");

  // 6. Меню "🔍 Пошук"
  const searchMenu = ui.createMenu("🔍 Пошук")
    .addItem('Гнучкий пошук по всіх листах', 'showGlobalFuzzySearchDialog');

  // Главное меню
  ui.createMenu("📋 Головне меню")
    .addSubMenu(mainMenu)
    .addSubMenu(formattingMenu)
    .addSubMenu(wordExportMenu)
    .addSubMenu(validationMenu)
    .addSubMenu(logMenu)
    .addSubMenu(searchMenu)
    .addToUi();

  if (typeof setupLogSheet === 'function') {
    setupLogSheet();
  }
}
