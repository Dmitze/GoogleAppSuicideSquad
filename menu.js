// ===== ОСНОВНІ ФУНКЦІЇ =====
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu("📋 Головне меню")
      .addSubMenu(ui.createMenu("📏 Форматування")
        .addItem("💡 Підсвітити збіги ID БпЛА", "highlightMatchingValues")
        .addItem("💡 Очистити кольори ID БпЛА", "clearHighlights")
        .addItem("🧩 Інструменти", "showSidebar"))
      .addSubMenu(ui.createMenu("📄 Експорт до Word")
        .addItem("📑 Скласти звіт у Word...", "showWordExportFullForm")
        .addItem("📘 Експортувати до Word...", "showExportToWordDialog")
        .addItem("📖 Експортувати виділений діапазон до Word", "exportSheetRangeToWord"))
      .addSubMenu(ui.createMenu("🕵️‍♂️ Перевірка та логи")
        .addItem("📊 Звіт по діям користувачів", "showUsersActionReport")
        .addItem("Ввести ім'я", "promptForUsername"))
      .addSubMenu(ui.createMenu("📦 Логи та архівація")
        .addItem("Сделать бекап в CSV", "backupLogAsCSV")
        .addItem("Открыть форму импорта", "showBackupForm"))
      .addSubMenu(ui.createMenu("🔍 Пошук")
        .addItem('Гнучкий пошук по всіх листах', 'showGlobalFuzzySearchDialog'))
      .addSubMenu(ui.createMenu('Синхронизация')
        .addItem('🔍 Перевірити ідентифікатори', 'checkIdentifiers')
}
