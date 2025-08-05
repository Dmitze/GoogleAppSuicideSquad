// ===== ОСНОВНІ ФУНКЦІЇ =====
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("📋 Головне меню")
      // Форматування
      .addSubMenu(ui.createMenu("📏 Форматування")
        .addItem("💡 Підсвітити збіги ID БпЛА", "highlightMatchingValues")
        .addItem("💡 Очистити кольори ID БпЛА", "clearHighlights")
        .addItem("🧩 Інструменти (Sidebar)", "showSidebar"))
      // Експорт
      .addSubMenu(ui.createMenu("📄 Експорт / Звіти")
        .addItem("📑 Скласти звіт у Word/PDF/Excel...", "showWordExportFullForm")
        .addItem("📘 Експортувати до Word/PDF/Excel...", "showExportToWordDialog")
        .addItem("📖 Експортувати виділений діапазон", "exportSheetRangeToWord"))
      // Аналітика та логи
      .addSubMenu(ui.createMenu("🕵️‍♂️ Перевірка та логи")
        .addItem("📊 Дашборд дій користувачів", "showUsersActionReport")
        .addItem("Ввести ім'я", "promptForUsername"))
      // Логи та архівація
      .addSubMenu(ui.createMenu("📦 Логи та архівація")
        .addItem("💾 Зробити бекап в CSV", "backupLogAsCSV")
        .addItem("📥 Відкрити форму імпорту", "showBackupForm"))
      // Пошук
      .addSubMenu(ui.createMenu("🔍 Пошук")
        .addItem("Гнучкий пошук по всіх листах", "showGlobalFuzzySearchDialog"))
      // Синхронізація
      .addSubMenu(ui.createMenu("🔄 Синхронізація")
        .addItem("🔍 Перевірити ідентифікатори", "checkIdentifiers"))
      .addToUi();
  } catch (e) {
    SpreadsheetApp.getUi().alert("Помилка при створенні меню: " + e);
  }
}
