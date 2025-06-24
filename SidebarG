function setUniformRowHeight(rowHeight) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const lastRow = dataRange.getLastRow();

  if (!rowHeight || isNaN(rowHeight) || rowHeight <= 0) {
    SpreadsheetApp.getUi().alert("❌ Помилка\n\nВведіть коректне число для висоти.");
    return;
  }

  sheet.setRowHeights(1, lastRow, rowHeight);

  SpreadsheetApp.getUi().alert("✅ Готово\n\nВисота рядків встановлена на " + rowHeight + " px");
}
