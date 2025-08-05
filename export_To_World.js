// === КОНФИГУРАЦИЯ ===
const DEFAULT_DOC_NAME = "Експортований лист";
const DEFAULT_WORD_FILE_NAME = "ExportedSheet.docx";

// === Импортируем универсальные функции ===
// (В Google Apps Script просто убедитесь, что export_utils.js подключён в проекте)

function showWordExportFullForm() {
  const html = HtmlService.createHtmlOutputFromFile('WordExportForm')
    .setWidth(1200).setHeight(1600);
  SpreadsheetApp.getUi().showModalDialog(html, "Генератор Word-звіту");
}

// === ГЕНЕРАЦИЯ DOCX, PDF, XLSX С ОТЧЁТОМ ===
function generateWordReport(formData) {
  try {
    const doc = DocumentApp.create(formData.title || DEFAULT_DOC_NAME);
    const body = doc.getBody();

    // Шапка
    if (formData.header) {
      if (formData.header.boss) body.appendParagraph(`Начальник: ${formData.header.boss}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (formData.header.date) body.appendParagraph(`Дата: ${formData.header.date}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (formData.header.order) body.appendParagraph(`Наказ: ${formData.header.order}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph('');
    }
    if (formData.title) body.appendParagraph(formData.title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (formData.description) body.appendParagraph(formData.description).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Таблицы
    if (formData.tables && Array.isArray(formData.tables)) {
      formData.tables.forEach((t, idx) => {
        if (!t.sheet || !t.range) return;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheet);
        if (!sheet) return;
        const values = sheet.getRange(t.range).getValues();
        body.appendParagraph(`Таблиця ${idx+1}: ${t.sheet} ${t.range}`);
        body.appendTable(values);
        body.appendParagraph('');
      });
    }
    doc.saveAndClose();

    const token = ScriptApp.getOAuthToken();
    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const blob = response.getBlob().setName((formData.title || DEFAULT_WORD_FILE_NAME).replace(/[^a-zA-Zа-яА-Я0-9_.-]/g, "_") + ".docx");
    const file = DriveApp.createFile(blob);

    return file.getUrl();
  } catch (e) {
    throw new Error(e && e.message ? e.message : e);
  }
}

function generatePdfReport(formData) {
  try {
    const doc = DocumentApp.create(formData.title || DEFAULT_DOC_NAME);
    const body = doc.getBody();

    if (formData.header) {
      if (formData.header.boss) body.appendParagraph(`Начальник: ${formData.header.boss}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (formData.header.date) body.appendParagraph(`Дата: ${formData.header.date}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (formData.header.order) body.appendParagraph(`Наказ: ${formData.header.order}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph('');
    }
    if (formData.title) body.appendParagraph(formData.title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (formData.description) body.appendParagraph(formData.description).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (formData.tables && Array.isArray(formData.tables)) {
      formData.tables.forEach((t, idx) => {
        if (!t.sheet || !t.range) return;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheet);
        if (!sheet) return;
        const values = sheet.getRange(t.range).getValues();
        body.appendParagraph(`Таблиця ${idx+1}: ${t.sheet} ${t.range}`);
        body.appendTable(values);
        body.appendParagraph('');
      });
    }

    doc.saveAndClose();

    const file = DriveApp.getFileById(doc.getId());
    const pdfBlob = file.getAs('application/pdf').setName((formData.title || "WordReport") + ".pdf");
    const exported = DriveApp.createFile(pdfBlob);

    file.setTrashed(true);

    return exported.getUrl();
  } catch (e) {
    throw new Error(e && e.message ? e.message : e);
  }
}

function generateExcelReport(formData) {
  try {
    const ss = SpreadsheetApp.create(formData.title || "ExcelReport");
    if (formData.tables && Array.isArray(formData.tables)) {
      formData.tables.forEach((t, idx) => {
        if (!t.sheet || !t.range) return;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheet);
        if (!sheet) return;
        const values = sheet.getRange(t.range).getValues();
        let newSheet = idx === 0 ? ss.getSheets()[0] : ss.insertSheet();
        newSheet.setName(`${t.sheet}_${idx+1}`);
        newSheet.getRange(1,1,values.length,values[0].length).setValues(values);
      });
    }
    const file = DriveApp.getFileById(ss.getId());
    const blob = file.getBlob().setName((formData.title || "ExcelReport") + ".xlsx");
    const exported = DriveApp.createFile(blob);
    file.setTrashed(true);
    return exported.getUrl();
  } catch (e) {
    throw new Error(e && e.message ? e.message : e);
  }
}
