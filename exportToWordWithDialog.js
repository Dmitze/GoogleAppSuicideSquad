// === КОНФИГУРАЦИЯ ===
const DEFAULT_DOC_NAME = "Експортований лист";
const DEFAULT_WORD_FILE_NAME = "ExportedSheet.docx";

// Получение списка начальников для автозаполнения
function getBossesList() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Довідники');
  if (!sheet) return [];
  let data = sheet.getRange('A2:A').getValues().flat().filter(String);
  return data;
}

// === Показать диалог выбора листа и діапазона ===
function showExportToWordDialog() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetOptions = sheets.map(s => `<option value="${s.getName()}">${s.getName()}</option>`).join("");
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif;">
      <h3>Експорт до Word / PDF / Excel</h3>
      <label>Лист:</label>
      <select id="sheetName">${sheetOptions}</select>
      <button onclick="updateSheets()" style="margin-left:10px;">🔄 Оновити</button>
      <br><br>
      <label>Діапазон (наприклад, A1:K27):</label>
      <input type="text" id="range" value="A1:K27" style="width:120px;">
      <br><br>
      <label>Ім'я файлу (без розширення):</label>
      <input type="text" id="fileName" value="ExportedSheet" style="width:180px;">
      <br><br>
      <button onclick="exportNow('docx')" style="font-size:1.1em;">📄 Word</button>
      <button onclick="exportNow('pdf')" style="font-size:1.1em;margin-left:10px;">🔽 PDF</button>
      <button onclick="exportNow('xlsx')" style="font-size:1.1em;margin-left:10px;">📊 Excel</button>
      <div id="status" style="margin-top:15px;"></div>
      <script>
        function updateSheets() {
          google.script.run.withSuccessHandler(function(names){
            var sel = document.getElementById('sheetName');
            var curr = sel.value;
            sel.innerHTML = names.map(function(n){return '<option value="'+n+'">'+n+'</option>';}).join('');
            sel.value = curr || names[0];
          }).getSheetNames();
        }
        function exportNow(format) {
          const sheet = document.getElementById('sheetName').value;
          const range = document.getElementById('range').value;
          const fileName = document.getElementById('fileName').value || 'ExportedSheet';
          if (!sheet || !range || !fileName) {
            document.getElementById('status').innerHTML = '<b style="color:red;">❌ Заповніть всі поля!</b>';
            return;
          }
          document.getElementById('status').innerHTML = '<div style="width:100%;height:10px;background:#eee;border-radius:5px;overflow:hidden;"><div id="pb" style="height:10px;width:30%;background:linear-gradient(90deg,#4a90e2,#2ecc71,#f39c12,#e74c3c,#4a90e2);background-size:200% 100%;animation:progressmove 1s linear infinite;"></div></div><span>⏳ Зачекайте...</span>';
          if (format === 'docx') {
            google.script.run
              .withSuccessHandler(url => {
                document.getElementById('status').innerHTML =
                  '<b>✅ Word-файл створено!</b><br><a href="'+url+'" target="_blank">Завантажити</a>';
              })
              .withFailureHandler(err => {
                document.getElementById('status').innerHTML =
                  '<b style="color:red;">❌ ' + (err.message || err) + '</b>';
              })
              .exportSheetRangeToWordCustom(sheet, range, fileName + ".docx");
          } else if (format === 'pdf') {
            google.script.run
              .withSuccessHandler(url => {
                document.getElementById('status').innerHTML =
                  '<b>✅ PDF-файл створено!</b><br><a href="'+url+'" target="_blank">Завантажити</a>';
              })
              .withFailureHandler(err => {
                document.getElementById('status').innerHTML =
                  '<b style="color:red;">❌ ' + (err.message || err) + '</b>';
              })
              .exportSheetRangeToPdfCustom(sheet, range, fileName + ".pdf");
          } else if (format === 'xlsx') {
            google.script.run
              .withSuccessHandler(url => {
                document.getElementById('status').innerHTML =
                  '<b>✅ Excel-файл створено!</b><br><a href="'+url+'" target="_blank">Завантажити</a>';
              })
              .withFailureHandler(err => {
                document.getElementById('status').innerHTML =
                  '<b style="color:red;">❌ ' + (err.message || err) + '</b>';
              })
              .exportSheetRangeToExcelCustom(sheet, range, fileName + ".xlsx");
          }
        }
      </script>
      <style>
        @keyframes progressmove {
          0% {background-position: 0% 50%;}
          100% {background-position: 100% 50%;}
        }
      </style>
    </div>
  `).setWidth(700).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, "Експорт до Word / PDF / Excel");
}

// === Экспорт в Word ===
function exportSheetRangeToWordCustom(sheetName, rangeA1, wordFileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Лист "${sheetName}" не знайдено!`);
    const values = sheet.getRange(rangeA1).getValues();
    if (!values || !values.length) throw new Error("Діапазон порожній або невірний!");

    const doc = DocumentApp.create(wordFileName.replace(/\.docx$/i, ''));
    doc.getBody().appendTable(values);
    doc.saveAndClose();

    const token = ScriptApp.getOAuthToken();
    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=docx`;
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const blob = response.getBlob().setName(wordFileName.endsWith('.docx') ? wordFileName : wordFileName + ".docx");
    const file = DriveApp.createFile(blob);

    return file.getUrl();
  } catch (e) {
    throw new Error(e && e.message ? e.message : e);
  }
}

function generatePdfReport(formData) {
  // Создаём Google Doc
  const doc = DocumentApp.create('Word звіт');
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

  // Экспортируем как PDF
  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs('application/pdf').setName("WordReport.pdf");
  const exported = DriveApp.createFile(pdfBlob);

  // Очищаем временный Google Doc
  file.setTrashed(true);

  return exported.getUrl();
}

function generateExcelReport(formData) {
  const ss = SpreadsheetApp.create('ExcelReport');
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
  const blob = file.getBlob().setName("ExcelReport.xlsx");
  const exported = DriveApp.createFile(blob);
  file.setTrashed(true);
  return exported.getUrl();
}

// Открытие HTML-формы (WordExportForm.html)
function showWordExportFullForm() {
  const html = HtmlService.createHtmlOutputFromFile('WordExportForm')
    .setWidth(1200).setHeight(1600);
  SpreadsheetApp.getUi().showModalDialog(html, "Генератор Word-звіту");
}

// Получение списка всех листов для формы
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// Получение данных диапазона для предварительного просмотра
function getPreviewData(sheetName, rangeA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(rangeA1).getValues();
}
