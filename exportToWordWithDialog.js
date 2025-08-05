const DEFAULT_DOC_NAME = "Експортований лист";
const DEFAULT_WORD_FILE_NAME = "ExportedSheet.docx";

function getBossesList() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Довідники');
  if (!sheet) return [];
  let data = sheet.getRange('A2:A').getValues().flat().filter(String);
  return data;
}

function showExportToWordDialog() {
  const sheets = getSheetNames();
  const sheetOptions = sheets.map(s => `<option value="${s}">${s}</option>`).join("");
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif;">
      <h3>Експорт до Word / PDF / Excel</h3>
      <label>Лист:</label>
      <select id="sheetName">${sheetOptions}</select>
      <button onclick="updateSheets()" style="margin-left:10px;">🔄 Оновити</button>
      <button onclick="useSelectedRange()" style="margin-left:10px;">📋 Використати виділений діапазон</button>
      <br><br>
      <label>Діапазон (наприклад, A1:K27):</label>
      <input type="text" id="range" value="A1:K27" style="width:120px;">
      <br><br>
      <label>Ім'я файлу (без розширення):</label>
      <input type="text" id="fileName" value="ExportedSheet" style="width:180px;">
      <br><br>
      <label>Email для відправки (необов'язково):</label>
      <input type="email" id="email" placeholder="user@email.com" style="width:180px;">
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
        function useSelectedRange() {
          google.script.run.withSuccessHandler(function(a1){
            if(a1) document.getElementById('range').value = a1;
          }).getActiveRangeA1();
        }
        function exportNow(format) {
          const sheet = document.getElementById('sheetName').value;
          const range = document.getElementById('range').value;
          const fileName = document.getElementById('fileName').value || 'ExportedSheet';
          const email = document.getElementById('email').value;
          if (!sheet || !range || !fileName) {
            document.getElementById('status').innerHTML = '<b style="color:red;">❌ Заповніть всі поля!</b>';
            return;
          }
          document.getElementById('status').innerHTML = '<span>⏳ Зачекайте...</span>';
          let exportFunc = format === 'docx' ? 'exportRangeToWord' : (format === 'pdf' ? 'exportRangeToPdf' : 'exportRangeToExcel');
          google.script.run.withSuccessHandler(function(url){
            document.getElementById('status').innerHTML = '<b>✅ Файл створено!</b><br><a href="'+url+'" target="_blank">Завантажити</a>';
            if(email) {
              google.script.run.sendFileByEmail(email, url, fileName + '.' + format, 'Ваш файл у вкладенні.');
            }
          }).withFailureHandler(function(err){
            document.getElementById('status').innerHTML = '<b style="color:red;">❌ ' + (err.message || err) + '</b>';
          })[exportFunc](sheet, range, fileName + '.' + format);
        }
      </script>
    </div>
  `).setWidth(700).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, "Експорт до Word / PDF / Excel");
}
 
function generatePdfReport(formData) {
  const doc = DocumentApp.create('Word звіт');
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
  const pdfBlob = file.getAs('application/pdf').setName("WordReport.pdf");
  const exported = DriveApp.createFile(pdfBlob);
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

function showWordExportFullForm() {
  const html = HtmlService.createHtmlOutputFromFile('WordExportForm')
    .setWidth(1200).setHeight(1600);
  SpreadsheetApp.getUi().showModalDialog(html, "Генератор Word-звіту");
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

function getPreviewData(sheetName, rangeA1) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(rangeA1).getValues();
}
