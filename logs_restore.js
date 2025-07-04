<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Архіви логів</title>
  <style>
    body { font-family: Arial; }
    table { border-collapse: collapse; width: 100%; margin-top: 1em; }
    th, td { border: 1px solid #ddd; padding: 4px 8px; font-size: 90%; }
    th { background: #f2f2f2; }
    tr:hover { background: #fafad2; }
    #preview { max-height: 400px; overflow: auto; margin-top: 1em; }
    .download-btn { background: #b6d7a8; border: none; cursor: pointer; padding: 2px 7px; }
  </style>
</head>
<body>
  <h2>Архіви логів у Drive</h2>
  <div id="status"></div>
  <table id="files-table">
    <thead>
      <tr>
        <th>Назва</th>
        <th>Дата</th>
        <th>Перегляд</th>
        <th>Скачати</th>
        <th>Копія .xlsx</th>
      </tr>
    </thead>
    <tbody>
    </tbody>
  </table>
  <div id="preview"></div>
<script>
let archives = [];
function loadArchives() {
  document.getElementById('status').textContent = 'Завантаження...';
  google.script.run.withSuccessHandler(function(list){
    archives = list;
    renderTable(list);
    document.getElementById('status').textContent = 'Знайдено ' + list.length + ' архівів';
  }).getLogArchivesList();
}

function renderTable(list) {
  const tbody = document.getElementById('files-table').querySelector('tbody');
  tbody.innerHTML = "";
  list.forEach(arch => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${arch.name}</td>
      <td>${new Date(arch.date).toLocaleString()}</td>
      <td>
        ${arch.name.endsWith('.csv') ? `<button onclick="previewCsv('${arch.id}')">Переглянути</button>` : ''}
      </td>
      <td>
        <a href="${arch.url}" target="_blank">Скачати</a>
      </td>
      <td>
        <button onclick="makeCopy('${arch.id}')">Зробити копію .xlsx</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function previewCsv(fileId) {
  document.getElementById('preview').innerHTML = 'Завантаження...';
  google.script.run.withSuccessHandler(function(rows){
    let html = '<table>';
    rows.forEach(row => {
      html += '<tr>' + row.map(cell => `<td>${escapeHtml(cell)}</td>`).join('') + '</tr>';
    });
    html += '</table>';
    document.getElementById('preview').innerHTML = html;
  }).getCsvPreview(fileId, 100);
}

function makeCopy(fileId) {
  document.getElementById('status').textContent = 'Створюємо копію...';
  google.script.run.withSuccessHandler(function(url){
    document.getElementById('status').textContent = 'Копія створена! <a href="'+url+'" target="_blank">Відкрити</a>';
  }).makeCopyOfArchive(fileId);
}

function escapeHtml(s) { return String(s).replace(/[<>&"]/g, c=>({"<":"&lt;",">":"&gt;","&":"&amp;",'"':"&quot;"}[c])); }

window.onload = loadArchives;
</script>
</body>
</html>

