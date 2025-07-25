
function showUsersActionReport() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("Лист 'Лог змін' не знайдено.");
    return;
  }
  const data = logSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Немає записаних змін для звіту.");
    return;
  }

  // Индексы колонок
  const IDX = {date:0, user:1, sheet:2, cell:3, type:4, oldV:5, newV:6};

  // Сбор статистики
  const userStats = {};
  const userActions = {};
  const dayStats = {};   // статистика по дням
  const typeStats = {};  // статистика по типам изменений

  for (let i = 1; i < data.length; i++) {
    const user = data[i][IDX.user] || "[невідомо]";
    const type = (data[i][IDX.type] || "").trim();
    const dtRaw = (data[i][IDX.date] || "");
    const sheet = (data[i][IDX.sheet] || "");
    const cell = (data[i][IDX.cell] || "");
    const oldV = (data[i][IDX.oldV] || "");
    const newV = (data[i][IDX.newV] || "");
    const day = dtRaw ? String(dtRaw).split(" ")[0] : "";

    if (!userStats[user]) userStats[user] = { "Додано значення":0, "Видалено значення":0, "Змінено":0, "Додано рядок":0, "Видалено рядок":0, "Додано стовпець":0, "Видалено стовпець":0, "Зміна значення":0, _total:0 };
    if (!userActions[user]) userActions[user] = [];
    if (!dayStats[day]) dayStats[day] = 0;
    if (!typeStats[type]) typeStats[type] = 0;

    // Типы
    if (userStats[user][type] !== undefined) userStats[user][type]++;
    else if (type) userStats[user][type] = 1;
    userStats[user]._total++;

    // По дням/типам
    dayStats[day]++;
    typeStats[type]++;

    // Детали
    userActions[user].push(`${dtRaw}: [${sheet} ${cell}] ${type} | було: "${oldV}" → стало: "${newV}"`);
  }

  // Сортировка пользователей по активности
  const sortedUsers = Object.keys(userStats).sort((a, b) => userStats[b]._total - userStats[a]._total);

  // Генерируем HTML-отчет
  let html = `
  <style>
    body { font-family: Arial, sans-serif; font-size: 14px; }
    table { border-collapse: collapse; margin-bottom: 20px; }
    th, td { border: 1px solid #b4c7e7; padding: 4px 8px; text-align: left; }
    th { background: #e3f2fd; }
    .user-table { margin-bottom: 30px; }
    .details { font-size: 12px; color: #555; max-height: 180px; overflow-y: auto;}
    .exp-btn { margin: 8px 0 24px 0; padding: 7px 18px; font-size: 13px; border-radius: 8px; border: 1px solid #b4c7e7; background: #f9fafc; cursor: pointer;}
    .exp-btn:hover { background: #e3f2fd; }
  </style>
  <h2>Звіт по діям користувачів</h2>
  <button class="exp-btn" onclick="exportUserReport()">⬇️ Експортувати у CSV</button>
  <table class="user-table">
    <tr>
      <th>Користувач</th>
      <th>Додано</th>
      <th>Видалено</th>
      <th>Змінено</th>
      <th>Додано рядків</th>
      <th>Видалено рядків</th>
      <th>Додано стовпців</th>
      <th>Видалено стовпців</th>
      <th>Всього дій</th>
    </tr>
  `;

  for (const user of sortedUsers) {
    html += `
    <tr>
      <td>${user}</td>
      <td>${userStats[user]["Додано значення"] || 0}</td>
      <td>${userStats[user]["Видалено значення"] || 0}</td>
      <td>${userStats[user]["Змінено"] || 0}</td>
      <td>${userStats[user]["Додано рядок"] || 0}</td>
      <td>${userStats[user]["Видалено рядок"] || 0}</td>
      <td>${userStats[user]["Додано стовпець"] || 0}</td>
      <td>${userStats[user]["Видалено стовпець"] || 0}</td>
      <td>${userStats[user]._total || 0}</td>
    </tr>
    `;
  }
  html += "</table>";

  // Статистика по дням (график тренда)
  html += `<h3>Активність по днях</h3><table><tr><th>Дата</th><th>Кількість дій</th></tr>`;
  Object.keys(dayStats).sort().forEach(day => {
    html += `<tr><td>${day}</td><td>${dayStats[day]}</td></tr>`;
  });
  html += "</table>";

  // Статистика по типам змін
  html += `<h3>Статистика за типами дій</h3><table><tr><th>Тип дії</th><th>Кількість</th></tr>`;
  Object.keys(typeStats).sort((a,b) => typeStats[b] - typeStats[a]).forEach(type => {
    html += `<tr><td>${type}</td><td>${typeStats[type]}</td></tr>`;
  });
  html += "</table>";

  // Детализированный список
  html += `<h3>Деталізований перелік по користувачах</h3>`;
  for (const user of sortedUsers) {
    html += `<b>${user}:</b><div class="details"><ul>`;
    for (const action of userActions[user]) {
      html += `<li>${action}</li>`;
    }
    html += `</ul></div>`;
  }

  // Экспорт в CSV (прямо из диалога)
  html += `
  <script>
    function exportUserReport() {
      const table = document.querySelector('.user-table');
      let csv = '';
      for (const row of table.rows) {
        const cells = Array.from(row.cells).map(cell => '"'+cell.innerText.replace(/"/g,'""')+'"');
        csv += cells.join(',') + '\\r\\n';
      }
      const blob = new Blob([csv], {type:'text/csv'});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'user_action_report.csv';
      document.body.appendChild(a);
      a.click();
      setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},600);
    }
  </script>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(850).setHeight(720),
    "Звіт по діям користувачів"
  );
}
