const showGlobalFuzzySearchDialog = () => {
  const html = HtmlService.createHtmlOutputFromFile('GlobalFuzzySearch')
    .setWidth(650)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Гнучкий пошук по всіх листах');
};

const globalFuzzySheetSearch = (query) => {
  if (!query || !query.trim()) return [];
  query = query.trim().toLowerCase();
  const queryWords = query.split(/\s+/);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const results = [];
  const MAX_LEV_DIST = 2;

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const data = sheet.getDataRange().getValues();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = data[r][c];
        if (val === "" || val === null) continue;
        const valStr = String(val).toLowerCase();

        if (valStr === query) {
          results.push(makeHit(sheet, r, c, val, "Точний збіг"));
          continue;
        }
        if (valStr.includes(query)) {
          results.push(makeHit(sheet, r, c, val, "Знайдено підрядок"));
          continue;
        }
        const allWords = queryWords.every(word => valStr.includes(word));
        if (allWords && queryWords.length > 1) {
          results.push(makeHit(sheet, r, c, val, "Всі слова зустрічаються"));
          continue;
        }
        let fuzzy = false;
        let minDist = 99;
        for (const qw of queryWords) {
          const wordsInCell = valStr.split(/\s+/);
          for (const wc of wordsInCell) {
            const dist = levenshtein(qw, wc);
            if (dist < minDist) minDist = dist;
            if (dist <= MAX_LEV_DIST) fuzzy = true;
          }
        }
        if (fuzzy) {
          results.push(makeHit(sheet, r, c, val, "Схоже значення (помилка/опечатка)"));
          continue;
        }
      }
    }
  });
  return results;
};

const makeHit = (sheet, r, c, val, reason) => ({
  sheet: sheet.getName(),
  row: r + 1,
  col: c + 1,
  a1: getA1Notation(r + 1, c + 1),
  value: String(val),
  link: `#gid=${sheet.getSheetId()}&range=${getA1Notation(r + 1, c + 1)}`,
  reason
});

const getA1Notation = (row, col) => {
  let letters = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    col = Math.floor((col - 1) / 26);
  }
  return letters + row;
};

const levenshtein = (a, b) => {
  if (a === b) return 0;
  const al = a.length, bl = b.length;
  if (al === 0) return bl;
  if (bl === 0) return al;
  const v0 = new Array(bl + 1);
  const v1 = new Array(bl + 1);
  for (let i = 0; i <= bl; i++) v0[i] = i;
  for (let i = 0; i < al; i++) {
    v1[0] = i + 1;
    for (let j = 0; j < bl; j++) {
      const cost = (a[i] === b[j]) ? 0 : 1;
      v1[j + 1] = Math.min(
        v1[j] + 1,
        v0[j + 1] + 1,
        v0[j] + cost
      );
    }
    for (let j = 0; j <= bl; j++) v0[j] = v1[j];
  }
  return v1[bl];
};

// Вернет прямую ссылку на таблицу (Google Sheets)
const getSpreadsheetUrl = () => SpreadsheetApp.getActiveSpreadsheet().getUrl();
