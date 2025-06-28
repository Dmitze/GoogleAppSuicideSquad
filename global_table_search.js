function showGlobalFuzzySearchDialog() {
  var html = HtmlService.createHtmlOutputFromFile('GlobalFuzzySearch')
    .setWidth(650)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Гнучкий пошук по всіх листах');
}

function globalFuzzySheetSearch(query) {
  if (!query || !query.trim()) return [];
  query = query.trim().toLowerCase();
  var queryWords = query.split(/\s+/);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var results = [];
  var MAX_LEV_DIST = 2;

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var data = sheet.getDataRange().getValues();
    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        var val = data[r][c];
        if (val === "" || val === null) continue;
        var valStr = String(val).toLowerCase();

        if (valStr === query) {
          results.push(makeHit(sheet, r, c, val, "Точний збіг"));
          continue;
        }
        if (valStr.indexOf(query) !== -1) {
          results.push(makeHit(sheet, r, c, val, "Знайдено підрядок"));
          continue;
        }
        let allWords = queryWords.every(word => valStr.indexOf(word) !== -1);
        if (allWords && queryWords.length > 1) {
          results.push(makeHit(sheet, r, c, val, "Всі слова зустрічаються"));
          continue;
        }
        let fuzzy = false;
        let minDist = 99;
        for (let qw of queryWords) {
          let wordsInCell = valStr.split(/\s+/);
          for (let wc of wordsInCell) {
            let dist = levenshtein(qw, wc);
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
}

function makeHit(sheet, r, c, val, reason) {
  return {
    sheet: sheet.getName(),
    row: r + 1,
    col: c + 1,
    a1: getA1Notation(r + 1, c + 1),
    value: String(val),
    link: `#gid=${sheet.getSheetId()}&range=${getA1Notation(r + 1, c + 1)}`,
    reason: reason
  };
}

function getA1Notation(row, col) {
  var letters = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    col = Math.floor((col - 1) / 26);
  }
  return letters + row;
}

function levenshtein(a, b) {
  if (a === b) return 0;
  var al = a.length, bl = b.length;
  if (al === 0) return bl;
  if (bl === 0) return al;
  var v0 = new Array(bl + 1), v1 = new Array(bl + 1);
  for (var i = 0; i <= bl; i++) v0[i] = i;
  for (var i = 0; i < al; i++) {
    v1[0] = i + 1;
    for (var j = 0; j < bl; j++) {
      var cost = (a[i] === b[j]) ? 0 : 1;
      v1[j + 1] = Math.min(v1[j] + 1, v0[j + 1] + 1, v0[j] + cost);
    }
    for (var j = 0; j <= bl; j++) v0[j] = v1[j];
  }
  return v1[bl];
}
