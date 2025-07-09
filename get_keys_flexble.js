function generateKeysWithCustomColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("На аркуші немає даних!");
    return;
  }
  const header = data[0];
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Вкажіть до 3 назв колонок через кому (наприклад: Марка,Модель,VIN):',
    'Виберіть колонки, які будуть використані для створення ключа. Враховуйте точність написання!',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const keyCols = result.getResponseText().split(',').map(s => s.trim()).filter(Boolean).slice(0, 3);

  if (keyCols.length === 0) {
    ui.alert("Не вказано жодної колонки!");
    return;
  }

  generateKeysWithPermanentId(sheet, header, keyCols);
}

function generateKeysWithPermanentId(sheet, header, keyCols) {
  if (!sheet || typeof sheet.getDataRange !== 'function') {
    throw new Error('sheet не передано або це не обʼєкт аркуша Google Таблиць!');
  }
  if (!Array.isArray(header) || header.length === 0) {
    throw new Error('header не передано або він порожній!');
  }
  if (!Array.isArray(keyCols) || keyCols.length === 0) {
    throw new Error('keyCols не передано або він порожній!');
  }

  const data = sheet.getDataRange().getValues();
  const idCol = findOrCreateColumn(sheet, header, 'Ідентифікатор');
  const permCol = findOrCreateColumn(sheet, header, 'Постійний ID', true);
  const keyColIdxs = keyCols.map(name =>
    header.findIndex(h =>
      typeof h === 'string' &&
      h.trim().toLowerCase() === name.trim().toLowerCase()
    )
  );
  for (let row = 1; row < data.length; row++) {
    const keyBase = keyColIdxs.map(idx => (idx !== -1 ? (data[row][idx] || '') : '')).join('_');
    const hasData = keyColIdxs.every(idx => idx !== -1 && data[row][idx] && String(data[row][idx]).trim().length > 0);
    let permId = data[row][permCol];
    if (hasData && (!permId || permId === '')) {
      permId = Utilities.getUuid();
      sheet.getRange(row + 1, permCol + 1).setValue(permId);
    }
    if (hasData) {
      const id = generateProductEncryptedId(keyBase);
      sheet.getRange(row + 1, idCol + 1).setValue(id);
    } else {
      sheet.getRange(row + 1, idCol + 1).setValue('');
      sheet.getRange(row + 1, permCol + 1).setValue('');
    }
  }
}

function generateQRCodesForSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const header = data[0];
  const permCol = findOrCreateColumn(sheet, header, 'Постійний ID', true);
  const qrCol = findOrCreateColumn(sheet, header, 'QR-код');
  const qrLinkCol = findOrCreateColumn(sheet, header, 'Посилання на QR');
  for (let row = 1; row < data.length; row++) {
    const permId = data[row][permCol];
    if (permId && String(permId).trim().length > 0) {
      const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=" + encodeURIComponent(permId);
      const qrUrlBig = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(permId);
      sheet.getRange(row + 1, qrCol + 1).setFormula('=IMAGE("' + qrUrl + '")');
      sheet.getRange(row + 1, qrLinkCol + 1).setFormula('=HYPERLINK("' + qrUrlBig + '"; "Відкрити великий QR")');
    } else {
      sheet.getRange(row + 1, qrCol + 1).setValue('');
      sheet.getRange(row + 1, qrLinkCol + 1).setValue('');
    }
  }
}

function generateFullInfoQRCodesForSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("Немає даних для генерації QR-кодів!");
    return;
  }
  const header = data[0];

  clearColumnByHeader(sheet, header, 'QR');
  clearColumnByHeader(sheet, header, 'link');

  const qrCol = findOrCreateColumn(sheet, header, 'QR');
  const linkCol = findOrCreateColumn(sheet, header, 'link');
  for (let row = 1; row < data.length; row++) {
    let text = header
      .map((col, idx) =>
        (col && data[row][idx] && String(data[row][idx]).trim().length > 0)
          ? (col + ": " + data[row][idx])
          : null
      )
      .filter(Boolean)
      .join('\n');
    
    if (text.length > 400) {
      text = text.substring(0, 397) + '...';
    }

    if (text.trim().length > 0) {
      const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + encodeURIComponent(text);
      const qrUrlBig = "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=" + encodeURIComponent(text);
      sheet.getRange(row + 1, qrCol + 1).setFormula('=IMAGE("' + qrUrl + '")');
      sheet.getRange(row + 1, linkCol + 1).setFormula('=HYPERLINK("' + qrUrlBig + '"; "link")');
    } else {
      sheet.getRange(row + 1, qrCol + 1).setValue('');
      sheet.getRange(row + 1, linkCol + 1).setValue('');
    }
  }
}

function clearColumnByHeader(sheet, header, colName) {
  let idx = header.findIndex(
    h => typeof h === 'string' && h.trim().toLowerCase() === colName.trim().toLowerCase()
  );
  if (idx !== -1 && sheet.getLastRow() > 1) {
    sheet.getRange(2, idx + 1, sheet.getLastRow() - 1, 1).clearContent();
  }
}

function findOrCreateColumn(sheet, header, colName, isHidden) {
  let idx = header.findIndex(
    h => typeof h === 'string' && h.trim().toLowerCase() === colName.trim().toLowerCase()
  );
  if (idx === -1) {
    const lastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue(colName);
    if (isHidden) sheet.hideColumns(lastCol + 1);
    idx = lastCol;
  }
  return idx;
}

function exportProductsWithKeysToCSV() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Немає даних для експорту!");
    return;
  }

  const header = data[0];
  const exportCols = ['Марка', 'Модель', 'VIN', 'Ідентифікатор', 'Постійний ID', 'QR-код'];
  const colIndexes = exportCols.map(name =>
    header.findIndex(h => typeof h === 'string' && h.toLowerCase().includes(name.toLowerCase()))
  );

  const rows = [exportCols];
  for (let i = 1; i < data.length; i++) {
    const row = colIndexes.map(idx => (idx !== -1 ? data[i][idx] : ''));
    rows.push(row);
  }

  const csv = rows.map(r => r.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\r\n');

  const folder = DriveApp.getRootFolder();
  const fileName = 'eksport_tovariv_z_kliuchamy_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
  const file = folder.createFile(fileName, csv, MimeType.CSV);

  SpreadsheetApp.getUi().alert(`Файл створено!\n${file.getUrl()}`);
  return file.getUrl();
}

function generateProductEncryptedId(rawString) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawString + new Date().getFullYear());
  const base64 = Utilities.base64Encode(digest);
  return base64.replace(/[^A-Za-z0-9]/g, '').substring(0, 12);
}
