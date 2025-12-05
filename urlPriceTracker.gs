/**
 * Main entry point to check all tracked product URLs on sheet "Blad1".
 */
function checkTrackedProductUrls() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Blad1');
  if (!sheet) {
    Logger.log('Sheet "Blad1" not found.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows found.');
    return;
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 17);
  var rows = dataRange.getValues();
  var updates = [];
  var updateRowNumbers = [];
  var changes = [];
  var timestamp = new Date();

  for (var i = 0; i < rows.length; i++) {
    var rowIndex = i + 2; // Sheet rows start at 1, data starts at row 2
    var row = rows[i];

    var url = row[0];
    var category = row[1];
    var product = row[2];
    var productId = row[3];
    var brand = row[4];
    var model = row[5];
    var lastPrice = toNumberOrNull_(row[6]);
    var previousPrice = toNumberOrNull_(row[7]);
    var lowestPrice = toNumberOrNull_(row[8]);
    var watchFlag = row[10];

    if (!url || !isWatchEnabled_(watchFlag)) {
      continue;
    }

    try {
      var response = UrlFetchApp.fetch(url, {
        followRedirects: true,
        muteHttpExceptions: true,
      });
      var status = response.getResponseCode();
      if (status !== 200) {
        Logger.log('HTTP error for row %s (%s): %s', rowIndex, url, status);
        appendUrlPriceHistory_(ss, {
          timestamp: timestamp,
          rowIndex: rowIndex,
          url: url,
          product: product,
          brand: brand,
          model: model,
          price: null,
          note: 'HTTP ' + status,
        });
        continue;
      }

      var html = response.getContentText();
      var newPrice = parsePriceFromHtml_(html);
      if (!isFinite(newPrice)) {
        Logger.log('Could not parse price for row %s (%s)', rowIndex, url);
        appendUrlPriceHistory_(ss, {
          timestamp: timestamp,
          rowIndex: rowIndex,
          url: url,
          product: product,
          brand: brand,
          model: model,
          price: null,
          note: 'Price not found',
        });
        continue;
      }

      var lowestBeforeUpdate = isFinite(lowestPrice) ? lowestPrice : null;
      var newLowest = isFinite(lowestPrice) ? Math.min(lowestPrice, newPrice) : newPrice;
      var daling = isFinite(previousPrice) && previousPrice > 0
        ? (previousPrice - newPrice) / previousPrice
        : 0;

      var rowUpdate = row.slice(6, 17);
      rowUpdate[1] = lastPrice; // H: Vorige prijs
      rowUpdate[0] = newPrice; // G: Laatste prijs
      rowUpdate[2] = newLowest; // I: Laagste prijs
      rowUpdate[3] = daling; // J: Daling %
      rowUpdate[5] = timestamp; // L: Laatste check
      rowUpdate[9] = newPrice; // P: Scraped price
      rowUpdate[10] = timestamp; // Q: Scraped lastchecked

      updates.push(rowUpdate);
      updateRowNumbers.push(rowIndex);

      appendUrlPriceHistory_(ss, {
        timestamp: timestamp,
        rowIndex: rowIndex,
        url: url,
        product: product,
        brand: brand,
        model: model,
        price: newPrice,
        note: 'OK',
      });

      if (lowestBeforeUpdate !== null && newPrice < lowestBeforeUpdate && daling >= 0.05) {
        changes.push({
          rowIndex: rowIndex,
          category: category,
          product: product,
          productId: productId,
          brand: brand,
          model: model,
          url: url,
          previousPrice: previousPrice,
          newPrice: newPrice,
          lowestPriceBeforeUpdate: lowestBeforeUpdate,
          newLowestPrice: newLowest,
          percentageDrop: daling,
        });
      }
    } catch (err) {
      Logger.log('Error for row %s (%s): %s', rowIndex, url, err);
      appendUrlPriceHistory_(ss, {
        timestamp: timestamp,
        rowIndex: rowIndex,
        url: url,
        product: product,
        brand: brand,
        model: model,
        price: null,
        note: 'Error: ' + err,
      });
    }
  }

  // Apply batched updates to the sheet
  for (var j = 0; j < updates.length; j++) {
    var rowNumber = updateRowNumbers[j];
    var values = [updates[j]];
    sheet.getRange(rowNumber, 7, 1, values[0].length).setValues(values);
  }

  if (changes.length > 0) {
    notifyPriceDropsForUrls_(changes);
  }
}

/**
 * Parse a price in EUR from arbitrary HTML text.
 * Returns a Number or null if no price is detected.
 */
function parsePriceFromHtml_(html) {
  if (!html) {
    return null;
  }
  var cleaned = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/\s+/g, ' ');

  var pricePattern = /(?:€|&euro;|eur\b)\s*([\d]{1,3}(?:[\.\s']\d{3})*(?:,\d{2})?|\d+(?:,\d{2})?)/i;
  var match = pricePattern.exec(cleaned);
  if (!match) {
    return null;
  }

  var numeric = match[1]
    .replace(/[\.\s']/g, '')
    .replace(',', '.');
  var value = parseFloat(numeric);
  return isFinite(value) ? value : null;
}

/**
 * Send email notifications for products that hit new lows with significant drops.
 */
function notifyPriceDropsForUrls_(changes) {
  if (!changes || !changes.length) {
    return;
  }

  var userEmail = Session.getActiveUser().getEmail();
  var subject = 'Nieuwe laagste prijs voor gevolgd product (URL)';

  var lines = ['Er zijn nieuwe prijsdalingen gevonden:', ''];
  changes.forEach(function(change) {
    var dropPct = (change.percentageDrop * 100).toFixed(2) + '%';
    lines.push([
      'Product: ' + (change.product || 'Onbekend'),
      'Categorie: ' + (change.category || ''),
      'Merk/Model: ' + [change.brand, change.model].filter(Boolean).join(' '),
      'URL: ' + change.url,
      'Vorige prijs: ' + formatEuro_(change.previousPrice),
      'Nieuwe prijs: ' + formatEuro_(change.newPrice),
      'Laagste (voorheen): ' + formatEuro_(change.lowestPriceBeforeUpdate),
      'Nieuwe laagste: ' + formatEuro_(change.newLowestPrice),
      'Daling: ' + dropPct,
      '',
    ].join('\n'));
  });

  var body = lines.join('\n');
  MailApp.sendEmail(userEmail, subject, body);
}

function isWatchEnabled_(flag) {
  if (flag === true) {
    return true;
  }
  if (typeof flag === 'string') {
    var text = flag.trim().toLowerCase();
    return ['ja', 'yes', 'y', 'true'].indexOf(text) !== -1;
  }
  return false;
}

function toNumberOrNull_(value) {
  var num = typeof value === 'number' ? value : parseFloat(value);
  return isFinite(num) ? num : null;
}

function formatEuro_(value) {
  if (!isFinite(value)) {
    return '-';
  }
  return '€ ' + value.toFixed(2);
}

function ensureHistorySheet_(ss) {
  var sheet = ss.getSheetByName('Prijshistorie_URLs');
  if (!sheet) {
    sheet = ss.insertSheet('Prijshistorie_URLs');
    sheet.appendRow(['Timestamp', 'Row', 'URL', 'Product', 'Brand', 'Model', 'Price', 'Note']);
  }
  return sheet;
}

function appendUrlPriceHistory_(ss, entry) {
  var sheet = ensureHistorySheet_(ss);
  sheet.appendRow([
    entry.timestamp,
    entry.rowIndex,
    entry.url,
    entry.product,
    entry.brand,
    entry.model,
    entry.price,
    entry.note || '',
  ]);
}
