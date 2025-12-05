var CATEGORY_SHEET_NAME = 'Categorieën';
var HISTORY_SHEET_NAME = 'Prijshistorie_URLs';

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

  var timestamp = new Date();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 17);
  var rows = dataRange.getValues();
  var updates = [];
  var updateRowNumbers = [];
  var historyRows = [];
  var rowsData = [];
  var categoriesConfig = loadCategories_();

  for (var i = 0; i < rows.length; i++) {
    var rowIndex = i + 2; // 1-based row number
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
    var scrapedTitle = row[12];
    var scrapedBrand = row[13];
    var scrapedIds = row[14];

    if (!url || !isWatchEnabled_(watchFlag)) {
      continue;
    }

    var updatedRow = row.slice();
    var status = 'OK';
    var newPrice = null;
    var newTitle = '';
    var newBrand = '';
    var newIds = '';
    var lowestBefore = isFinite(lowestPrice) ? lowestPrice : null;
    var lowestAfter = lowestBefore;
    var daling = 0;

    var html = fetchHtml_(url);
    if (!html) {
      status = 'HTTP_ERROR';
      historyRows.push(buildHistoryRow_({
        timestamp: timestamp,
        category: category,
        product: product,
        productId: productId,
        url: url,
        previousPrice: null,
        newPrice: null,
        lowestAfter: null,
        daling: null,
        brand: brand,
        model: model,
        status: status,
      }));
      continue;
    }

    try {
      newPrice = parsePriceFromHtml_(html);
      newTitle = parseTitleFromHtml_(html) || scrapedTitle || product || '';
      newBrand = parseBrandFromHtml_(html) || scrapedBrand || brand || '';
      newIds = parseIdsFromHtml_(html);

      // Update scraped info
      if (newTitle) {
        updatedRow[12] = newTitle;
      }
      if (newBrand) {
        updatedRow[13] = newBrand;
      }
      updatedRow[14] = newIds || '';

      // Auto-fill product columns only if empty
      if (!product && newTitle) {
        updatedRow[2] = newTitle;
        product = newTitle;
      }
      if (!brand && newBrand) {
        updatedRow[4] = newBrand;
        brand = newBrand;
      }
      if (!productId && newIds) {
        var primaryId = extractPrimaryId_(newIds);
        if (primaryId) {
          updatedRow[3] = primaryId;
          productId = primaryId;
        }
      }

      // Auto-categorize if needed
      if (!category) {
        category = autoCategorizeRow_(categoriesConfig, {
          product: product,
          title: newTitle,
          url: url,
        });
        if (category) {
          updatedRow[1] = category;
        }
      }

      if (!isFinite(newPrice)) {
        status = 'PRICE_NOT_FOUND';
      } else {
        var oldLastPrice = lastPrice;
        updatedRow[7] = oldLastPrice; // H: previous price
        updatedRow[6] = newPrice; // G: last price

        if (lowestBefore === null) {
          lowestAfter = newPrice;
        } else {
          lowestAfter = Math.min(lowestBefore, newPrice);
        }
        updatedRow[8] = lowestAfter; // I: lowest price

        if (isFinite(previousPrice) && previousPrice > 0) {
          daling = (previousPrice - newPrice) / previousPrice;
        }
        updatedRow[9] = daling; // J: daling %

        updatedRow[11] = timestamp; // L: last check
        updatedRow[15] = newPrice; // P: scraped price
        updatedRow[16] = timestamp; // Q: scraped lastchecked

        rowsData.push({
          rowIndex: rowIndex,
          groupKey: buildGroupKey_(productId, brand, model, product || newTitle),
          category: category,
          product: product || newTitle,
          productId: productId,
          brand: brand,
          model: model,
          url: url,
          shopName: getShopNameFromUrl_(url),
          previousPrice: previousPrice,
          newPrice: newPrice,
          lowestBefore: lowestBefore,
          lowestAfter: lowestAfter,
          percentageDrop: daling,
        });
      }
    } catch (err) {
      Logger.log('Error for row %s (%s): %s', rowIndex, url, err);
      status = 'HTTP_ERROR';
    }

    // Persist row updates if anything changed
    if (!arraysEqual_(row, updatedRow)) {
      updates.push(updatedRow);
      updateRowNumbers.push(rowIndex);
    }

    // Build history row
    var historyPrevious = status === 'OK' && isFinite(previousPrice) ? previousPrice : null;
    var historyNew = status === 'OK' && isFinite(newPrice) ? newPrice : null;
    var historyLowest = status === 'OK' && isFinite(lowestAfter) ? lowestAfter : null;
    var historyDrop = status === 'OK' ? daling : null;

    historyRows.push(buildHistoryRow_({
      timestamp: timestamp,
      category: category,
      product: product,
      productId: productId,
      url: url,
      previousPrice: historyPrevious,
      newPrice: historyNew,
      lowestAfter: historyLowest,
      daling: historyDrop,
      brand: brand,
      model: model,
      status: status,
    }));
  }

  // Apply updates
  for (var u = 0; u < updates.length; u++) {
    sheet.getRange(updateRowNumbers[u], 1, 1, 17).setValues([updates[u]]);
  }

  appendHistoryRows_(ss, historyRows);

  var changes = determineGroupedChanges_(rowsData);
  if (changes.length > 0) {
    notifyPriceDropsForUrls_(changes);
  }

  applyFormatting_(sheet);
}

function determineGroupedChanges_(rowsData) {
  var changes = [];
  if (!rowsData.length) {
    return changes;
  }

  var groups = new Map();
  rowsData.forEach(function(r) {
    if (!groups.has(r.groupKey)) {
      groups.set(r.groupKey, []);
    }
    groups.get(r.groupKey).push(r);
  });

  groups.forEach(function(groupRows, groupKey) {
    var sorted = groupRows
      .filter(function(r) { return isFinite(r.newPrice); })
      .sort(function(a, b) { return a.newPrice - b.newPrice; });
    if (!sorted.length) {
      return;
    }

    var newCheapest = sorted[0];
    var qualifiesLowest = isFinite(newCheapest.lowestBefore) && newCheapest.newPrice < newCheapest.lowestBefore;
    var qualifiesDrop = newCheapest.percentageDrop >= 0.05;
    if (qualifiesLowest && qualifiesDrop) {
      changes.push({
        groupKey: groupKey,
        category: newCheapest.category,
        product: newCheapest.product,
        productId: newCheapest.productId,
        brand: newCheapest.brand,
        model: newCheapest.model,
        newCheapest: newCheapest,
        allShops: sorted,
      });
    }
  });

  return changes;
}

function buildHistoryRow_(entry) {
  return [
    entry.timestamp,
    entry.category,
    entry.product,
    entry.productId,
    entry.url,
    entry.previousPrice,
    entry.newPrice,
    entry.lowestAfter,
    entry.daling,
    entry.brand,
    entry.model,
    entry.status || 'OK',
  ];
}

function appendHistoryRows_(ss, rows) {
  if (!rows || !rows.length) {
    return;
  }
  var sheet = ensureSheet_(ss, HISTORY_SHEET_NAME, [
    'Tijdstip',
    'Categorie',
    'Product',
    'ProductId',
    'URL',
    'Oude prijs',
    'Nieuwe prijs',
    'Laagste prijs ooit',
    '% daling',
    'Merk',
    'Model',
    'Status',
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function parsePriceFromHtml_(html) {
  if (!html) {
    return null;
  }

  var cleaned = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/\s+/g, ' ');

  var patterns = [
    /"price"\s*:\s*"?([\d.,]+)"?/i,
    /<meta[^>]*itemprop=["']price["'][^>]*content=["']([\d.,]+)["'][^>]*>/i,
    /data-(?:price|product-price|price-amount)\s*=\s*"?([\d.,]+)"?/i,
    /(?:€|&euro;|eur\b)\s*([\d]{1,3}(?:[\.\s']\d{3})*(?:,\d{2})?|\d+(?:,\d{2})?)/i,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = patterns[i].exec(cleaned);
    if (match && match[1]) {
      var numeric = match[1].replace(/[\.\s']/g, '').replace(',', '.');
      var value = parseFloat(numeric);
      if (isFinite(value)) {
        return value;
      }
    }
  }

  return null;
}

function parseTitleFromHtml_(html) {
  if (!html) {
    return '';
  }
  var titleMatch = /<title>([^<]{3,})<\/title>/i.exec(html);
  if (titleMatch) {
    return titleMatch[1].trim();
  }
  var ogMatch = /<meta[^>]*property=["']og:title["'][^>]*content=["']([^"']+)["'][^>]*>/i.exec(html);
  return ogMatch ? ogMatch[1].trim() : '';
}

function parseBrandFromHtml_(html) {
  if (!html) {
    return '';
  }
  var ldMatch = /"brand"\s*:\s*"([^"}]+)"/i.exec(html);
  if (ldMatch) {
    return ldMatch[1].trim();
  }
  var metaMatch = /<meta[^>]*itemprop=["']brand["'][^>]*content=["']([^"']+)["'][^>]*>/i.exec(html);
  return metaMatch ? metaMatch[1].trim() : '';
}

function parseIdsFromHtml_(html) {
  if (!html) {
    return '';
  }
  var ids = [];
  var sku = /"sku"\s*:\s*"([^"}]+)"/i.exec(html);
  if (sku && sku[1]) {
    ids.push('sku:' + sku[1].trim());
  }
  var gtin = /"gtin(8|12|13|14)"\s*:\s*"?([^"}]+)"?/i.exec(html);
  if (gtin && gtin[2]) {
    ids.push('gtin:' + gtin[2].trim());
  }
  var mpn = /"mpn"\s*:\s*"([^"}]+)"/i.exec(html);
  if (mpn && mpn[1]) {
    ids.push('mpn:' + mpn[1].trim());
  }
  return ids.join(';');
}

function extractPrimaryId_(idsString) {
  if (!idsString) {
    return '';
  }
  var parts = idsString.split(';');
  for (var i = 0; i < parts.length; i++) {
    var pair = parts[i].split(':');
    if (pair.length === 2 && pair[1]) {
      return pair[1];
    }
  }
  return '';
}

function loadCategories_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
  if (!sheet) {
    return [];
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var configs = [];
  rows.forEach(function(r) {
    var name = r[0];
    var pattern = r[2];
    var priority = toNumberOrNull_(r[3]);
    var active = r[4] === true;
    if (!active || !name || !pattern) {
      return;
    }
    var cleanPattern = String(pattern).replace(/^\(\?i\)/i, '');
    var regex;
    try {
      regex = new RegExp(cleanPattern, 'i');
    } catch (e) {
      Logger.log('Invalid regex in categories: %s', pattern);
      return;
    }
    configs.push({
      name: name,
      regex: regex,
      priority: isFinite(priority) ? priority : 9999,
      active: active,
    });
  });

  configs.sort(function(a, b) { return a.priority - b.priority; });
  return configs;
}

function autoCategorizeRow_(categoriesConfig, rowData) {
  if (!categoriesConfig || !categoriesConfig.length) {
    return '';
  }
  var textToMatch = [(rowData.product || ''), (rowData.title || ''), (rowData.url || '')]
    .join(' ')
    .toLowerCase();
  for (var i = 0; i < categoriesConfig.length; i++) {
    var cfg = categoriesConfig[i];
    try {
      if (cfg.regex.test(textToMatch)) {
        return cfg.name;
      }
    } catch (e) {
      Logger.log('Error testing category regex: %s', e);
    }
  }
  return '';
}

function getShopNameFromUrl_(url) {
  try {
    var host = new URL(url).hostname;
    return host.replace(/^www\./i, '');
  } catch (e) {
    return url;
  }
}

function buildGroupKey_(productId, brand, model, product) {
  if (productId) {
    return 'ID:' + String(productId).trim();
  }
  return ('MMN:' + (brand || '') + '|' + (model || '') + '|' + (product || ''))
    .toLowerCase();
}

function notifyPriceDropsForUrls_(changes) {
  if (!changes || !changes.length) {
    return;
  }

  var userEmail = Session.getActiveUser().getEmail();
  var subject = 'Nieuwe laagste prijs voor gevolgd product (URL)';
  var lines = ['Er zijn nieuwe prijsdalingen gevonden:', ''];

  changes.forEach(function(change) {
    var cheapest = change.newCheapest;
    lines.push('Product: ' + (change.product || 'Onbekend'));
    if (change.category) {
      lines.push('Categorie: ' + change.category);
    }
    lines.push('Merk/Model: ' + [change.brand, change.model].filter(Boolean).join(' '));
    lines.push('Nieuwe laagste prijs bij ' + cheapest.shopName + ': ' + formatEuro_(cheapest.newPrice));
    lines.push('Vorige prijs: ' + formatEuro_(cheapest.previousPrice));
    lines.push('Daling: ' + (cheapest.percentageDrop * 100).toFixed(2) + '%');
    lines.push('URL: ' + cheapest.url);
    lines.push('Alle prijzen:');
    change.allShops.forEach(function(shop) {
      lines.push(' - ' + shop.shopName + ': ' + formatEuro_(shop.newPrice));
    });
    lines.push('');
  });

  var body = lines.join('\n');
  MailApp.sendEmail(userEmail, subject, body);
}

function applyFormatting_(sheet) {
  try {
    sheet.getRange('G:I').setNumberFormat('€ #,##0.00');
    sheet.getRange('P:P').setNumberFormat('€ #,##0.00');
    sheet.getRange('J:J').setNumberFormat('0,0%');
    sheet.getRange('L:L').setNumberFormat('dd-mm-yyyy hh:mm');
    sheet.getRange('Q:Q').setNumberFormat('dd-mm-yyyy hh:mm');
  } catch (e) {
    Logger.log('Formatting error: %s', e);
  }
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (headers && headers.length) {
    var existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var needsHeader = existing.join('') === '';
    if (needsHeader) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
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
  var num = toNumberOrNull_(value);
  if (!isFinite(num)) {
    return '-';
  }
  return '€ ' + num.toFixed(2);
}

function arraysEqual_(a, b) {
  if (a === b) return true;
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

function fetchHtml_(url) {
  var options = {
    followRedirects: true,
    muteHttpExceptions: true,
    method: 'get',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7',
    },
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();

  if (code !== 200) {
    Logger.log('HTTP error ' + code + ' for ' + url);
    return null;
  }
  return resp.getContentText();
}
