var CATEGORY_SHEET_NAME = 'Categorieën';
var HISTORY_SHEET_NAME = 'Prijshistorie_URLs';

/**
 * Main entry point to check all tracked product URLs on sheet "Blad1".
 */
function checkTrackedProductUrls() {
  processTrackedProductUrls_();
}

/**
 * Processes only the currently selected data rows on "Blad1".
 */
function checkSelectedTrackedProductUrls() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Blad1');
  if (!sheet) {
    Logger.log('Sheet "Blad1" not found.');
    return;
  }

  var range = ss.getActiveRange();
  if (!range || range.getSheet().getName() !== sheet.getName()) {
    Logger.log('Selecteer eerst rijen op Blad1.');
    return;
  }

  var rowSelectionMap = {};
  var startRow = range.getRow();
  var endRow = startRow + range.getNumRows() - 1;
  for (var r = startRow; r <= endRow; r++) {
    if (r >= 2) { // skip header row
      rowSelectionMap[r] = true;
    }
  }

  if (Object.keys(rowSelectionMap).length === 0) {
    Logger.log('Geen datarijen geselecteerd.');
    return;
  }

  processTrackedProductUrls_(rowSelectionMap);
}

/**
 * Adds a custom menu to launch sync actions.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Prijswatcher')
    .addItem('Sync alle URLs', 'checkTrackedProductUrls')
    .addItem('Sync geselecteerde rijen', 'checkSelectedTrackedProductUrls')
    .addSeparator()
    .addItem('Pas formattering toe', 'applyFormattingMenu_')
    .addToUi();
}

/**
 * Processes all rows or a specific selection of rows on "Blad1".
 * @param {Object<string, boolean>} [rowSelectionMap] Optional map of row indexes (2-based) to limit processing.
 */
function processTrackedProductUrls_(rowSelectionMap) {
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
    if (rowSelectionMap && !rowSelectionMap[rowIndex]) {
      continue;
    }
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

/**
 * Probeert een prijs uit de HTML te halen en normaliseert die via parsePrice_.
 * Geeft een Number terug (bijv. 399.95) of null als er niets gevonden wordt.
 */
function parsePriceFromHtml_(html) {
  if (!html) return null;

  // 1) PROBEER JSON-LD / inline JSON MET "price"
  try {
    // Pak alle <script type="application/ld+json"> blokken
    const scriptRegex = /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
    let match;
    while ((match = scriptRegex.exec(html)) !== null) {
      const jsonText = match[1].trim();
      try {
        const data = JSON.parse(jsonText);

        // data kan object of array zijn → normaliseer naar array
        const items = Array.isArray(data) ? data : [data];

        for (let item of items) {
          const priceCandidate = extractPriceFromObject_(item);
          if (priceCandidate != null) {
            const n = parsePrice_(priceCandidate);
            if (typeof n === 'number' && !isNaN(n)) {
              return n;
            }
          }
        }
      } catch (e) {
        // als 1 JSON-blok faalt, ga gewoon naar de volgende
      }
    }
  } catch (e) {
    // negeren, we gaan naar volgende strategie
  }

  // 2) META TAGS MET itemprop="price"
  try {
    const metaPriceRegex =
      /<meta[^>]+itemprop=["']price["'][^>]*content=["']([^"']+)["'][^>]*>/i;
    const metaMatch = metaPriceRegex.exec(html);
    if (metaMatch && metaMatch[1]) {
      const n = parsePrice_(metaMatch[1]);
      if (typeof n === 'number' && !isNaN(n)) {
        return n;
      }
    }
  } catch (e) {
    // laat maar, volgende stap
  }

  // 3) DATA-ATTRIBUTES (data-price, data-product-price etc.)
  try {
    const dataAttrRegex =
      /data-(?:price|product-price|price-amount)=["']([^"']+)["']/gi;
    let m;
    while ((m = dataAttrRegex.exec(html)) !== null) {
      const n = parsePrice_(m[1]);
      if (typeof n === 'number' && !isNaN(n)) {
        return n;
      }
    }
  } catch (e) {
    // volgende stap
  }

  // 4) FALLBACK: zoek laatste "€ 1.234,56" achtige notatie in de HTML
  try {
    const euroRegex = /(?:€|\&euro;)\s*([\d\.\,]+)/gi;
    let m;
    let lastValue = null;
    while ((m = euroRegex.exec(html)) !== null) {
      lastValue = m[1];
    }
    if (lastValue != null) {
      const n = parsePrice_(lastValue);
      if (typeof n === 'number' && !isNaN(n)) {
        return n;
      }
    }
  } catch (e) {
    // niets
  }

  // Als alle strategieën falen:
  return null;
}

/**
 * Helper voor parsePriceFromHtml_:
 * loopt een JSON-object af en zoekt naar keys als "price", "offers.price", etc.
 */
function extractPriceFromObject_(obj) {
  if (!obj || typeof obj !== 'object') return null;

  // expliciete velden
  if (obj.price != null) return obj.price;
  if (obj.offers && obj.offers.price != null) return obj.offers.price;
  if (Array.isArray(obj.offers)) {
    for (let offer of obj.offers) {
      if (offer && offer.price != null) return offer.price;
    }
  }

  // generieke deep search op key "price"
  for (let key in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
    const val = obj[key];

    if (key.toLowerCase() === 'price' && val != null) {
      return val;
    }

    if (typeof val === 'object') {
      const nested = extractPriceFromObject_(val);
      if (nested != null) return nested;
    }
  }
  return null;
}

function parsePrice_(v) {
  if (v == null) return null;

  // 1) Als het al een Number is (vaak uit JSON-LD)
  if (typeof v === 'number') {
    let n = v;

    // Heuristiek: sommige shops geven prijzen in centen terug (bijv. 39995 voor € 399,95)
    // Voor babyproducten is een prijs van > 5.000 euro extreem onwaarschijnlijk.
    // Als het een groot, heel getal is, behandelen we het als centen.
    if (Number.isInteger(n) && n >= 5000) {
      n = n / 100; // 39995 -> 399.95
    }
    return n;
  }

  // 2) Als het een string is
  let s = String(v).trim();

  // Case: pure digits zonder komma/punt, en minstens 4 cijfers (bijv. "39995")
  // Grote kans dat dit centen zijn.
  if (/^\d{4,}$/.test(s)) {
    const raw = parseInt(s, 10);
    if (!isNaN(raw)) {
      if (raw >= 5000) {
        return raw / 100; // 39995 -> 399.95
      }
      return raw; // kleinere getallen (bijv. "999") laten we staan als euro's
    }
  }

  // 3) Normale NL/EN prijsstrings ("€ 399,95", "399.95", etc.)
  s = s
    .replace(/[^\d,\.]/g, '')  // alleen cijfers, komma en punt overhouden
    .replace(/\./g, '')        // alle punten weg (duizend-separators)
    .replace(',', '.');        // komma als decimaal

  const n = parseFloat(s);
  return isNaN(n) ? null : n;
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

function applyFormattingMenu_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Blad1');
  if (!sheet) {
    Logger.log('Sheet "Blad1" not found.');
    return;
  }
  applyFormatting_(sheet);
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
