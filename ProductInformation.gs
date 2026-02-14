function stepA_populateMainFromUrl() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MAIN);
  if (!sh) throw new Error('Tabblad ' + SHEET_MAIN + ' ontbreekt.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    dbg('[A] geen data in', SHEET_MAIN);
    return;
  }

  // Lees bestaande headers
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let map = headerMap_(headers);

  dbg('[A] START headers=', JSON.stringify(headers));
  dbg('[A] START headerMap=', JSON.stringify(map));

  // Zorg dat scraped-kolommen bestaan
  const neededKeys = [
    { key: 'scraped titel',   label: 'Scraped titel' },
    { key: 'scraped brand',   label: 'Scraped brand' },
    { key: 'scraped ids',     label: 'Scraped ids' },
    { key: 'scraped price',   label: 'Scraped price' },
    { key: 'scraped lastchecked', label: 'Scraped lastchecked' }
  ];

  let changed = false;
  neededKeys.forEach(obj => {
    if (map[obj.key] == null) {
      headers.push(obj.label);
      changed = true;
    }
  });

  if (changed) {
    dbg('[A] uitbreiden headers met scraped kolommen →', JSON.stringify(headers));
    sh.getRange(1, 1, 1, headers.length).clearContent();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    map = headerMap_(headers); // opnieuw opbouwen
    dbg('[A] nieuwe headerMap=', JSON.stringify(map));
  }

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const idxUrl   = map['url'];
  const idxStit  = map['scraped titel'];
  const idxSbr   = map['scraped brand'];
  const idxSids  = map['scraped ids'];
  const idxSpr   = map['scraped price'];
  const idxSlast = map['scraped lastchecked'];

  if (idxUrl == null) {
    throw new Error('[A] Kolom "URL" niet gevonden in Blad1');
  }

  dbg('[A] indexen:',
      'idxUrl=', idxUrl,
      'idxStit=', idxStit,
      'idxSbr=', idxSbr,
      'idxSids=', idxSids,
      'idxSpr=', idxSpr,
      'idxSlast=', idxSlast);

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const row = data[i];

    const url = String(row[idxUrl] || '').trim();
    if (!url) {
      dbg('[A]', rowNum, 'SKIP (geen URL)');
      continue;
    }

    dbg('---------------------------');
    dbg('[A-ROW]', rowNum, 'URL=', url);

    const html = safeFetch_(url, 'A-main-' + rowNum);
    if (!html) {
      dbg('[A]', rowNum, 'geen HTML of HTTP != 200');
      continue;
    }

    dbg('[A]', rowNum, 'HTML length=', html.length);

    const prod = parseProductPage_(html);
    if (!prod) {
      dbg('[A]', rowNum, 'parseProductPage_ → null');
      continue;
    }

    dbg('[A]', rowNum, 'PARSED:',
        'name=', prod.name,
        'brand=', prod.brand,
        'ids=', JSON.stringify(prod.ids),
        'price=', prod.price);

    if (idxStit  != null) row[idxStit]  = prod.name  || '';
    if (idxSbr   != null) row[idxSbr]   = prod.brand || '';
    if (idxSids  != null) row[idxSids]  = (prod.ids || []).join(', ');
    if (idxSpr   != null) row[idxSpr]   = prod.price != null ? prod.price : '';
    if (idxSlast != null) row[idxSlast] = new Date();

    // Schrijf de hele rij terug met alle kolommen (incl. nieuwe scraped-data)
    sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
    dbg('[A]', rowNum, 'ROW UPDATED');
  }

  dbg('[A] klaar');
}

/************************************
 * GLOBALE CONFIG
 ************************************/

const DEBUG_ENABLED = true;

function dbg() {
  if (!DEBUG_ENABLED) return;
  const args = Array.prototype.slice.call(arguments);
  Logger.log(args.join(' '));
}

const SHEET_MAIN       = 'Blad1';
const SHEET_SEARCHLIST = 'Searchlist';
const SHEET_COMPETITOR = 'Concurrenten-2';

// Als je SerpAPI straks (weer) wilt gebruiken:
const SERPAPI_KEY = PropertiesService.getScriptProperties().getProperty('SERPAPI_API_KEY');


/************************************
 * SHEET HELPERS
 ************************************/

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    dbg('[SHEET] create', name);
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    dbg('[SHEET] init headers for', name);
    sh.appendRow(headers);
  } else if (headers && headers.length) {
    const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const mismatch = headers.some((h, i) => firstRow[i] !== h);
    if (mismatch) {
      dbg('[SHEET] reset headers for', name);
      sh.clearContents();
      sh.appendRow(headers);
    }
  }
  return sh;
}

function extractProductLinksFromSearch_(html, baseUrl, shopLabel, maxCandidates) {
  maxCandidates = maxCandidates || 10;

  // basis-URL schoonmaken (zonder trailing slash)
  baseUrl = (baseUrl || '').replace(/\/+$/, '');

  const results = [];
  const seen = {};

  // alle hrefs uit de HTML trekken
  const re = /<a\b[^>]*href=["']([^"']+)["'][^>]*>/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    let href = m[1];

    // skip lege / ankers
    if (!href || href === '#' || href.indexOf('javascript:') === 0) continue;

    // naar absolute URL omzetten indien relatief
    if (href.indexOf('http://') !== 0 && href.indexOf('https://') !== 0) {
      href = baseUrl + '/' + href.replace(/^\/+/, '');
    }

    // alleen URLs binnen dezelfde site houden
    if (baseUrl && href.indexOf(baseUrl) !== 0) continue;

    // ruwe path (zonder query/hash) pakken voor filtering
    const urlObj = href.split('#')[0];
    const pathPart = urlObj.split('://')[1].split('/').slice(1).join('/');

    const lower = pathPart.toLowerCase();

    // heel generieke / irrelevante pagina's skippen
    if (
      !lower || 
      lower === '' ||
      lower === '/' ||
      lower === 'customer/account' ||
      lower.indexOf('customer/account') !== -1 ||
      lower.indexOf('login') !== -1 ||
      lower.indexOf('register') !== -1 ||
      lower.indexOf('wishlist') !== -1 ||
      lower.indexOf('cart') !== -1 ||
      lower.indexOf('checkout') !== -1 ||
      lower.indexOf('account') !== -1 ||
      lower.indexOf('klant') !== -1
    ) {
      continue;
    }

    // beetje product-achtig pad: minstens 1 segment met letters + cijfers
    const segments = lower.split('/');
    const hasProductishSegment = segments.some(seg => /[a-z]/.test(seg) && /\d/.test(seg));
    if (!hasProductishSegment) continue;

    // duplicates verwijderen
    if (seen[href]) continue;
    seen[href] = true;
    results.push(href);

    if (results.length >= maxCandidates) break;
  }

  dbg(
    '[B-SEARCH-EXTRACT]',
    shopLabel || '',
    'baseUrl=', baseUrl,
    '→', results.length, 'product-link(s) gevonden, voorbeeld=',
    results.slice(0, 3).join(', ')
  );

  return results;
}

function headerMap_(headers) {
  const map = {};
  headers.forEach(function(h, i) {
    const key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

/************************************
 * FETCH & PARSE HELPERS
 ************************************/

function safeFetch_(url, label) {
  dbg('[FETCH]', label || '', '→', url);
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const code = resp.getResponseCode();
    dbg('[FETCH-RESP]', label || '', 'HTTP', code);
    if (code !== 200) return null;
    return resp.getContentText();
  } catch (e) {
    dbg('[FETCH-ERR]', label || '', e);
    return null;
  }
}

function parseProductPage_(html) {
  if (!html) return null;

  let name = '';
  let price = null;
  let brand = '';
  let ids = [];

  // 1) JSON-LD blokken zoeken
  let ldMatches = [];
  try {
    ldMatches = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi) || [];
  } catch (e) {
    dbg('[PARSE] ld+json regex error', e);
    ldMatches = [];
  }

  dbg('[PARSE] ld+json blocks=', ldMatches.length);

  try {
    for (let i = 0; i < ldMatches.length; i++) {
      // verwijder script tags
      const inner = ldMatches[i].replace(/^<script[^>]*>/i, '').replace(/<\/script>$/i, '');
      // sommige sites HTML-escapen JSON, probeer dat te fixen
      const txt = inner.replace(/&quot;/g, '"').replace(/&apos;/g, "'").trim();

      let obj;
      try {
        obj = JSON.parse(txt);
      } catch (e) {
        dbg('[PARSE] JSON parse error blok', i+1, e);
        continue;
      }

      const prod = unwrapJsonLdProduct_(obj);
      if (prod) {
        dbg('[PARSE] JSON-LD product gevonden in blok', i+1,
            'name=', prod.name,
            'brand=', prod.brand,
            'price=', prod.price,
            'ids=', JSON.stringify(prod.ids));
        if (prod.name) name = prod.name;
        if (prod.price != null) price = prod.price;
        if (prod.brand) brand = prod.brand;
        if (prod.ids && prod.ids.length) ids = mergeUnique_(ids, prod.ids);
        break; // eerste product is vaak goed
      }
    }
  } catch (e) {
    dbg('[PARSE] JSON-LD loop error', e);
  }

  // 2) Fallback: OG title / H1 / title
  if (!name) {
    const ogTitle = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i);
    if (ogTitle) {
      name = ogTitle[1].trim();
      dbg('[PARSE] fallback og:title →', name);
    }
  }

  if (!name) {
    const h1 = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    if (h1) {
      name = h1[1].replace(/<[^>]*>/g, '').trim();
      dbg('[PARSE] fallback h1 →', name);
    }
  }

  if (!name) {
    const t = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (t) {
      name = t[1].replace(/<[^>]*>/g, '').trim();
      dbg('[PARSE] fallback <title> →', name);
    }
  }

 // 3) Extra numeric IDs uit HTML
const extraIds = extractNumericIdsFromHtml_(html);
if (extraIds.length) {
  dbg('[PARSE] extra numeric ids uit HTML (max 20 getoond):',
      JSON.stringify(extraIds.slice(0, 20)));
  ids = mergeUnique_(ids, extraIds);
}

dbg('[PARSE-RESULT]',
    'name=',  name,
    'brand=', brand,
    'price=', price,
    'ids=', JSON.stringify(ids.slice(0, 20)), // alleen eerste 20 in log
);

  // 4) Fallback prijs als JSON-LD geen prijs heeft
  if (price == null) {
    // zoek typische euro-formatten
    const m = html.match(/€\s*([\d\.\,]+)/);
    if (m && m[1]) {
      const raw = m[1].replace(/\./g, '').replace(',', '.');
      const n = parseFloat(raw);
      if (!isNaN(n)) {
        price = n;
        dbg('[PARSE] fallback prijs uit HTML →', price);
      }
    }
  }

  dbg('[PARSE-RESULT]',
      'name=', name,
      'brand=', brand,
      'price=', price,
      'ids=', JSON.stringify(ids));

  if (!name && price == null && !ids.length) {
    // echt niets bruikbaars
    return null;
  }

  return { name: name, price: price, brand: brand, ids: ids };
}

function unwrapJsonLdProduct_(obj) {
  if (!obj) return null;

  // als array → itereren
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      const r = unwrapJsonLdProduct_(obj[i]);
      if (r) return r;
    }
    return null;
  }

  // @graph → itereren
  if (obj['@graph']) {
    return unwrapJsonLdProduct_(obj['@graph']);
  }

  // type bepalen
  const type = obj['@type'];
  const isProductType =
    type === 'Product' ||
    type === 'ProductModel' ||
    (Array.isArray(type) && (type.indexOf('Product') !== -1 || type.indexOf('ProductModel') !== -1));

  if (!isProductType) return null;

  const ids = [];
  ['gtin', 'gtin8', 'gtin12', 'gtin13', 'gtin14', 'sku', 'productId'].forEach(k => {
    const v = obj[k];
    if (v && /^\d{8,14}$/.test(String(v))) ids.push(String(v));
  });

  let price = null;
  if (obj.offers) {
    if (Array.isArray(obj.offers)) {
      const o = obj.offers[0];
      if (o && o.price != null) price = Number(o.price);
      else if (o && o.priceSpecification && o.priceSpecification.price != null) {
        price = Number(o.priceSpecification.price);
      }
    } else {
      if (obj.offers.price != null) {
        price = Number(obj.offers.price);
      } else if (obj.offers.priceSpecification && obj.offers.priceSpecification.price != null) {
        price = Number(obj.offers.priceSpecification.price);
      }
    }
  }

  let brand = '';
  if (obj.brand) {
    if (typeof obj.brand === 'string') brand = obj.brand;
    else if (obj.brand.name) brand = obj.brand.name;
  }

  return {
    name: obj.name || '',
    price: price,
    brand: brand,
    ids: ids
  };
}

function extractNumericIdsFromHtml_(html) {
  const out = [];
  const re = /\b(\d{8,14})\b/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    out.push(m[1]);
    if (out.length >= 200) break; // hard cap om gekke explosies te voorkomen
  }
  return out;
}

function mergeUnique_(arr1, arr2) {
  const seen = {};
  const out = [];
  (arr1 || []).concat(arr2 || []).forEach(v => {
    const s = String(v || '').trim();
    if (!s) return;
    if (seen[s]) return;
    seen[s] = true;
    out.push(s);
  });
  return out;
}

// Naam-similarity voor later (B/C)
function normalizeName_(s) {
  if (!s) return '';
  return s.toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function nameSimilarity_(a, b) {
  const na = normalizeName_(a);
  const nb = normalizeName_(b);
  if (!na || !nb) return 0;
  const ta = na.split(/\s+/).filter(t => t.length > 2);
  const tb = nb.split(/\s+/).filter(t => t.length > 2);
  if (!ta.length || !tb.length) return 0;
  const setB = {};
  tb.forEach(t => { setB[t] = true; });
  let overlap = 0;
  ta.forEach(t => { if (setB[t]) overlap++; });
  const denom = Math.max(ta.length, tb.length);
  return denom ? overlap / denom : 0;
}
function debugOneRowA(rowNumber) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MAIN);
  if (!sh) throw new Error('Tabblad ' + SHEET_MAIN + ' ontbreekt.');
  const lastRow = sh.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) {
    throw new Error('RowNumber buiten bereik');
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);
  dbg('[A-DEBUG-HEADERS]', JSON.stringify(headers));
  dbg('[A-DEBUG-MAP]', JSON.stringify(map));

  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxUrl = map['url'];
  const url = String(row[idxUrl] || '').trim();
  dbg('[A-DEBUG] row=', rowNumber, 'URL=', url);

  const html = safeFetch_(url, 'A-debug-' + rowNumber);
  if (!html) {
    dbg('[A-DEBUG] geen HTML');
    return;
  }
  dbg('[A-DEBUG] HTML length=', html.length);

  const prod = parseProductPage_(html);
  dbg('[A-DEBUG] parse result=', JSON.stringify(prod));
}

/***************************************************
 * COMPONENT B – Searchlist vullen op basis van
 *  - Scraped titel (naam)
 *  - fallback op IDs
 *
 * Resultaat in SHEET_SEARCHLIST
 ***************************************************/

const SHOP_CONFIGS = [
  {
    code: 'Babydump',
    label: 'Baby-Dump',
    domain: 'baby-dump.nl',
    baseUrl: 'https://www.baby-dump.nl',
    buildSearchUrl: q => 'https://www.baby-dump.nl/search/?q=' + encodeURIComponent(q)
  },
  {
    code: 'Prenatal',
    label: 'Prénatal',
    domain: 'prenatal.nl',
    baseUrl: 'https://www.prenatal.nl',
    buildSearchUrl: q => 'https://www.prenatal.nl/catalogsearch/result/?q=' + encodeURIComponent(q)
  },
  {
    code: 'BabyTiener',
    label: 'Baby & Tiener',
    domain: 'babyentiener.nl',
    baseUrl: 'https://www.babyentiener.nl',
    buildSearchUrl: q => 'https://www.babyentiener.nl/searchresults?q=' + encodeURIComponent(q)
  },
  {
    code: 'BabyPlanet',
    label: 'BabyPlanet',
    domain: 'babyplanet.nl',
    baseUrl: 'https://www.babyplanet.nl',
    buildSearchUrl: q => 'https://www.babyplanet.nl/catalogsearch/result/?q=' + encodeURIComponent(q)
  },
  {
    code: 'MamaLoes',
    label: 'MamaLoes',
    domain: 'mamaloes.nl',
    baseUrl: 'https://www.mamaloes.nl',
    buildSearchUrl: q => 'https://www.mamaloes.nl/search?query=' + encodeURIComponent(q)
  },
  {
    code: 'VanAsten',
    label: 'Van Asten Babysuperstore',
    domain: 'vanastenbabysuperstore.nl',
    baseUrl: 'https://www.vanastenbabysuperstore.nl',
    buildSearchUrl: q => 'https://www.vanastenbabysuperstore.nl/search?q=' + encodeURIComponent(q)
  }
];

// Limieten & drempels
const B_MAX_CANDIDATES_PER_SEARCH = 5;
const B_SCORE_THRESHOLD            = 0.45; // minimaal benodigde matchscore

// Headers voor Searchlist
const SEARCHLIST_HEADERS = [
  'MainRow',
  'MainURL',
  'MainProduct',
  'MainBrand',
  'MainIds',
  'Shop',
  'ResultURL',
  'ResultTitle',
  'ResultBrand',
  'ResultIds',
  'ResultPrice',
  'NameSim',
  'IdOverlapCount',
  'MatchScore',
  'SearchType',       // 'name' of 'id'
  'SearchQueryUsed',
  'LastChecked'
];

/**
 * Hoofdfunctie voor B
 * - leest Blad1
 * - zoekt per rij en per shop naar beste match
 * - schrijft alles naar Searchlist
 */
function stepB_buildSearchlist() {
  const ss = SpreadsheetApp.getActive();
  const mainSh = ss.getSheetByName(SHEET_MAIN);
  if (!mainSh) throw new Error('Tabblad ' + SHEET_MAIN + ' ontbreekt.');

  // Searchlist tab initialiseren
  const searchSh = ensureSheet_(ss, SHEET_SEARCHLIST, SEARCHLIST_HEADERS);
  // Voor nu: volledig opnieuw opbouwen
  searchSh.clearContents();
  searchSh.appendRow(SEARCHLIST_HEADERS);

  const lastRow = mainSh.getLastRow();
  if (lastRow < 2) {
    dbg('[B] geen data in', SHEET_MAIN);
    return;
  }

  const headers = mainSh.getRange(1, 1, 1, mainSh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);

  const idxUrl         = map['url'];
  const idxProduct     = map['product'];
  const idxScrTitle    = map['scraped titel'];
  const idxScrBrand    = map['scraped brand'];
  const idxScrIds      = map['scraped ids'];
  const idxScrLast     = map['scraped lastchecked'];

  if (idxUrl == null) throw new Error('[B] Kolom "URL" niet gevonden in ' + SHEET_MAIN);

  dbg('[B] START – lastRow=', lastRow, 'headers=', JSON.stringify(map));

  const data = mainSh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const outRows = [];

  for (let r = 0; r < data.length; r++) {
    const rowNum = r + 2;
    const row = data[r];

    const url = String(row[idxUrl] || '').trim();
    if (!url) {
      dbg('[B]', rowNum, 'SKIP (geen URL)');
      continue;
    }

    // Bepaal mainName en brand
    const mainName =
      (idxScrTitle != null && row[idxScrTitle]) ?
        String(row[idxScrTitle]).trim() :
        (idxProduct != null ? String(row[idxProduct] || '').trim() : '');

    if (!mainName) {
      dbg('[B]', rowNum, 'SKIP (geen mainName)');
      continue;
    }

    const mainBrand =
      (idxScrBrand != null && row[idxScrBrand]) ?
        String(row[idxScrBrand]).trim() :
        '';

    const mainIdsStr = (idxScrIds != null && row[idxScrIds]) ? String(row[idxScrIds]) : '';
    const mainIdsArr = parseIdsString_(mainIdsStr);

    dbg('---------------------------');
    dbg('[B-ROW]', rowNum, 'URL=', url, 'mainName=', mainName, 'ids=', JSON.stringify(mainIdsArr));

    // Per shop: beste match zoeken
    SHOP_CONFIGS.forEach(shop => {
      const cand = findBestCompetitorForProduct_(
        mainName,
        mainBrand,
        mainIdsArr,
        shop
      );
      if (!cand) {
        dbg('[B]', rowNum, shop.code, '→ geen geschikte match');
        return;
      }

      dbg('[B]', rowNum, shop.code, 'BEST:',
          'score=', cand.score,
          'nameSim=', cand.nameSim,
          'idOverlap=', cand.idOverlap,
          'url=', cand.url);

      outRows.push([
        rowNum,          // MainRow
        url,             // MainURL
        mainName,        // MainProduct
        mainBrand,       // MainBrand
        mainIdsStr,      // MainIds
        shop.label,      // Shop
        cand.url,        // ResultURL
        cand.title || '',// ResultTitle
        cand.brand || '',// ResultBrand
        (cand.ids || []).join(', '), // ResultIds
        cand.price != null ? cand.price : '', // ResultPrice
        cand.nameSim,    // NameSim
        cand.idOverlap,  // IdOverlapCount
        cand.score,      // MatchScore
        cand.searchType, // SearchType
        cand.query,      // SearchQueryUsed
        new Date()       // LastChecked
      ]);
    });
  }

  if (outRows.length) {
    searchSh.getRange(2, 1, outRows.length, SEARCHLIST_HEADERS.length).setValues(outRows);
  }

  dbg('[B] klaar –', outRows.length, 'regels in', SHEET_SEARCHLIST);
}

/**
 * Berekent beste competitor voor 1 product en 1 shop.
 * - eerst search op naam
 * - fallback op ID(s) als nodig
 */
function findBestCompetitorForProduct_(mainName, mainBrand, mainIdsArr, shop) {
  let best = null;

  // 1) Search op naam
  const qName = mainName;
  dbg('[B-SEARCH-NAME]', shop.code, 'q=', qName);
  const candName = searchAndScoreForQuery_(qName, 'name', mainName, mainBrand, mainIdsArr, shop);
  if (candName && candName.score >= B_SCORE_THRESHOLD) {
    best = candName;
  }

  // 2) Fallback: search op ID als:
  // - geen best, OF score onder threshold
  if ((!best || best.score < B_SCORE_THRESHOLD) && mainIdsArr && mainIdsArr.length) {
    const idCandidates = pickBestIdsForSearch_(mainIdsArr);
    for (let i = 0; i < idCandidates.length; i++) {
      const qId = idCandidates[i];
      dbg('[B-SEARCH-ID]', shop.code, 'id=', qId);
      const candId = searchAndScoreForQuery_(qId, 'id', mainName, mainBrand, mainIdsArr, shop);
      if (!candId) continue;
      if (!best || candId.score > best.score) {
        best = candId;
      }
    }
  }

  // Nog steeds niets bruikbaars?
  if (!best || best.score < B_SCORE_THRESHOLD) return null;
  return best;
}

/**
 * Voor een bepaalde query (naam of id) de searchpagina ophalen,
 * kandidaat-URL's extraheren, per productpagina scoren en beste teruggeven.
 */
function searchAndScoreForQuery_(query, searchType, mainName, mainBrand, mainIdsArr, shop) {
  const searchUrl = shop.buildSearchUrl(query);
  const html = safeFetch_(searchUrl, 'B-search-' + shop.code + '-' + searchType);
  if (!html) {
    dbg('[B-SEARCH]', shop.code, searchType, 'geen HTML voor', searchUrl);
    return null;
  }

  const productLinks = extractProductLinksFromSearch_(
    html,
    shop.baseUrl,
    shop.label,
    B_MAX_CANDIDATES_PER_SEARCH
  );
  dbg('[B-SEARCH]', shop.code, searchType, 'links gevonden=', productLinks.length,
      'sample=', JSON.stringify(productLinks.slice(0, 3)));

  if (!productLinks.length) return null;

  let best = null;

  for (let i = 0; i < productLinks.length; i++) {
    const prodUrl = productLinks[i];
    const prodHtml = safeFetch_(prodUrl, 'B-prod-' + shop.code + '-' + (i+1));
    if (!prodHtml) continue;

    const p = parseProductPage_(prodHtml);
    if (!p) {
      dbg('[B-PARSE]', shop.code, 'geen productdata voor', prodUrl);
      continue;
    }

    const candName  = p.name || '';
    const candBrand = p.brand || '';
    const candPrice = p.price;
    const candIds   = p.ids || [];

    const nameSim   = nameSimilarity_(mainName, candName);
    const idOverlap = countIdOverlap_(mainIdsArr, candIds);
    const brandMatch = mainBrand && candBrand &&
                       mainBrand.toLowerCase() === candBrand.toLowerCase();

    const score =
      0.60 * nameSim +
      0.30 * (idOverlap > 0 ? 1 : 0) +
      0.10 * (brandMatch ? 1 : 0);

    dbg('[B-CAND]', shop.code,
        'url=', prodUrl,
        'nameSim=', nameSim.toFixed(2),
        'idOverlap=', idOverlap,
        'brandMatch=', brandMatch ? 1 : 0,
        'score=', score.toFixed(2),
        'title=', candName);

    if (!best || score > best.score) {
      best = {
        url: prodUrl,
        title: candName,
        brand: candBrand,
        price: candPrice,
        ids: candIds,
        nameSim: nameSim,
        idOverlap: idOverlap,
        score: score,
        searchType: searchType,
        query: query
      };
    }
  }

  return best;
}

/**
 * Uit search-HTML productlinks halen voor een bepaalde domain.
 * We pakken gewoon alle href's, filteren op domein, dedupliceren,
 * en nemen max maxCount.
 */
function extractProductLinksFromSearchHtml_(html, domain, maxCount) {
  const out = [];
  const seen = {};
  const re = /<a[^>]+href=["']([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    let href = m[1];
    if (!href) continue;

    // relatieve links → absolute
    if (href.indexOf('http') !== 0) {
      href = 'https://' + domain.replace(/^www\./, '') + href;
    }

    // domain check
    if (href.indexOf(domain) === -1) continue;

    // simpele filters: geen anchors, geen zoeklinks zelf
    if (href.indexOf('#') !== -1) continue;
    if (/\/search|\/zoeken|\/catalogsearch/.test(href) && href.indexOf('.html') === -1) continue;

    const key = href.split('?')[0]; // querystring eraf
    if (seen[key]) continue;
    seen[key] = true;
    out.push(key);
    if (out.length >= (maxCount || 10)) break;
  }
  return out;
}

/**
 * Parsed ids-string (zoals in Scraped ids) → array met nette numerieke IDs.
 */
function parseIdsString_(s) {
  if (!s) return [];
  const out = [];
  const re = /\b(\d{8,18})\b/g;
  let m;
  while ((m = re.exec(String(s))) !== null) {
    const id = m[1];
    // rommel filteren
    if (id === '00000000000003' ||
        id === '00000000' ||
        id === '99999999' ||
        id === '66666666666666') continue;
    out.push(id);
  }
  return mergeUnique_([], out);
}

/**
 * Kies 1–2 "beste" IDs om op te zoeken:
 * - prefer 13–14 cijfers (EAN/GTIN)
 * - dan 12
 * - dan de rest
 */
function pickBestIdsForSearch_(idsArr) {
  if (!idsArr || !idsArr.length) return [];
  const byLen = {};
  idsArr.forEach(id => {
    const len = id.length;
    if (!byLen[len]) byLen[len] = [];
    byLen[len].push(id);
  });

  const order = [14, 13, 12, 11, 10, 9, 8];
  const selected = [];

  for (let i = 0; i < order.length; i++) {
    const len = order[i];
    if (byLen[len] && byLen[len].length) {
      byLen[len].forEach(id => selected.push(id));
      if (selected.length >= 2) break; // max 2 ID's gebruiken
    }
  }

  if (!selected.length) return idsArr.slice(0, 2);
  return selected.slice(0, 2);
}

/**
 * Aantal overlappende IDs tussen main en candidate.
 */
function countIdOverlap_(a, b) {
  if (!a || !b || !a.length || !b.length) return 0;
  const setB = {};
  b.forEach(id => { setB[id] = true; });
  let cnt = 0;
  a.forEach(id => { if (setB[id]) cnt++; });
  return cnt;
}
