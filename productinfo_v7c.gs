/** PRODUCTINFO & PRODUCTFIELD — v7c (decode HTML entities + robust EAN) **/

/* -------- Patches -------- */
const CACHE_VER      = "v7c";    // bump cache to refresh
const MAX_DESC       = 1200;
const MAX_TITLE      = 300;
const SAFE_CELL_MAX  = 49000;

function decodeHtml_(s) {
  if (!s) return "";
  let out = String(s);
  out = out.replace(/&#(\d+);/g, (_,d) => String.fromCharCode(parseInt(d,10)));
  out = out.replace(/&#x([0-9a-fA-F]+);/g, (_,h) => String.fromCharCode(parseInt(h,16)));
  const map = {
    "&nbsp;":" ", "&amp;":"&", "&lt;":"<", "&gt;":">",
    "&quot;":"\"", "&apos;":"'", "&lsquo;":"'", "&rsquo;":"'",
    "&ldquo;":"\"", "&rdquo;":"\"", "&hellip;":"…", "&ndash;":"–", "&mdash;":"—",
    "&euro;":"€"
  };
  out = out.replace(/&(nbsp|amp|lt|gt|quot|apos|lsquo|rsquo|ldquo|rdquo|hellip|ndash|mdash|euro);/g,(m)=>map[m]||m);
  return out;
}
function stripHtml_(s){ if(!s)return ""; const txt=String(s).replace(/<[^>]*>/g," "); return txt.replace(/\s+/g," ").trim();}
function clamp_(s,n){ if(!s)return ""; const clean=stripHtml_(decodeHtml_(s)); return clean.length>n?clean.slice(0,n-1)+"…":clean;}
function safeCell_(v,n=SAFE_CELL_MAX){ if(v===null||v===undefined)return ""; const s=String(v); return s.length>n?s.slice(0,n-1)+"…":s;}
function normalizeHtml_(txt) {
  if (!txt) return "";
  let s = String(txt)
    // strip alle scripts BEHALVE ld+json
    .replace(/<script(?![^>]*type=["']application\/ld\+json["'])[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");
  s = s.replace(/\s{2,}/g, " ");
  return s;
}

const PRODUCTINFO_HEADERS=[
  "title","brand","model","color","sku","mpn","gtin",
  "price","currency","availability","seller",
  "rating","reviews","image","description","canonical","domain"
];

function PRODUCTINFO(url){
  if(!url) return [[""]];
  try{
    const cacheKey=CACHE_VER+":"+url;
    const cached=CacheService.getScriptCache().get(cacheKey);
    let data;
    if(cached){ data=JSON.parse(cached); }
    else{
      const html=fetchHtml_(url);
      data=extractProductData_(html,url);
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(data), 21600); // 6 uur
    }
    const row=PRODUCTINFO_HEADERS.map(h=>safeCell_(data[h]??""));
    return [row];
  }catch(e){ return [[`ERROR: ${e.message}`]]; }
}

function PRODUCTFIELD(url, field){
  const arr=PRODUCTINFO(url);
  const idx=PRODUCTINFO_HEADERS.indexOf(String(field||"").toLowerCase());
  if(idx===-1) return [["ERROR: onbekend veld"]];
  return [[safeCell_(arr[0][idx])]];
}

/* ---------------- helpers ---------------- */

function fetchHtml_(url){
  const res=UrlFetchApp.fetch(url,{
    muteHttpExceptions:true, followRedirects:true, validateHttpsCertificates:true, method:"get",
    headers:{
      "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
      "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      "Accept-Language":"nl-NL,nl;q=0.9,en;q=0.8",
      "Cache-Control":"no-cache"
    }
  });
  const code=res.getResponseCode();
  if(code<200||code>=400) throw new Error(`HTTP ${code}`);
  return res.getContentText();
}

// Vervanger voor new URL()
function getHostname_(u){
  try{ const m=String(u).match(/^[a-z]+:\/\/([^\/?#:]+)(?::\d+)?/i); return m?m[1].replace(/^www\./,""):""; }
  catch(e){ return ""; }
}

/* ---------- EAN helpers (buiten elke functie) ---------- */
function sanitizeEAN_(s){
  const t=String(s||"").replace(/\D/g,"");             // alleen cijfers
  return (t.length>=8 && t.length<=14) ? t : "";       // GTIN-8/12/13/14
}

function extractEANFromHtml_(html){
  if(!html) return "";
  const candidates=[];
  // itemprop/meta/og varianten
  const reItemprop=/<(?:meta|span|div)[^>]+(?:itemprop|property|name)=["'](?:gtin|gtin13|gtin14|gtin12|ean|ean13|product:ean)["'][^>]+?(?:content|value)?=["']?([0-9\.\-\s]{8,20})["']?/gi;
  let m;
  while((m=reItemprop.exec(html))!==null) candidates.push(m[1]);
  // data- attributen
  const reData=/data-(?:ean|gtin|gtin13|gtin14)\s*=\s*["']([0-9\.\-\s]{8,20})["']/gi;
  while((m=reData.exec(html))!==null) candidates.push(m[1]);
  // zichtbare tekst
  const reText=/\b(?:EAN|GTIN(?:\s*1[234])?)\b[^0-9]{0,10}([0-9][0-9\.\-\s]{6,18}[0-9])/gi;
  while((m=reText.exec(html))!==null) candidates.push(m[1]);
  // expliciet "barcode" veld uit ruwe JSON/inline definities
  const reBarcode=/["']barcode["']\s*[:=]\s*["']?([0-9\.\-\s]{8,14})["']?/gi;
  while((m=reBarcode.exec(html))!==null) candidates.push(m[1]);

  // normaliseer naar cijfers en accepteer GTIN-8/12/13/14
  const cleaned=[...new Set(
    candidates
      .map(v => String(v||"").replace(/\D/g,""))
      .filter(s => s.length>=8 && s.length<=14)
  )];
  if(!cleaned.length) return "";
  cleaned.sort((a,b)=>b.length-a.length);
  return cleaned[0];
}
/* ------------------------------------------------------- */

function extractProductData_(html, url){
  const doc=normalizeHtml_(html);
  const domain=getHostname_(url);

  let productNode=null, offerNode=null, ratingNode=null;

  // PATCH: JSON-LD uit ruwe html
  const jsonldRegex=/<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m; const nodes=[];
  while((m=jsonldRegex.exec(html))!==null){  // <-- html i.p.v. doc
    parseJsonLdBlock_(m[1]).forEach(n=>nodes.push(n));
  }

  const flat=flattenNodes_(nodes);
  productNode = flat.find(n=>typeIncludes_(n,"Product")) || null;
  offerNode   = flat.find(n=>typeIncludes_(n,"Offer"))   || (productNode && pickOffer_(productNode)) || null;
  ratingNode  = flat.find(n=>typeIncludes_(n,"AggregateRating")) || (productNode && productNode.aggregateRating) || null;

  const og   = extractOG_(doc);
  const meta = extractMetaItemprops_(doc);

  const out={};
  out.title = clamp_(firstNonEmpty_([
    getDeep_(productNode,"name"), og["og:title"], meta["title"], extractH1_(doc), extractTitleTag_(doc)
  ]), MAX_TITLE);

  const brandObj=getDeep_(productNode,"brand");
  out.brand = firstNonEmpty_([ brandObj && (brandObj.name||brandObj.brand||brandObj), og["product:brand"], meta["brand"] ]);
  out.model = firstNonEmpty_([ getDeep_(productNode,"model"), getDeep_(productNode,"mpn"), meta["model"] ]);
  out.color = firstNonEmpty_([ getDeep_(productNode,"color"), meta["color"] ]);
  out.sku   = firstNonEmpty_([ getDeep_(productNode,"sku"),  meta["sku"] ]);
  out.mpn   = firstNonEmpty_([ getDeep_(productNode,"mpn"),  meta["mpn"] ]);

  // --- ROBUST GTIN/EAN ---
  out.gtin = sanitizeEAN_(firstNonEmpty_([
    getDeep_(productNode,"gtin13"),
    getDeep_(productNode,"gtin14"),
    getDeep_(productNode,"gtin12"),
    getDeep_(productNode,"gtin"),
    meta["gtin"], meta["gtin13"], meta["gtin14"], meta["gtin12"],
    meta["ean"], meta["ean13"], meta["ean-13"],
    og["product:ean"]
  ])) || sanitizeEAN_(extractEANFromHtml_(html)); // <-- html i.p.v. doc

  // prijs & valuta (met fallback)
  let priceStr=firstNonEmpty_([ getDeep_(offerNode,"price"), og["product:price:amount"], meta["price"] ]);
  let priceNum=normalizePrice_(priceStr);
  if(priceNum==="") priceNum=fallbackPriceFromHtml_(doc);
  out.price=priceNum;

  out.currency = firstNonEmpty_([ getDeep_(offerNode,"priceCurrency"), og["product:price:currency"], meta["pricecurrency"] ])
                 || guessCurrencyFromText_(doc) || "";

  const availRaw=firstNonEmpty_([ getDeep_(offerNode,"availability"), meta["availability"] ]);
  out.availability=simplifyAvailability_(availRaw);

  const sellerObj=getDeep_(offerNode,"seller");
  out.seller = sellerObj && (sellerObj.name||sellerObj.seller||sellerObj) || "";

  out.rating  = firstNonEmpty_([ getDeep_(ratingNode,"ratingValue"), og["og:rating"], meta["ratingvalue"] ]) || "";
  out.reviews = firstNonEmpty_([ getDeep_(ratingNode,"reviewCount"),  meta["reviewcount"] ]) || "";

  const img=getDeep_(productNode,"image");
  out.image = Array.isArray(img) ? img[0] : (img || og["og:image"] || "");

  out.description = clamp_(firstNonEmpty_([ getDeep_(productNode,"description"), og["og:description"], meta["description"] ]), MAX_DESC);

  out.canonical = og["og:url"] || extractCanonical_(doc) || url;
  out.domain    = domain;

  if(!out.title) out.title = clamp_(extractTitleTag_(doc), MAX_TITLE);
  return out;
}

function extractH1_(html){ const m=html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i); return m?stripHtml_(m[1]):""; }
function typeIncludes_(node,typeName){ if(!node||!node["@type"])return false; const t=node["@type"]; return (Array.isArray(t)?t:[t]).some(x=>String(x).toLowerCase()===String(typeName).toLowerCase()); }
function pickOffer_(product){ const offers=product&&product.offers; if(!offers) return null; if(Array.isArray(offers)) return offers.find(o=>o.price)||offers[0]; return offers; }
function flattenNodes_(x){ const out=[];(function walk(n){ if(!n||typeof n!=="object")return; out.push(n); Object.keys(n).forEach(k=>{ const v=n[k]; if(v&&typeof v==="object"){ if(Array.isArray(v)) v.forEach(walk); else walk(v);} });})(x); return out; }
function extractOG_(html){ const out={}; const re=/<meta\s+(?:property|name)=["']([^"']+)["']\s+content=["']([^"']*)["'][^>]*>/gi; let m; while((m=re.exec(html))!==null) out[m[1].toLowerCase()]=m[2]; return out; }
function extractMetaItemprops_(html){ const out={}; const re=/<meta\s+itemprop=["']([^"']+)["']\s+content=["']([^"']*)["'][^>]*>/gi; let m; while((m=re.exec(html))!==null) out[m[1].toLowerCase()]=m[2]; return out; }
function extractCanonical_(html){ const m=html.match(/<link[^>]+rel=["']canonical["'][^>]+href=["']([^"']+)["']/i); return m?m[1]:""; }
function extractTitleTag_(html){ const m=html.match(/<title[^>]*>([\s\S]*?)<\/title>/i); return m?m[1].trim():""; }
function getDeep_(obj,path){ if(!obj||!path)return undefined; const v=obj[path]; return Array.isArray(v)?v[0]:v; }
function firstNonEmpty_(arr){ for(let v of arr){ if(v===null||v===undefined)continue; const s=String(v).trim(); if(s) return s; } return ""; }
function normalizePrice_(p){ if(!p)return ""; let s=String(p).replace(/[^\d,.\-]/g,"").trim(); if(s.indexOf(",")>-1&&s.indexOf(".")>-1) s=s.replace(/\./g,"").replace(",","."); else if(s.indexOf(",")>-1) s=s.replace(",","."); const num=parseFloat(s); return isNaN(num)?"":num; }
function guessCurrencyFromText_(html){ if(/\bEUR?\b|€/.test(html))return "EUR"; if(/\bUSD?\b|\$/.test(html))return "USD"; if(/\bGBP\b|£/.test(html))return "GBP"; return ""; }
function simplifyAvailability_(a){ if(!a)return ""; const s=String(a).toLowerCase(); if(s.includes("instock")||s.includes("in stock")||s.includes("op voorraad"))return "inStock"; if(s.includes("outofstock")||s.includes("out of stock")||s.includes("niet op voorraad"))return "outOfStock"; if(s.includes("preorder"))return "preOrder"; return a; }
function safeJson_(txt){ return decodeHtml_(txt||"").replace(/^\ufeff/,''); }
function parseJsonLdBlock_(raw){
  const cleaned=safeJson_(raw).trim();
  const attempts=[cleaned];
  if(/}\s*{/.test(cleaned)) attempts.push(`[${cleaned.replace(/}\s*{/g,'},{')}]`);
  for(const candidate of attempts){
    try{
      const parsed=JSON.parse(candidate);
      return Array.isArray(parsed)?parsed:[parsed];
    }catch(e){ /* try next */ }
  }
  return [];
}

/* ---- prijs-fallback uit HTML ---- */
function fallbackPriceFromHtml_(html){
  let m=html.match(/"price"\s*:\s*"?(?<p>\d+[.,]\d{2})"?/i);
  if(m&&m.groups&&m.groups.p) return normalizePrice_(m.groups.p);
  m=html.match(/data-(?:price|amount|product-price)\s*=\s*"(?<p>\d+[.,]\d{2})"/i);
  if(m&&m.groups&&m.groups.p) return normalizePrice_(m.groups.p);
  m=html.match(/(?:prijs|price)[^€]{0,40}€\s*(?<p>\d+[.,]\d{2})/i);
  if(m&&m.groups&&m.groups.p) return normalizePrice_(m.groups.p);
  m=html.match(/€\s*(?<p>\d{1,4}(?:[.,]\d{3})*[.,]\d{2})/);
  if(m&&m.groups&&m.groups.p) return normalizePrice_(m.groups.p);
  return "";
}

//********* PRICE WATCHER V3 – afgestemd op G:H:I:J:K:L *********/

const CONFIG = {
  SHEET_NAME: 'Blad1',   // pas aan als je tab anders heet (bijv. 'Babyproducten')
  FIRST_ROW: 2,          // eerste data-rij
  COL_URL: 1,            // A = URL
  COL_GTIN: 4,           // D = GTIN (voor tab "Concurrenten")
  COL_LAST_PRICE: 7,     // G = Laatste prijs
  COL_PREV_PRICE: 8,     // H = Vorige prijs
  COL_LOWEST_PRICE: 9,   // I = Laagste prijs
  COL_WATCH: 11,         // K = Watch? (checkbox of 'Ja')
  COL_LAST_CHECK: 12,    // L = Laatste check (datum/tijd)
  EMAIL: 'jij@voorbeeld.nl',  // zet hier je echte adres
  MIN_DROP_PERCENT: 5    // drempel voor mailtje
};

function refreshPricesAndNotify() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Sheet niet gevonden: ' + CONFIG.SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.FIRST_ROW) return;

  const numRows = lastRow - CONFIG.FIRST_ROW + 1;

  // alles in één keer inlezen
  const urls   = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_URL,          numRows, 1).getValues();
  const gtins  = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_GTIN,         numRows, 1).getValues();
  const watchs = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_WATCH,        numRows, 1).getValues();
  const lastPs = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_LAST_PRICE,   numRows, 1).getValues();
  const prevPs = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_PREV_PRICE,   numRows, 1).getValues();
  const lowPs  = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_LOWEST_PRICE, numRows, 1).getValues();

  const drops = [];

  for (let i = 0; i < numRows; i++) {
    const row = CONFIG.FIRST_ROW + i;
    const url = urls[i][0];
    const rawWatch = watchs[i][0];

    // alleen rijen met URL én watch == TRUE of 'ja'
    const watch = (rawWatch === true) || (String(rawWatch).toLowerCase() === 'ja');
    if (!url || !watch) continue;

    const oldLast  = Number(lastPs[i][0]) || 0;
    const oldPrev  = Number(prevPs[i][0]) || 0;
    const oldLow   = Number(lowPs[i][0])  || 0;

    // productinfo één keer ophalen (met cache in PRODUCTINFO)
    let info;
    try {
      const arr = PRODUCTINFO(url);
      info = Array.isArray(arr) ? arr[0] : arr;
    } catch (e) {
      Logger.log('Fout bij PRODUCTFIELD voor rij ' + row + ' (' + url + '): ' + e);
      continue;
    }

    if (!info) continue;

    const priceIndex = PRODUCTINFO_HEADERS.indexOf('price');
    const gtinIndex  = PRODUCTINFO_HEADERS.indexOf('gtin');
    const rawPrice   = priceIndex >= 0 ? info[priceIndex] : '';
    const newGtin    = gtinIndex >= 0 ? sanitizeEAN_(info[gtinIndex]) : '';

    // string → getal
    let raw = rawPrice;
    if (typeof raw === 'string') {
      raw = raw.replace(/[^\d,.\-]/g, '').trim();
      if (raw.indexOf(',') > -1 && raw.indexOf('.') > -1) {
        raw = raw.replace(/\./g, '').replace(',', '.');
      } else if (raw.indexOf(',') > -1) {
        raw = raw.replace(',', '.');
      }
    }
    const newPrice = Number(raw);
    if (!newPrice || isNaN(newPrice)) {
      Logger.log('Geen geldige prijs voor rij ' + row + ': ' + raw);
      continue;
    }

    // vorige prijs = oude "laatste", tenzij die leeg was
    const prevPrice = oldLast > 0 ? oldLast : (oldPrev > 0 ? oldPrev : newPrice);

    // laagste prijs ooit
    const lowest = oldLow > 0 ? Math.min(oldLow, newPrice) : newPrice;

    // dalingspercentage t.o.v. oude prijs
    let dropPct = 0;
    if (oldLast > 0) {
      dropPct = ((oldLast - newPrice) / oldLast) * 100;
    }

    const isDrop = oldLast > 0 && newPrice < oldLast && dropPct >= CONFIG.MIN_DROP_PERCENT;
    if (isDrop) {
      drops.push({ row, url, oldPrice: oldLast, newPrice, dropPct });
    }

    // ✅ Alleen deze kolommen overschrijven
    sh.getRange(row, CONFIG.COL_PREV_PRICE).setValue(prevPrice);
    sh.getRange(row, CONFIG.COL_LAST_PRICE).setValue(newPrice);
    sh.getRange(row, CONFIG.COL_LOWEST_PRICE).setValue(lowest);
    sh.getRange(row, CONFIG.COL_LAST_CHECK).setValue(new Date());

    // GTIN (D) aanvullen of bijwerken voor tab "Concurrenten"
    const currentGtin = String(gtins[i][0] || '').trim();
    if (newGtin && newGtin !== currentGtin) {
      sh.getRange(row, CONFIG.COL_GTIN).setValue(newGtin);
    }
  }

  if (drops.length > 0) {
    sendDropEmail_(drops);
  }
}

function sendDropEmail_(drops) {
  if (!CONFIG.EMAIL) return;

  let body = 'Er zijn prijsdalingen bij je babyproducten:\n\n';
  drops.forEach(d => {
    body +=
      'Rij: ' + d.row + '\n' +
      'URL: ' + d.url + '\n' +
      'Oude prijs: € ' + d.oldPrice.toFixed(2) + '\n' +
      'Nieuwe prijs: € ' + d.newPrice.toFixed(2) + '\n' +
      'Daling: ' + d.dropPct.toFixed(1) + ' %\n\n';
  });

  MailApp.sendEmail({
    to: CONFIG.EMAIL,
    subject: 'Prijsdaling babyproducten 🎁',
    body
  });
}

// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 Babytools')
    .addItem('Prijzen verversen', 'refreshPricesAndNotify')
    .addToUi();
}
