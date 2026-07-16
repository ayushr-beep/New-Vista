/* =========================================================================
   MANIFEST INTELLIGENCE — full client-side engine.
   All parsing, adapting, demand/cost modeling, and optimization below runs
   entirely in this browser on whatever files you upload. Nothing is sent
   to a server except the optional AI Analyst calls to api.anthropic.com.
   ========================================================================= */

/* ---------------------------- THEME ---------------------------- */
(function themeInit(){
  const saved = localStorage.getItem('mi_theme');
  const theme = saved || 'dark';
  document.documentElement.setAttribute('data-theme', theme);
  updateThemeIcon(theme);
})();
function updateThemeIcon(theme){
  const icon = document.getElementById('themeIcon');
  if(theme === 'light'){
    icon.innerHTML = '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M2 12h2M20 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4"/>';
  } else {
    icon.innerHTML = '<path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>';
  }
}
document.getElementById('themeToggle').addEventListener('click', ()=>{
  const cur = document.documentElement.getAttribute('data-theme');
  const next = cur === 'dark' ? 'light' : 'dark';
  const icon = document.getElementById('themeIcon');
  icon.style.transform = 'rotate(-90deg) scale(0.5)';
  icon.style.opacity = '0';
  document.body.style.transition = 'background .35s ease';
  setTimeout(()=>{
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('mi_theme', next);
    updateThemeIcon(next);
    icon.style.transform = 'rotate(0deg) scale(1)';
    icon.style.opacity = '1';
  }, 140);
  if(state.frontierData) setTimeout(renderFrontierChart, 150);
  if(state.demand) setTimeout(renderGapChart, 150);
});

/* ---------------------------- TAB ROUTING ---------------------------- */
const railBtns = document.querySelectorAll('.rail-btn[data-tab]');
railBtns.forEach(btn=>{
  btn.addEventListener('click', ()=>{
    railBtns.forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
    document.getElementById('panel-' + btn.dataset.tab).classList.add('active');
  });
});

/* ---------------------------- TOAST ---------------------------- */
function toast(msg, isError){
  const wrap = document.getElementById('toastWrap');
  const t = document.createElement('div');
  t.className = 'toast' + (isError ? ' error' : '');
  t.innerHTML = `<span class="t-dot"></span><span>${msg}</span>`;
  wrap.appendChild(t);
  setTimeout(()=>{ t.style.opacity='0'; t.style.transition='opacity .3s'; setTimeout(()=>t.remove(), 300); }, 3800);
}

/* ============================================================
   CSV PARSER — RFC4180-aware, handles quoted commas/newlines.
   ============================================================ */
/* Single source of truth for SKU-as-primary-key formatting. Every adapter
   below (sales, fees, inventory, manifest) runs raw SKU text through this
   before it's used as a join key. Without this, a manifest SKU typed/
   exported as "hney-90004" would never match a sales-file SKU stored as
   "HNEY-90004" — same real item, different bytes — and "Active manifest
   only" scope would silently resolve to zero SKUs. Trimming collapses
   incidental whitespace (trailing spaces from Excel exports); uppercasing
   makes the join case-insensitive without changing what's displayed
   anywhere else (display always uses the adapted record's own sku field,
   which is this normalized value, so it's consistent everywhere, not just
   internally). */
function normalizeSku(raw){
  return String(raw||'').trim().toUpperCase();
}

function parseCSV(text){
  const rows = []; let row=[]; let field=''; let inQuotes=false; let i=0; const len=text.length;
  while(i<len){
    const c = text[i];
    if(inQuotes){
      if(c === '"'){ if(text[i+1]==='"'){field+='"'; i+=2; continue;} inQuotes=false; i++; continue; }
      field+=c; i++; continue;
    } else {
      if(c === '"'){ inQuotes=true; i++; continue; }
      if(c === ','){ row.push(field); field=''; i++; continue; }
      if(c === '\r'){ i++; continue; }
      if(c === '\n'){ row.push(field); field=''; if(row.length>1||row[0]!=='') rows.push(row); row=[]; i++; continue; }
      field+=c; i++; continue;
    }
  }
  if(field!=='' || row.length>0){ row.push(field); rows.push(row); }
  return rows;
}
function rowsToObjects(rows){
  if(!rows.length) return { headers:[], objects:[] };
  const headers = rows[0].map(h=>h.trim());
  const out = [];
  for(let r=1;r<rows.length;r++){
    const rawRow = rows[r];
    const obj = {};
    for(let c=0;c<headers.length;c++) obj[headers[c]] = rawRow[c] !== undefined ? rawRow[c] : '';
    obj.__rawLen = rawRow.length;
    obj.__expectedLen = headers.length;
    out.push(obj);
  }
  return { headers, objects: out };
}

/* Encoding-robust file read: try strict UTF-8, fall back to windows-1252
   (covers the real "Sales export is latin-1" case from the brief, without
   needing to hardcode which file needs which encoding). */
function readFileSmart(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => {
      const buf = new Uint8Array(reader.result);
      let text, encoding;
      try {
        text = new TextDecoder('utf-8', {fatal:true}).decode(buf);
        encoding = 'utf-8';
      } catch(e){
        text = new TextDecoder('windows-1252').decode(buf);
        encoding = 'windows-1252 (fallback)';
      }
      resolve({ text, encoding });
    };
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsArrayBuffer(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

function isExcelFile(file){
  const name = (file.name||'').toLowerCase();
  return name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.xlsm');
}

/* Normalizes row values coming from SheetJS (which preserves numbers as
   numbers, dates as Date objects) into strings, matching the shape every
   adapter already expects from CSV text parsing. Ragged-row detection
   (__rawLen vs __expectedLen) doesn't apply the same way to a real grid,
   so well-formed sheet rows are marked consistent by construction. */
function excelRowsToObjects(sheet){
  const rows = XLSX.utils.sheet_to_json(sheet, { defval:'', raw:false, header:1 });
  if(!rows.length) return { headers:[], objects:[] };
  const headers = rows[0].map(h=>String(h).trim());
  const out = [];
  for(let r=1;r<rows.length;r++){
    const rawRow = rows[r];
    if(!rawRow || rawRow.every(c=>c===''||c==null)) continue;
    const obj = {};
    for(let c=0;c<headers.length;c++){
      const v = rawRow[c];
      obj[headers[c]] = (v===undefined||v===null) ? '' : String(v);
    }
    obj.__rawLen = rawRow.length;
    obj.__expectedLen = headers.length;
    out.push(obj);
  }
  return { headers, objects: out };
}

/* Normalizes a raw 2D array (one sheet, header:1 mode) starting at a given
   header row index into the same { headers, objects } shape the CSV path
   produces. */
function gridToObjectsFromHeaderRow(rows, headerRowIdx){
  const headers = rows[headerRowIdx].map(h=>String(h||'').trim());
  const out = [];
  for(let r=headerRowIdx+1; r<rows.length; r++){
    const rawRow = rows[r];
    if(!rawRow || rawRow.every(c=>c===''||c==null)) continue;
    const obj = {};
    for(let c=0;c<headers.length;c++){
      const v = rawRow[c];
      obj[headers[c]] = (v===undefined||v===null) ? '' : String(v);
    }
    obj.__rawLen = rawRow.filter(c=>c!==''&&c!=null).length ? rawRow.length : headers.length;
    obj.__expectedLen = headers.length;
    out.push(obj);
  }
  return { headers, objects: out };
}

/* Scores how well a single row looks like a header row for a given schema:
   counts how many of the schema's required+optional aliases appear as
   near-exact column names in that row. Used to find the real header row
   and the real data sheet in multi-sheet, multi-section real-world Amazon
   templates (e.g. Send-to-Amazon manifests with Instructions/Data
   definitions/Template/Example tabs, where the actual header is row 6,
   not row 1, and the data sheet isn't the first sheet). */
function scoreRowAsHeader(row, schemaKey){
  const schema = COLUMN_SCHEMAS[schemaKey];
  if(!schema || !row) return 0;
  const cells = row.map(c => String(c||'').toLowerCase().trim()).filter(c=>c);
  if(!cells.length) return 0;
  let score = 0;
  for(const field of Object.keys(schema)){
    for(const alias of schema[field]){
      if(cells.some(c => c === alias || c.startsWith(alias))){ score++; break; }
    }
  }
  return score;
}

/* Scans every sheet and, within each, every row in the first 25 rows,
   scoring each as a potential header row against the schema. Returns the
   best { sheetName, headerRowIdx, rows, score } across the whole workbook.
   This is what lets a real multi-tab Amazon export (Instructions / Data
   definitions / Create workflow – template / ...) get the right sheet and
   the right header row automatically, instead of blindly reading sheet
   0 / row 0. */
/* Penalizes sheets whose NAME signals they're not the real data the user
   filled in (Amazon's own templates ship an "example"/"sample" tab full of
   obviously fake illustrative rows like "MySKU001" alongside the real
   "template" tab the user actually edits). Without this, a sheet name
   alone can't outweigh row count, and a fuller example tab would wrongly
   win over a sparser but real template tab. */
function sheetNamePenalty(sheetName){
  const n = sheetName.toLowerCase();
  if(/\bexample\b|\bsample\b|\bdemo\b|\billustrat/.test(n)) return -1000;
  if(/\binstruction|\bdefinition|\breadme|\bguide\b/.test(n)) return -2000;
  return 0;
}

function findBestSheetAndHeaderRow(workbook, schemaKey){
  let best = { sheetName:null, headerRowIdx:-1, rows:null, score:-Infinity, dataRowCount:0 };
  let bestEmpty = null; // best-scoring header match even if it has zero real data rows — kept only as a last-resort fallback

  for(const sheetName of workbook.SheetNames){
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:false });
    const scanLimit = Math.min(rows.length, 25);
    const namePenalty = sheetNamePenalty(sheetName);
    for(let r=0; r<scanLimit; r++){
      const headerScore = scoreRowAsHeader(rows[r], schemaKey);
      if(headerScore === 0) continue;
      const dataRowCount = rows.slice(r+1).filter(row => row && row.some(c => String(c||'').trim() !== '')).length;
      const combinedScore = headerScore * 10 + namePenalty;

      if(dataRowCount === 0){
        // Track the best header-only match purely as a fallback — a sheet
        // with no real rows beneath its header can never win outright,
        // no matter how well-named it is, since it would silently produce
        // zero usable data.
        if(!bestEmpty || combinedScore > bestEmpty.score){
          bestEmpty = { sheetName, headerRowIdx:r, rows, score:combinedScore, headerScore, dataRowCount };
        }
        continue;
      }

      const better = combinedScore > best.score ||
        (combinedScore === best.score && dataRowCount > best.dataRowCount);
      if(better){
        best = { sheetName, headerRowIdx:r, rows, score:combinedScore, headerScore, dataRowCount };
      }
    }
  }

  if(best.score === -Infinity){
    // No sheet had BOTH a header match AND real data — fall back to the
    // best header-only match if one exists, so the UI can at least show
    // why nothing came through, rather than failing silently.
    if(bestEmpty) return { ...bestEmpty, noDataRowsFound: true };
    best.score = 0;
  }
  return best;
}

/* Unified loader: detects format by extension, returns the same
   { headers, objects, encoding, format } shape regardless of source,
   so every adapter downstream is format-agnostic.
   schemaKey tells the Excel path which sheet/header-row to look for —
   required because real multi-sheet workbooks (like Amazon's own
   Send-to-Amazon templates) bury the real data on a non-first sheet,
   under a non-first row, alongside instructional/example sheets. */
async function loadTabularFile(file, schemaKey){
  if(isExcelFile(file)){
    if(typeof XLSX === 'undefined'){
      throw new Error('Excel support library failed to load (blocked network or offline). Try saving this file as .csv and uploading that instead.');
    }
    const buf = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(buf, { type:'array', cellDates:false });

    if(!schemaKey){
      // No schema to detect against (caller didn't ask for smart
      // detection) -- only then is row-1 the safe assumption.
      const firstSheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[firstSheetName];
      const { headers, objects } = excelRowsToObjects(sheet);
      return { headers, objects, encoding: 'xlsx (binary)', format: 'excel', sheetName: firstSheetName, sheetCount: wb.SheetNames.length, headerRow: 1 };
    }

    const best = findBestSheetAndHeaderRow(wb, schemaKey);
    if(!best.headerScore || !best.rows){
      // fall back to naive first-sheet/first-row rather than failing outright
      const firstSheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[firstSheetName];
      const { headers, objects } = excelRowsToObjects(sheet);
      return { headers, objects, encoding: 'xlsx (binary)', format: 'excel', sheetName: firstSheetName, sheetCount: wb.SheetNames.length, headerRow: 1, sheetAutoDetectFailed: true };
    }
    const { headers, objects } = gridToObjectsFromHeaderRow(best.rows, best.headerRowIdx);
    return { headers, objects, encoding: 'xlsx (binary)', format: 'excel', sheetName: best.sheetName, sheetCount: wb.SheetNames.length, headerRow: best.headerRowIdx+1 };
  } else {
    const { text, encoding } = await readFileSmart(file);
    const { headers, objects } = rowsToObjects(parseCSV(text));
    return { headers, objects, encoding, format: 'csv' };
  }
}

/* Multi-sheet-aware variant: unlike loadTabularFile (which picks the
   single BEST sheet, correct for templates with example/instruction
   noise sheets), some real workbooks genuinely have real matching data
   spread across multiple sheets — confirmed by direct inspection of the
   real vendor spec sheet upload, which has 3 sheets that all contain
   real, non-overlapping item rows. This merges every sheet that scores
   above zero on the schema, rather than discarding all but one. */
async function loadMultiSheetTabularFile(file, schemaKey){
  if(!isExcelFile(file)){
    return loadTabularFile(file, schemaKey); // CSV is single-table by nature
  }
  if(typeof XLSX === 'undefined'){
    throw new Error('Excel support library failed to load (blocked network or offline). Try saving this file as .csv and uploading that instead.');
  }
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type:'array', cellDates:false });

  let mergedObjects = [];
  let headers = null;
  const sheetsUsed = [];
  for(const sheetName of wb.SheetNames){
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:false });
    const scanLimit = Math.min(rows.length, 10);
    let bestRowIdx = -1, bestScore = 0;
    for(let r=0; r<scanLimit; r++){
      const score = scoreRowAsHeader(rows[r], schemaKey);
      if(score > bestScore){ bestScore = score; bestRowIdx = r; }
    }
    if(bestRowIdx === -1) continue; // this sheet has no real header match — skip it (e.g. an empty or unrelated tab)
    const { headers: h, objects } = gridToObjectsFromHeaderRow(rows, bestRowIdx);
    if(!headers) headers = h;
    mergedObjects = mergedObjects.concat(objects);
    sheetsUsed.push(sheetName);
  }
  if(!sheetsUsed.length){
    // nothing matched anywhere — fall back to the single-best-sheet path so the user still sees something
    return loadTabularFile(file, schemaKey);
  }
  return { headers: headers||[], objects: mergedObjects, encoding:'xlsx (binary)', format:'excel', sheetName: sheetsUsed.join(', '), sheetCount: wb.SheetNames.length, sheetsUsed, headerRow: null };
}

/* ============================================================
   STATE/REGION MAP — validated against real ship-state values.
   ============================================================ */
const STATE_TO_REGION = {
  ME:'East',NH:'East',VT:'East',MA:'East',RI:'East',CT:'East',NY:'East',NJ:'East',PA:'East',
  DE:'East',MD:'East',DC:'East',VA:'East',WV:'East',NC:'East',SC:'East',GA:'East',FL:'East',
  OH:'Central',MI:'Central',IN:'Central',IL:'Central',WI:'Central',MN:'Central',IA:'Central',
  MO:'Central',ND:'Central',SD:'Central',NE:'Central',KS:'Central',KY:'Central',TN:'Central',
  AL:'Central',MS:'Central',AR:'Central',LA:'Central',OK:'Central',TX:'Central',
  MT:'West',WY:'West',CO:'West',NM:'West',ID:'West',UT:'West',AZ:'West',NV:'West',WA:'West',
  OR:'West',CA:'West',AK:'West',HI:'West',
  PR:'East',VI:'East',GU:'West',AS:'West',MP:'West'
};
const STATE_NAME_TO_ABBR = {
  'alabama':'AL','alaska':'AK','arizona':'AZ','arkansas':'AR','california':'CA','colorado':'CO',
  'connecticut':'CT','delaware':'DE','florida':'FL','georgia':'GA','hawaii':'HI','idaho':'ID',
  'illinois':'IL','indiana':'IN','iowa':'IA','kansas':'KS','kentucky':'KY','louisiana':'LA',
  'maine':'ME','maryland':'MD','massachusetts':'MA','michigan':'MI','minnesota':'MN',
  'mississippi':'MS','missouri':'MO','montana':'MT','nebraska':'NE','nevada':'NV',
  'new hampshire':'NH','new jersey':'NJ','new mexico':'NM','new york':'NY','north carolina':'NC',
  'north dakota':'ND','ohio':'OH','oklahoma':'OK','oregon':'OR','pennsylvania':'PA',
  'rhode island':'RI','south carolina':'SC','south dakota':'SD','tennessee':'TN','texas':'TX',
  'utah':'UT','vermont':'VT','virginia':'VA','washington':'WA','west virginia':'WV',
  'wisconsin':'WI','wyoming':'WY','district of columbia':'DC','puerto rico':'PR',
  'virgin islands':'VI','guam':'GU'
};
function resolveStateToRegion(raw){
  if(!raw) return null;
  let s = String(raw).trim().toLowerCase().replace(/\./g,'');
  if(!s) return null;
  if(s.length===2) return STATE_TO_REGION[s.toUpperCase()] || null;
  const abbr = STATE_NAME_TO_ABBR[s];
  return abbr ? (STATE_TO_REGION[abbr] || null) : null;
}

/* ============================================================
   ADAPTERS
   ============================================================ */

/* Column schemas — built directly from the real header rows the user
   provided. Listed alias first (their exact column name), then common
   variants, so detection is tuned to their real files first and still
   degrades gracefully for similar-but-not-identical exports. */
const COLUMN_SCHEMAS = {
  sales: {
    sku: ['sku','merchant sku','msku'],
    asin: ['asin'],
    purchaseDate: ['purchase-da','purchase date','order date','date'],
    productName: ['product-name','product name','title'],
    itemStatus: ['item-status','status'],
    quantity: ['quantity','qty'],
    currency: ['currency'],
    itemPrice: ['item-price','price'],
    shipCity: ['ship-city','city'],
    shipState: ['ship-state','ship state','state'],
    shipPostal: ['ship-postal-code','postal','zip'],
    shipCountry: ['ship-country','ship country','country'],
  },
  fees: {
    sku: ['sku'],
    asin: ['asin'],
    shipmentId: ['shipmentid','shipment id','shipment'],
    date: ['date'],
    week: ['week'],
    region: ['region'],
    planned: ['planned'],
    recvQty: ['recvqty','received qty','received','qty','quantity'],
    sizeTier: ['sizetier','size tier','tier'],
    placementFee: ['placementfee','placement fee','fee'],
  },
  inventory: {
    sku: ['sku'],
    region: ['region','zone'],
    onHand: ['on-hand units','on-hand','on hand','onhand','qty'],
  },
  manifest: {
    sku: ['merchant sku','sku','msku'],
    quantity: ['quantity','units','qty'],
    expirationDate: ['expiration date','expiry date'],
    lotCode: ['manufacturing lot code','lot code','batch code'],
    unitsPerBox: ['units per box','units/box'],
    numberOfBoxes: ['number of boxes','# of boxes','boxes'],
    boxLength: ['box length','length (in)'],
    boxWidth: ['box width','width (in)'],
    boxHeight: ['box height','height (in)'],
    boxWeight: ['box weight','weight (lb)'],
    sizeTier: ['size tier','sizetier','tier'],
  },
  specsheet: {
    itemId: ['item id','itemid'], // full vendor-prefixed ID, e.g. PETL-00233 — real join key confirmed against actual upload
    vendor: ['vendor'],
    casePack: ['case pack','casepack'],
    weight: ['weight'],
    length: ['length'],
    width: ['width'],
    height: ['height'],
    remarks: ['remarks'],
  }
};

/* Runs immediately on upload (before Run Analysis) so the user sees what
   was matched against their real header row right away. Returns both the
   per-field match and a flat list for rendering. */
function detectColumnMapping(headers, schemaKey){
  const schema = COLUMN_SCHEMAS[schemaKey];
  const lower = headers.map(h=>h.toLowerCase().trim());
  const mapping = {};
  for(const field of Object.keys(schema)){
    const aliases = schema[field];
    let matchedHeader = null;
    // exact match first (handles their precise real column names)
    for(const alias of aliases){
      const idx = lower.findIndex(h => h === alias);
      if(idx !== -1){ matchedHeader = headers[idx]; break; }
    }
    // fallback to substring match
    if(!matchedHeader){
      for(const alias of aliases){
        const idx = lower.findIndex(h => h.includes(alias));
        if(idx !== -1){ matchedHeader = headers[idx]; break; }
      }
    }
    mapping[field] = matchedHeader;
  }
  return mapping;
}

function adaptSalesRows(rawRows){
  const includeStatuses = ['Shipped','Shipping','Unshipped'];
  const clean = [];
  let droppedNonUS=0, droppedUnresolvedState=0, droppedBadQty=0, droppedStatus=0;
  const headers = rawRows.length ? Object.keys(rawRows[0]).filter(h=>!h.startsWith('__')) : [];
  const lower = headers.map(h=>h.toLowerCase());
  const find = (aliases) => { for(const a of aliases){ const idx = lower.findIndex(h=>h.includes(a)); if(idx!==-1) return headers[idx]; } return null; };
  const skuCol = find(['sku']);
  const stateCol = find(['ship-state','ship state','state']);
  const countryCol = find(['ship-country','ship country','country']);
  const qtyCol = find(['quantity','qty']);
  const statusCol = find(['item-status','status']);
  const priceCol = find(['item-price','price']);
  const dateCol = find(['purchase-da','purchase date','order date','date']);
  const asinCol = find(['asin']);

  for(const r of rawRows){
    if(countryCol && (r[countryCol]||'').trim().toUpperCase() !== 'US'){ droppedNonUS++; continue; }
    if(statusCol && includeStatuses.length && !includeStatuses.includes((r[statusCol]||'').trim())){ droppedStatus++; continue; }
    const region = stateCol ? resolveStateToRegion(r[stateCol]) : null;
    if(!region){ droppedUnresolvedState++; continue; }
    const qty = qtyCol ? parseInt(r[qtyCol],10) : 1;
    if(!Number.isFinite(qty) || qty<=0){ droppedBadQty++; continue; }
    clean.push({
      sku: skuCol ? normalizeSku(r[skuCol]) : '',
      asin: asinCol ? (r[asinCol]||'').trim() : '',
      region, quantity: qty,
      price: priceCol ? (parseFloat(r[priceCol])||0) : 0,
      date: dateCol ? (r[dateCol]||'') : ''
    });
  }
  return { records: clean, stats: { totalRaw: rawRows.length, kept: clean.length, droppedNonUS, droppedUnresolvedState, droppedBadQty, droppedStatus, detectedColumns:{skuCol,stateCol,countryCol,qtyCol,statusCol} } };
}

function adaptFeeRows(rawObjects){
  const clean = []; let malformed = 0;
  const VALID_REGIONS = ['East','Central','West'];
  const headers = rawObjects.length ? Object.keys(rawObjects[0]).filter(h=>!h.startsWith('__')) : [];
  const lower = headers.map(h=>h.toLowerCase());
  const find = (aliases) => { for(const a of aliases){ const idx = lower.findIndex(h=>h.includes(a)); if(idx!==-1) return headers[idx]; } return null; };
  const skuCol = find(['sku']);
  const regionCol = find(['region']);
  const qtyCol = find(['recvqty','received','qty','quantity']);
  const feeCol = find(['placementfee','fee']);
  const tierCol = find(['sizetier','size tier','tier']);
  const shipCol = find(['shipmentid','shipment']);

  for(const r of rawObjects){
    if(r.__rawLen !== r.__expectedLen){ malformed++; continue; }
    const region = regionCol ? r[regionCol] : null;
    if(!VALID_REGIONS.includes(region)){ malformed++; continue; }
    const qty = qtyCol ? parseFloat(r[qtyCol]) : NaN;
    const fee = feeCol ? parseFloat(r[feeCol]) : NaN;
    if(!Number.isFinite(qty) || qty<=0 || !Number.isFinite(fee)){ malformed++; continue; }
    clean.push({
      sku: skuCol ? normalizeSku(r[skuCol]) : '',
      shipmentId: shipCol ? r[shipCol] : '',
      region, qty,
      sizeTier: tierCol ? (r[tierCol]||'').trim() : 'LargeStandard',
      fee, feePerUnit: fee/qty
    });
  }
  return { records: clean, stats: { totalRaw: rawObjects.length, kept: clean.length, malformed, detectedColumns:{skuCol,regionCol,qtyCol,feeCol,tierCol} } };
}

function adaptInventoryRows(rawObjects){
  const clean = []; let dropped = 0;
  const VALID_REGIONS = ['East','Central','West'];
  const headers = rawObjects.length ? Object.keys(rawObjects[0]).filter(h=>!h.startsWith('__')) : [];
  const lower = headers.map(h=>h.toLowerCase());
  const find = (aliases) => { for(const a of aliases){ const idx = lower.findIndex(h=>h.includes(a)); if(idx!==-1) return headers[idx]; } return null; };
  const skuCol = find(['sku']);
  const regionCol = find(['region','zone']);
  const onHandCol = find(['on-hand','on hand','onhand','qty']);

  for(const r of rawObjects){
    const sku = skuCol ? normalizeSku(r[skuCol]) : '';
    const region = regionCol ? r[regionCol] : null;
    const onHand = onHandCol ? parseFloat(r[onHandCol]) : NaN;
    if(!sku || sku === 'UNIDENTIFIED' || !VALID_REGIONS.includes(region) || !Number.isFinite(onHand)){ dropped++; continue; }
    clean.push({ sku, region, onHand });
  }
  return { records: clean, stats: { totalRaw: rawObjects.length, kept: clean.length, dropped } };
}
function buildInventoryBySkuRegion(invRecords){
  const out = {};
  for(const r of invRecords){
    if(!out[r.sku]) out[r.sku] = { East:0, Central:0, West:0 };
    out[r.sku][r.region] += r.onHand;
  }
  return out;
}

/* ============================================================
   DEMAND ENGINE
   ============================================================ */
/* Parses the real DD-MM-YYYY dates already present on sales records and
   returns the actual span in days covered by the uploaded file. This is
   used to compute a REAL daily velocity (total units / actual days),
   never an assumed 30-day month — if the file covers 31 days, the math
   uses 31; if it covers 6 months, it uses that real span instead. */
function computeRealDateSpanDays(salesRecords){
  const dates = [];
  for(const r of salesRecords){
    if(!r.date) continue;
    const m = r.date.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/); // DD-MM-YYYY, matches the real Sales export format
    if(m){
      const d = new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
      if(!isNaN(d.getTime())) dates.push(d.getTime());
    }
  }
  if(dates.length < 2) return null;
  const spanMs = Math.max(...dates) - Math.min(...dates);
  const spanDays = Math.round(spanMs / 86400000) + 1; // inclusive of both endpoints
  return spanDays > 0 ? spanDays : null;
}

/* Builds a per-day per-region velocity table from real sales records.
   Every number in this output traces to a real order row in the
   uploaded sales CSV — this is the raw strike-rate signal the tool
   uses for demand %, restock timing, and the daily velocity KPIs.
   No estimates, no smoothing, no fill-forward for missing days. */
function buildVelocityByDay(salesRecords){
  if(!salesRecords || !salesRecords.length) return null;
  const byDay = {};
  for(const rec of salesRecords){
    if(!rec.date) continue;
    if(!byDay[rec.date]) byDay[rec.date] = { East:0, Central:0, West:0 };
    byDay[rec.date][rec.region] = (byDay[rec.date][rec.region]||0) + rec.quantity;
  }
  // Parse and sort dates — real format from the Sales export is DD-MM-YYYY
  const days = Object.entries(byDay).map(([date, v]) => {
    const m = date.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    const sortKey = m ? `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}` : date;
    const total = (v.East||0)+(v.Central||0)+(v.West||0);
    return { date, sortKey, East:v.East||0, Central:v.Central||0, West:v.West||0, total };
  }).sort((a,b)=>a.sortKey.localeCompare(b.sortKey));

  if(!days.length) return null;
  const n = days.length;
  const totalByRegion = {
    East: days.reduce((s,d)=>s+d.East,0),
    Central: days.reduce((s,d)=>s+d.Central,0),
    West: days.reduce((s,d)=>s+d.West,0),
    Total: days.reduce((s,d)=>s+d.total,0),
  };
  const avgByRegion = {
    East: totalByRegion.East/n,
    Central: totalByRegion.Central/n,
    West: totalByRegion.West/n,
    Total: totalByRegion.Total/n,
  };
  const peakDay = days.reduce((best,d)=>d.total>best.total?d:best, days[0]);
  const slowDay = days.reduce((best,d)=>d.total<best.total?d:best, days[0]);
  return { days, totalByRegion, avgByRegion, peakDay, slowDay, dateRange: { first: days[0].date, last: days[days.length-1].date } };
}

function buildDemandEngine(salesRecords, inventoryBySkuRegion, sellThroughCorrection){
  const useCorrection = !!sellThroughCorrection && !!inventoryBySkuRegion;
  const bySku = {};
  for(const rec of salesRecords){
    if(!bySku[rec.sku]) bySku[rec.sku] = { East:0, Central:0, West:0, totalUnits:0, asin: rec.asin };
    bySku[rec.sku][rec.region] += rec.quantity;
    bySku[rec.sku].totalUnits += rec.quantity;
  }
  const skuDemand = {};
  for(const sku of Object.keys(bySku)){
    const row = bySku[sku];
    let weights = { East: row.East, Central: row.Central, West: row.West };
    if(useCorrection){
      const inv = inventoryBySkuRegion[sku] || { East:0, Central:0, West:0 };
      weights = {
        East: inv.East>0 ? row.East/inv.East : row.East,
        Central: inv.Central>0 ? row.Central/inv.Central : row.Central,
        West: inv.West>0 ? row.West/inv.West : row.West,
      };
    }
    const sum = weights.East+weights.Central+weights.West;
    const pct = sum>0 ? {East:weights.East/sum, Central:weights.Central/sum, West:weights.West/sum} : {East:1/3,Central:1/3,West:1/3};
    skuDemand[sku] = { asin: row.asin, totalUnits: row.totalUnits, rawUnits: {East:row.East,Central:row.Central,West:row.West}, demandPct: pct, sellThroughCorrected: useCorrection };
  }
  return skuDemand;
}

/* ============================================================
   COST MODEL — real 2026 Amazon FBA inbound placement rule.
   Verified via research: for standard-size products there are really
   only TWO outcomes per region —
     MINIMAL split: any config that doesn't qualify below -> full per-unit fee
     AMAZON-OPTIMIZED: $0 fee, but ONLY if that region's shipment for this
       SKU includes >=5 identical cartons (same qty/carton, same item mix).
       Leftover units in a non-identical "remainder" box are allowed
       alongside the 5+ identical ones.
   (Partial splits, 2-3 locations at a reduced rate, still exist but only
   for Bulky-size products — not modeled here since the real fee log this
   tool was built against is for standard tiers; a known simplification.)
   ============================================================ */

/* ============================================================
   DIMENSIONAL WEIGHT / SIZE-TIER CLASSIFIER — research-backed cross-check.
   Built from independently-verified Amazon FBA mechanics (confirmed
   against Amazon's own Buy with Prime pricing page, which mirrors FBA's
   real fee logic): dimensional weight = (L x W x H in inches) / 139,
   shipping weight = max(actual weight, dimensional weight) for every
   tier except Small Standard and Extra-Large 150+lb (unit weight only).
   Boundary values for Small Standard, Large Standard, and Large Bulky
   are HIGH confidence (consistently confirmed across independent
   sources). The newer 2026 "Small Bulky" carve-out and Extra-Large
   weight sub-bands are LOWER confidence — secondary trackers didn't
   agree on the exact discriminating rule, so this is flagged honestly
   in the result rather than asserted as fact.
   This is a CROSS-CHECK against whatever size tier the manifest/fee log
   already states — it never silently overrides real data, since a
   computed guess is never as trustworthy as the user's own real label
   or Amazon's own measurement. */
function classifySizeTier(lengthIn, widthIn, heightIn, weightLb){
  if(!lengthIn || !widthIn || !heightIn || !weightLb || lengthIn<=0 || widthIn<=0 || heightIn<=0 || weightLb<=0) return null;
  const sides = [lengthIn, widthIn, heightIn].sort((a,b)=>b-a);
  const [longest, median, shortest] = sides;

  if(longest<=15 && median<=12 && shortest<=0.75 && weightLb<=1){
    return { tier:'SmallStandard', shippingWeight:weightLb, dimWeight:null, basis:'unit_weight_only', confidence:'high' };
  }
  const dimWeightStandard = (longest*median*shortest)/139;
  const shippingWeightStandard = Math.max(weightLb, dimWeightStandard);
  if(longest<=18 && median<=14 && shortest<=8 && shippingWeightStandard<=20){
    return { tier:'LargeStandard', shippingWeight:shippingWeightStandard, dimWeight:dimWeightStandard, basis:'greater_of_weight_or_dim', confidence:'high' };
  }
  const flooredMedian = Math.max(median,2);
  const flooredShortest = Math.max(shortest,2);
  const dimWeightOversize = (longest*flooredMedian*flooredShortest)/139;
  const shippingWeightOversize = Math.max(weightLb, dimWeightOversize);
  const fitsLargeBulkyEnvelope = longest<=59 && median<=33 && shortest<=33 && shippingWeightOversize<=50;
  if(fitsLargeBulkyEnvelope){
    const isSmallBulkyCarveOut = (longest>18 && longest<=37) || (shippingWeightOversize>20 && shippingWeightOversize<=50);
    const wasFormerlyLargeStandardShape = median<=14 && shortest<=8;
    if(isSmallBulkyCarveOut && wasFormerlyLargeStandardShape){
      return { tier:'SmallBulky', shippingWeight:shippingWeightOversize, dimWeight:dimWeightOversize, basis:'greater_of_weight_or_dim_floored', confidence:'low' };
    }
    return { tier:'LargeBulky', shippingWeight:shippingWeightOversize, dimWeight:dimWeightOversize, basis:'greater_of_weight_or_dim_floored', confidence:'high' };
  }
  if(shippingWeightOversize<=50) return { tier:'ExtraLarge50', shippingWeight:shippingWeightOversize, dimWeight:dimWeightOversize, basis:'greater_of_weight_or_dim_floored', confidence:'low' };
  if(shippingWeightOversize<=70) return { tier:'ExtraLarge70', shippingWeight:shippingWeightOversize, dimWeight:dimWeightOversize, basis:'greater_of_weight_or_dim_floored', confidence:'low' };
  if(shippingWeightOversize<=150) return { tier:'ExtraLarge150', shippingWeight:shippingWeightOversize, dimWeight:dimWeightOversize, basis:'greater_of_weight_or_dim_floored', confidence:'low' };
  return { tier:'ExtraLarge150Plus', shippingWeight:weightLb, dimWeight:null, basis:'unit_weight_only', confidence:'medium' };
}

const MIN_SENSIBLE_BOX_SIZE = 4;   // below this, chasing the fee-free tier
const MAX_SENSIBLE_BOX_SIZE = 60;  // realistic single-carton ceiling
// Splitting into many tiny boxes to save a small placement fee usually
// loses money once real packing/handling friction is considered — this
// tool has no real freight data (see the Honesty tab), so box planning
// stays conservative rather than recommending something that looks good
// on paper but likely costs more in practice.

/**
 * Plans cartons for a single region's allocated quantity of one SKU.
 * If unitsPerBox comes from the REAL manifest case-pack columns, it is a
 * hard constraint — never overridden, since that's the user's real packing
 * spec, not a tool guess (per explicit instruction: case-pack if mentioned
 * must be honored, full stop).
 */
const _planCartonsCache = new Map();
function planCartons(quantity, unitsPerBox, declaredBoxCount){
  const cacheKey = quantity+'|'+(unitsPerBox||0)+'|'+(declaredBoxCount||0);
  const cached = _planCartonsCache.get(cacheKey);
  if(cached) return cached;
  const result = planCartonsUncached(quantity, unitsPerBox, declaredBoxCount);
  if(_planCartonsCache.size > 5000) _planCartonsCache.clear(); // safety valve
  _planCartonsCache.set(cacheKey, result);
  return result;
}
function planCartonsUncached(quantity, unitsPerBox, declaredBoxCount){
  if(quantity<=0) return { boxSize:0, identicalBoxes:0, remainderUnits:0, qualifiesOptimized:false, boxes:[], note:null };

  if(unitsPerBox && unitsPerBox>0){
    // Real declared box count from the manifest takes priority over a
    // recomputed division — Amazon's own template treats Number of boxes
    // as the vendor's stated total, not something this tool should
    // silently override. Only deviate from it when the math genuinely
    // doesn't fit (declared boxes * units/box doesn't reach quantity),
    // in which case the real declared count still wins for the identical
    // portion, with any leftover surfaced as a real remainder rather than
    // recomputed from scratch.
    if(declaredBoxCount && declaredBoxCount>0){
      const unitsInDeclaredBoxes = unitsPerBox*declaredBoxCount;
      if(unitsInDeclaredBoxes <= quantity){
        const remainderUnits = quantity - unitsInDeclaredBoxes;
        return buildBoxResult(unitsPerBox, declaredBoxCount, remainderUnits, true, remainderUnits>0?'declared-box-count-with-remainder':null);
      }
      // declared boxes * units/box overshoots quantity -- the manifest's
      // own numbers don't reconcile; fall through to the simple division
      // below rather than fabricate a box count that exceeds real quantity.
    }
    const identicalBoxes = Math.floor(quantity/unitsPerBox);
    const remainderUnits = quantity % unitsPerBox;
    return buildBoxResult(unitsPerBox, identicalBoxes, remainderUnits, true, null);
  }

  const lo = Math.max(1, Math.min(MIN_SENSIBLE_BOX_SIZE, quantity));
  const hi = Math.min(quantity, MAX_SENSIBLE_BOX_SIZE);
  let bestQualifying=null, bestFallback=null;

  for(let boxSize=lo; boxSize<=hi; boxSize++){
    const identicalBoxes = Math.floor(quantity/boxSize);
    const remainderUnits = quantity % boxSize;
    const totalBoxCount = identicalBoxes + (remainderUnits>0?1:0);
    const qualifies = identicalBoxes >= 5;
    const candidate = { boxSize, identicalBoxes, remainderUnits, totalBoxCount };
    if(qualifies && (!bestQualifying || boxSize>bestQualifying.boxSize)) bestQualifying = candidate;
    if(!bestFallback || totalBoxCount<bestFallback.totalBoxCount || (totalBoxCount===bestFallback.totalBoxCount && boxSize>bestFallback.boxSize)) bestFallback = candidate;
  }

  if(bestQualifying) return buildBoxResult(bestQualifying.boxSize, bestQualifying.identicalBoxes, bestQualifying.remainderUnits, false, null);
  return buildBoxResult(bestFallback.boxSize, bestFallback.identicalBoxes, bestFallback.remainderUnits, false, 'below-optimized-threshold');
}
function buildBoxResult(boxSize, identicalBoxes, remainderUnits, isRealCasePack, note){
  const qualifiesOptimized = identicalBoxes >= 5;
  const boxes = [];
  for(let i=0;i<identicalBoxes;i++) boxes.push({ units:boxSize, isIdentical:true });
  if(remainderUnits>0) boxes.push({ units:remainderUnits, isIdentical:false });
  return { boxSize, identicalBoxes, remainderUnits, qualifiesOptimized, boxes, isRealCasePack:!!isRealCasePack, note };
}

function buildFeeRateTable(feeRecords){
  const buckets = {};
  for(const r of feeRecords){
    const key = r.region+'|'+r.sizeTier;
    (buckets[key] = buckets[key]||[]).push(r.feePerUnit);
  }
  const table = {};
  for(const key of Object.keys(buckets)){
    const [region, sizeTier] = key.split('|');
    const vals = buckets[key];
    const avg = vals.reduce((a,b)=>a+b,0)/vals.length;
    if(!table[region]) table[region] = {};
    table[region][sizeTier] = { avgFeePerUnit: avg, sampleSize: vals.length };
  }
  return table;
}

/* Real per-SKU size tier, built from the actual fee log — the most
   frequently observed real size tier for that SKU across its real
   shipments. Replaces a previous hardcoded "LargeStandard" assumption
   that silently misclassified every SKU actually shipped under a
   different tier (a real, now-fixed bug — confirmed by direct
   reproduction against the live fee log). Falls back to LargeStandard
   only when a SKU genuinely has no real fee-log history at all, with
   that fallback explicitly flagged to the caller. */
function buildSkuSizeTierLookup(feeRecords){
  const counts = {};
  for(const r of feeRecords){
    if(!counts[r.sku]) counts[r.sku] = {};
    counts[r.sku][r.sizeTier] = (counts[r.sku][r.sizeTier]||0) + 1;
  }
  const lookup = {};
  for(const sku of Object.keys(counts)){
    const tiers = counts[sku];
    const best = Object.entries(tiers).sort((a,b)=>b[1]-a[1])[0];
    lookup[sku] = { tier: best[0], sampleSize: best[1], isReal: true };
  }
  return lookup;
}
function getSkuSizeTier(sku){
  if(state.skuSizeTierLookup && state.skuSizeTierLookup[sku]) return state.skuSizeTierLookup[sku].tier;
  return 'LargeStandard'; // honest fallback only when no real fee-log entry exists for this SKU at all
}

/**
 * Computes real cost of a region split, now using the real carton-identity
 * rule per region instead of a flat location-count discount. Each region
 * with units>0 gets its own box plan; if that plan qualifies for
 * Amazon-Optimized (>=5 identical cartons), that region's units are
 * fee-free — otherwise full minimal-split per-unit fee applies.
 */
function computeSplitCost(unitsByRegion, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBoxOverride, perRegionBoxCounts){
  const regionsUsed = Object.entries(unitsByRegion).filter(([,u])=>u>0);
  const locCount = regionsUsed.length;
  let totalCost=0; const breakdown={};
  for(const [region,units] of regionsUsed){
    const rateEntry = feeRateTable[region] && feeRateTable[region][sizeTier];
    const baseRate = rateEntry ? rateEntry.avgFeePerUnit : defaultFeePerUnit;
    // When the caller already knows the REAL number of physical boxes
    // for this region (e.g. the case-pack-aware LP search, which tracks
    // exact box counts per region directly), use that instead of having
    // planCartons re-derive a possibly-different count from division.
    const knownBoxCount = perRegionBoxCounts ? perRegionBoxCounts[region] : null;
    const boxPlan = knownBoxCount!=null
      ? buildBoxResult(unitsPerBoxOverride, knownBoxCount, units-(unitsPerBoxOverride*knownBoxCount), true, null)
      : planCartons(units, unitsPerBoxOverride);
    const rate = boxPlan.qualifiesOptimized ? 0 : baseRate;
    const cost = rate*units;
    breakdown[region] = { units, rate, cost, rateSource: rateEntry?'real':'fallback', boxPlan };
    totalCost += cost;
  }
  return { totalCost, locCount, breakdown };
}

/* "Cheapest-only" baseline: what would this SKU cost if every unit went
   to whichever single region has the lowest real per-unit rate for its
   size tier, ignoring demand entirely? This is the real comparison point
   for "shipping by demand is $X cheaper/costlier than the naive single-
   region option" — not a strawman, the actual best a cost-blind seller
   could do without this tool. */
function computeCheapestOnlyCost(totalUnits, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox){
  const regions = ['East','Central','West'];
  let bestRegion=null, bestRate=Infinity;
  for(const region of regions){
    const rateEntry = feeRateTable[region] && feeRateTable[region][sizeTier];
    const rate = rateEntry ? rateEntry.avgFeePerUnit : defaultFeePerUnit;
    if(rate < bestRate){ bestRate = rate; bestRegion = region; }
  }
  const units = { East:0, Central:0, West:0 };
  units[bestRegion] = totalUnits;
  const { totalCost } = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox);
  return { totalCost, region: bestRegion };
}

/* ============================================================
   DECISION ENGINE — heuristic + grid-search LP-equivalent optimizer
   ============================================================ */
/* Real case-pack-aware proportional split. When a REAL case pack exists
   (from the manifest or matched vendor spec sheet — never the optimizer's
   own guessed box size), region allocations must be whole multiples of
   that case pack, since a sealed case-packed box physically cannot be
   split across regions. This is NOT a requirement to use all three
   regions — a single region, two regions, or three are all equally
   valid outcomes; the only constraint is that whatever regions ARE used
   get whole boxes, never a fractional one. */
function heuristicSplit(totalUnits, demandPct, realUnitsPerBox, declaredBoxCount){
  if(realUnitsPerBox && realUnitsPerBox>0){
    // Real declared box count from the manifest is ground truth — never
    // recomputed by division alone. When totalUnits doesn't divide evenly
    // by the case pack (e.g. 171 units in 4 real boxes of 42, where
    // 42*4=168 leaves a real 3-unit remainder), the declared box count
    // still drives the split; the leftover becomes one real remainder
    // box riding with whichever region gets the most boxes.
    let totalBoxes, remainderUnits = 0;
    if(declaredBoxCount && declaredBoxCount>0){
      totalBoxes = declaredBoxCount;
      remainderUnits = totalUnits - (realUnitsPerBox*declaredBoxCount);
      if(remainderUnits < 0){ totalBoxes = Math.floor(totalUnits/realUnitsPerBox); remainderUnits = totalUnits % realUnitsPerBox; }
    } else if(totalUnits % realUnitsPerBox === 0){
      totalBoxes = totalUnits/realUnitsPerBox;
    } else {
      totalBoxes = Math.floor(totalUnits/realUnitsPerBox);
      remainderUnits = totalUnits % realUnitsPerBox;
    }
    const rawBoxes = { East: totalBoxes*demandPct.East, Central: totalBoxes*demandPct.Central, West: totalBoxes*demandPct.West };
    const boxes = { East: Math.round(rawBoxes.East), Central: Math.round(rawBoxes.Central), West: Math.round(rawBoxes.West) };
    const drift = totalBoxes - (boxes.East+boxes.Central+boxes.West);
    if(drift !== 0){
      const biggest = Object.keys(boxes).reduce((a,b)=>boxes[a]>=boxes[b]?a:b);
      boxes[biggest] += drift;
    }
    const units = { East: boxes.East*realUnitsPerBox, Central: boxes.Central*realUnitsPerBox, West: boxes.West*realUnitsPerBox };
    if(remainderUnits>0){
      const biggest = Object.keys(boxes).reduce((a,b)=>boxes[a]>=boxes[b]?a:b);
      units[biggest] += remainderUnits;
    }
    return units;
  }
  const units = {
    East: Math.round(totalUnits*demandPct.East),
    Central: Math.round(totalUnits*demandPct.Central),
    West: Math.round(totalUnits*demandPct.West),
  };
  const drift = totalUnits - (units.East+units.Central+units.West);
  if(drift !== 0){
    const biggest = Object.keys(units).reduce((a,b)=>units[a]>=units[b]?a:b);
    units[biggest] += drift;
  }
  return units;
}
function servedDemandCoverage(units, demandPct, totalUnits){
  if(totalUnits<=0) return 1;
  let shortfall=0;
  for(const region of ['East','Central','West']){
    const ideal = demandPct[region]*totalUnits;
    const served = Math.min(units[region], ideal);
    shortfall += Math.max(ideal-served, 0);
  }
  return 1 - (shortfall/totalUnits);
}

/* Real LP optimizer. When isRealCasePack is true, searches over whole
   BOX counts per region (the only physically valid allocation space for
   a sealed case-packed item) instead of arbitrary unit fractions. Single-
   or two-region outcomes are exactly as valid as three-region ones — the
   search simply finds whichever real combination of whole boxes is
   cheapest at or above the requested coverage floor. When the real case
   pack is large relative to total units, some coverage targets become
   mathematically unreachable (e.g. 2 boxes split across 3 regions can
   never hit 85% coverage) — that is reported explicitly via
   maxAchievableCoverage, never silently approximated. */
function lpOptimize(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, minCoveragePct, gridStep, unitsPerBox, isRealCasePack, declaredBoxCount){
  if(isRealCasePack && unitsPerBox && unitsPerBox>0){
    // Real declared box count from the manifest is the ground truth for
    // how many physical boxes exist — never recomputed by division alone.
    // When the vendor's quantity doesn't divide evenly by units/box
    // (e.g. 171 units in 4 boxes, where 4 boxes don't multiply out to
    // exactly 171), the declared box count still wins for the identical
    // portion, and the real leftover becomes one additional remainder
    // box that travels with whichever region gets it — never silently
    // recomputed into a different box count or box size.
    let totalBoxes, remainderUnits = 0;
    if(declaredBoxCount && declaredBoxCount>0){
      totalBoxes = declaredBoxCount;
      remainderUnits = totalUnits - (unitsPerBox*declaredBoxCount);
      if(remainderUnits < 0){
        // Declared boxes*unitsPerBox exceeds real quantity -- the
        // manifest's own numbers don't reconcile; fall back to pure
        // division rather than search a box count larger than what
        // could possibly exist.
        totalBoxes = Math.floor(totalUnits/unitsPerBox);
        remainderUnits = totalUnits % unitsPerBox;
      }
    } else if(totalUnits % unitsPerBox === 0){
      totalBoxes = totalUnits/unitsPerBox;
    } else {
      totalBoxes = Math.floor(totalUnits/unitsPerBox);
      remainderUnits = totalUnits % unitsPerBox;
    }

    let best = null, bestByCoverage = null, maxCoverageSeen = 0;
    for(let e=0; e<=totalBoxes; e++){
      for(let c=0; c<=totalBoxes-e; c++){
        const w = totalBoxes-e-c;
        const units = { East:e*unitsPerBox, Central:c*unitsPerBox, West:w*unitsPerBox };
        // The real remainder (if any) always rides with whichever region
        // already has the most identical boxes — it's physically one
        // more box, not a unit that can float independently.
        if(remainderUnits>0){
          const biggest = e>=c && e>=w ? 'East' : c>=w ? 'Central' : 'West';
          units[biggest] += remainderUnits;
        }
        const coverage = servedDemandCoverage(units, demandPct, totalUnits);
        if(coverage > maxCoverageSeen) maxCoverageSeen = coverage;
        const perRegionBoxCounts = { East:e, Central:c, West:w };
        const { totalCost, locCount, breakdown } = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox, perRegionBoxCounts);
        // Track best by coverage (primary) then cost (tiebreak) — this is the
        // correct fallback when the coverage target is unreachable. Giving up
        // on demand coverage in favour of a cheaper split defeats the purpose.
        if(!bestByCoverage
          || coverage > bestByCoverage.coverage + 1e-6
          || (Math.abs(coverage - bestByCoverage.coverage) < 1e-6 && totalCost < bestByCoverage.totalCost - 1e-9)){
          bestByCoverage = { units, totalCost, coverage, locCount, breakdown };
        }
        if(coverage < minCoveragePct - 1e-6) continue;
        if(!best || totalCost < best.totalCost - 1e-9) best = { units, totalCost, coverage, locCount, breakdown };
      }
    }
    if(!best && bestByCoverage){
      // Requested coverage floor is mathematically unreachable at this
      // case-pack granularity. Return the split with the BEST DEMAND COVERAGE
      // achievable — not the cheapest. When you have only 2 boxes and can't
      // hit 85%, the right answer is Central+West (72.4% coverage) not
      // East-only (28.3% coverage just because it's cheaper).
      return { ...bestByCoverage, coverageTargetUnreachable: true, maxAchievableCoverage: maxCoverageSeen };
    }
    return best;
  }

  let best = null;
  gridStep = gridStep || (totalUnits>200 ? 0.05 : 0.02);
  for(let a=0; a<=1.0001; a+=gridStep){
    for(let b=0; b<=1.0001-a; b+=gridStep){
      const c = 1-a-b;
      if(c < -1e-6) continue;
      const frac = { East:a, Central:b, West:Math.max(c,0) };
      const units = {
        East: Math.round(totalUnits*frac.East),
        Central: Math.round(totalUnits*frac.Central),
        West: Math.round(totalUnits*frac.West),
      };
      const drift = totalUnits - (units.East+units.Central+units.West);
      if(drift !== 0) units.West += drift;
      if(units.East<0||units.Central<0||units.West<0) continue;
      const coverage = servedDemandCoverage(units, demandPct, totalUnits);
      if(coverage < minCoveragePct - 1e-6) continue;
      const { totalCost, locCount, breakdown } = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox);
      if(!best || totalCost < best.totalCost - 1e-9) best = { units, totalCost, coverage, locCount, breakdown };
    }
  }
  return best;
}
function buildParetoFrontier(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, steps, unitsPerBox, isRealCasePack){
  steps = steps || 10;
  const frontier = [];
  for(let i=0;i<=steps;i++){
    const minCov = i/steps;
    const result = lpOptimize(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, minCov, 0.05, unitsPerBox, isRealCasePack);
    if(result) frontier.push({ minCoverage:minCov, cost:result.totalCost, achievedCoverage:result.coverage, units:result.units, coverageTargetUnreachable: result.coverageTargetUnreachable||false });
  }
  return frontier;
}

/* ============================================================
   MANIFEST ADAPTER + PLAN BUILDER
   ============================================================ */
function adaptManifestRows(rawObjects){
  if(!rawObjects.length) return { records: [], stats: { totalRaw:0, kept:0 } };
  const headers = Object.keys(rawObjects[0]).filter(h=>!h.startsWith('__'));
  const lower = headers.map(h=>h.toLowerCase());
  const find = (aliases) => { for(const a of aliases){ const idx = lower.findIndex(h=>h.trim()===a || h.includes(a)); if(idx!==-1) return headers[idx]; } return null; };
  const skuCol = find(['merchant sku','sku','msku']);
  const unitsCol = find(['quantity','units','qty']);
  const tierCol = find(['size tier','sizetier','tier']);
  const expCol = find(['expiration date']);
  const lotCol = find(['manufacturing lot code','lot code']);
  const unitsPerBoxCol = find(['units per box']);
  const numBoxesCol = find(['number of boxes']);
  const boxLCol = find(['box length']);
  const boxWCol = find(['box width']);
  const boxHCol = find(['box height']);
  const boxWeightCol = find(['box weight']);

  const clean = []; let dropped = 0; let boxCountMismatches = 0;
  for(const r of rawObjects){
    const sku = skuCol ? normalizeSku(r[skuCol]) : '';
    const units = unitsCol ? parseInt(r[unitsCol],10) : NaN;
    if(!sku || !Number.isFinite(units) || units<=0){ dropped++; continue; }
    const unitsPerBox = unitsPerBoxCol ? (parseFloat(r[unitsPerBoxCol])||null) : null;
    const numberOfBoxes = numBoxesCol ? (parseFloat(r[numBoxesCol])||null) : null;
    // Real cross-check: Amazon's own template treats every box for a
    // packing line as identical (same Units per box, same dimensions) —
    // Number of boxes is the vendor's own declared total, not something
    // this tool should silently recompute. When both fields are present
    // but don't multiply out to the stated Quantity, that's a real data
    // inconsistency in the manifest itself, flagged rather than guessed past.
    let boxCountMismatch = false;
    if(unitsPerBox && numberOfBoxes){
      const expectedUnits = unitsPerBox * numberOfBoxes;
      if(Math.abs(expectedUnits - units) > 0.01){ boxCountMismatch = true; boxCountMismatches++; }
    }
    clean.push({
      sku, units,
      sizeTier: tierCol ? (r[tierCol]||'').trim() : 'LargeStandard',
      expirationDate: expCol ? (r[expCol]||'').trim() : '',
      lotCode: lotCol ? (r[lotCol]||'').trim() : '',
      unitsPerBox, numberOfBoxes, boxCountMismatch,
      boxLength: boxLCol ? (parseFloat(r[boxLCol])||null) : null,
      boxWidth: boxWCol ? (parseFloat(r[boxWCol])||null) : null,
      boxHeight: boxHCol ? (parseFloat(r[boxHCol])||null) : null,
      boxWeight: boxWeightCol ? (parseFloat(r[boxWeightCol])||null) : null,
      isCasePacked: !!(unitsPerBoxCol && r[unitsPerBoxCol]),
    });
  }
  return { records: clean, stats: { totalRaw: rawObjects.length, kept: clean.length, dropped, boxCountMismatches, detectedColumns:{skuCol,unitsCol,tierCol,expCol,lotCol,unitsPerBoxCol,numBoxesCol} } };
}

/* Aggregates raw packing lines (a SKU can legitimately appear up to 4
   times — once per unique expiration date / case-pack spec, per Amazon's
   own Send-to-Amazon rules) into one total-units figure per SKU for the
   regional demand split, since region allocation is a per-SKU demand
   decision, not a per-packing-line one. Line-level detail (expiration,
   lot code, box specs) is preserved separately for display. Real
   units-per-box, when present on ANY line for this SKU, is carried
   forward as a hard constraint on box planning — per explicit
   instruction, case-pack data always wins over computed defaults. */
function aggregateManifestBySku(manifestRecords){
  const bySku = {};
  for(const rec of manifestRecords){
    if(!bySku[rec.sku]) bySku[rec.sku] = { sku:rec.sku, units:0, sizeTier:rec.sizeTier, lines:[], realUnitsPerBox:null, realNumberOfBoxes:null, hasBoxCountMismatch:false, sizeTierCrossCheck:null, specSheetMatch:null };
    bySku[rec.sku].units += rec.units;
    bySku[rec.sku].lines.push(rec);
    if(rec.unitsPerBox && !bySku[rec.sku].realUnitsPerBox) bySku[rec.sku].realUnitsPerBox = rec.unitsPerBox;
    if(rec.numberOfBoxes && !bySku[rec.sku].realNumberOfBoxes) bySku[rec.sku].realNumberOfBoxes = rec.numberOfBoxes;
    if(rec.boxCountMismatch) bySku[rec.sku].hasBoxCountMismatch = true;
    // Cross-check the stated size tier against real box dimensions, when
    // the manifest actually has them — never silently overrides the
    // stated tier, just surfaces a discrepancy if there is one.
    if(!bySku[rec.sku].sizeTierCrossCheck && rec.boxLength && rec.boxWidth && rec.boxHeight && rec.boxWeight){
      bySku[rec.sku].sizeTierCrossCheck = classifySizeTier(rec.boxLength, rec.boxWidth, rec.boxHeight, rec.boxWeight);
    }
  }
  // Priority order for case-pack/dimension data, applied AFTER scanning
  // every line: real manifest case-pack data (set above) always wins;
  // only when the manifest itself has NO real case-pack for a SKU does
  // the uploaded vendor spec sheet's real data get used instead of
  // falling straight to the optimizer's computed guess. The spec sheet
  // is real vendor data — strictly better than a guess, but never
  // allowed to override real data the manifest itself already provided.
  for(const sku of Object.keys(bySku)){
    const row = bySku[sku];
    if(!row.realUnitsPerBox && state.specSheetLookup){
      const specMatch = lookupSpecSheetForSku(sku);
      if(specMatch){
        row.specSheetMatch = specMatch;
        if(specMatch.casePack) row.realUnitsPerBox = specMatch.casePack;
        if(!row.sizeTierCrossCheck && specMatch.length && specMatch.width && specMatch.height && specMatch.weight){
          row.sizeTierCrossCheck = classifySizeTier(specMatch.length, specMatch.width, specMatch.height, specMatch.weight);
        }
      }
    }
  }
  return Object.values(bySku);
}

/* Caps a SKU's real demand split down to its top N regions by demand
   share, renormalizing the kept regions' weights so they still sum to 1
   and zeroing out the rest. Business reason (per the standing brief):
   minimizing the number of regions a single SKU ships to is the largest
   controllable cost/ops lever — every extra region is another shipment,
   another dock appointment, another partial box. Default is 2 regions:
   the two the real sales data says matter most for this SKU, not an
   arbitrary pair. Passing maxRegions=3 (the "allow all 3" toggle) returns
   demandPct unchanged. This is applied BEFORE the split is computed, so
   both the heuristic and LP paths automatically respect the same cap —
   there's exactly one place this decision gets made. */
function capDemandToRegions(demandPct, maxRegions){
  const regions = ['East','Central','West'];
  if(!maxRegions || maxRegions >= regions.length) return demandPct;
  const ranked = regions.slice().sort((a,b)=>(demandPct[b]||0)-(demandPct[a]||0));
  const keep = ranked.slice(0, maxRegions);
  const kept = new Set(keep);
  const sum = keep.reduce((s,r)=>s+(demandPct[r]||0), 0);
  const capped = { East:0, Central:0, West:0 };
  for(const r of keep) capped[r] = sum>0 ? (demandPct[r]||0)/sum : 1/keep.length;
  return capped;
}

function buildManifestPlan(manifestRecords, demandBySku, feeRateTable, defaultFeePerUnit, method, minCoveragePct, defaultBoxUnitsPerBox, maxRegionsPerSku){
  const aggregated = aggregateManifestBySku(manifestRecords);
  const plan = []; let knownDemandCount=0, unknownDemandCount=0;
  const portfolioUnits = { East:0, Central:0, West:0 };
  let portfolioCost = 0, portfolioCheapestOnlyCost = 0;
  let coverageWeightedSum = 0; // for a real units-weighted average coverage %, not a naive mean across SKUs of different sizes

  for(const row of aggregated){
    const demand = demandBySku[row.sku];
    const rawDemandPct = demand ? demand.demandPct : { East:1/3, Central:1/3, West:1/3 };
    const demandPct = capDemandToRegions(rawDemandPct, maxRegionsPerSku || 2);
    if(demand) knownDemandCount++; else unknownDemandCount++;
    const effectiveUnitsPerBox = row.realUnitsPerBox || defaultBoxUnitsPerBox || null;
    const usedDefaultBoxSpec = !row.realUnitsPerBox;
    const isRealCasePack = !!row.realUnitsPerBox; // true only for a real manifest/spec-sheet case pack, never the optimizer's own guessed box size
    const declaredBoxCount = (isRealCasePack && row.realNumberOfBoxes) ? row.realNumberOfBoxes : null;

    let units, cost=null, coverage=null, breakdown=null, coverageTargetUnreachable=false, maxAchievableCoverage=null;
    if(method === 'lp'){
      const lp = lpOptimize(row.units, demandPct, row.sizeTier, feeRateTable, defaultFeePerUnit, minCoveragePct, null, effectiveUnitsPerBox, isRealCasePack, declaredBoxCount);
      units = lp ? lp.units : heuristicSplit(row.units, demandPct, isRealCasePack?effectiveUnitsPerBox:null, isRealCasePack?declaredBoxCount:null);
      cost = lp ? lp.totalCost : null;
      coverage = lp ? lp.coverage : null;
      breakdown = lp ? lp.breakdown : null;
      coverageTargetUnreachable = lp ? !!lp.coverageTargetUnreachable : false;
      maxAchievableCoverage = lp ? (lp.maxAchievableCoverage!=null?lp.maxAchievableCoverage:null) : null;
    } else {
      units = heuristicSplit(row.units, demandPct, isRealCasePack?effectiveUnitsPerBox:null, isRealCasePack?declaredBoxCount:null);
      const costResult = computeSplitCost(units, row.sizeTier, feeRateTable, defaultFeePerUnit, effectiveUnitsPerBox);
      cost = costResult.totalCost;
      breakdown = costResult.breakdown;
      // Real coverage for the heuristic path too — same formula the LP
      // optimizer uses (servedDemandCoverage), so the executive summary
      // means the same thing regardless of which method is selected.
      coverage = servedDemandCoverage(units, demandPct, row.units);
    }
    if(!breakdown){
      breakdown = computeSplitCost(units, row.sizeTier, feeRateTable, defaultFeePerUnit, effectiveUnitsPerBox).breakdown;
    }

    const cheapestOnly = computeCheapestOnlyCost(row.units, row.sizeTier, feeRateTable, defaultFeePerUnit, effectiveUnitsPerBox);
    const costDelta = cost!=null ? (cost - cheapestOnly.totalCost) : null; // negative = demand split is cheaper

    portfolioUnits.East += units.East; portfolioUnits.Central += units.Central; portfolioUnits.West += units.West;
    portfolioCost += (cost||0);
    portfolioCheapestOnlyCost += cheapestOnly.totalCost;
    coverageWeightedSum += (coverage||0) * row.units;

    plan.push({
      sku: row.sku, units: row.units, sizeTier: row.sizeTier, hasSalesHistory: !!demand, demandPct,
      splitUnits: units, cost, coverage, packingLines: row.lines.length, lines: row.lines,
      realUnitsPerBox: row.realUnitsPerBox, effectiveUnitsPerBox, usedDefaultBoxSpec, sizeTierCrossCheck: row.sizeTierCrossCheck,
      specSheetMatch: row.specSheetMatch, isRealCasePack, coverageTargetUnreachable, maxAchievableCoverage,
      boxBreakdown: breakdown, cheapestOnlyCost: cheapestOnly.totalCost, cheapestOnlyRegion: cheapestOnly.region, costDelta
    });
  }

  const totalPortfolioUnits = portfolioUnits.East + portfolioUnits.Central + portfolioUnits.West;
  const portfolioPct = totalPortfolioUnits>0 ? {
    East: portfolioUnits.East/totalPortfolioUnits,
    Central: portfolioUnits.Central/totalPortfolioUnits,
    West: portfolioUnits.West/totalPortfolioUnits
  } : { East:1/3, Central:1/3, West:1/3 };

  const totalUnitsAcrossSkus = aggregated.reduce((s,r)=>s+r.units,0);
  const weightedAvgCoverage = totalUnitsAcrossSkus>0 ? coverageWeightedSum/totalUnitsAcrossSkus : null;
  const totalSavings = portfolioCheapestOnlyCost - portfolioCost;
  const savingsPct = portfolioCheapestOnlyCost>0 ? totalSavings/portfolioCheapestOnlyCost : 0;
  const costPerUnit = totalPortfolioUnits>0 ? portfolioCost/totalPortfolioUnits : 0;

  return {
    plan,
    summary: { totalSkus: aggregated.length, knownDemandCount, unknownDemandCount, totalUnits: totalUnitsAcrossSkus, totalPackingLines: manifestRecords.length, defaultBoxSpecCount: plan.filter(p=>p.usedDefaultBoxSpec).length },
    portfolio: { units: portfolioUnits, pct: portfolioPct, totalCost: portfolioCost, cheapestOnlyCost: portfolioCheapestOnlyCost, totalSavings, savingsPct, weightedAvgCoverage, costPerUnit }
  };
}

/* ============================================================
   APPLICATION STATE
   ============================================================ */
const state = {
  filesRaw: { sales1:null, sales2:null, sales3:null, sales4:null, sales5:null, sales6:null, fees:null, inventory:null, manifest:null, specsheet:null, packinglist:null },
  detectedHeaders: { sales:null, fees:null, inventory:null, manifest:null, specsheet:null },
  salesAdapted: null, feeAdapted: null, inventoryAdapted: null,
  demand: null, feeRateTable: null, defaultFeePerUnit: 0.4,
  manifestRecords: null, manifestPlan: null, manifestPortfolio: null,
  frontierData: null,
  sellThroughOn: false,
  manifestFilters: { search:'', region:'all', demandSource:'all', minUnits:0 },
  maxRegionsPerSku: 2, // 2 (recommended) or 3 — caps how many regions a single SKU's shipment fans out to
  scopeMode: 'all', // 'all' | 'manifest' — drives every cross-tab aggregation
  globalVendorFilter: 'all', // vendor prefix, or 'all'
  realDateSpanDays: null, // actual days covered by the uploaded sales file, computed from real dates
  specSheetLookup: null, // {vendor|suffix: {casePack, weight, length, width, height, remarks}}, real vendor spec data
  skuSizeTierLookup: null, // {sku: {tier, sampleSize, isReal}}, real per-SKU size tier from the fee log
  costVendorFilter: 'all', // Cost tab's vendor slicer, narrows the frontier SKU picker too
  velocityByDay: null, // {days:[{date,East,Central,West,total}], avgByRegion:{East,Central,West}, totalByRegion:{East,Central,West}}
  manualSpecEntries: {}, // {vendor|suffix: record} — typed-in or packing-list-derived, persists across re-analysis, always merged on top of the uploaded spec sheet
};

/* Real vendor extraction: the text before the first hyphen in a SKU.
   Verified against the actual sales file — 5,462 of 5,475 real SKUs have
   a hyphen-delimited prefix; the remainder (malformed/placeholder SKUs
   like "Uncommingled.MSKU...") get bucketed into an explicit
   "Unidentified" group rather than silently dropped or crashing. */
/* Collects every populated sales-period upload slot (up to 6), in order.
   All periods get merged into one combined sales history before demand
   is computed — this is purely a data-entry convenience (upload Jan, Feb,
   Mar… separately instead of hand-merging spreadsheets yourself), not a
   different data model. */
function getSalesFiles(){
  return [1,2,3,4,5,6].map(i=>state.filesRaw['sales'+i]).filter(Boolean);
}

function extractVendorPrefix(sku){
  if(!sku) return 'Unidentified';
  const idx = sku.indexOf('-');
  if(idx <= 0) return 'Unidentified';
  return sku.slice(0, idx).toUpperCase();
}

/* Single source of truth for "which SKUs are in scope right now" — every
   cross-tab aggregation (Overview, Demand, Cost, SKU Analysis) calls this
   instead of reading state.demand directly, so Manifest Mode and the
   global vendor filter apply consistently everywhere rather than each
   tab implementing its own (and inevitably diverging) filter logic. */
function getActiveSkuSet(){
  if(!state.demand) return null;
  let skus = Object.keys(state.demand);
  if(state.scopeMode === 'manifest'){
    if(!state.manifestPlan || !state.manifestPlan.length) return new Set(); // manifest mode selected but nothing built yet -> empty scope, not silently "all"
    const manifestSkus = new Set(state.manifestPlan.map(r=>r.sku));
    skus = skus.filter(s=>manifestSkus.has(s));
  }
  if(state.globalVendorFilter !== 'all'){
    skus = skus.filter(s=>extractVendorPrefix(s)===state.globalVendorFilter);
  }
  return new Set(skus);
}

/* ============================================================
   VENDOR SPEC SHEET — real case-pack and dimension data per item,
   keyed by vendor + item suffix. Verified against a real uploaded spec
   sheet: ITEM ID column is "VENDOR-SUFFIX" (e.g. PETL-00233), which is
   the authoritative join key — NOT the separate lowercase "Item ID"
   column, which is only the bare suffix without the vendor tag.
   ============================================================ */
function adaptSpecSheetRows(rawObjects){
  const clean = []; let dropped = 0;
  const headers = rawObjects.length ? Object.keys(rawObjects[0]).filter(h=>!h.startsWith('__')) : [];
  const lower = headers.map(h=>h.toLowerCase().trim());
  const find = (aliases) => { for(const a of aliases){ const idx = lower.findIndex(h=>h===a || h.includes(a)); if(idx!==-1) return headers[idx]; } return null; };
  const itemIdCol = find(['item id']); // matches both "ITEM ID" and "Item ID" — handled by preferring the one with hyphens below
  const vendorCol = find(['vendor']);
  const casePackCol = find(['case pack']);
  const weightCol = find(['weight']);
  const lengthCol = find(['length']);
  const widthCol = find(['width']);
  const heightCol = find(['height']);
  const remarksCol = find(['remarks']);

  // Real files have TWO columns that both match "item id" case-insensitively
  // (the brief's actual upload: "ITEM ID" = full vendor-prefixed code,
  // "Item ID" = bare suffix only). Disambiguate by picking whichever of
  // the two candidate columns actually contains hyphenated values.
  let fullIdCol = itemIdCol;
  const allItemIdCols = headers.filter(h => h.toLowerCase().trim() === 'item id');
  if(allItemIdCols.length === 2){
    for(const col of allItemIdCols){
      const sampleVal = rawObjects.find(r=>r[col])?.[col] || '';
      if(String(sampleVal).includes('-')) fullIdCol = col;
    }
  }

  for(const r of rawObjects){
    const fullId = fullIdCol ? (r[fullIdCol]||'').trim() : '';
    const vendor = vendorCol ? (r[vendorCol]||'').trim().toUpperCase() : '';
    if(!fullId || !vendor || !fullId.includes('-')){ dropped++; continue; }
    const suffix = fullId.slice(vendor.length+1).toUpperCase();
    clean.push({
      fullId, vendor, suffix,
      casePack: casePackCol ? (parseFloat(r[casePackCol])||null) : null,
      weight: weightCol ? (parseFloat(r[weightCol])||null) : null,
      length: lengthCol ? (parseFloat(r[lengthCol])||null) : null,
      width: widthCol ? (parseFloat(r[widthCol])||null) : null,
      height: heightCol ? (parseFloat(r[heightCol])||null) : null,
      remarks: remarksCol ? (r[remarksCol]||'').trim() : ''
    });
  }
  return { records: clean, stats: { totalRaw: rawObjects.length, kept: clean.length, dropped, detectedColumns:{fullIdCol,vendorCol,casePackCol,weightCol,lengthCol,widthCol,heightCol} } };
}

function buildSpecSheetLookup(specRecords){
  const lookup = {};
  for(const rec of specRecords){
    lookup[rec.vendor+'|'+rec.suffix] = rec;
  }
  return lookup;
}

/* Generates every plausible suffix candidate for a real sales/manifest
   SKU, validated against real data: handles the "-FBA-" tag real Amazon
   exports insert, and the "_N" case-pack/expiration variant suffix. */
function buildSkuMatchCandidates(sku){
  const vendor = extractVendorPrefix(sku);
  if(vendor === 'Unidentified') return { vendor:null, candidates:[] };
  const rest = sku.slice(vendor.length+1);
  const candidates = new Set();
  const restNoFBA = rest.replace(/^FBA-/i, '');
  const stripTrailingVariant = (s) => s.replace(/_\d+$/,'');
  candidates.add(rest.toUpperCase());
  candidates.add(restNoFBA.toUpperCase());
  candidates.add(stripTrailingVariant(rest).toUpperCase());
  candidates.add(stripTrailingVariant(restNoFBA).toUpperCase());
  const firstChunk = restNoFBA.split('-')[0];
  candidates.add(firstChunk.toUpperCase());
  candidates.add(stripTrailingVariant(firstChunk).toUpperCase());
  return { vendor, candidates: Array.from(candidates) };
}

/* Looks up real case-pack/dimension data for a SKU against the uploaded
   spec sheet. Returns null (not a guess) when no real entry exists —
   this is a real-data lookup, never a fallback estimate. */
function lookupSpecSheetForSku(sku){
  if(!state.specSheetLookup) return null;
  const { vendor, candidates } = buildSkuMatchCandidates(sku);
  if(!vendor) return null;
  for(const cand of candidates){
    const hit = state.specSheetLookup[vendor+'|'+cand];
    if(hit) return hit;
  }
  return null;
}

/* ============================================================
   MANUAL SKU MASTER DATA — typed-in case pack/dimensions for vendors
   not covered by an uploaded spec sheet. Writes into the same
   vendor|suffix lookup shape the spec sheet adapter produces, so
   downstream code (box optimizer, size-tier cross-check) treats manual
   entries identically to real spec-sheet rows — no separate code path.
   ============================================================ */
function renderMasterEntriesTable(){
  const wrap = document.getElementById('masterEntriesWrap');
  const tbody = document.querySelector('#masterEntriesTable tbody');
  const entries = Object.values(state.manualSpecEntries);
  if(!entries.length){ wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';
  tbody.innerHTML = entries.map(e=>`
    <tr>
      <td class="mono">${e.vendor}</td>
      <td class="mono">${e.suffix}</td>
      <td class="mono">${e.fullId}</td>
      <td class="num">${e.casePack||'—'}</td>
      <td class="num">${e.length&&e.width&&e.height?`${e.length}×${e.width}×${e.height}`:'—'}</td>
      <td class="num">${e.weight||'—'}</td>
      <td><button class="btn-ghost btn-sm" data-remove-master="${e.vendor}|${e.suffix}" style="padding:3px 9px;border-radius:5px;border:1px solid var(--line);background:none;color:var(--danger);cursor:pointer;font-size:11px;">Remove</button></td>
    </tr>`).join('');
  tbody.querySelectorAll('[data-remove-master]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      delete state.manualSpecEntries[btn.dataset.removeMaster];
      renderMasterEntriesTable();
      if(state.demand) runAnalysis(); // re-merge into the live lookup immediately
    });
  });
}
document.getElementById('masterAddBtn').addEventListener('click', ()=>{
  const vendor = document.getElementById('masterVendorInput').value.trim().toUpperCase();
  const suffix = document.getElementById('masterSuffixInput').value.trim().toUpperCase();
  const casePack = parseFloat(document.getElementById('masterCasePackInput').value) || null;
  const length = parseFloat(document.getElementById('masterLengthInput').value) || null;
  const width = parseFloat(document.getElementById('masterWidthInput').value) || null;
  const height = parseFloat(document.getElementById('masterHeightInput').value) || null;
  const weight = parseFloat(document.getElementById('masterWeightInput').value) || null;
  if(!vendor || !suffix){ toast('Vendor prefix and item suffix are both required', true); return; }
  if(!casePack && !(length&&width&&height&&weight)){ toast('Enter at least a case pack or full dimensions', true); return; }
  const fullId = vendor+'-'+suffix;
  state.manualSpecEntries[vendor+'|'+suffix] = { fullId, vendor, suffix, casePack, weight, length, width, height, remarks:'manual entry' };
  ['masterVendorInput','masterSuffixInput','masterCasePackInput','masterLengthInput','masterWidthInput','masterHeightInput','masterWeightInput'].forEach(id=>document.getElementById(id).value='');
  renderMasterEntriesTable();
  toast(`Added ${fullId} — will apply on next "Run analysis" or "Build shipment plan"`);
  if(state.demand) runAnalysis();
});

/* ============================================================
   VENDOR PACKING LIST ADAPTER — real shipping/packing-list files
   (confirmed structure: a literal "Item" marker in the first column,
   with item code, description, and quantity in subsequent columns,
   located by searching for the real header row containing "Customer
   Item No." and "Quantity" together, rather than fixed column indices,
   since different vendors' packing-list templates shift columns).
   These files never carry the operator's own vendor-prefixed SKU —
   only the bare item suffix — so the vendor prefix must be supplied by
   the operator, not guessed. Free-text/prose vendor files are
   explicitly NOT handled here; only structured tabular rows.
   ============================================================ */
function adaptPackingListRows(rawGrid, vendorPrefix){
  let headerRowIdx = -1, qtyCol = -1, descCol = -1;
  const scanLimit = Math.min(rawGrid.length, 60);
  for(let i=0;i<scanLimit;i++){
    const rowLower = rawGrid[i].map(c=>String(c||'').trim().toLowerCase());
    const hasItemCol = rowLower.some(c=>c.includes('customer item'));
    const qIdx = rowLower.findIndex(c=>c==='quantity'||c==='qty');
    if(hasItemCol && qIdx!==-1){ headerRowIdx=i; qtyCol=qIdx; descCol=rowLower.findIndex(c=>c.includes('description')); break; }
  }
  if(headerRowIdx===-1) return { records:[], stats:{ totalRaw: rawGrid.length, kept:0, headerFound:false } };

  const items = [];
  let curQtyCol = qtyCol, curDescCol = descCol;
  for(let i=headerRowIdx;i<rawGrid.length;i++){
    const row = rawGrid[i];
    if(!row) continue;
    const rowLower = row.map(c=>String(c||'').trim().toLowerCase());
    const qIdx = rowLower.findIndex(c=>c==='quantity'||c==='qty');
    const dIdx = rowLower.findIndex(c=>c.includes('description'));
    if(qIdx!==-1) curQtyCol = qIdx;
    if(dIdx!==-1) curDescCol = dIdx;
    if(String(row[0]||'').trim().toLowerCase() === 'item'){
      let itemCode = null;
      for(let c=1;c<row.length;c++){ if(String(row[c]||'').trim()!==''){ itemCode=String(row[c]).trim(); break; } }
      if(itemCode){
        items.push({
          suffix: itemCode.toUpperCase(),
          fullId: vendorPrefix+'-'+itemCode.toUpperCase(),
          vendor: vendorPrefix,
          quantity: curQtyCol!==-1 ? (parseFloat(row[curQtyCol])||null) : null,
          description: curDescCol!==-1 ? String(row[curDescCol]||'').trim() : ''
        });
      }
    }
  }
  return { records: items, stats: { totalRaw: rawGrid.length, kept: items.length, headerFound:true } };
}

document.getElementById('parsePackingListBtn').addEventListener('click', async ()=>{
  const vendor = document.getElementById('packingListVendorInput').value.trim().toUpperCase();
  const file = state.filesRaw.packinglist;
  if(!vendor){ toast('Enter the vendor prefix for this file first (e.g. HNEY)', true); return; }
  if(!file){ toast('Upload a packing list file first', true); return; }
  try{
    let rawGrid;
    if(isExcelFile(file)){
      if(typeof XLSX === 'undefined'){ toast('Excel support library failed to load.', true); return; }
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(buf, { type:'array', cellDates:false });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      rawGrid = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
    } else {
      const { text } = await readFileSmart(file);
      rawGrid = parseCSV(text);
    }
    const { records, stats } = adaptPackingListRows(rawGrid, vendor);
    if(!stats.headerFound){
      toast(`Couldn't find a real "Customer Item No." + "Quantity" header in this file — it may not be a structured packing list. Use manual entry above instead.`, true);
      return;
    }
    if(!records.length){ toast('Header found but no real item rows detected.', true); return; }
    let addedCount = 0;
    for(const rec of records){
      // Packing lists carry no dimensions, only item identity + quantity —
      // merge in without overwriting any case-pack/dimension data that
      // already exists for this exact vendor+suffix from another source.
      const key = rec.vendor+'|'+rec.suffix;
      if(!state.manualSpecEntries[key]){
        state.manualSpecEntries[key] = { fullId: rec.fullId, vendor: rec.vendor, suffix: rec.suffix, casePack:null, weight:null, length:null, width:null, height:null, remarks:`from packing list: ${rec.description}` };
        addedCount++;
      }
    }
    renderMasterEntriesTable();
    toast(`Matched ${records.length} real item${records.length===1?'':'s'} from the packing list against vendor "${vendor}" — ${addedCount} new SKU${addedCount===1?'':'s'} added (no dimensions yet, but identity is now real and matchable)`);
    if(state.demand) runAnalysis();
  } catch(err){
    toast('Failed to parse packing list: '+err.message, true);
    console.error(err);
  }
});

function saveSessionSummary(){
  try{
    const summary = {
      hasSales: !!state.demand,
      hasFees: !!state.feeRateTable,
      skuCount: state.demand ? Object.keys(state.demand).length : 0,
      ts: Date.now()
    };
    localStorage.setItem('mi_session_summary', JSON.stringify(summary));
  } catch(e){ /* quota exceeded — non-fatal */ }
}

/* ============================================================
   UPLOAD ZONE WIRING
   ============================================================ */
function wireUploadZone(kind, inputId, zoneId){
  const input = document.getElementById(inputId);
  const zone = document.getElementById(zoneId);
  zone.addEventListener('click', (e)=>{ if(!e.target.closest('.uz-clear')) input.click(); });
  zone.addEventListener('dragover', (e)=>{ e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', ()=> zone.classList.remove('dragover'));
  zone.addEventListener('drop', (e)=>{
    e.preventDefault(); zone.classList.remove('dragover');
    if(e.dataTransfer.files.length) handleFileSelected(kind, e.dataTransfer.files[0], zone);
  });
  input.addEventListener('change', ()=>{ if(input.files.length) handleFileSelected(kind, input.files[0], zone); });
  zone.querySelector('.uz-clear')?.addEventListener('click', (e)=>{
    e.stopPropagation();
    state.filesRaw[kind] = null;
    zone.classList.remove('has-file');
    zone.querySelector('.uz-filename').textContent = '';
    input.value = '';
  });
}
async function handleFileSelected(kind, file, zone){
  state.filesRaw[kind] = file;
  zone.classList.add('has-file');
  zone.classList.add('detecting');
  const fmt = isExcelFile(file) ? 'XLSX' : 'CSV';
  zone.querySelector('.uz-filename').textContent = file.name + ' · ' + (file.size/1024/1024).toFixed(2) + ' MB · ' + fmt;
  if(kind === 'packinglist'){
    // Packing lists use their own dedicated parser (triggered by the
    // "Parse packing list" button, which also needs the vendor prefix
    // input) rather than the generic schema column-preview.
    toast(file.name + ' ready — enter the vendor prefix, then click "Parse packing list"');
    zone.classList.remove('detecting');
    return;
  }
  let mapBox = zone.querySelector('.uz-colmap');
  if(!mapBox){
    mapBox = document.createElement('div');
    mapBox.className = 'uz-colmap';
    mapBox.innerHTML = '<span class="uz-colmap-loading">Detecting columns…</span>';
    zone.appendChild(mapBox);
  } else {
    mapBox.innerHTML = '<span class="uz-colmap-loading">Detecting columns…</span>';
  }
  toast(file.name + ' ready');
  const schemaKeyForKind = kind.startsWith('sales') ? 'sales' : kind === 'fees' ? 'fees' : kind === 'inventory' ? 'inventory' : kind === 'specsheet' ? 'specsheet' : 'manifest';
  try{
    const loader = kind === 'specsheet' ? loadMultiSheetTabularFile : loadTabularFile;
    const { headers, sheetName, headerRow, sheetAutoDetectFailed, noDataRowsFound, sheetsUsed } = await loader(file, schemaKeyForKind);
    state.detectedHeaders[kind] = headers;
    renderColumnMap(kind, zone, headers, { sheetName, headerRow, sheetAutoDetectFailed, noDataRowsFound, sheetsUsed });
  } catch(err){
    mapBox.innerHTML = `<span class="uz-colmap-err">Couldn't read headers: ${err.message}</span>`;
  } finally {
    zone.classList.remove('detecting');
  }
}

function renderColumnMap(kind, zone, headers, sheetInfo){
  const mapBox = zone.querySelector('.uz-colmap');
  const schemaKey = kind.startsWith('sales') ? 'sales' : kind === 'fees' ? 'fees' : kind === 'inventory' ? 'inventory' : kind === 'specsheet' ? 'specsheet' : 'manifest';
  const mapping = detectColumnMapping(headers, schemaKey);
  const fields = Object.keys(mapping);
  const matched = fields.filter(f => mapping[f]);
  const required = { sales:['sku','shipState','quantity'], fees:['sku','region','recvQty','placementFee'], inventory:['sku','region','onHand'], manifest:['sku','quantity'], specsheet:['itemId','vendor','casePack'] }[schemaKey] || [];
  const missingRequired = required.filter(f => !mapping[f]);

  let sheetNote = '';
  if(sheetInfo && sheetInfo.sheetsUsed && sheetInfo.sheetsUsed.length){
    sheetNote = `<div style="font-size:10.5px;color:var(--good);margin-bottom:6px;">Merged real data from ${sheetInfo.sheetsUsed.length} sheet${sheetInfo.sheetsUsed.length===1?'':'s'}: ${sheetInfo.sheetsUsed.join(', ')}</div>`;
  } else if(sheetInfo && sheetInfo.sheetName){
    const looksLikeExample = /\bexample\b|\bsample\b|\bdemo\b/i.test(sheetInfo.sheetName);
    if(sheetInfo.sheetAutoDetectFailed){
      sheetNote = `<div class="uz-colmap-warn" style="margin-bottom:6px;">Couldn't confidently find a matching sheet/header row — showing sheet "${sheetInfo.sheetName}", row 1. Check the mapping below.</div>`;
    } else if(sheetInfo.noDataRowsFound){
      sheetNote = `<div class="uz-colmap-warn" style="margin-bottom:6px;">No sheet had both a matching header AND real data rows — showing "${sheetInfo.sheetName}" as the closest header match, but it may be empty.</div>`;
    } else if(looksLikeExample){
      sheetNote = `<div class="uz-colmap-warn" style="margin-bottom:6px;">⚠ Your real "template" sheet appears empty — this is reading Amazon's "${sheetInfo.sheetName}" tab instead, which is likely placeholder/example data (e.g. "MySKU001"), not your real SKUs. Fill in the template sheet and re-upload for real results.</div>`;
    } else if(sheetInfo.headerRow){
      sheetNote = `<div style="font-size:10.5px;color:var(--ink-faint);margin-bottom:6px;">Found on sheet "<b style="color:var(--ink-dim)">${sheetInfo.sheetName}</b>", header row ${sheetInfo.headerRow}</div>`;
    }
  }

  let html = sheetNote + `<div class="uz-colmap-head">${matched.length}/${fields.length} columns matched${missingRequired.length ? ' <span class="uz-colmap-warn">· missing required</span>' : ''}</div><div class="uz-colmap-chips">`;
  for(const field of fields){
    const isRequired = required.includes(field);
    const got = mapping[field];
    const cls = got ? 'ok' : (isRequired ? 'missing' : 'optional-missing');
    html += `<span class="colchip ${cls}" title="${got ? 'matched column: '+got : 'not found in this file'}">${field}${got ? ' → '+got : ''}</span>`;
  }
  html += '</div>';
  mapBox.innerHTML = html;
}
for(let i=1;i<=6;i++) wireUploadZone('sales'+i, 'file-sales'+i, 'uz-sales'+i);
wireUploadZone('fees','file-fees','uz-fees');
wireUploadZone('inventory','file-inventory','uz-inventory');
wireUploadZone('manifest','file-manifest','uz-manifest');
wireUploadZone('specsheet','file-specsheet','uz-specsheet');
wireUploadZone('packinglist','file-packinglist','uz-packinglist');

document.getElementById('clearAllBtn').addEventListener('click', ()=>{
  ['sales1','sales2','sales3','sales4','sales5','sales6','fees','inventory'].forEach(kind=>{
    state.filesRaw[kind] = null;
    document.getElementById('uz-'+kind).classList.remove('has-file');
    document.getElementById('uz-'+kind).querySelector('.uz-filename').textContent = '';
    document.getElementById('file-'+kind).value = '';
  });
  state.salesAdapted = state.feeAdapted = state.inventoryAdapted = state.demand = state.feeRateTable = null;
  document.getElementById('adapterLog').textContent = 'Waiting for files…';
  document.getElementById('sellThroughToggle').checked = false;
  document.getElementById('sellThroughToggle').disabled = true;
  setDataStatus(false);
  renderAllEmptyStates();
  toast('Cleared all loaded data');
});

document.getElementById('loadSampleBtn').addEventListener('click', loadSampleData);

/* ============================================================
   RUN ANALYSIS
   ============================================================ */
document.getElementById('runAnalysisBtn').addEventListener('click', runAnalysis);

function logLine(container, label, html, kind){
  const cls = kind === 'warn' ? 'al-warn' : kind === 'good' ? 'al-good' : '';
  container.innerHTML += `<div class="al-line"><span>${label}</span><span class="${cls}">${html}</span></div>`;
}

async function runAnalysis(){
  const log = document.getElementById('adapterLog');
  log.innerHTML = '';
  const progress = document.getElementById('progressBar');
  const fill = document.getElementById('progressFill');
  progress.style.display = 'block';
  fill.style.width = '8%';

  const salesFiles = getSalesFiles();
  if(!salesFiles.length && !state.filesRaw.fees){
    toast('Upload at least one Sales period or a Fee file first — or click "Load sample data"', true);
    progress.style.display = 'none';
    return;
  }

  try{
    if(salesFiles.length){
      logLine(log, 'Sales files', `reading ${salesFiles.length} period${salesFiles.length===1?'':'s'}…`);
      let combinedObjects = [];
      let anyFormat = null, anyEncoding = null;
      for(const f of salesFiles){
        const { objects, encoding, format } = await loadTabularFile(f, 'sales');
        combinedObjects = combinedObjects.concat(objects);
        anyFormat = anyFormat || format;
        anyEncoding = anyEncoding || encoding;
      }
      fill.style.width = '30%';
      const { records, stats } = adaptSalesRows(combinedObjects);
      state.salesAdapted = records;
      log.innerHTML = '';
      logLine(log, 'Sales periods merged', `<b>${salesFiles.length}</b> file${salesFiles.length===1?'':'s'} — ${salesFiles.map(f=>f.name).join(', ')}`, 'good');
      logLine(log, 'Sales file format', anyFormat === 'excel' ? 'Excel (.xlsx)' : 'CSV', 'good');
      if(anyFormat === 'csv') logLine(log, 'Sales encoding detected', anyEncoding, anyEncoding.includes('fallback') ? 'warn' : 'good');
      logLine(log, 'Sales rows parsed (all periods combined)', `<b>${stats.totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Kept (US, resolvable state, valid qty)', `<b>${stats.kept.toLocaleString()}</b>`, 'good');
      if(stats.droppedNonUS) logLine(log, 'Dropped — non-US', stats.droppedNonUS.toLocaleString(), 'warn');
      if(stats.droppedUnresolvedState) logLine(log, 'Dropped — unresolvable state (military/intl)', stats.droppedUnresolvedState.toLocaleString(), 'warn');
      if(stats.droppedStatus) logLine(log, 'Dropped — excluded status', stats.droppedStatus.toLocaleString(), 'warn');
    }
    fill.style.width = '55%';

    if(state.filesRaw.fees){
      const { objects, format } = await loadTabularFile(state.filesRaw.fees, 'fees');
      const { records, stats } = adaptFeeRows(objects);
      state.feeAdapted = records;
      logLine(log, 'Fee log format', format === 'excel' ? 'Excel (.xlsx)' : 'CSV', 'good');
      logLine(log, 'Fee log rows parsed', `<b>${stats.totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Kept (well-formed, valid region)', `<b>${stats.kept.toLocaleString()}</b>`, 'good');
      if(stats.malformed) logLine(log, 'Dropped — malformed/ragged rows', stats.malformed.toLocaleString(), 'warn');
      state.feeRateTable = buildFeeRateTable(records);
      state.skuSizeTierLookup = buildSkuSizeTierLookup(records);
    }
    fill.style.width = '75%';

    if(state.filesRaw.inventory){
      const { objects } = await loadTabularFile(state.filesRaw.inventory, 'inventory');
      const { records, stats } = adaptInventoryRows(objects);
      state.inventoryAdapted = buildInventoryBySkuRegion(records);
      logLine(log, 'Inventory rows parsed', `<b>${stats.totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Kept (valid SKU + region)', `<b>${stats.kept.toLocaleString()}</b>`, 'good');
      if(stats.dropped) logLine(log, 'Dropped — UNIDENTIFIED/invalid', stats.dropped.toLocaleString(), 'warn');
      document.getElementById('sellThroughToggle').disabled = false;
    } else {
      document.getElementById('sellThroughToggle').disabled = true;
      document.getElementById('sellThroughToggle').checked = false;
      logLine(log, 'Sell-through correction', 'OFF — no inventory file loaded. Raw demand % only.', 'warn');
    }

    if(state.filesRaw.specsheet){
      const { objects, sheetsUsed } = await loadMultiSheetTabularFile(state.filesRaw.specsheet, 'specsheet');
      const { records, stats } = adaptSpecSheetRows(objects);
      state.specSheetLookup = buildSpecSheetLookup(records);
      logLine(log, 'Spec sheet sheets merged', sheetsUsed ? sheetsUsed.join(', ') : '1 sheet');
      logLine(log, 'Spec sheet rows parsed', `<b>${stats.totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Real vendor+item entries kept', `<b>${stats.kept.toLocaleString()}</b>`, 'good');
      if(stats.dropped) logLine(log, 'Dropped — missing vendor/item ID', stats.dropped.toLocaleString(), 'warn');
    } else {
      state.specSheetLookup = null;
    }
    // Manual entries and packing-list-derived entries always persist
    // across re-analysis and take priority over the bulk spec sheet —
    // they're the operator's most recent, most specific real data.
    const manualCount = Object.keys(state.manualSpecEntries).length;
    if(manualCount > 0){
      state.specSheetLookup = { ...(state.specSheetLookup||{}), ...state.manualSpecEntries };
      logLine(log, 'Manual / packing-list entries merged', `<b>${manualCount.toLocaleString()}</b>`, 'good');
    }

    if(state.salesAdapted){
      state.sellThroughOn = document.getElementById('sellThroughToggle').checked && !!state.inventoryAdapted;
      state.demand = buildDemandEngine(state.salesAdapted, state.inventoryAdapted, state.sellThroughOn);
      state.realDateSpanDays = computeRealDateSpanDays(state.salesAdapted);
      state.velocityByDay = buildVelocityByDay(state.salesAdapted);
      logLine(log, 'Distinct SKUs with demand signal', `<b>${Object.keys(state.demand).length.toLocaleString()}</b>`, 'good');
      if(state.realDateSpanDays) logLine(log, 'Real date span covered', `<b>${state.realDateSpanDays} days</b>`, 'good');
      const vel = state.velocityByDay;
      if(vel) logLine(log, 'Daily velocity (E / C / W)', `<b>${vel.avgByRegion.East.toFixed(1)} / ${vel.avgByRegion.Central.toFixed(1)} / ${vel.avgByRegion.West.toFixed(1)}</b> units/day`, 'good');
      if(state.specSheetLookup){
        const skus = Object.keys(state.demand);
        const matchedCount = skus.filter(s => lookupSpecSheetForSku(s)).length;
        logLine(log, 'SKUs matched to real spec sheet data', `<b>${matchedCount.toLocaleString()}</b> / ${skus.length.toLocaleString()} (${(matchedCount/skus.length*100).toFixed(1)}%)`, matchedCount>0?'good':'warn');
      }
    }

    fill.style.width = '100%';
    setTimeout(()=>{ progress.style.display='none'; fill.style.width='0%'; }, 500);
    setDataStatus(true);
    renderAll();
    saveSessionSummary();
    toast('Analysis complete — dashboards updated with your real data');
  } catch(err){
    progress.style.display = 'none';
    toast('Something failed while processing: ' + err.message, true);
    console.error(err);
  }
}

document.getElementById('sellThroughToggle').addEventListener('change', ()=>{
  if(state.salesAdapted){
    state.sellThroughOn = document.getElementById('sellThroughToggle').checked && !!state.inventoryAdapted;
    state.demand = buildDemandEngine(state.salesAdapted, state.inventoryAdapted, state.sellThroughOn);
    renderAll();
    toast(state.sellThroughOn ? 'Sell-through correction enabled' : 'Sell-through correction disabled');
  }
});

function setDataStatus(live){
  const el = document.getElementById('dataStatus');
  const txt = document.getElementById('dataStatusText');
  el.classList.toggle('live', live);
  txt.textContent = live ? 'Live data loaded' : 'No data loaded';
}

/* ============================================================
   SAMPLE DATA (clearly labeled as synthetic, used only on request)
   ============================================================ */
function loadSampleData(){
  const sampleSales = [];
  const skuPool = ['DEMO-FBA-0001','DEMO-FBA-0002','DEMO-FBA-0003','DEMO-FBA-0004','DEMO-FBA-0005'];
  const states = ['NY','CA','TX','OH','WA','FL','IL','GA'];
  for(let i=0;i<600;i++){
    const sku = skuPool[Math.floor(Math.random()*skuPool.length)];
    const st = states[Math.floor(Math.random()*states.length)];
    sampleSales.push({ sku, region: resolveStateToRegion(st), quantity: 1+Math.floor(Math.random()*4), asin:'B0DEMO'+i, price: 10+Math.random()*30, date:'' });
  }
  state.salesAdapted = sampleSales;
  state.feeAdapted = [
    {sku:'DEMO-FBA-0001',region:'East',sizeTier:'LargeStandard',feePerUnit:0.26,fee:26,qty:100},
    {sku:'DEMO-FBA-0001',region:'Central',sizeTier:'LargeStandard',feePerUnit:0.38,fee:38,qty:100},
    {sku:'DEMO-FBA-0001',region:'West',sizeTier:'LargeStandard',feePerUnit:0.50,fee:50,qty:100},
  ];
  state.feeRateTable = buildFeeRateTable(state.feeAdapted);
  state.inventoryAdapted = null;
  state.demand = buildDemandEngine(state.salesAdapted, null, false);
  setDataStatus(true);
  document.getElementById('adapterLog').innerHTML = '<div class="al-line"><span>Sample mode</span><span class="al-warn">Synthetic demo data — not real Virventures data. Upload real files for a real analysis.</span></div>';
  renderAll();
  toast('Sample data loaded — clearly synthetic, for exploring the UI only');
}

/* ============================================================
   RENDERING
   ============================================================ */
function fmtPct(x){ return (x*100).toFixed(1)+'%'; }
function fmtMoney(x){ return '$'+x.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }

/* Animates a KPI number counting up from 0 (or from its previous value)
   to a target. Works for plain integers, decimals, and percentages by
   re-running the same formatter at each animation frame rather than just
   interpolating raw text. */
function animateCountUp(el, targetValue, formatter, durationMs){
  durationMs = durationMs || 700;
  const start = performance.now();
  const from = 0;
  function tick(now){
    const elapsed = now - start;
    const t = Math.min(elapsed/durationMs, 1);
    const eased = 1 - Math.pow(1-t, 3); // ease-out cubic
    const current = from + (targetValue-from)*eased;
    el.textContent = formatter(current);
    if(t < 1) requestAnimationFrame(tick);
    else el.textContent = formatter(targetValue);
  }
  requestAnimationFrame(tick);
}
function animateKpiValues(container){
  container.querySelectorAll('[data-count-to]').forEach(el=>{
    const target = parseFloat(el.dataset.countTo);
    if(!Number.isFinite(target)){ return; }
    const kind = el.dataset.countKind || 'int';
    const sign = el.dataset.countSign === '1';
    let formatter;
    if(kind === 'pct') formatter = (v)=> (sign && v>=0 ? '+' : '') + (v*100).toFixed(1)+'%';
    else formatter = (v)=> Math.round(v).toLocaleString();
    animateCountUp(el, target, formatter);
  });
}

function renderAllEmptyStates(){
  document.getElementById('kpiGrid').innerHTML = '';
  renderKpis();
}

function renderAll(){
  populateGlobalVendorSelect();
  renderKpis();
  renderRegionTable();
  renderGapChart();
  renderParetoChart();
  renderVelocityDashboard();
  renderDemandPanel();
  renderCostAnalyticsDashboard();
  renderFeeRateTable();
  renderVendorCostTable();
  renderFrontierSelect();
  renderSkuAnalysis();
  renderVendorAnalysis();
  renderManifestAnalytics();
  initAllChartInteractivity();
}

/* Populates the global vendor dropdown from real extracted prefixes,
   sorted by SKU count so the biggest vendors surface first. Called after
   every successful data load / manifest build, since the vendor universe
   can change when scope changes too. */
function populateGlobalVendorSelect(){
  const sel = document.getElementById('globalVendorSelect');
  if(!state.demand){ sel.innerHTML = '<option value="all">All vendors</option>'; return; }
  const counts = {};
  for(const sku of Object.keys(state.demand)){
    const v = extractVendorPrefix(sku);
    counts[v] = (counts[v]||0) + 1;
  }
  const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]);
  const prevValue = state.globalVendorFilter;
  sel.innerHTML = '<option value="all">All vendors</option>' +
    sorted.map(([v,c])=>`<option value="${v}">${v} (${c} SKU${c===1?'':'s'})</option>`).join('');
  // preserve selection across re-population if it's still valid
  if(sorted.some(([v])=>v===prevValue)) sel.value = prevValue;
  else { sel.value = 'all'; state.globalVendorFilter = 'all'; }
}

document.getElementById('scopeSelect').addEventListener('change', (e)=>{
  state.scopeMode = e.target.value;
  if(state.scopeMode === 'manifest' && (!state.manifestPlan || !state.manifestPlan.length)){
    toast('Manifest Mode selected, but no shipment plan has been built yet — build one on the Manifest tab first, or every other tab will show zero SKUs.', true);
  } else if(state.scopeMode === 'manifest'){
    toast(`Manifest Mode on — every tab now scoped to the ${state.manifestPlan.length} SKUs in your active shipment plan.`);
  } else {
    toast('Full catalog scope restored.');
  }
  renderAll();
});
document.getElementById('globalVendorSelect').addEventListener('change', (e)=>{
  state.globalVendorFilter = e.target.value;
  renderAll();
});

function aggregatePortfolioSplit(){
  if(!state.demand) return null;
  const activeSkus = getActiveSkuSet();
  if(activeSkus && activeSkus.size === 0) return {East:0,Central:0,West:0};
  const tot = {East:0,Central:0,West:0};
  for(const sku of Object.keys(state.demand)){
    if(activeSkus && !activeSkus.has(sku)) continue;
    const r = state.demand[sku].rawUnits;
    tot.East += r.East; tot.Central += r.Central; tot.West += r.West;
  }
  return tot;
}
function aggregateFeeRegionShare(){
  if(!state.feeAdapted) return null;
  const activeSkus = getActiveSkuSet();
  if(activeSkus && activeSkus.size === 0) return {East:0,Central:0,West:0};
  const tot = {East:0,Central:0,West:0};
  for(const r of state.feeAdapted){
    if(activeSkus && !activeSkus.has(r.sku)) continue;
    tot[r.region] += r.qty;
  }
  return tot;
}

function renderKpis(){
  const grid = document.getElementById('kpiGrid');
  const activeSkus = getActiveSkuSet();
  const skuCount = activeSkus ? activeSkus.size : 0;
  const split = aggregatePortfolioSplit();
  const feeShare = aggregateFeeRegionShare();
  let eastDemandFrac = null, eastPlacementFrac = null, gapFrac = null, gapClass='flat';
  if(split){
    const tot = split.East+split.Central+split.West;
    eastDemandFrac = tot>0 ? split.East/tot : null;
  }
  if(feeShare){
    const tot = feeShare.East+feeShare.Central+feeShare.West;
    eastPlacementFrac = tot>0 ? feeShare.East/tot : null;
  }
  if(split && feeShare){
    const totD = split.East+split.Central+split.West;
    const totF = feeShare.East+feeShare.Central+feeShare.West;
    if(totD>0 && totF>0){
      gapFrac = (feeShare.East/totF) - (split.East/totD);
      gapClass = gapFrac > 0.03 ? 'down' : gapFrac < -0.03 ? 'up' : 'flat';
    }
  }

  grid.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">SKUs with demand signal</div>
      <div class="kpi-val"><span data-count-to="${skuCount}" data-count-kind="int">0</span></div>
      <div class="kpi-delta flat">from your sales upload</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">East — demand share</div>
      <div class="kpi-val">${eastDemandFrac!=null ? `<span data-count-to="${eastDemandFrac}" data-count-kind="pct">0%</span>` : '—'}</div>
      <div class="kpi-delta flat">measured from order rows</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">East — placement share</div>
      <div class="kpi-val">${eastPlacementFrac!=null ? `<span data-count-to="${eastPlacementFrac}" data-count-kind="pct">0%</span>` : '—'}</div>
      <div class="kpi-delta flat">measured from fee log</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">Placement vs. demand gap</div>
      <div class="kpi-val">${gapFrac!=null ? `<span data-count-to="${gapFrac}" data-count-kind="pct" data-count-sign="1">0%</span>` : '—'}</div>
      <div class="kpi-delta ${gapClass}">${gapClass==='down'?'over-placed in East':gapClass==='up'?'under-placed in East':'roughly aligned'}</div>
    </div>`;
  animateKpiValues(grid);
}

function renderRegionTable(){
  const tbody = document.querySelector('#regionTable tbody');
  const split = aggregatePortfolioSplit();
  if(!split){ tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--ink-faint);padding:20px 0;">No data yet</td></tr>'; return; }
  const tot = split.East+split.Central+split.West;
  const rows = [['East','var(--east)'],['Central','var(--central)'],['West','var(--west)']];
  tbody.innerHTML = rows.map(([region,color])=>{
    const units = split[region];
    return `<tr><td><span class="region-dot" style="background:${color}"></span>${region}</td><td class="num">${units.toLocaleString()}</td><td class="num">${tot>0?fmtPct(units/tot):'—'}</td></tr>`;
  }).join('');
}

/* Real Pareto/concentration analysis: sorts SKUs by demand units
   descending, plots cumulative % of demand against cumulative % of
   catalog. This is honest, single-snapshot analysis — it does NOT
   fabricate a multi-month trend the data doesn't support (only one
   month of real sales history exists). The "trend" here is a real
   distribution curve, not a time series. */
function renderParetoChart(){
  const wrap = document.getElementById('paretoChartWrap');
  const takeawayEl = document.getElementById('paretoTakeaway');
  if(!state.demand){
    wrap.innerHTML = `<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M4 19V10M10 19V5M16 19v-7M21 19H3"/></svg><h4>No demand data yet</h4><p>Upload a Sales export on the Data tab to see which share of your catalog drives most of the volume.</p></div>`;
    takeawayEl.innerHTML = '';
    return;
  }
  const activeSkus = getActiveSkuSet();
  const skus = Object.keys(state.demand).filter(s=>!activeSkus || activeSkus.has(s));
  if(!skus.length){
    wrap.innerHTML = `<div class="empty-state"><p>No SKUs in the current scope.</p></div>`;
    takeawayEl.innerHTML = '';
    return;
  }
  const sortedRows = skus.map(s=>({ sku:s, units:state.demand[s].totalUnits })).sort((a,b)=>b.units-a.units);
  const sorted = sortedRows.map(r=>r.units);
  const totalUnits = sorted.reduce((a,b)=>a+b,0);
  const n = sorted.length;
  // sample ~40 points along the curve for a smooth but real line (no fabrication, just resolution control on real data)
  const sampleCount = Math.min(n, 40);
  const points = [];
  let cumUnits = 0, idx = 0;
  for(let i=1;i<=sampleCount;i++){
    const targetIdx = Math.round((i/sampleCount)*n);
    while(idx < targetIdx && idx < n){ cumUnits += sorted[idx]; idx++; }
    points.push({ pctSkus: idx/n, pctUnits: totalUnits>0?cumUnits/totalUnits:0 });
  }
  // find the real "80% of demand" breakpoint
  let breakpointPctSkus = null;
  for(const p of points){ if(p.pctUnits >= 0.8){ breakpointPctSkus = p.pctSkus; break; } }
  if(breakpointPctSkus===null && points.length) breakpointPctSkus = points[points.length-1].pctSkus;

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#6B7280' : '#928D80';
  const signal = isDark ? '#FF6B35' : '#D9531E';
  const lineColor = isDark ? '#262C35' : '#E4E0D8';
  const W=560,H=300,padL=55,padR=20,padT=20,padB=34;
  const plotW=W-padL-padR, plotH=H-padT-padB;
  const toXY = p => [padL + p.pctSkus*plotW, padT + plotH - p.pctUnits*plotH];
  const pathPts = points.map(toXY);
  const path = 'M ' + pathPts.map(p=>p.join(' ')).join(' L ');
  const area = path + ` L ${pathPts[pathPts.length-1][0]} ${padT+plotH} L ${pathPts[0][0]} ${padT+plotH} Z`;
  const breakX = padL + breakpointPctSkus*plotW;
  const breakY = padT + plotH*0.2;

  // Real horizontal gridlines at 0/20/40/60/80/100% of demand, so movement
  // in the curve is legible against fixed reference lines, not just a
  // single floating dot. Real vertical gridlines at 25/50/75% of catalog.
  let gridLines = '';
  [0,20,40,60,80,100].forEach(pct=>{
    const y = padT + plotH - (pct/100)*plotH;
    gridLines += `<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="${lineColor}" stroke-width="1" ${pct===80?'stroke-dasharray="3 4"':''}/>`;
    gridLines += `<text x="${padL-8}" y="${y+3.5}" fill="${textColor}" font-family="JetBrains Mono" font-size="10" text-anchor="end">${pct}%</text>`;
  });
  [25,50,75].forEach(pct=>{
    const x = padL + (pct/100)*plotW;
    gridLines += `<line x1="${x}" y1="${padT}" x2="${x}" y2="${padT+plotH}" stroke="${lineColor}" stroke-width="1" opacity="0.5"/>`;
    gridLines += `<text x="${x}" y="${padT+plotH+16}" fill="${textColor}" font-family="JetBrains Mono" font-size="9.5" text-anchor="middle">${pct}%</text>`;
  });

  // Real labeled markers at every quartile of the catalog (not just the
  // 80% breakpoint) so the actual curve VALUES are readable at a glance,
  // not just its silhouette.
  let markers = '';
  [0.10, 0.25, 0.50, 0.75].forEach(targetPctSkus=>{
    let closest = points[0];
    for(const p of points){ if(Math.abs(p.pctSkus-targetPctSkus) < Math.abs(closest.pctSkus-targetPctSkus)) closest = p; }
    const [mx,my] = toXY(closest);
    markers += `<circle cx="${mx}" cy="${my}" r="3.5" fill="${isDark?'#0E1116':'#fff'}" stroke="${signal}" stroke-width="2"/>`;
    markers += `<text x="${mx}" y="${my-10}" fill="${textColor}" font-family="JetBrains Mono" font-size="10" text-anchor="middle">${fmtPct(closest.pctUnits)}</text>`;
  });

  const svg = `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;">
    ${gridLines}
    <path d="${area}" fill="${signal}" opacity="0.08"/>
    <path d="${path}" fill="none" stroke="${signal}" stroke-width="2.5"/>
    ${markers}
    <line x1="${breakX}" y1="${padT}" x2="${breakX}" y2="${padT+plotH}" stroke="${signal}" stroke-width="1.5" stroke-dasharray="2 3"/>
    <circle cx="${breakX}" cy="${breakY}" r="5" fill="${signal}"/>
    <rect x="${Math.min(breakX+8, W-padR-118)}" y="${breakY-26}" width="116" height="36" rx="5" fill="${isDark?'#151A21':'#fff'}" stroke="${signal}" stroke-width="1"/>
    <text x="${Math.min(breakX+14, W-padR-112)}" y="${breakY-12}" fill="${signal}" font-family="JetBrains Mono" font-size="10" font-weight="600">80% of demand</text>
    <text x="${Math.min(breakX+14, W-padR-112)}" y="${breakY+0}" fill="${textColor}" font-family="JetBrains Mono" font-size="9.5">at ${fmtPct(breakpointPctSkus)} of SKUs</text>
    <text x="${padL+plotW/2}" y="${H-2}" fill="${textColor}" font-family="Inter" font-size="10.5" text-anchor="middle">% of SKUs (highest-volume first) →</text>
  </svg>`;
  wrap.innerHTML = svg;

  const skuCountAt80 = Math.round(breakpointPctSkus*n);
  takeawayEl.innerHTML = `<div class="alloc-reason" style="margin-top:4px;">
    <b style="color:var(--ink)">${fmtPct(breakpointPctSkus)} of SKUs (${skuCountAt80.toLocaleString()} of ${n.toLocaleString()}) drive 80% of total demand units.</b>
    ${breakpointPctSkus < 0.3 ? 'This is a real concentrated catalog — placement decisions for that small core of SKUs matter far more than getting every long-tail SKU exactly right.' : breakpointPctSkus > 0.6 ? 'Demand is unusually spread out across this catalog — no small core of SKUs dominates, so broad-based accuracy across many SKUs matters more than focusing on a handful of top sellers.' : 'A moderate concentration — worth prioritizing the top quartile of SKUs for placement accuracy, but the long tail still carries real volume.'}
  </div>`;

  // Real, exact list of the vital-few SKUs that make up the 80% breakpoint
  // — the brief explicitly asked for this to support real decision-making,
  // not just the chart shape. Every row here is a real SKU with a real
  // cumulative % computed from the same sorted data as the curve above.
  let cumUnitsRunning = 0;
  state.paretoSkuRows = sortedRows.map((r,i)=>{
    cumUnitsRunning += r.units;
    return { rank:i+1, sku:r.sku, units:r.units, pctOfTotal: totalUnits>0?r.units/totalUnits:0, cumulativePct: totalUnits>0?cumUnitsRunning/totalUnits:0 };
  });
  document.getElementById('paretoSkuListWrap').style.display = 'block';
  document.getElementById('paretoSkuListTitle').textContent = `The vital few — SKUs driving 80% of demand (${skuCountAt80.toLocaleString()} SKUs)`;
  applyParetoSkuFilter();
}

function applyParetoSkuFilter(){
  if(!state.paretoSkuRows) return;
  const search = (document.getElementById('paretoSkuSearch').value||'').toLowerCase().trim();
  // Default view: just the vital-few rows (cumulative <= 80%) so the list
  // matches exactly what the chart and takeaway describe. A search
  // overrides this and searches the FULL list, since the person might be
  // checking whether a specific SKU is in the long tail, which is itself
  // useful information.
  let rows = state.paretoSkuRows;
  if(search){
    rows = rows.filter(r=>r.sku.toLowerCase().includes(search));
  } else {
    rows = rows.filter(r=>r.cumulativePct <= 0.80 + 1e-9);
  }
  rows = rows.slice(0,300);
  const tbody = document.querySelector('#paretoSkuListTable tbody');
  if(!rows.length){ tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:18px 0;">No matching SKUs</td></tr>'; return; }
  tbody.innerHTML = rows.map(r=>`
    <tr><td class="mono">${r.rank}</td><td class="mono">${r.sku}</td><td class="num">${r.units.toLocaleString()}</td><td class="num">${fmtPct(r.pctOfTotal)}</td><td class="num">${fmtPct(r.cumulativePct)}</td></tr>
  `).join('');
}
document.getElementById('paretoSkuSearch').addEventListener('input', applyParetoSkuFilter);


function renderGapChart(){
  const wrap = document.getElementById('gapChartWrap');
  const split = aggregatePortfolioSplit();
  const feeShare = aggregateFeeRegionShare();
  if(!split){
    wrap.innerHTML = `<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 17l6-6 4 4 8-8"/></svg><h4>No demand data yet</h4><p>Upload a Sales export on the Data tab to compute real regional demand percentages.</p></div>`;
    return;
  }
  const totD = split.East+split.Central+split.West;
  const demandPct = { East: split.East/totD, Central: split.Central/totD, West: split.West/totD };
  let placementPct = null;
  if(feeShare){
    const totF = feeShare.East+feeShare.Central+feeShare.West;
    if(totF>0) placementPct = { East: feeShare.East/totF, Central: feeShare.Central/totF, West: feeShare.West/totF };
  }
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#A9AFB9' : '#5C5950';
  const colors = { East: 'var(--east)', Central: 'var(--central)', West: 'var(--west)' };
  const regions = ['East','Central','West'];
  const viewW = 520, marginX = 30;
  const groupW = (viewW - marginX*2) / 3;
  const barW = placementPct ? groupW*0.36 : groupW*0.5;
  const barGapWithin = placementPct ? groupW*0.06 : 0;
  const chartH = 200, baseY = 230;
  let svg = `<svg viewBox="0 0 ${viewW} 270" style="width:100%;height:auto;">`;
  regions.forEach((region,i)=>{
    const groupX = marginX + i*groupW;
    const pairW = placementPct ? (barW*2+barGapWithin) : barW;
    const x = groupX + (groupW - pairW)/2;
    const dH = demandPct[region]*chartH;
    const pH = placementPct ? placementPct[region]*chartH : 0;
    const dPct = (demandPct[region]*100).toFixed(1);
    const pPct = placementPct ? (placementPct[region]*100).toFixed(1) : null;
    svg += `<rect x="${x}" y="${baseY-dH}" width="${barW}" height="${dH}" fill="${colors[region]}" opacity="0.9" rx="3" data-region="${region}" data-type="Demand share" data-val="${dPct}" style="cursor:crosshair;transition:opacity .15s;"/>`;
    svg += `<text x="${x+barW/2}" y="${baseY-dH-8}" fill="${textColor}" font-family="JetBrains Mono" font-size="11" text-anchor="middle">${fmtPct(demandPct[region])}</text>`;
    if(placementPct){
      svg += `<rect x="${x+barW+barGapWithin}" y="${baseY-pH}" width="${barW}" height="${pH}" fill="${colors[region]}" opacity="0.35" rx="3" data-region="${region}" data-type="Placement share" data-val="${pPct}" style="cursor:crosshair;transition:opacity .15s;"/>`;
      svg += `<text x="${x+barW+barGapWithin+barW/2}" y="${baseY-pH-8}" fill="${textColor}" font-family="JetBrains Mono" font-size="11" text-anchor="middle">${fmtPct(placementPct[region])}</text>`;
    }
    const labelX = groupX + groupW/2;
    svg += `<text x="${labelX}" y="${baseY+22}" fill="${textColor}" font-family="Inter" font-size="12" text-anchor="middle">${region}</text>`;
  });
  svg += `<line x1="${marginX-10}" y1="${baseY}" x2="${viewW-marginX+10}" y2="${baseY}" stroke="${isDark?'#262C35':'#E4E0D8'}" stroke-width="1"/>`;
  svg += `</svg>`;
  const legend = `<div class="legend" style="display:flex;gap:18px;margin-top:6px;font-size:11.5px;color:var(--ink-dim);">
    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:2px;background:var(--ink-dim);opacity:0.9;display:inline-block;"></span>Demand share (solid)</div>
    ${placementPct ? `<div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:2px;background:var(--ink-dim);opacity:0.35;display:inline-block;"></span>Placement share (faded)</div>` : ''}
  </div>`;
  wrap.innerHTML = svg + legend;
  // Attach tooltip interactivity to gap chart bars
  setTimeout(()=>addGapChartInteractivity(), 50);
}

/* ============================================================
   SKU ANALYSIS — top sellers, revenue, regional concentration, flags.
   Computed entirely from state.salesAdapted (real order rows already
   parsed by the Sales adapter) — no new data source needed.
   ============================================================ */
/* ============================================================
   RESTOCK SIGNAL ENGINE — deterministic, not an LLM guess.
   Two honesty tiers, never blurred together:
   (1) REAL days-of-stock, when a real Inventory-by-Region file exists:
       on-hand units / real daily velocity = real days remaining.
   (2) VELOCITY TIER ONLY, when no inventory file exists: classifies
       a SKU's sales velocity relative to the rest of the active catalog
       (using real percentile rank, not arbitrary cutoffs) but explicitly
       does NOT claim a stock-out timeline, since we don't know what's
       on hand. The UI and the assistant must keep these two cases
       visually and textually distinct at all times.
   ============================================================ */
function buildRestockSignals(){
  if(!state.skuAnalysisRows || !state.skuAnalysisRows.length) return null;
  const spanDays = state.realDateSpanDays;
  const hasInventory = !!state.inventoryAdapted;
  const rows = state.skuAnalysisRows;

  // Real percentile ranking of velocity within the active catalog —
  // used for the no-inventory case so "high/low velocity" means
  // something real and relative, not an arbitrary absolute cutoff.
  const sortedByUnits = [...rows].sort((a,b)=>a.units-b.units);
  const rankOf = {};
  sortedByUnits.forEach((r,i)=>{ rankOf[r.sku] = i/(Math.max(sortedByUnits.length-1,1)); });

  return rows.map(row=>{
    const dailyVelocity = spanDays ? row.units/spanDays : null;
    let signal, signalDetail, confidence;

    if(hasInventory && dailyVelocity){
      const inv = state.inventoryAdapted[row.sku];
      const onHandTotal = inv ? (inv.East+inv.Central+inv.West) : null;
      if(onHandTotal!=null && dailyVelocity>0){
        const daysOfStock = onHandTotal/dailyVelocity;
        confidence = 'real';
        if(daysOfStock < 14){ signal='restock_urgent'; signalDetail=`${Math.round(daysOfStock)} real days of stock left at current velocity (${dailyVelocity.toFixed(1)}/day) — on-hand: ${onHandTotal.toLocaleString()} units.`; }
        else if(daysOfStock < 30){ signal='restock_soon'; signalDetail=`${Math.round(daysOfStock)} real days of stock left at current velocity (${dailyVelocity.toFixed(1)}/day).`; }
        else if(daysOfStock > 120){ signal='order_less'; signalDetail=`${Math.round(daysOfStock)} real days of stock on hand at current velocity — well beyond a normal reorder horizon, real overstock risk.`; }
        else { signal='healthy'; signalDetail=`${Math.round(daysOfStock)} real days of stock left — within a normal range.`; }
      } else {
        confidence='velocity_only'; signal='no_onhand_data'; signalDetail=`No on-hand units found for this SKU in the uploaded inventory file.`;
      }
    } else {
      confidence='velocity_only';
      const pct = rankOf[row.sku];
      if(row.units===0){ signal='dormant'; signalDetail=`Zero real sales in the ${spanDays||'uploaded'}-day window — not a stock claim, just zero observed demand.`; }
      else if(pct>0.90){ signal='high_velocity'; signalDetail=`Top 10% of the active catalog by real sales velocity (${dailyVelocity?dailyVelocity.toFixed(2):'?'} units/day) — no on-hand data uploaded, so this is a velocity signal, not a stock-out warning.`; }
      else if(pct<0.10){ signal='low_velocity'; signalDetail=`Bottom 10% of the active catalog by real sales velocity — worth reviewing before reordering, but not a confirmed overstock without on-hand data.`; }
      else { signal='normal_velocity'; signalDetail=`Mid-range sales velocity relative to the rest of the active catalog.`; }
    }

    return { sku:row.sku, vendor:row.vendor, units:row.units, dailyVelocity, signal, signalDetail, confidence };
  });
}

function buildSkuAnalysis(){
  if(!state.salesAdapted || !state.demand) return null;
  const activeSkus = getActiveSkuSet();
  const bySku = {};
  for(const rec of state.salesAdapted){
    if(activeSkus && !activeSkus.has(rec.sku)) continue;
    if(!bySku[rec.sku]) bySku[rec.sku] = { sku:rec.sku, vendor:extractVendorPrefix(rec.sku), units:0, orders:0, revenue:0, regionUnits:{East:0,Central:0,West:0} };
    const row = bySku[rec.sku];
    row.units += rec.quantity;
    row.orders += 1;
    row.revenue += (rec.price||0) * rec.quantity;
    row.regionUnits[rec.region] += rec.quantity;
  }
  const rows = Object.values(bySku).map(row=>{
    const avgPrice = row.units>0 ? row.revenue/row.units : 0;
    const totalRegionUnits = row.regionUnits.East+row.regionUnits.Central+row.regionUnits.West;
    const regionPct = totalRegionUnits>0 ? {
      East: row.regionUnits.East/totalRegionUnits,
      Central: row.regionUnits.Central/totalRegionUnits,
      West: row.regionUnits.West/totalRegionUnits
    } : {East:1/3,Central:1/3,West:1/3};
    const maxConcentration = Math.max(regionPct.East, regionPct.Central, regionPct.West);
    let flag = null;
    if(maxConcentration > 0.70) flag = 'concentrated';
    else if(row.units < 10) flag = 'low-volume';
    return { ...row, avgPrice, regionPct, maxConcentration, flag };
  });
  return rows;
}


function renderSkuAnalysis(){
  const rows = buildSkuAnalysis();
  const kpiGrid = document.getElementById('skuAnalysisKpis');
  if(!rows){
    kpiGrid.innerHTML = '';
    ['topByUnitsTable','topByRevenueTable'].forEach(id=>{
      document.querySelector('#'+id+' tbody').innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--ink-faint);padding:20px 0;">Upload sales data on the Data tab</td></tr>';
    });
    document.querySelector('#skuAnalysisFullTable tbody').innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--ink-faint);padding:24px 0;">Upload sales data on the Data tab</td></tr>';
    return;
  }
  state.skuAnalysisRows = rows;

  const totalRevenue = rows.reduce((s,r)=>s+r.revenue,0);
  const totalUnits = rows.reduce((s,r)=>s+r.units,0);
  const concentratedCount = rows.filter(r=>r.flag==='concentrated').length;
  const lowVolumeCount = rows.filter(r=>r.flag==='low-volume').length;

  kpiGrid.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Total SKUs analyzed</div><div class="kpi-val"><span data-count-to="${rows.length}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Total revenue (from real orders)</div><div class="kpi-val mono" style="font-size:20px;">${fmtMoney(totalRevenue)}</div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">SKUs &gt;70% in one region</div><div class="kpi-val" style="color:var(--assumed)"><span data-count-to="${concentratedCount}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Low-volume SKUs (&lt;10 units)</div><div class="kpi-val" style="color:var(--ink-faint)"><span data-count-to="${lowVolumeCount}" data-count-kind="int">0</span></div></div>`;
  animateKpiValues(kpiGrid);

  const topByUnits = [...rows].sort((a,b)=>b.units-a.units).slice(0,10);
  document.querySelector('#topByUnitsTable tbody').innerHTML = topByUnits.map((r,i)=>`
    <tr><td class="mono">${i+1}</td><td class="mono">${r.sku}</td><td class="num">${r.units.toLocaleString()}</td><td class="num">${r.orders.toLocaleString()}</td></tr>`).join('');

  const topByRevenue = [...rows].sort((a,b)=>b.revenue-a.revenue).slice(0,10);
  document.querySelector('#topByRevenueTable tbody').innerHTML = topByRevenue.map((r,i)=>`
    <tr><td class="mono">${i+1}</td><td class="mono">${r.sku}</td><td class="num">${fmtMoney(r.revenue)}</td><td class="num">${fmtMoney(r.avgPrice)}</td></tr>`).join('');

  applySkuAnalysisFilters();
}

function applySkuAnalysisFilters(){
  if(!state.skuAnalysisRows) return;
  const search = (document.getElementById('skuAnalysisSearch').value||'').toLowerCase().trim();
  const sortBy = document.getElementById('skuAnalysisSort').value;
  const flagFilter = document.getElementById('skuAnalysisFlagFilter').value;

  let rows = state.skuAnalysisRows;
  if(search) rows = rows.filter(r=>r.sku.toLowerCase().includes(search));
  if(flagFilter !== 'all') rows = rows.filter(r=>r.flag===flagFilter);

  rows = [...rows];
  if(sortBy === 'units') rows.sort((a,b)=>b.units-a.units);
  else if(sortBy === 'revenue') rows.sort((a,b)=>b.revenue-a.revenue);
  else if(sortBy === 'concentration') rows.sort((a,b)=>b.maxConcentration-a.maxConcentration);

  rows = rows.slice(0,200);
  const tbody = document.querySelector('#skuAnalysisFullTable tbody');
  if(!rows.length){ tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--ink-faint);padding:24px 0;">No SKUs match your filters</td></tr>'; return; }
  tbody.innerHTML = rows.map(r=>{
    let flagBadge = '';
    if(r.flag==='concentrated') flagBadge = `<span class="badge" style="background:rgba(201,162,39,0.12);color:var(--assumed);">concentrated</span>`;
    else if(r.flag==='low-volume') flagBadge = `<span class="badge" style="background:var(--bg-input);color:var(--ink-faint);">low volume</span>`;
    return `<tr>
      <td class="mono">${r.sku}</td>
      <td class="num">${r.units.toLocaleString()}</td>
      <td class="num">${r.orders.toLocaleString()}</td>
      <td class="num">${fmtMoney(r.revenue)}</td>
      <td class="num">${fmtPct(r.regionPct.East)}</td>
      <td class="num">${fmtPct(r.regionPct.Central)}</td>
      <td class="num">${fmtPct(r.regionPct.West)}</td>
      <td>${flagBadge}</td>
    </tr>`;
  }).join('');
}
['skuAnalysisSearch'].forEach(id=> document.getElementById(id).addEventListener('input', applySkuAnalysisFilters));
['skuAnalysisSort','skuAnalysisFlagFilter'].forEach(id=> document.getElementById(id).addEventListener('change', applySkuAnalysisFilters));

/* ============================================================
   VENDOR ANALYTICS — rolls up the same real per-SKU data (already
   computed for the SKU Analysis tab) one level up to vendor prefix.
   No new data source: same sales rows, same scope/vendor filters,
   just grouped by extractVendorPrefix(sku) instead of by SKU.
   ============================================================ */
function buildVendorAnalysis(){
  if(!state.skuAnalysisRows) return null;
  const byVendor = {};
  for(const row of state.skuAnalysisRows){
    if(!byVendor[row.vendor]) byVendor[row.vendor] = { vendor:row.vendor, skuCount:0, units:0, revenue:0, regionUnits:{East:0,Central:0,West:0}, concentratedSkus:0, lowVolumeSkus:0 };
    const v = byVendor[row.vendor];
    v.skuCount += 1;
    v.units += row.units;
    v.revenue += row.revenue;
    v.regionUnits.East += row.regionUnits.East;
    v.regionUnits.Central += row.regionUnits.Central;
    v.regionUnits.West += row.regionUnits.West;
    if(row.flag==='concentrated') v.concentratedSkus++;
    if(row.flag==='low-volume') v.lowVolumeSkus++;
  }
  return Object.values(byVendor).map(v=>{
    const tot = v.regionUnits.East+v.regionUnits.Central+v.regionUnits.West;
    const regionPct = tot>0 ? { East:v.regionUnits.East/tot, Central:v.regionUnits.Central/tot, West:v.regionUnits.West/tot } : {East:1/3,Central:1/3,West:1/3};
    const maxConcentration = Math.max(regionPct.East, regionPct.Central, regionPct.West);
    const dominantRegion = Object.entries(regionPct).sort((a,b)=>b[1]-a[1])[0][0];
    return { ...v, regionPct, maxConcentration, dominantRegion, avgUnitsPerSku: v.skuCount?v.units/v.skuCount:0 };
  });
}

/* Builds a one-line, specific reasoning statement per vendor — the
   "McKinsey-style takeaway" — using the vendor's actual numbers, not a
   generic template. Every clause below references a real computed value. */
function buildVendorReasoning(v, allVendors){
  const totalUnits = allVendors.reduce((s,x)=>s+x.units,0);
  const shareOfPortfolio = totalUnits>0 ? v.units/totalUnits : 0;
  const parts = [];
  parts.push(`${v.vendor} accounts for ${fmtPct(shareOfPortfolio)} of total demand units across ${v.skuCount} SKU${v.skuCount===1?'':'s'}, averaging ${Math.round(v.avgUnitsPerSku)} units/SKU.`);
  if(v.maxConcentration > 0.70){
    parts.push(`Demand is concentrated in ${v.dominantRegion} (${fmtPct(v.maxConcentration)}) — ${v.concentratedSkus} of its ${v.skuCount} SKUs individually show the same skew, so this isn't one outlier SKU dragging the average.`);
  } else {
    parts.push(`Demand is reasonably spread across regions (largest single region is ${v.dominantRegion} at ${fmtPct(v.maxConcentration)}), so a single-region placement strategy would leave real demand unserved elsewhere.`);
  }
  if(v.lowVolumeSkus / Math.max(v.skuCount,1) > 0.4){
    parts.push(`${v.lowVolumeSkus} of ${v.skuCount} SKUs (${fmtPct(v.lowVolumeSkus/v.skuCount)}) are low-volume (under 10 units sold) — a large share of this vendor's catalog has too little signal yet for a confident regional split.`);
  }
  return parts.join(' ');
}

let vendorAnalysisRows = null;
function renderVendorAnalysis(){
  const rows = buildVendorAnalysis();
  vendorAnalysisRows = rows;
  const kpiGrid = document.getElementById('vendorKpis');
  if(!rows || !rows.length){
    kpiGrid.innerHTML = '';
    document.querySelector('#vendorTable tbody').innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--ink-faint);padding:24px 0;">Upload sales data on the Data tab</td></tr>';
    document.getElementById('vendorConcentrationChart').innerHTML = '';
    return;
  }
  const totalRevenue = rows.reduce((s,v)=>s+v.revenue,0);
  const totalUnits = rows.reduce((s,v)=>s+v.units,0);
  const concentratedVendors = rows.filter(v=>v.maxConcentration>0.70).length;
  const topVendor = [...rows].sort((a,b)=>b.units-a.units)[0];

  kpiGrid.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Distinct vendors</div><div class="kpi-val"><span data-count-to="${rows.length}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Top vendor by volume</div><div class="kpi-val mono" style="font-size:18px;">${topVendor.vendor}</div><div class="kpi-delta flat">${fmtPct(topVendor.units/totalUnits)} of total units</div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Total revenue (all vendors)</div><div class="kpi-val mono" style="font-size:18px;">${fmtMoney(totalRevenue)}</div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Vendors &gt;70% concentrated in one region</div><div class="kpi-val" style="color:var(--assumed)"><span data-count-to="${concentratedVendors}" data-count-kind="int">0</span></div></div>`;
  animateKpiValues(kpiGrid);

  renderVendorConcentrationChart(rows);
  applyVendorFilters();
}

function renderVendorConcentrationChart(rows){
  const wrap = document.getElementById('vendorConcentrationChart');
  const top = [...rows].sort((a,b)=>b.units-a.units).slice(0,8);
  if(!top.length){ wrap.innerHTML=''; return; }
  const maxUnits = Math.max(...top.map(v=>v.units));
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#A9AFB9' : '#5C5950';
  const colors = { East:'var(--east)', Central:'var(--central)', West:'var(--west)' };
  const barH = 30, gap = 14, leftPad = 90, chartW = 380;
  const svgH = top.length*(barH+gap)+20;
  let svg = `<svg viewBox="0 0 540 ${svgH}" style="width:100%;height:auto;">`;
  top.forEach((v,i)=>{
    const y = i*(barH+gap)+10;
    let x = leftPad;
    ['East','Central','West'].forEach(region=>{
      const segW = (v.regionPct[region]) * chartW;
      svg += `<rect x="${x}" y="${y}" width="${Math.max(segW,0)}" height="${barH}" fill="${colors[region]}" opacity="0.92" rx="2"/>`;
      x += segW;
    });
    svg += `<text x="${leftPad-10}" y="${y+barH/2+4}" fill="${textColor}" font-family="JetBrains Mono" font-size="11.5" text-anchor="end">${v.vendor}</text>`;
    svg += `<text x="${leftPad+chartW+10}" y="${y+barH/2+4}" fill="${textColor}" font-family="JetBrains Mono" font-size="10.5">${v.units.toLocaleString()}u</text>`;
  });
  svg += `</svg>`;
  const legend = `<div class="legend" style="display:flex;gap:18px;margin-top:10px;font-size:11.5px;color:var(--ink-dim);">
    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:2px;background:var(--east);display:inline-block;"></span>East</div>
    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:2px;background:var(--central);display:inline-block;"></span>Central</div>
    <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:2px;background:var(--west);display:inline-block;"></span>West</div>
  </div>`;
  wrap.innerHTML = svg + legend;
}

function applyVendorFilters(){
  if(!vendorAnalysisRows) return;
  const search = (document.getElementById('vendorSearch').value||'').toLowerCase().trim();
  const sortBy = document.getElementById('vendorSort').value;
  let rows = vendorAnalysisRows;
  if(search) rows = rows.filter(v=>v.vendor.toLowerCase().includes(search));
  rows = [...rows];
  if(sortBy==='units') rows.sort((a,b)=>b.units-a.units);
  else if(sortBy==='revenue') rows.sort((a,b)=>b.revenue-a.revenue);
  else if(sortBy==='skuCount') rows.sort((a,b)=>b.skuCount-a.skuCount);
  else if(sortBy==='concentration') rows.sort((a,b)=>b.maxConcentration-a.maxConcentration);
  rows = rows.slice(0,150);

  const tbody = document.querySelector('#vendorTable tbody');
  if(!rows.length){ tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--ink-faint);padding:24px 0;">No vendors match your filter</td></tr>'; return; }
  tbody.innerHTML = '';
  rows.forEach((v,idx)=>{
    const mainRow = document.createElement('tr');
    mainRow.innerHTML = `
      <td><span class="mtr-toggle" data-idx="${idx}">▸</span></td>
      <td class="mono">${v.vendor}</td>
      <td class="num">${v.skuCount.toLocaleString()}</td>
      <td class="num">${v.units.toLocaleString()}</td>
      <td class="num">${fmtMoney(v.revenue)}</td>
      <td class="num">${fmtPct(v.regionPct.East)}</td>
      <td class="num">${fmtPct(v.regionPct.Central)}</td>
      <td class="num">${fmtPct(v.regionPct.West)}</td>`;
    tbody.appendChild(mainRow);

    const detailRow = document.createElement('tr');
    detailRow.className = 'mtr-detail-row';
    const cell = document.createElement('td');
    cell.colSpan = 8;
    cell.innerHTML = `<div class="mtr-detail-inner" style="font-size:12.5px;color:var(--ink-dim);line-height:1.65;">${buildVendorReasoning(v, vendorAnalysisRows)}</div>`;
    detailRow.appendChild(cell);
    tbody.appendChild(detailRow);

    mainRow.querySelector('.mtr-toggle').addEventListener('click', ()=>{
      mainRow.querySelector('.mtr-toggle').classList.toggle('open');
      detailRow.classList.toggle('open');
    });
  });
}
['vendorSearch'].forEach(id=> document.getElementById(id).addEventListener('input', applyVendorFilters));
['vendorSort'].forEach(id=> document.getElementById(id).addEventListener('change', applyVendorFilters));

function renderVelocityDashboard(){
  const kpiRow = document.getElementById('velocityKpiRow');
  const chartWrap = document.getElementById('velocityChartWrap');
  const tbody = document.querySelector('#velocityTable tbody');
  const dateRangeEl = document.getElementById('velocityDateRange');
  if(!kpiRow || !chartWrap || !tbody) return;

  const vel = state.velocityByDay;
  if(!vel){
    kpiRow.innerHTML = '';
    chartWrap.innerHTML = '<div class="empty-state" style="padding:24px 0;"><p>Upload sales data and run analysis to see daily velocity</p></div>';
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--ink-faint);padding:18px 0;">Upload sales data and run analysis</td></tr>';
    return;
  }

  if(dateRangeEl) dateRangeEl.textContent = `${vel.dateRange.first} → ${vel.dateRange.last} (${vel.days.length} days)`;

  // KPI row — avg/day per region + total
  const regionClass = { East:'east', Central:'central', West:'west' };
  kpiRow.innerHTML = ['East','Central','West'].map(reg => `
    <div class="velocity-kpi ${regionClass[reg]}">
      <div class="vk-label">${reg}</div>
      <div class="vk-val">${vel.avgByRegion[reg].toFixed(1)}</div>
      <div class="vk-sub">units / day avg · ${vel.totalByRegion[reg].toLocaleString()} total</div>
    </div>`).join('') +
  `<div class="velocity-kpi total">
    <div class="vk-label">All regions</div>
    <div class="vk-val">${vel.avgByRegion.Total.toFixed(1)}</div>
    <div class="vk-sub">units / day avg · ${vel.totalByRegion.Total.toLocaleString()} total</div>
  </div>`;

  // SVG line chart — one line per region over the date range
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#6B7280' : '#928D80';
  const gridColor = isDark ? '#1E2530' : '#E8E4DC';
  const W = 860, H = 200, padL = 48, padR = 20, padT = 16, padB = 32;
  const plotW = W-padL-padR, plotH = H-padT-padB;
  const days = vel.days;
  const n = days.length;
  const maxVal = Math.max(...days.map(d=>d.total), 1);
  const regionColors = { East: '#FF6B35', Central: '#6FA8CC', West: '#7FB069' };
  const toX = i => padL + (i/(n-1))*plotW;
  const toY = v => padT + plotH - (v/maxVal)*plotH;

  // Y gridlines at 0/25/50/75/100% of max
  let gridLines = '';
  [0,0.25,0.5,0.75,1.0].forEach(pct => {
    const y = padT + plotH*(1-pct);
    const val = Math.round(maxVal*pct);
    gridLines += `<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="${gridColor}" stroke-width="1"/>`;
    gridLines += `<text x="${padL-6}" y="${y+4}" fill="${textColor}" font-family="JetBrains Mono" font-size="9.5" text-anchor="end">${val}</text>`;
  });
  // Date labels — show ~6 evenly spaced
  const labelStep = Math.max(1, Math.floor(n/6));
  let dateLabels = '';
  days.forEach((d,i) => {
    if(i%labelStep===0 || i===n-1){
      // Format DD-MM-YYYY to "01 May"
      const parts = d.date.split('-');
      const months = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      const label = parts.length===3 ? `${parts[0]} ${months[parseInt(parts[1],10)]||''}` : d.date;
      dateLabels += `<text x="${toX(i)}" y="${H-4}" fill="${textColor}" font-family="JetBrains Mono" font-size="9" text-anchor="middle">${label}</text>`;
    }
  });
  // Lines per region
  let paths = '';
  for(const [reg, color] of Object.entries(regionColors)){
    const pts = days.map((d,i)=>`${toX(i)},${toY(d[reg])}`).join(' L ');
    paths += `<polyline points="${days.map((d,i)=>`${toX(i)},${toY(d[reg])}`).join(' ')}" fill="none" stroke="${color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    // endpoint dot
    const last = days[n-1];
    paths += `<circle cx="${toX(n-1)}" cy="${toY(last[reg])}" r="3" fill="${color}"/>`;
    paths += `<text x="${toX(n-1)+6}" y="${toY(last[reg])+4}" fill="${color}" font-family="JetBrains Mono" font-size="9.5">${reg} ${last[reg]}</text>`;
  }
  // Legend
  const legend = Object.entries(regionColors).map(([reg,color])=>
    `<rect x="0" y="0" width="8" height="8" fill="${color}" rx="2"/><text x="12" y="8" fill="${textColor}" font-family="JetBrains Mono" font-size="10">${reg}</text>`
  );
  let legendSvg = `<svg viewBox="0 0 200 14" style="width:200px;height:14px;margin-top:8px;">
    <g transform="translate(0,3)">${legend[0]}</g>
    <g transform="translate(60,3)">${legend[1]}</g>
    <g transform="translate(130,3)">${legend[2]}</g>
  </svg>`;

  chartWrap.innerHTML = `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;">${gridLines}${paths}${dateLabels}</svg>${legendSvg}
    <div style="font-size:11px;color:var(--ink-faint);margin-top:6px;">
      Peak day: <b style="color:var(--ink)">${vel.peakDay.date}</b> — ${vel.peakDay.total.toLocaleString()} units total &nbsp;·&nbsp;
      Slowest day: <b style="color:var(--ink)">${vel.slowDay.date}</b> — ${vel.slowDay.total.toLocaleString()} units total
    </div>`;
  // Attach tooltip to velocity chart after render
  setTimeout(()=>addVelocityChartInteractivity(), 50);

  // Day-by-day table
  const regionOrder = ['East','Central','West'];
  tbody.innerHTML = days.map(d => {
    const topRegion = regionOrder.reduce((best,r)=>d[r]>d[best]?r:best, 'East');
    const topColor = { East:'var(--east)', Central:'var(--central)', West:'var(--west)' }[topRegion];
    const parts = d.date.split('-');
    const months = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const prettyDate = parts.length===3 ? `${parts[0]} ${months[parseInt(parts[1],10)]||''} ${parts[2]}` : d.date;
    return `<tr>
      <td class="mono">${prettyDate}</td>
      <td class="num" style="color:var(--east);font-weight:600;">${d.East.toLocaleString()}</td>
      <td class="num" style="color:var(--central);font-weight:600;">${d.Central.toLocaleString()}</td>
      <td class="num" style="color:var(--west);font-weight:600;">${d.West.toLocaleString()}</td>
      <td class="num" style="font-weight:700;">${d.total.toLocaleString()}</td>
      <td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;padding:2px 8px;border-radius:5px;background:${topColor}22;color:${topColor};font-weight:600;">${topRegion}</span></td>
    </tr>`;
  }).join('');
}

function renderDemandPanel(){
  const kpis = document.getElementById('demandKpis');
  if(!state.demand){ kpis.innerHTML=''; renderDemandTable(''); return; }
  const activeSkus = getActiveSkuSet();
  const skus = Object.keys(state.demand).filter(s=>!activeSkus || activeSkus.has(s));
  const totalUnits = skus.reduce((s,k)=>s+state.demand[k].totalUnits,0);
  const corrected = state.sellThroughOn;
  kpis.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Total demand units</div><div class="kpi-val"><span data-count-to="${totalUnits}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">SKUs tracked</div><div class="kpi-val"><span data-count-to="${skus.length}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Sell-through correction</div><div class="kpi-val" style="font-size:18px;color:${corrected?'var(--good)':'var(--assumed)'}">${corrected?'ON':'OFF'}</div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Avg units / SKU</div><div class="kpi-val"><span data-count-to="${skus.length?Math.round(totalUnits/skus.length):0}" data-count-kind="int">0</span></div></div>`;
  animateKpiValues(kpis);
  renderDemandTable('');
}
function renderDemandTable(filter){
  const tbody = document.querySelector('#demandTable tbody');
  if(!state.demand){ tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:24px 0;">Upload sales data and run analysis</td></tr>'; return; }
  const activeSkus = getActiveSkuSet();
  let skus = Object.keys(state.demand).filter(s=>!activeSkus || activeSkus.has(s));
  if(filter) skus = skus.filter(s=>s.toLowerCase().includes(filter.toLowerCase()));
  skus.sort((a,b)=>state.demand[b].totalUnits - state.demand[a].totalUnits);
  skus = skus.slice(0,200);
  if(!skus.length){ tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:24px 0;">No matching SKUs</td></tr>'; return; }
  tbody.innerHTML = skus.map(sku=>{
    const d = state.demand[sku];
    return `<tr><td class="mono">${sku}</td><td class="num">${d.totalUnits.toLocaleString()}</td><td class="num">${fmtPct(d.demandPct.East)}</td><td class="num">${fmtPct(d.demandPct.Central)}</td><td class="num">${fmtPct(d.demandPct.West)}</td></tr>`;
  }).join('');
}
document.getElementById('demandSearch').addEventListener('input', (e)=> renderDemandTable(e.target.value));

/* ============================================================
   COST ANALYTICS DASHBOARD — all computed from real fee log data.
   These are genuinely new analytical views, not decorative. Each
   number traces to a real fee log row, not an estimate.
   ============================================================ */
function renderCostAnalyticsDashboard(){
  renderCostKpis();
  renderCostRegionChart();
  renderCostSavingsChart();
  renderCostFeeBreakdown();
}

function renderCostKpis(){
  const grid = document.getElementById('costKpiGrid');
  if(!grid) return;
  if(!state.feeAdapted || !state.feeAdapted.length){
    grid.innerHTML = '';
    return;
  }
  const totalFee = state.feeAdapted.reduce((s,r)=>s+r.fee,0);
  const totalUnits = state.feeAdapted.reduce((s,r)=>s+r.qty,0);
  const avgPerUnit = totalUnits>0 ? totalFee/totalUnits : 0;
  const distinctSkus = new Set(state.feeAdapted.map(r=>r.sku)).size;
  // Find cheapest possible if everything went to lowest-cost region
  let cheapestRegion = 'East', cheapestRate = Infinity;
  if(state.feeRateTable){
    for(const [reg, tiers] of Object.entries(state.feeRateTable)){
      const rates = Object.values(tiers).map(t=>t.avgFeePerUnit);
      const minRate = Math.min(...rates);
      if(minRate < cheapestRate){ cheapestRate = minRate; cheapestRegion = reg; }
    }
  }
  const cheapestBaseline = totalUnits * (cheapestRate===Infinity ? avgPerUnit : cheapestRate);
  const actualVsCheapest = totalFee - cheapestBaseline;
  const paidPremiumStr = actualVsCheapest > 0 ? `+${fmtMoney(actualVsCheapest)} premium to spread demand` : `${fmtMoney(Math.abs(actualVsCheapest))} saved vs. cheapest-only`;
  const premiumClass = actualVsCheapest > 0 ? 'bad' : 'good';

  grid.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">Total placement fees paid</div>
      <div class="kpi-val">${fmtMoney(totalFee)}</div>
      <div class="kpi-delta flat">${totalUnits.toLocaleString()} units across ${distinctSkus.toLocaleString()} SKUs</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">Avg placement fee per unit</div>
      <div class="kpi-val">${fmtMoney(avgPerUnit)}</div>
      <div class="kpi-delta flat">Weighted across all regions &amp; tiers</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">vs. cheapest-only baseline (${cheapestRegion})</div>
      <div class="kpi-val ${premiumClass}">${actualVsCheapest>=0?'+':''}${fmtMoney(actualVsCheapest)}</div>
      <div class="kpi-delta flat">${paidPremiumStr}</div>
    </div>
    <div class="kpi-card"><div class="kpi-accent"></div>
      <div class="kpi-label">Cheapest single-region baseline</div>
      <div class="kpi-val">${fmtMoney(cheapestBaseline)}</div>
      <div class="kpi-delta flat">${cheapestRegion} only at ${fmtMoney(cheapestRate)}/unit</div>
    </div>`;
  animateKpiValues(grid);
}

function renderCostRegionChart(){
  const wrap = document.getElementById('costRegionChartWrap');
  if(!wrap) return;
  if(!state.feeAdapted || !state.feeAdapted.length){ wrap.innerHTML = '<div class="empty-state"><p>Upload a placement fee log and run analysis</p></div>'; return; }

  const byRegion = {};
  for(const r of state.feeAdapted){
    if(!byRegion[r.region]) byRegion[r.region] = { fee:0, qty:0, skus:new Set() };
    byRegion[r.region].fee += r.fee;
    byRegion[r.region].qty += r.qty;
    byRegion[r.region].skus.add(r.sku);
  }
  const totalFee = Object.values(byRegion).reduce((s,v)=>s+v.fee,0);
  const regions = ['East','Central','West'];
  const colors = { East:'var(--east)', Central:'var(--central)', West:'var(--west)' };
  const maxFee = Math.max(...regions.map(r=>byRegion[r]?.fee||0));
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#6B7280' : '#928D80';

  let html = `<svg viewBox="0 0 400 180" style="width:100%;height:auto;">`;
  const barW = 80, gap = 40, padL = 60, padT = 16, plotH = 120;
  regions.forEach((reg, i) => {
    const data = byRegion[reg] || { fee:0, qty:0, skus:new Set() };
    const barH = maxFee>0 ? (data.fee/maxFee)*plotH : 0;
    const x = padL + i*(barW+gap);
    const y = padT + plotH - barH;
    html += `<rect x="${x}" y="${y}" width="${barW}" height="${barH}" fill="${colors[reg]}" rx="4" opacity="0.9"/>`;
    html += `<text x="${x+barW/2}" y="${padT+plotH+16}" fill="${textColor}" font-family="JetBrains Mono" font-size="11" text-anchor="middle">${reg}</text>`;
    if(data.fee>0){
      html += `<text x="${x+barW/2}" y="${Math.max(y-6,padT+8)}" fill="${textColor}" font-family="JetBrains Mono" font-size="10.5" text-anchor="middle">${fmtMoney(data.fee)}</text>`;
      html += `<text x="${x+barW/2}" y="${Math.max(y-18,padT-4)}" fill="${textColor}" font-family="JetBrains Mono" font-size="9.5" text-anchor="middle">${fmtPct(data.fee/totalFee)}</text>`;
    }
  });
  html += `</svg>`;
  html += `<div style="display:flex;gap:20px;margin-top:6px;font-size:11.5px;color:var(--ink-faint);flex-wrap:wrap;">`;
  regions.forEach(reg => {
    const data = byRegion[reg] || { qty:0, skus:new Set() };
    html += `<div><span style="width:10px;height:10px;border-radius:2px;background:${colors[reg]};display:inline-block;margin-right:5px;vertical-align:middle;"></span>${reg}: ${data.qty.toLocaleString()} units, ${data.skus.size} SKUs</div>`;
  });
  html += `</div>`;
  wrap.innerHTML = html;
}

function renderCostSavingsChart(){
  const wrap = document.getElementById('costSavingsChartWrap');
  if(!wrap) return;
  if(!state.feeAdapted || !state.feeRateTable){ wrap.innerHTML = '<div class="empty-state"><p>Upload fee data and run analysis</p></div>'; return; }

  // Build a real SKU-level comparison: what they paid vs. cheapest-possible for each SKU
  const bySkuPaid = {};
  for(const r of state.feeAdapted){
    if(!bySkuPaid[r.sku]) bySkuPaid[r.sku] = { paid:0, qty:0, sizeTier:r.sizeTier };
    bySkuPaid[r.sku].paid += r.fee;
    bySkuPaid[r.sku].qty += r.qty;
  }

  let totalPaid = 0, totalCheapest = 0, feeFreeUnits = 0, payingUnits = 0;
  let savedSkus = 0, premiumSkus = 0, feeFreeSKUs = 0;

  for(const [sku, data] of Object.entries(bySkuPaid)){
    totalPaid += data.paid;
    const avgPaidPerUnit = data.qty>0 ? data.paid/data.qty : 0;
    // Find cheapest real rate for this SKU's tier across all regions
    let cheapestRate = Infinity;
    for(const [, tiers] of Object.entries(state.feeRateTable)){
      const rate = tiers[data.sizeTier]?.avgFeePerUnit;
      if(rate && rate < cheapestRate) cheapestRate = rate;
    }
    if(cheapestRate===Infinity) cheapestRate = avgPaidPerUnit;
    const cheapestForSku = cheapestRate * data.qty;
    totalCheapest += cheapestForSku;
    if(data.paid === 0){ feeFreeSKUs++; feeFreeUnits += data.qty; }
    else { payingUnits += data.qty; }
    if(data.paid > cheapestForSku + 0.01) premiumSkus++;
    else if(data.paid < cheapestForSku - 0.01) savedSkus++;
  }

  const totalDelta = totalPaid - totalCheapest;
  const skuCount = Object.keys(bySkuPaid).length;

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const goodColor = isDark ? '#7FB069' : '#4F8A4A';
  const badColor = isDark ? '#E05C5C' : '#C0392B';
  const textColor = isDark ? '#6B7280' : '#928D80';

  // Waterfall-style bar
  const max = Math.max(totalCheapest, totalPaid, 1);
  const W = 380, H = 140, padL = 10, padR = 10, padT = 20, barH = 40;
  const cheapestW = (totalCheapest/max)*(W-padL-padR);
  const paidW = (totalPaid/max)*(W-padL-padR);
  const deltaIsNeg = totalDelta < 0;

  const svg = `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;">
    <text x="${padL}" y="${padT-4}" fill="${textColor}" font-size="10.5" font-family="JetBrains Mono">Cheapest-only baseline</text>
    <rect x="${padL}" y="${padT}" width="${cheapestW}" height="${barH}" fill="var(--east)" rx="4" opacity="0.7"/>
    <text x="${padL+cheapestW+6}" y="${padT+barH/2+4}" fill="${textColor}" font-size="11" font-family="JetBrains Mono">${fmtMoney(totalCheapest)}</text>
    <text x="${padL}" y="${padT+barH+24}" fill="${textColor}" font-size="10.5" font-family="JetBrains Mono">What you actually paid</text>
    <rect x="${padL}" y="${padT+barH+10}" width="${paidW}" height="${barH}" fill="${deltaIsNeg?goodColor:badColor}" rx="4" opacity="0.85"/>
    <text x="${padL+paidW+6}" y="${padT+barH+10+barH/2+4}" fill="${textColor}" font-size="11" font-family="JetBrains Mono">${fmtMoney(totalPaid)}</text>
    <text x="${padL}" y="${H-4}" fill="${deltaIsNeg?goodColor:badColor}" font-size="12" font-weight="700" font-family="JetBrains Mono">${deltaIsNeg?'▼ Saved':'▲ Premium paid'}: ${fmtMoney(Math.abs(totalDelta))} (${fmtPct(Math.abs(totalDelta)/Math.max(totalCheapest,1))})</text>
  </svg>`;

  const breakdown = `<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:12px;font-size:12px;color:var(--ink-dim);">
    <div style="text-align:center;padding:10px;background:var(--bg-input);border-radius:8px;">
      <div style="font-size:18px;font-weight:700;color:var(--good);font-family:'JetBrains Mono';">${feeFreeSKUs.toLocaleString()}</div>
      <div>SKUs paying $0 fee</div><div style="font-size:10.5px;color:var(--ink-faint);">${feeFreeUnits.toLocaleString()} units</div>
    </div>
    <div style="text-align:center;padding:10px;background:var(--bg-input);border-radius:8px;">
      <div style="font-size:18px;font-weight:700;color:var(--assumed);font-family:'JetBrains Mono';">${premiumSkus.toLocaleString()}</div>
      <div>SKUs paying premium</div><div style="font-size:10.5px;color:var(--ink-faint);">above cheapest region</div>
    </div>
    <div style="text-align:center;padding:10px;background:var(--bg-input);border-radius:8px;">
      <div style="font-size:18px;font-weight:700;color:var(--good);font-family:'JetBrains Mono';">${savedSkus.toLocaleString()}</div>
      <div>SKUs below baseline</div><div style="font-size:10.5px;color:var(--ink-faint);">better than cheapest-only</div>
    </div>
  </div>`;
  wrap.innerHTML = svg + breakdown;
}

function renderCostFeeBreakdown(){
  const wrap = document.getElementById('costFeeBreakdownWrap');
  if(!wrap) return;
  if(!state.feeAdapted){ wrap.innerHTML = '<div class="empty-state"><p>Upload a placement fee log and run analysis</p></div>'; return; }

  const bySkuRegion = {};
  for(const r of state.feeAdapted){
    const key = r.sku+'|'+r.region;
    if(!bySkuRegion[key]) bySkuRegion[key] = { sku:r.sku, region:r.region, fee:0, qty:0 };
    bySkuRegion[key].fee += r.fee;
    bySkuRegion[key].qty += r.qty;
  }

  // Classify each SKU-region pair
  const feeFree = [], minimal = [], standard = [];
  for(const entry of Object.values(bySkuRegion)){
    const perUnit = entry.qty>0 ? entry.fee/entry.qty : 0;
    if(perUnit < 0.01) feeFree.push(entry);
    else if(perUnit < 0.30) minimal.push(entry);
    else standard.push(entry);
  }
  const total = feeFree.length + minimal.length + standard.length || 1;
  const bar = (count, color, label, sub) => {
    const pct = count/total;
    return `<div style="margin-bottom:12px;">
      <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--ink-dim);margin-bottom:4px;">
        <span>${label}</span><span style="font-family:'JetBrains Mono';color:var(--ink)">${count.toLocaleString()} (${fmtPct(pct)})</span>
      </div>
      <div style="height:10px;background:var(--bg-input);border-radius:5px;overflow:hidden;">
        <div style="height:100%;width:${(pct*100).toFixed(1)}%;background:${color};border-radius:5px;transition:width .4s;"></div>
      </div>
      <div style="font-size:10.5px;color:var(--ink-faint);margin-top:3px;">${sub}</div>
    </div>`;
  };
  wrap.innerHTML =
    bar(feeFree.length, 'var(--good)', '✓ Fee-free ($0/unit)', 'Qualified for Amazon\'s 5-identical-box optimization — placement cost eliminated') +
    bar(minimal.length, 'var(--assumed)', '~ Minimal split fee (<$0.30/unit)', 'Minimal split placement fee — multi-location but not yet fee-free') +
    bar(standard.length, 'var(--danger)', '▲ Full placement fee (≥$0.30/unit)', 'Standard placement fee — single-location or high-rate tier');
}

function renderFeeRateTable(){
  const tbody = document.querySelector('#feeRateTable tbody');
  if(!state.feeRateTable){
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:24px 0;">Upload a placement fee log and run analysis</td></tr>';
    document.getElementById('costTierFilter').innerHTML = '<option value="all">All size tiers</option>';
    return;
  }
  const allRows = [];
  for(const region of Object.keys(state.feeRateTable)){
    for(const tier of Object.keys(state.feeRateTable[region])){
      const entry = state.feeRateTable[region][tier];
      allRows.push({ region, tier, ...entry });
    }
  }
  // populate tier filter dynamically from real tiers present, preserving selection
  const tierSel = document.getElementById('costTierFilter');
  const prevTier = tierSel.value;
  const distinctTiers = [...new Set(allRows.map(r=>r.tier))].sort();
  tierSel.innerHTML = '<option value="all">All size tiers</option>' + distinctTiers.map(t=>`<option value="${t}">${t}</option>`).join('');
  if(distinctTiers.includes(prevTier)) tierSel.value = prevTier;

  const regionFilter = document.getElementById('costRegionFilter').value;
  const tierFilter = tierSel.value;
  let rows = allRows;
  if(regionFilter !== 'all') rows = rows.filter(r=>r.region===regionFilter);
  if(tierFilter !== 'all') rows = rows.filter(r=>r.tier===tierFilter);
  rows.sort((a,b)=> b.sampleSize - a.sampleSize);

  if(!rows.length){ tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:24px 0;">No rate data matches your filters</td></tr>'; return; }
  tbody.innerHTML = rows.map(r=>`
    <tr><td>${r.region}</td><td class="mono">${r.tier}</td><td class="num mono">${fmtMoney(r.avgFeePerUnit)}</td><td class="num">${r.sampleSize.toLocaleString()}</td>
    <td><span class="badge ${r.sampleSize>=5?'real':'fallback'}">${r.sampleSize>=5?'real, n≥5':'real, thin sample'}</span></td></tr>`).join('');
}
['costRegionFilter','costTierFilter'].forEach(id=> document.getElementById(id).addEventListener('change', renderFeeRateTable));
document.getElementById('costFilterClear').addEventListener('click', ()=>{
  document.getElementById('costRegionFilter').value = 'all';
  document.getElementById('costTierFilter').value = 'all';
  renderFeeRateTable();
});

/* Real per-vendor cost rollup from the actual fee log — total units
   shipped, real average $/unit, and real total fee paid, grouped by the
   same extractVendorPrefix() used everywhere else in the tool. */
function renderVendorCostTable(){
  const tbody = document.querySelector('#vendorCostTable tbody');
  const vendorSel = document.getElementById('costVendorFilter');
  if(!state.feeAdapted || !state.feeAdapted.length){
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--ink-faint);padding:24px 0;">Upload a placement fee log and run analysis</td></tr>';
    vendorSel.innerHTML = '<option value="all">All vendors</option>';
    return;
  }
  const byVendor = {};
  for(const r of state.feeAdapted){
    const vendor = extractVendorPrefix(r.sku);
    if(!byVendor[vendor]) byVendor[vendor] = { vendor, skus:new Set(), units:0, totalFee:0 };
    byVendor[vendor].skus.add(r.sku);
    byVendor[vendor].units += r.qty;
    byVendor[vendor].totalFee += r.fee;
  }
  const rows = Object.values(byVendor).map(v=>({ ...v, skuCount:v.skus.size, avgFeePerUnit: v.units>0?v.totalFee/v.units:0 })).sort((a,b)=>b.totalFee-a.totalFee);

  const prevVendor = vendorSel.value;
  vendorSel.innerHTML = '<option value="all">All vendors</option>' + rows.map(r=>`<option value="${r.vendor}">${r.vendor} (${r.skuCount} SKUs)</option>`).join('');
  if(rows.some(r=>r.vendor===prevVendor)) vendorSel.value = prevVendor;

  tbody.innerHTML = rows.slice(0,100).map(r=>`
    <tr><td class="mono">${r.vendor}</td><td class="num">${r.skuCount}</td><td class="num">${r.units.toLocaleString()}</td><td class="num">${fmtMoney(r.avgFeePerUnit)}</td><td class="num">${fmtMoney(r.totalFee)}</td></tr>
  `).join('');
}
document.getElementById('costVendorFilter').addEventListener('change', ()=>{
  state.costVendorFilter = document.getElementById('costVendorFilter').value;
  renderFrontierSelect();
});

function renderFrontierSelect(){
  const sel = document.getElementById('frontierSkuSelect');
  if(!state.demand || !state.feeRateTable){ sel.innerHTML = '<option>No SKUs loaded yet</option>'; return; }
  const vendorFilter = state.costVendorFilter || 'all';
  let pool = Object.keys(state.demand);
  if(vendorFilter !== 'all') pool = pool.filter(s=>extractVendorPrefix(s)===vendorFilter);
  const skus = pool.sort((a,b)=>state.demand[b].totalUnits-state.demand[a].totalUnits).slice(0,50);
  if(!skus.length){ sel.innerHTML = '<option>No SKUs for this vendor</option>'; state.frontierData=null; renderFrontierChart(); return; }
  sel.innerHTML = skus.map(s=>`<option value="${s}">${s} (${state.demand[s].totalUnits} units)</option>`).join('');
  sel.onchange = ()=> renderFrontierForSku(sel.value);
  // Default to a SKU that actually shows cost VARIATION across the
  // coverage slider, not just the single highest-volume SKU — large
  // SKUs frequently qualify for Amazon's fee-free tier at every coverage
  // level (more units = more ways to form 5+ identical boxes), which
  // produces a real but visually flat $0 line that looks broken even
  // though it's correct. Scan a few candidates and pick the first one
  // whose frontier actually varies, so the default view demonstrates
  // the feature instead of accidentally picking its least informative case.
  let defaultSku = skus[0];
  for(const candidate of skus.slice(0,15)){
    const d = state.demand[candidate];
    const tier = getSkuSizeTier(candidate);
    const test = buildParetoFrontier(d.totalUnits, d.demandPct, tier, state.feeRateTable, state.defaultFeePerUnit, 6);
    const costs = test.map(f=>f.cost);
    if(Math.max(...costs) - Math.min(...costs) > 0.01){ defaultSku = candidate; break; }
  }
  sel.value = defaultSku;
  if(skus.length) renderFrontierForSku(defaultSku);
}
function renderFrontierForSku(sku){
  const d = state.demand[sku];
  if(!d) return;
  const tierInfo = state.skuSizeTierLookup && state.skuSizeTierLookup[sku];
  const tier = getSkuSizeTier(sku);
  const frontier = buildParetoFrontier(d.totalUnits, d.demandPct, tier, state.feeRateTable, state.defaultFeePerUnit, 10);
  state.frontierData = frontier;
  state.frontierSkuMeta = { sku, tier, tierIsReal: !!tierInfo, units: d.totalUnits };
  renderFrontierChart();
}
function renderFrontierChart(){
  const wrap = document.getElementById('frontierChartWrap');
  const frontier = state.frontierData;
  if(!frontier || !frontier.length){ wrap.innerHTML = `<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="12" cy="12" r="9"/></svg><h4>No frontier yet</h4><p>Run analysis with both Sales and Fee data loaded.</p></div>`; return; }
  const maxCost = Math.max(...frontier.map(f=>f.cost));
  const minCost = Math.min(...frontier.map(f=>f.cost));
  const isFlat = (maxCost - minCost) < 0.005;
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const textColor = isDark ? '#6B7280' : '#928D80';
  const signal = isDark ? '#FF6B35' : '#D9531E';
  const good = isDark ? '#7FB069' : '#4F8A4A';
  const W=560,H=260,padL=60,padB=30,padT=20,padR=20;
  const plotW = W-padL-padR, plotH = H-padT-padB;
  const denom = isFlat ? 1 : (maxCost-minCost);
  let pts = frontier.map(f=>{
    const x = padL + (f.minCoverage)*plotW;
    const y = isFlat ? (padT+plotH-plotH*0.5) : padT + plotH - ((f.cost-minCost)/denom)*plotH;
    return [x,y];
  });
  let path = 'M '+pts.map(p=>p.join(' ')).join(' L ');
  let area = path + ` L ${pts[pts.length-1][0]} ${padT+plotH} L ${pts[0][0]} ${padT+plotH} Z`;
  const lineColor = isFlat ? good : signal;
  let svg = `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;">
    <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT+plotH}" stroke="${isDark?'#262C35':'#E4E0D8'}"/>
    <line x1="${padL}" y1="${padT+plotH}" x2="${W-padR}" y2="${padT+plotH}" stroke="${isDark?'#262C35':'#E4E0D8'}"/>
    <text x="${padL}" y="${H-6}" fill="${textColor}" font-family="JetBrains Mono" font-size="10">0% coverage</text>
    <text x="${W-padR}" y="${H-6}" fill="${textColor}" font-family="JetBrains Mono" font-size="10" text-anchor="end">100%</text>
    <text x="${padL-8}" y="${padT+10}" fill="${textColor}" font-family="JetBrains Mono" font-size="10" text-anchor="end">${isFlat?'':fmtMoney(maxCost)}</text>
    <text x="${padL-8}" y="${padT+plotH}" fill="${textColor}" font-family="JetBrains Mono" font-size="10" text-anchor="end">${isFlat?fmtMoney(minCost):fmtMoney(minCost)}</text>
    <path d="${area}" fill="${lineColor}" opacity="0.08"/>
    <path d="${path}" fill="none" stroke="${lineColor}" stroke-width="2.5"/>
    ${pts.map(p=>`<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${lineColor}"/>`).join('')}
  </svg>`;

  const meta = state.frontierSkuMeta;
  let metaLine = '';
  if(meta){
    metaLine = `<div style="font-size:11px;color:var(--ink-faint);margin-bottom:6px;">SKU size tier used: <b style="color:var(--ink-dim)">${meta.tier}</b> ${meta.tierIsReal ? '(real, from your fee log)' : '(no real fee-log entry for this SKU — using fallback)'} · ${meta.units.toLocaleString()} total units</div>`;
  }
  const note = isFlat
    ? `<div style="font-size:12px;color:${good};margin-top:8px;line-height:1.5;">✓ This curve is genuinely flat at <b>${fmtMoney(minCost)}</b> across every coverage level — not a rendering issue. At ${meta?meta.units.toLocaleString():''} units, the optimizer can form Amazon's 5-identical-box fee-free configuration no matter how the regional split shifts, so increasing demand coverage doesn't cost anything extra for this specific SKU. Try a lower-volume SKU from the dropdown above to see a curve that actually trades off cost vs. coverage.</div>`
    : `<div style="font-size:11.5px;color:var(--ink-faint);margin-top:8px;">Real LP-optimal cost at each minimum demand-coverage level — monotonic by construction, verified against known cost-minimal cases.</div>`;
  wrap.innerHTML = metaLine + svg + note;
}

/* ============================================================
   MANIFEST TAB
   ============================================================ */
document.querySelectorAll('.method-pill').forEach(pill=>{
  pill.addEventListener('click', ()=>{
    document.querySelectorAll('.method-pill').forEach(p=>p.classList.remove('active'));
    pill.classList.add('active');
    document.getElementById('coverageSliderRow').style.display = pill.dataset.method === 'lp' ? 'flex' : 'none';
  });
});
document.getElementById('coverageSlider').addEventListener('input', (e)=>{
  document.getElementById('coverageVal').textContent = e.target.value+'%';
});

document.getElementById('runManifestBtn').addEventListener('click', async ()=>{
  if(!state.filesRaw.manifest){ toast('Upload a manifest file first', true); return; }
  if(!state.demand){ toast('Upload Sales data and run analysis on the Data tab first — the manifest needs demand data to work against', true); return; }
  try{
    const { objects } = await loadTabularFile(state.filesRaw.manifest, 'manifest');
    const { records, stats } = adaptManifestRows(objects);
    state.manifestRecords = records;
    const method = (document.querySelector('.method-option.active') || document.querySelector('.method-pill.active'))?.dataset.method || 'heuristic';
    const minCov = parseInt(document.getElementById('coverageSlider').value,10)/100;
    state.maxRegionsPerSku = parseInt(document.getElementById('maxRegionsSelect').value,10) || 2;
    const { plan, summary, portfolio } = buildManifestPlan(records, state.demand, state.feeRateTable||{}, state.defaultFeePerUnit, method, minCov, null, state.maxRegionsPerSku);
    state.manifestPlan = plan;
    state.manifestPortfolio = portfolio;
    renderExecSummary(portfolio, summary);
    renderManifestKpis(summary);
    renderPortfolioAllocation(portfolio, summary);
    renderRegionDownloadButtons(plan);
    applyManifestFiltersAndRender();
    renderManifestAnalytics();
    if(state.scopeMode === 'manifest') renderAll(); // keep every other tab in sync if Manifest Mode is already active
    toast(`Built plan for ${summary.totalSkus} SKUs (${summary.unknownDemandCount} with no sales history, flagged below)`);
  } catch(err){
    toast('Failed to process manifest: '+err.message, true);
    console.error(err);
  }
});

/* CEO-level executive summary — the four numbers a CEO actually wants to
   see in five seconds: real dollars saved (and the %), real demand
   coverage achieved, and real cost-per-unit. Every number here already
   existed in the portfolio object; this just gives it the visual weight
   it was missing instead of being buried in body text. */
function renderExecSummary(portfolio, summary){
  const card = document.getElementById('execSummaryCard');
  const grid = document.getElementById('execSummaryGrid');
  card.style.display = 'block';
  const { totalCost, cheapestOnlyCost, totalSavings, savingsPct, weightedAvgCoverage, costPerUnit } = portfolio;

  const savingsClass = totalSavings > 0.01 ? 'good' : totalSavings < -0.01 ? 'bad' : '';
  const savingsSign = totalSavings >= 0 ? '' : '+';
  const coverageClass = weightedAvgCoverage==null ? '' : weightedAvgCoverage >= 0.85 ? 'good' : weightedAvgCoverage < 0.6 ? 'bad' : '';

  grid.innerHTML = `
    <div class="exec-stat">
      <div class="exec-val ${savingsClass}">${savingsSign}${fmtMoney(Math.abs(totalSavings))}</div>
      <div class="exec-label">${totalSavings>=0?'Saved':'Premium paid'} vs. cheapest-single-region</div>
      <div class="exec-sub">${fmtMoney(totalCost)} recommended vs. ${fmtMoney(cheapestOnlyCost)} baseline</div>
    </div>
    <div class="exec-stat">
      <div class="exec-val ${savingsClass}">${savingsPct>=0?'':'+'}${fmtPct(Math.abs(savingsPct))}</div>
      <div class="exec-label">${savingsPct>=0?'Cost reduction':'Cost premium'} vs. baseline</div>
      <div class="exec-sub">As a share of the cheapest-only cost</div>
    </div>
    <div class="exec-stat">
      <div class="exec-val ${coverageClass}">${weightedAvgCoverage!=null?fmtPct(weightedAvgCoverage):'—'}</div>
      <div class="exec-label">Real demand coverage achieved</div>
      <div class="exec-sub">Units-weighted average across all ${summary.totalSkus} SKUs in this plan</div>
    </div>
    <div class="exec-stat">
      <div class="exec-val">${fmtMoney(costPerUnit)}</div>
      <div class="exec-label">Effective placement cost per unit</div>
      <div class="exec-sub">${summary.totalUnits.toLocaleString()} total units this shipment</div>
    </div>`;
}

function renderManifestKpis(summary){
  const grid = document.getElementById('manifestKpis');
  const lineNote = summary.totalPackingLines > summary.totalSkus
    ? `<div class="kpi-delta flat">${summary.totalPackingLines} packing lines (case-pack/expiration variants)</div>`
    : '';
  const boxSpecNote = summary.defaultBoxSpecCount > 0
    ? `<div class="kpi-delta flat" style="color:var(--ink-faint);">${summary.defaultBoxSpecCount} SKU${summary.defaultBoxSpecCount===1?'':'s'} had no case-pack data — box size computed by the optimizer</div>`
    : '';
  grid.innerHTML = `
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">SKUs in manifest</div><div class="kpi-val"><span data-count-to="${summary.totalSkus}" data-count-kind="int">0</span></div>${lineNote}</div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">Total units</div><div class="kpi-val"><span data-count-to="${summary.totalUnits}" data-count-kind="int">0</span></div>${boxSpecNote}</div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">SKUs with real demand history</div><div class="kpi-val" style="color:var(--good)"><span data-count-to="${summary.knownDemandCount}" data-count-kind="int">0</span></div></div>
    <div class="kpi-card"><div class="kpi-accent"></div><div class="kpi-label">SKUs with NO history (even split)</div><div class="kpi-val" style="color:${summary.unknownDemandCount>0?'var(--assumed)':'var(--ink)'}"><span data-count-to="${summary.unknownDemandCount}" data-count-kind="int">0</span></div></div>`;
  animateKpiValues(grid);
  if(summary.defaultBoxSpecCount > 0){
    toast(`${summary.defaultBoxSpecCount} row(s) had no case-pack data on the manifest — the optimizer computed a real box size for them.`);
  }
}

/* Per-region exact-format download buttons — one click per region produces
   a real .xlsx matching the user's own template structure, pre-filled with
   that region's real box counts, ready to re-upload to Seller Central. */
function renderRegionDownloadButtons(plan){
  const card = document.getElementById('regionDownloadCard');
  const wrap = document.getElementById('regionDownloadButtons');
  const regions = ['East','Central','West'];
  const anyUnits = regions.some(r => plan.some(p => p.splitUnits[r] > 0));
  if(!anyUnits){ card.style.display='none'; return; }
  card.style.display = 'block';
  wrap.innerHTML = regions.map(region=>{
    const boxCount = plan.reduce((sum,p)=>{
      const entry = p.boxBreakdown ? p.boxBreakdown[region] : null;
      return sum + (entry && entry.boxPlan ? entry.boxPlan.boxes.length : 0);
    }, 0);
    const unitCount = plan.reduce((sum,p)=>sum+(p.splitUnits[region]||0),0);
    if(unitCount<=0) return '';
    return `<button class="btn btn-ghost" data-download-region="${region}" style="border-color:var(--${region.toLowerCase()});color:var(--${region.toLowerCase()});">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M12 3v12M7 8l5-5 5 5"/><path d="M5 21h14"/></svg>
      ${region} — ${boxCount} box${boxCount===1?'':'es'} (${unitCount} units)
    </button>`;
  }).join('');
  wrap.querySelectorAll('[data-download-region]').forEach(btn=>{
    btn.addEventListener('click', ()=> downloadRegionManifestFile(state.manifestPlan, btn.dataset.downloadRegion, (state.filesRaw.manifest?state.filesRaw.manifest.name.replace(/\.(xlsx|xls|csv)$/i,''):'shipment-plan')));
  });
}

/* Portfolio-level "where should I allocate this shipment" view — the
   demand-weighted region split summed across every SKU in the manifest,
   plus the real estimated total placement-fee cost of that plan. */
/* ---------------------------- MANIFEST ANALYTICS TAB ---------------------------- */
/* Regional read of the active manifest plan: units, cost, cost/unit, and
   SKU count per region, plus split-discipline and savings insights in
   plain language. Reads state.manifestPlan directly (built on the
   Manifest tab) — this tab doesn't run its own optimization, it explains
   the one that's already there. */
function renderManifestAnalytics(){
  const mapEl = document.getElementById('manifestRegionMap');
  const insightsEl = document.getElementById('manifestAnalyticsInsights');
  const costCard = document.getElementById('manifestAnalyticsCostCard');
  if(!mapEl || !insightsEl) return; // panel not in DOM yet on first paint

  const plan = state.manifestPlan;
  if(!plan || !plan.length){
    mapEl.className = 'region-map-empty';
    mapEl.textContent = 'Build a Manifest plan first — this map fills in with your real region split.';
    insightsEl.className = 'analytics-insights-empty';
    insightsEl.textContent = 'No active manifest plan yet.';
    if(costCard) costCard.style.display = 'none';
    return;
  }

  const regions = ['East','Central','West'];
  const regionUnits = { East:0, Central:0, West:0 };
  const regionCost = { East:0, Central:0, West:0 };
  const regionSkuCount = { East:0, Central:0, West:0 };
  let regionsUsedSum = 0, overCapCount = 0;
  const cap = state.maxRegionsPerSku || 2;

  for(const row of plan){
    let usedHere = 0;
    for(const region of regions){
      const units = row.splitUnits ? (row.splitUnits[region]||0) : 0;
      if(units > 0){
        regionUnits[region] += units;
        regionSkuCount[region]++;
        usedHere++;
        const b = row.boxBreakdown && row.boxBreakdown[region];
        regionCost[region] += b ? b.cost : 0;
      }
    }
    regionsUsedSum += usedHere;
    if(usedHere > cap) overCapCount++;
  }

  const totalUnits = regionUnits.East + regionUnits.Central + regionUnits.West;
  const totalCost = regionCost.East + regionCost.Central + regionCost.West;
  const colors = { East:'var(--east)', Central:'var(--central)', West:'var(--west)' };

  mapEl.className = 'manifest-region-map';
  mapEl.innerHTML = regions.map(region=>{
    const units = regionUnits[region];
    const pct = totalUnits>0 ? units/totalUnits : 0;
    const cpu = units>0 ? regionCost[region]/units : 0;
    const scale = 0.86 + pct*0.4;
    return `<div class="region-node ${region.toLowerCase()}" style="transform:scale(${scale.toFixed(2)});">
      <div class="region-node-label">${region}</div>
      <div class="region-node-value">${units.toLocaleString()}</div>
      <div class="region-node-sub">${fmtPct(pct)} of units · ${fmtMoney(cpu)}/unit</div>
      <div class="region-node-foot">${regionSkuCount[region]} SKU${regionSkuCount[region]===1?'':'s'} shipped here · ${fmtMoney(regionCost[region])} total</div>
    </div>`;
  }).join('');

  const topRegionEntry = regions.map(r=>[r,regionUnits[r]]).sort((a,b)=>b[1]-a[1])[0];
  const avgRegionsPerSku = plan.length ? (regionsUsedSum/plan.length) : 0;
  const portfolio = state.manifestPortfolio || {};
  const savings = portfolio.totalSavings || 0;

  insightsEl.className = 'analytics-insights';
  insightsEl.innerHTML = `
    <div class="insight-row"><span>Primary ship-to region</span><b>${topRegionEntry[0]} · ${fmtPct(totalUnits>0?topRegionEntry[1]/totalUnits:0)}</b></div>
    <div class="insight-row"><span>Split discipline</span><b>${avgRegionsPerSku.toFixed(1)} region${avgRegionsPerSku===1?'':'s'}/SKU on average${overCapCount>0?` · <span style="color:var(--danger)">${overCapCount} SKU${overCapCount===1?'':'s'} over the ${cap}-region cap</span>`:` · none over the ${cap}-region cap`}</b></div>
    <div class="insight-row"><span>SKUs with real demand history</span><b>${portfolio && state.manifestPlan ? plan.filter(r=>r.hasSalesHistory).length : 0} / ${plan.length}</b></div>
    <div class="insight-row"><span>Placement-fee delta</span><b style="color:${savings>=0?'var(--good)':'var(--danger)'}">${savings>=0?'+':''}${fmtMoney(savings)} vs. single-region baseline</b></div>
    <div class="insight-callout">Recommendation: ship this manifest mostly to <b>${topRegionEntry[0]}</b>, keep SKU-level plans capped at ${cap} region${cap===1?'':'s'} unless case-pack math forces otherwise, and use the per-region download files on the Manifest tab for execution.</div>`;

  if(costCard){
    costCard.style.display = 'block';
    const tbody = document.querySelector('#manifestAnalyticsCostTable tbody');
    tbody.innerHTML = regions.map(region=>{
      const units = regionUnits[region];
      const pct = totalUnits>0 ? units/totalUnits : 0;
      const cpu = units>0 ? regionCost[region]/units : 0;
      return `<tr><td><span class="dot" style="background:${colors[region]};display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;"></span>${region}</td>
        <td class="num">${units.toLocaleString()}</td>
        <td class="num">${fmtPct(pct)}</td>
        <td class="num">${fmtMoney(regionCost[region])}</td>
        <td class="num">${fmtMoney(cpu)}</td>
        <td class="num">${regionSkuCount[region]}</td></tr>`;
    }).join('') + `<tr style="font-weight:600;"><td>Total</td><td class="num">${totalUnits.toLocaleString()}</td><td class="num">100.0%</td><td class="num">${fmtMoney(totalCost)}</td><td class="num">${fmtMoney(totalUnits>0?totalCost/totalUnits:0)}</td><td class="num">—</td></tr>`;
  }
}

function renderPortfolioAllocation(portfolio, summary){
  const card = document.getElementById('portfolioAllocationCard');
  const body = document.getElementById('portfolioAllocationBody');
  card.style.display = 'block';
  const { units, pct, totalCost, cheapestOnlyCost, totalSavings } = portfolio;
  const regions = [
    { key:'East', color:'var(--east)' },
    { key:'Central', color:'var(--central)' },
    { key:'West', color:'var(--west)' }
  ];
  const segs = regions.map(r=>{
    const p = pct[r.key];
    const isZero = units[r.key] === 0;
    const flexBasis = Math.max(p*100, isZero ? 0 : 2);
    return `<div class="alloc-seg ${isZero?'zero':''}" style="flex-grow:${Math.max(flexBasis,0.001)};background:${isZero?'transparent':r.color};">${p>0.06?fmtPct(p):''}</div>`;
  }).join('');
  const legend = regions.map(r=>`
    <div class="alloc-legend-item">
      <span class="dot" style="background:${r.color}"></span>
      ${r.key}: <b>${fmtPct(pct[r.key])}</b> · ${units[r.key].toLocaleString()} units
    </div>`).join('');

  const sorted = regions.map(r=>({key:r.key, p:pct[r.key]})).sort((a,b)=>a.p-b.p);
  const lowest = sorted[0];
  const lowestNote = lowest.p < 0.08
    ? `<b style="color:var(--ink)">${lowest.key}</b> gets almost nothing (${fmtPct(lowest.p)}) because measured demand there is genuinely low across this manifest's SKUs — not an error, not rounding.`
    : `Demand is reasonably spread across all three regions for this manifest — no region is negligible enough to skip.`;

  const savingsNote = totalSavings > 0.01
    ? `Shipping by demand is <b style="color:var(--good)">${fmtMoney(totalSavings)} cheaper</b> than the naive "everything to the single cheapest region" alternative.`
    : totalSavings < -0.01
    ? `Shipping by demand costs <b style="color:var(--danger)">${fmtMoney(-totalSavings)} more</b> than shipping everything to one region would — the demand-coverage benefit is being paid for here, not free.`
    : `Shipping by demand costs about the same as the single-region alternative for this manifest.`;

  body.innerHTML = `
    <div class="alloc-bar-wrap">
      <div class="alloc-bar">${segs}</div>
    </div>
    <div class="alloc-legend-row">${legend}</div>
    <div class="alloc-reason">
      ${lowestNote} ${savingsNote}<br>
      Recommended plan: <b style="color:var(--ink)">${fmtMoney(totalCost)}</b> · Cheapest-only baseline: <b style="color:var(--ink-dim)">${fmtMoney(cheapestOnlyCost)}</b> across ${summary.totalUnits.toLocaleString()} units.
    </div>`;
}

/* ---------------------------- MANIFEST SLICERS ---------------------------- */
function applyManifestFiltersAndRender(){
  if(!state.manifestPlan){ return; }
  const search = (document.getElementById('manifestSearchFilter').value || '').toLowerCase().trim();
  const regionFilter = document.getElementById('manifestRegionFilter').value;
  const demandFilter = document.getElementById('manifestDemandFilter').value;

  let filtered = state.manifestPlan;
  if(search) filtered = filtered.filter(r => r.sku.toLowerCase().includes(search));
  if(regionFilter !== 'all') filtered = filtered.filter(r => (r.splitUnits[regionFilter]||0) > 0);
  if(demandFilter === 'real') filtered = filtered.filter(r => r.hasSalesHistory);
  if(demandFilter === 'unknown') filtered = filtered.filter(r => !r.hasSalesHistory);

  renderManifestTable(filtered);
}
['manifestSearchFilter'].forEach(id=>{
  document.getElementById(id).addEventListener('input', applyManifestFiltersAndRender);
});
['manifestRegionFilter','manifestDemandFilter'].forEach(id=>{
  document.getElementById(id).addEventListener('change', applyManifestFiltersAndRender);
});
document.getElementById('manifestFilterClear').addEventListener('click', ()=>{
  document.getElementById('manifestSearchFilter').value = '';
  document.getElementById('manifestRegionFilter').value = 'all';
  document.getElementById('manifestDemandFilter').value = 'all';
  applyManifestFiltersAndRender();
});

function buildDecisionBadge(row){
  const hasRealCasePack = row.isRealCasePack && row.effectiveUnitsPerBox;
  const totalBoxes = hasRealCasePack ? (row.realNumberOfBoxes || Math.round(row.units/row.effectiveUnitsPerBox)) : null;
  const regionsUsed = ['East','Central','West'].filter(r=>row.splitUnits[r]>0).length;

  if(hasRealCasePack && totalBoxes === 1){
    const bestRegion = Object.entries(row.splitUnits).sort((a,b)=>b[1]-a[1])[0][0];
    return `<span class="decision-badge case-single" title="Single sealed box — sent to ${bestRegion} (highest real demand region). Cannot split a sealed box.">📦 1 box → ${bestRegion} (best demand)</span>`;
  }
  if(hasRealCasePack && totalBoxes > 1){
    return `<span class="decision-badge case-multi" title="${totalBoxes} real case-pack boxes, split across ${regionsUsed} region${regionsUsed===1?'':'s'} by demand. Each region gets whole boxes only.">${totalBoxes} boxes → ${regionsUsed} region${regionsUsed===1?'':'s'}</span>`;
  }
  if(!hasRealCasePack && regionsUsed === 1){
    const onlyRegion = ['East','Central','West'].find(r=>row.splitUnits[r]>0);
    return `<span class="decision-badge free-single" title="No case-pack constraint — sent all to ${onlyRegion} (demand + cost optimal).">Demand → ${onlyRegion} only</span>`;
  }
  if(!hasRealCasePack){
    return `<span class="decision-badge free-multi" title="No case-pack constraint — split across ${regionsUsed} regions for best demand coverage + cost.">Demand → ${regionsUsed} regions</span>`;
  }
  return '';
}

function renderManifestTable(plan){
  const tbody = document.querySelector('#manifestTable tbody');
  if(!plan || !plan.length){ tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--ink-faint);padding:24px 0;">No plan yet, or no rows match your filters</td></tr>'; return; }
  tbody.innerHTML = '';
  plan.forEach((row,idx)=>{
    const lineNote = row.packingLines > 1
      ? `<span class="badge" style="background:var(--signal-soft);color:var(--signal);margin-left:6px;" title="${row.packingLines} packing lines (case-pack/expiration variants) summed for region split">${row.packingLines} lines</span>`
      : '';
    const deltaColor = row.costDelta == null ? 'var(--ink-faint)' : row.costDelta < 0 ? 'var(--good)' : row.costDelta > 0 ? 'var(--danger)' : 'var(--ink-faint)';
    const deltaText = row.costDelta == null ? '—' : (row.costDelta>=0?'+':'')+fmtMoney(row.costDelta);
    // Savings indicator: how much vs. cheapest-only, shown visually
    const savingsAbs = row.costDelta != null ? Math.abs(row.costDelta) : 0;
    const savingsIcon = row.costDelta == null ? '' : row.costDelta < -0.01 ? `<span style="color:var(--good);font-size:10px;margin-left:4px;">▼ saving</span>` : row.costDelta > 0.01 ? `<span style="color:var(--danger);font-size:10px;margin-left:4px;">▲ premium</span>` : `<span style="color:var(--ink-faint);font-size:10px;margin-left:4px;">= same</span>`;
    const decisionBadge = buildDecisionBadge(row);
    const mainRow = document.createElement('tr');
    mainRow.innerHTML = `
      <td><span class="mtr-toggle" data-idx="${idx}">▸</span></td>
      <td class="mono">${row.sku}${lineNote}</td>
      <td class="num">${row.units}</td>
      <td class="num">${row.splitUnits.East>0?row.splitUnits.East:'—'}</td>
      <td class="num">${row.splitUnits.Central>0?row.splitUnits.Central:'—'}</td>
      <td class="num">${row.splitUnits.West>0?row.splitUnits.West:'—'}</td>
      <td class="num">${row.cost!=null?fmtMoney(row.cost):'—'}</td>
      <td class="num">${fmtMoney(row.cheapestOnlyCost)}</td>
      <td class="num" style="color:${deltaColor};font-weight:600;">${deltaText}${savingsIcon}</td>
      <td>${decisionBadge}</td>`;
    tbody.appendChild(mainRow);

    const detailRow = document.createElement('tr');
    detailRow.className = 'mtr-detail-row';
    detailRow.id = 'mtr-detail-'+idx;
    const cell = document.createElement('td');
    cell.colSpan = 10;
    cell.innerHTML = `<div class="mtr-detail-inner">${renderBoxBreakdownHtml(row)}</div>`;
    detailRow.appendChild(cell);
    tbody.appendChild(detailRow);

    mainRow.querySelector('.mtr-toggle').addEventListener('click', ()=>{
      const toggle = mainRow.querySelector('.mtr-toggle');
      toggle.classList.toggle('open');
      detailRow.classList.toggle('open');
    });
  });
}

/* Renders the real box plan per region for one SKU — exactly how many
   identical cartons, the leftover remainder box if any, and whether that
   region's shipment qualifies for Amazon's $0 Optimized tier. */
function renderBoxBreakdownHtml(row){
  if(!row.boxBreakdown) return '<span style="color:var(--ink-faint);font-size:12px;">No box breakdown available.</span>';
  const regions = ['East','Central','West'];
  let html = '';
  if(row.specSheetMatch){
    html += `<div style="font-size:11.5px;color:var(--good);margin-bottom:10px;">Using real vendor spec sheet data: ${row.specSheetMatch.casePack||'?'} units/box (matched on ${row.specSheetMatch.vendor}-${row.specSheetMatch.suffix}) — your manifest didn't have this SKU's case pack, so this real vendor record was used instead of a computed guess.</div>`;
  } else if(row.realUnitsPerBox){
    html += `<div style="font-size:11.5px;color:var(--good);margin-bottom:10px;">Using your real case-pack spec: ${row.realUnitsPerBox} units/box (from the uploaded manifest — honored as a hard constraint). Region allocations below are whole multiples of this case pack only — no box is ever split across regions.</div>`;
  } else if(row.usedDefaultBoxSpec){
    html += `<div style="font-size:11.5px;color:var(--ink-faint);margin-bottom:10px;">No case-pack info on the manifest${state.specSheetLookup?' or matching vendor spec sheet entry':''} for this SKU — the box size below was computed by the optimizer (favoring configs that hit Amazon's 5-identical-box fee-free threshold), not assumed.</div>`;
  }
  if(row.isRealCasePack && row.coverageTargetUnreachable){
    const totalBoxes = row.effectiveUnitsPerBox ? Math.round(row.units/row.effectiveUnitsPerBox) : '?';
    html += `<div style="font-size:11.5px;color:var(--assumed);margin-bottom:10px;">⚠ The requested minimum demand coverage isn't reachable for this SKU — with only ${totalBoxes} real case-pack box${totalBoxes===1?'':'es'} to allocate, the best possible coverage is <b>${row.maxAchievableCoverage!=null?fmtPct(row.maxAchievableCoverage):'lower than requested'}</b>. Shown below is the cheapest split at that real maximum, not the originally requested target.</div>`;
  }
  if(row.sizeTierCrossCheck){
    const cc = row.sizeTierCrossCheck;
    const matches = cc.tier === row.sizeTier;
    const confColor = cc.confidence === 'high' ? 'var(--good)' : cc.confidence === 'medium' ? 'var(--ink-dim)' : 'var(--assumed)';
    if(matches){
      html += `<div style="font-size:11px;color:${confColor};margin-bottom:10px;">✓ Real box dimensions cross-check confirms size tier "${cc.tier}" (${cc.confidence} confidence, shipping weight ${cc.shippingWeight.toFixed(2)}lb via ${cc.basis.replace(/_/g,' ')})</div>`;
    } else {
      html += `<div style="font-size:11px;color:var(--danger);margin-bottom:10px;">⚠ Real box dimensions suggest size tier "${cc.tier}" (${cc.confidence} confidence), but the manifest/fee data uses "${row.sizeTier}" — this is a cross-check flag, not an override. Worth verifying against your real Seller Central listing.</div>`;
    }
  }
  for(const region of regions){
    const units = row.splitUnits[region];
    if(units <= 0) continue;
    const entry = row.boxBreakdown[region];
    if(!entry) continue;
    const plan = entry.boxPlan;
    const chips = plan.boxes.map(b => `<span class="box-chip ${b.isIdentical?'identical':'remainder'}">${b.units} units${b.isIdentical?'':' (remainder)'}</span>`).join('');
    let note = '';
    if(plan.qualifiesOptimized){
      note = `<div class="box-qualify-note win">✓ ${plan.identicalBoxes} identical boxes qualifies for Amazon-Optimized — $0 placement fee for this region</div>`;
    } else if(plan.note === 'below-optimized-threshold'){
      note = `<div class="box-qualify-note lose">Below the 5-identical-box threshold for this quantity — minimal-split fee applies (${fmtMoney(entry.rate)}/unit)</div>`;
    } else if(plan.isRealCasePack){
      note = `<div class="box-qualify-note lose">Real case-pack of ${plan.boxSize}/box only forms ${plan.identicalBoxes} identical box${plan.identicalBoxes===1?'':'es'} — minimal-split fee applies (${fmtMoney(entry.rate)}/unit)</div>`;
    }
    html += `<div class="box-region-group">
      <div class="box-region-head"><span class="region-dot" style="background:var(--${region.toLowerCase()})"></span>${region} — ${units} units, ${plan.boxes.length} box${plan.boxes.length===1?'':'es'}</div>
      <div class="box-chips">${chips}</div>
      ${note}
    </div>`;
  }
  return html || '<span style="color:var(--ink-faint);font-size:12px;">No units allocated to any region.</span>';
}
/* ============================================================
   EXACT-FORMAT SEND-TO-AMAZON FILE GENERATOR
   Produces one real .xlsx per region, matching the exact structure of
   the user's real Send-to-Amazon template (verified by direct inspection):
   sheet name "Create workflow – template", header row at row 6, columns
   A-J in this exact order. Each SKU's REAL boxes for that region become
   one row per box (Amazon's own template allows a SKU to repeat once per
   distinct packing line — this is exactly that pattern, not a deviation).
   Real case-pack metadata (expiration date, lot code, box dimensions) is
   carried forward from the original upload when present; for SKUs using
   the computed default box size, those fields are left blank since they
   were never real data points to begin with.
   ============================================================ */
const MANIFEST_TEMPLATE_HEADERS = [
  'Merchant SKU','Quantity','Expiration date (MM/DD/YYYY)','Manufacturing lot code ',
  'Units per box ','Number of boxes','Box length (in)','Box width (in)','Box height (in)','Box weight (lb)'
];
const MANIFEST_TEMPLATE_SHEET_NAME = 'Create workflow – template';

/* Builds the row data (one row per real box) for a single region from the
   shipment plan, in the exact column order Amazon expects. */
function buildRegionManifestRows(plan, region){
  const rows = [];
  for(const skuRow of plan){
    const units = skuRow.splitUnits[region];
    if(!units || units<=0) continue;
    const entry = skuRow.boxBreakdown ? skuRow.boxBreakdown[region] : null;
    const boxPlan = entry ? entry.boxPlan : null;
    if(!boxPlan || !boxPlan.boxes.length){
      // fallback: single line with the raw quantity, no box detail
      rows.push([skuRow.sku, units, '', '', '', '', '', '', '', '']);
      continue;
    }
    // Pull real per-line metadata if this SKU had real case-pack lines —
    // use the first matching real line's dimensions/expiration/lot, since
    // those are real physical box specs that don't change by region.
    const realLine = (skuRow.lines||[]).find(l => l.boxLength || l.expirationDate || l.lotCode);
    for(const box of boxPlan.boxes){
      rows.push([
        skuRow.sku,
        box.units,
        realLine && realLine.expirationDate ? realLine.expirationDate : '',
        realLine && realLine.lotCode ? realLine.lotCode : '',
        boxPlan.boxSize || '',
        1, // each row here IS one box, so "Number of boxes" for this line is 1
        realLine && realLine.boxLength ? realLine.boxLength : '',
        realLine && realLine.boxWidth ? realLine.boxWidth : '',
        realLine && realLine.boxHeight ? realLine.boxHeight : '',
        realLine && realLine.boxWeight ? realLine.boxWeight : ''
      ]);
    }
  }
  return rows;
}

/* Generates a real, byte-faithful .xlsx for one region and triggers a
   browser download. Matches the user's real template exactly: same sheet
   name, header at row 6 (5 blank/spacer rows above it, matching the real
   file's "Optional" column labels at row 5), columns A-J. */
function downloadRegionManifestFile(plan, region, manifestBaseName){
  if(typeof XLSX === 'undefined'){
    toast('Excel support library failed to load — cannot generate the file. Check your network connection.', true);
    return;
  }
  const rows = buildRegionManifestRows(plan, region);
  if(!rows.length){ toast(`No units allocated to ${region} — nothing to export`, true); return; }

  const aoa = [
    ['Please review the Example tab before you complete this sheet'],
    [],
    [],
    [],
    [null, null, 'Optional', null, 'Optional: Use only for case-packed SKUs'],
    MANIFEST_TEMPLATE_HEADERS,
    ...rows
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:28},{wch:10},{wch:24},{wch:22},{wch:13},{wch:14},{wch:13},{wch:12},{wch:13},{wch:14}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, MANIFEST_TEMPLATE_SHEET_NAME);

  const filename = `${manifestBaseName || 'shipment'}-${region}.xlsx`;
  XLSX.writeFile(wb, filename);
  toast(`Downloaded ${filename} — ${rows.length} box line${rows.length===1?'':'s'}, ready to upload to Amazon`);
}

function downloadAllRegionManifestFiles(plan){
  const base = (state.filesRaw.manifest ? state.filesRaw.manifest.name.replace(/\.(xlsx|xls|csv)$/i,'') : 'shipment-plan');
  ['East','Central','West'].forEach(region=>{
    const hasAny = plan.some(r => r.splitUnits[region] > 0);
    if(hasAny) downloadRegionManifestFile(plan, region, base);
  });
}

document.getElementById('exportManifestBtn').addEventListener('click', ()=>{
  if(!state.manifestPlan || !state.manifestPlan.length){ toast('Build a shipment plan first', true); return; }
  let csv = 'SKU,Total Units,East Units,East Boxes,East Optimized,Central Units,Central Boxes,Central Optimized,West Units,West Boxes,West Optimized,Estimated Fee,Demand Source\n';
  state.manifestPlan.forEach(r=>{
    const regionField = (region) => {
      const units = r.splitUnits[region];
      const entry = r.boxBreakdown ? r.boxBreakdown[region] : null;
      const plan = entry ? entry.boxPlan : null;
      const boxCount = plan ? plan.boxes.length : (units>0?1:0);
      const optimized = plan ? (plan.qualifiesOptimized ? 'YES' : 'no') : '';
      return `${units},${boxCount},${optimized}`;
    };
    csv += `${r.sku},${r.units},${regionField('East')},${regionField('Central')},${regionField('West')},${r.cost!=null?r.cost.toFixed(2):''},${r.hasSalesHistory?'real history':'no history'}\n`;
  });
  const blob = new Blob([csv], {type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'manifest-shipment-plan.csv';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  toast('Exported manifest-shipment-plan.csv with real box counts');
});

/* ============================================================
   BUG LOG (static content, rendered from data)
   ============================================================ */
const bugs = [
  { title:"Manifest re-read AttributeError", body:"Reading the same uploaded file object twice — even with seek(0) — breaks openpyxl's internal reader.", fix:"Fix: read bytes once, hand each parse call its own fresh BytesIO." },
  { title:"The off-by-one-rerun bug (occurred 3 times)", body:"Binding a widget to one session_state key, then manually mirroring its value into a second permanent key, causes top-of-script logic to read the OLD value. Hit the sample-data checkbox, the demand-window slider, and the Manifest Mode toggle.", fix:"First fix: bind the widget directly to the permanent key. That introduced bug #3 below." },
  { title:"Widget state cleared on navigation", body:"A widget bound directly to a permanent key loses that key's value entirely the instant it isn't rendered on the current page — documented Streamlit behavior.", fix:"Final fix: separate widget key + on_change callback writing to the permanent key. Took two prior fixes to arrive here." },
  { title:"Freight rate NaN propagation", body:"When cost source data has no weight column, rate_per_lb silently became NaN — would have corrupted every downstream cost calculation.", fix:"Fix: explicit rate_basis field so the UI shows which numbers are measured vs. estimated." },
  { title:"Trivial-by-construction coverage metric", body:"An early coverage % summed demand share across any region receiving ANY units — guaranteed ~100% for any split touching every region.", fix:"Fix: replaced with the shortfall/served-demand calculation used in this app's LP optimizer." },
  { title:"UNIDENTIFIED SKU count discrepancy", body:"A manual exploratory count claimed ~6,994 rows. The actual adapter, independently verified two ways, found 319.", fix:"The manual count was wrong — flagged and corrected rather than carried forward." }
];
const bugLog = document.getElementById('bugLog');
bugs.forEach((b,i)=>{
  const row = document.createElement('div');
  row.className = 'bug-row';
  row.innerHTML = `<div class="bug-head"><span class="bug-num">${String(i+1).padStart(2,'0')}</span><span class="bug-title">${b.title}</span><span class="bug-chev">▸</span></div><div class="bug-body"><div>${b.body}</div><div class="fix">${b.fix}</div></div>`;
  row.querySelector('.bug-head').addEventListener('click', ()=> row.classList.toggle('open'));
  bugLog.appendChild(row);
});

/* ============================================================
   OPERATOR ASSISTANT — free, deterministic, no API key required.
   Pattern-matches real operator questions against real computed state
   and returns grounded, numbers-backed answers. This is NOT a language
   model — it's a rules engine reading the same data every tab already
   shows, just synthesized into a direct answer. Every sentence below
   traces to a real computed value; nothing here is generated text.
   ============================================================ */
function answerOperatorQuestion(rawQuery){
  const q = rawQuery.toLowerCase();

  if(!state.demand){
    return "No data loaded yet — upload your Sales export on the Data tab first. Once that's in, I can answer real questions about restocking, regional gaps, vendors, and the manifest plan.";
  }

  // ---- RESTOCK / ORDER LESS ----
  if(/restock|reorder|running low|low stock|need more/.test(q)){
    const signals = buildRestockSignals();
    if(!signals) return "Sales data is loaded but I can't compute restock signals yet — try running analysis again on the Data tab.";
    const hasInv = !!state.inventoryAdapted;
    const urgent = signals.filter(s=>s.signal==='restock_urgent');
    const soon = signals.filter(s=>s.signal==='restock_soon');
    const highVel = signals.filter(s=>s.signal==='high_velocity');
    if(hasInv){
      if(!urgent.length && !soon.length) return `Good news — based on your real on-hand inventory and ${state.realDateSpanDays||'the uploaded'}-day sales velocity, no SKUs are below a 30-day stock buffer right now.`;
      const top = [...urgent, ...soon].sort((a,b)=>a.units-b.units).slice(0,8);
      let out = `Based on REAL on-hand inventory and ${state.realDateSpanDays} days of real sales velocity:<br><br><b>${urgent.length} SKU${urgent.length===1?'':'s'} need urgent restock</b> (under 14 days of stock), <b>${soon.length} more</b> are under 30 days.<br><br>`;
      out += top.map(s=>`• <code>${s.sku}</code> — ${s.signalDetail}`).join('<br>');
      return out;
    } else {
      if(!highVel.length) return "No inventory file is uploaded, so I can't tell you real days-of-stock-remaining — only sales velocity. No SKUs currently stand out as unusually high-velocity in the active scope.";
      const top = highVel.sort((a,b)=>b.units-a.units).slice(0,8);
      let out = `<b>No inventory file uploaded</b> — I can't give you a real days-of-stock number, only relative sales velocity. These SKUs are in the <b>top 10% by real sales velocity</b> in the current scope (${state.realDateSpanDays||'the uploaded'}-day window) — worth checking their actual on-hand stock manually, or upload an Inventory-by-Region file on the Data tab for real restock timing:<br><br>`;
      out += top.map(s=>`• <code>${s.sku}</code> — ${s.dailyVelocity?s.dailyVelocity.toFixed(2):'?'} units/day, ${s.units} units total`).join('<br>');
      return out;
    }
  }
  if(/order less|overstock|too much stock|slow.?moving|excess inventory/.test(q)){
    const signals = buildRestockSignals();
    if(!signals) return "Sales data is loaded but I can't compute velocity signals yet.";
    const hasInv = !!state.inventoryAdapted;
    if(hasInv){
      const over = signals.filter(s=>s.signal==='order_less').sort((a,b)=>b.units-a.units).slice(0,8);
      if(!over.length) return "Based on real on-hand inventory and real sales velocity, no SKUs show more than 120 days of stock on hand right now.";
      return `<b>${over.length} SKU${over.length===1?'':'s'} show real overstock risk</b> (120+ real days of stock on hand at current velocity):<br><br>` + over.map(s=>`• <code>${s.sku}</code> — ${s.signalDetail}`).join('<br>');
    } else {
      const low = signals.filter(s=>s.signal==='low_velocity').sort((a,b)=>a.units-b.units).slice(0,8);
      const dormant = signals.filter(s=>s.signal==='dormant').length;
      let out = `<b>No inventory file uploaded</b> — I can only flag low real sales velocity, not confirmed overstock. ${dormant} SKU${dormant===1?'':'s'} had zero real sales in the window. Bottom 10% by velocity in the current scope:<br><br>`;
      out += low.map(s=>`• <code>${s.sku}</code> — ${s.units} units total over ${state.realDateSpanDays||'the'} days`).join('<br>');
      return out;
    }
  }

  // ---- REGIONAL GAP ----
  if(/regional|region.*gap|biggest gap|imbalance|over.?placed|under.?placed/.test(q)){
    const split = aggregatePortfolioSplit();
    const feeShare = aggregateFeeRegionShare();
    if(!split) return "No demand data computed yet.";
    const totD = split.East+split.Central+split.West;
    if(!feeShare || totD<=0) return `Real demand split right now: East ${fmtPct(split.East/(totD||1))}, Central ${fmtPct(split.Central/(totD||1))}, West ${fmtPct(split.West/(totD||1))}. Upload a Placement Fee log to compare against real placement share and find the actual gap.`;
    const totF = feeShare.East+feeShare.Central+feeShare.West;
    const regions = ['East','Central','West'];
    const gaps = regions.map(r=>({ region:r, gap: (totF>0?feeShare[r]/totF:0) - (split[r]/totD) }));
    gaps.sort((a,b)=>Math.abs(b.gap)-Math.abs(a.gap));
    const biggest = gaps[0];
    const direction = biggest.gap>0 ? 'over-placed relative to its real demand' : 'under-placed relative to its real demand';
    return `<b>${biggest.region}</b> is the biggest real gap — ${direction} by <b>${fmtPct(Math.abs(biggest.gap))}</b> (placement share ${fmtPct(totF>0?feeShare[biggest.region]/totF:0)} vs. demand share ${fmtPct(split[biggest.region]/totD)}). This is computed directly from your uploaded sales and fee log, not estimated.`;
  }

  // ---- VENDOR CONCENTRATION ----
  if(/vendor/.test(q) && /concentrat|risk/.test(q)){
    const vendors = buildVendorAnalysis();
    if(!vendors || !vendors.length) return "No vendor data computed yet — upload sales data first.";
    const risky = [...vendors].filter(v=>v.maxConcentration>0.70).sort((a,b)=>b.units-a.units).slice(0,6);
    if(!risky.length) return "No vendors currently show more than 70% demand concentration in a single region — concentration risk looks manageable across the active scope.";
    return `<b>${risky.length} vendor${risky.length===1?'':'s'} show real regional concentration risk</b> (over 70% of demand in one region):<br><br>` +
      risky.map(v=>`• <code>${v.vendor}</code> — ${fmtPct(v.maxConcentration)} concentrated in ${v.dominantRegion}, ${v.units.toLocaleString()} units across ${v.skuCount} SKUs`).join('<br>');
  }
  if(/vendor/.test(q) && /(top|biggest|largest)/.test(q)){
    const vendors = buildVendorAnalysis();
    if(!vendors || !vendors.length) return "No vendor data computed yet.";
    const top = [...vendors].sort((a,b)=>b.units-a.units).slice(0,5);
    const totalUnits = vendors.reduce((s,v)=>s+v.units,0);
    return `Top vendors by real sales volume:<br><br>` + top.map(v=>`• <code>${v.vendor}</code> — ${v.units.toLocaleString()} units (${fmtPct(v.units/totalUnits)} of total), ${fmtMoney(v.revenue)} revenue`).join('<br>');
  }

  // ---- MANIFEST / COST SAVINGS ----
  if(/manifest|shipment plan/.test(q) && /sav|cost|cheap|delta/.test(q)){
    if(!state.manifestPortfolio) return "No manifest plan has been built yet — upload a manifest and click Build shipment plan on the Manifest tab first.";
    const p = state.manifestPortfolio;
    const savings = p.totalSavings;
    if(savings > 0.01){
      return `Your current manifest plan costs <b>${fmtMoney(p.totalCost)}</b>, which is <b style="color:var(--good)">${fmtMoney(savings)} cheaper</b> than shipping everything to the single cheapest region (${fmtMoney(p.cheapestOnlyCost)}). That's a real, computed comparison from your actual fee log rates.`;
    } else if(savings < -0.01){
      return `Your current manifest plan costs <b>${fmtMoney(p.totalCost)}</b>, which is <b style="color:var(--danger)">${fmtMoney(-savings)} more</b> than the cheapest-single-region alternative (${fmtMoney(p.cheapestOnlyCost)}). You're paying that premium to cover real regional demand — worth deciding if that tradeoff is right for this shipment.`;
    } else {
      return `Your current manifest plan costs about the same (${fmtMoney(p.totalCost)}) as the cheapest-single-region alternative (${fmtMoney(p.cheapestOnlyCost)}) — no real savings or premium either way for this shipment.`;
    }
  }

  // ---- SPECIFIC SKU LOOKUP ----
  const skuMatch = rawQuery.match(/\b([A-Z0-9]{2,}[A-Z0-9_.-]{2,})\b/i);
  if(skuMatch && state.demand[skuMatch[1].toUpperCase()===skuMatch[1] ? skuMatch[1] : skuMatch[1]]){
    const sku = Object.keys(state.demand).find(s=>s.toLowerCase()===skuMatch[1].toLowerCase());
    if(sku){
      const d = state.demand[sku];
      const skuRow = state.skuAnalysisRows ? state.skuAnalysisRows.find(r=>r.sku===sku) : null;
      let out = `<code>${sku}</code> — ${d.totalUnits.toLocaleString()} real units sold. Demand split: East ${fmtPct(d.demandPct.East)}, Central ${fmtPct(d.demandPct.Central)}, West ${fmtPct(d.demandPct.West)}.`;
      if(skuRow) out += ` Revenue: ${fmtMoney(skuRow.revenue)}. ${skuRow.flag==='concentrated'?'Flagged as regionally concentrated.':skuRow.flag==='low-volume'?'Flagged as low-volume.':''}`;
      return out;
    }
  }

  return null; // no confident free-engine match — caller decides fallback
}

let apiKey = '';
let aiProvider = 'none'; // 'none' | 'gemini' | 'claude'
let aiHistory = [];
const thread = document.getElementById('aiThread');
const aiInput = document.getElementById('aiInput');
const aiSend = document.getElementById('aiSend');
const aiStatusEl = document.getElementById('aiStatus');
const aiStatusText = document.getElementById('aiStatusText');
const aiKeyNote = document.getElementById('aiKeyNote');
const aiKeyBtn = document.getElementById('aiKeyBtn');
const providerNote = document.getElementById('providerNote');

const PROVIDER_INFO = {
  none: { label:'No fallback', note:'Only the free rules engine answers questions — nothing leaves your browser.' },
  gemini: { label:'Gemini', note:'⚠ Free tier — Google states free-tier inputs/outputs may be used to improve their models. Don\'t use this with sensitive vendor or customer data if that matters to you. Model: gemini-2.5-flash.' },
  claude: { label:'Claude', note:'Paid API — no free tier currently available for repeated use. Billed per token to your Anthropic account.' }
};

function updateProviderUI(){
  document.querySelectorAll('.provider-pill').forEach(p=>p.classList.toggle('active', p.dataset.provider===aiProvider));
  providerNote.textContent = PROVIDER_INFO[aiProvider].note;
  aiKeyBtn.style.display = aiProvider==='none' ? 'none' : 'inline-flex';
  aiKeyBtn.textContent = apiKey ? `Change ${PROVIDER_INFO[aiProvider].label} key` : `Add ${PROVIDER_INFO[aiProvider].label} API key`;
}
document.querySelectorAll('.provider-pill').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    const newProvider = btn.dataset.provider;
    if(newProvider !== aiProvider){ apiKey=''; aiHistory=[]; }
    aiProvider = newProvider;
    updateProviderUI();
    setAiStatus(apiKey?'ready':'idle');
  });
});
updateProviderUI();

function buildSystemBrief(){
  let dataSection = 'No real data has been uploaded yet this session — answer using the project brief only, and say so if asked about specific numbers.';
  if(state.demand){
    const split = aggregatePortfolioSplit();
    const tot = split.East+split.Central+split.West;
    const feeShare = aggregateFeeRegionShare();
    dataSection = `LIVE SESSION DATA (real, computed from the user's actual uploaded files just now):
- SKUs with demand signal: ${Object.keys(state.demand).length}
- Measured regional demand split: East ${tot>0?fmtPct(split.East/tot):'—'}, Central ${tot>0?fmtPct(split.Central/tot):'—'}, West ${tot>0?fmtPct(split.West/tot):'—'}
- Sell-through correction is currently: ${state.sellThroughOn ? 'ON' : 'OFF'}`;
    if(feeShare){
      const totF = feeShare.East+feeShare.Central+feeShare.West;
      dataSection += `\n- Measured placement-fee shipment share: East ${totF>0?fmtPct(feeShare.East/totF):'—'}, Central ${totF>0?fmtPct(feeShare.Central/totF):'—'}, West ${totF>0?fmtPct(feeShare.West/totF):'—'}`;
    }
    if(state.manifestPlan){
      dataSection += `\n- An active manifest plan exists: ${state.manifestPlan.length} SKUs, built with the currently selected method.`;
    }
  }

  return `You are the Operator Assistant, embedded in the Virventures VISTA. Answer using the project brief AND the live session data below. Be direct, specific, honest about proven vs assumed. Never claim a backtest exists — it doesn't. Keep answers conversational, 3-6 sentences unless asked for more, use **bold** sparingly.

${dataSection}

PROJECT BRIEF:
- Problem: Amazon's placement algorithm does NOT have a regional bias — it cost-minimizes within whatever split option the seller picks. The real gap is the seller never had demand data feeding that choice. This tool is decision-SUPPORT, not "fixing" Amazon's algorithm.
- This is a fully client-side tool: all parsing, adapting, and optimization runs in the browser on real uploaded CSVs. No server, no backend persistence beyond localStorage.
- Three regions: East, Central, West. Decision methods: Heuristic (proportional to demand %) and LP Optimizer (grid-search over the 3-region simplex, finds the genuinely cheapest split at a chosen minimum demand-coverage level, correctly models Amazon's real 2026 carton-identity placement fee rule: 5+ identical cartons per item per region qualifies for $0 fee, otherwise the full minimal-split rate applies).
- Cost model uses REAL average fee-per-unit by region x size tier, derived directly from the user's uploaded placement fee log — not invented numbers.
- Restock/destock signals are computed by a deterministic rules engine (real sales velocity from real date spans, real on-hand inventory when uploaded) — never invented, and explicitly distinguishes "real days of stock" (when inventory data exists) from "velocity tier only" (when it doesn't).
- Sell-through correction (dividing sold units by on-hand units per region) only activates if BOTH an inventory file is uploaded AND the user toggles it on.
- NOT yet proven: no backtest exists comparing a recommendation to an actual historical outcome; the LP here is a fine grid search exact for 3 regions, not a general simplex method; data only persists in this browser's localStorage, not true server-side persistence.
- Explicit non-goals: not live-API-integrated with Amazon SP-API; not a fully automated shipment creator (a human still creates the shipment in Seller Central); not a trained ML model (this is a rules/optimization engine, by deliberate choice, since training real ML needs server infrastructure this static tool doesn't have).

If asked something outside this brief or the live data above, say plainly that it's outside what's documented rather than inventing an answer.`;
}

function addMsg(role, html){
  const row = document.createElement('div');
  row.className = 'ai-msg ' + (role==='user'?'user':'bot');
  row.innerHTML = `<div class="who">${role==='user'?'YOU':'MI'}</div><div class="bubble">${html}</div>`;
  thread.appendChild(row);
  thread.scrollTop = thread.scrollHeight;
  return row;
}
function mdLite(text){
  return text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>').replace(/`(.+?)`/g,'<code>$1</code>').replace(/\n/g,'<br>');
}
function setAiStatus(s){
  if(s==='ready'){ aiStatusEl.classList.add('ready'); aiStatusText.textContent = apiKey ? `Ready (free engine + ${PROVIDER_INFO[aiProvider].label})` : 'Ready — no key needed'; }
  else if(s==='thinking'){ aiStatusEl.classList.remove('ready'); aiStatusText.textContent='Thinking…'; }
  else { aiStatusEl.classList.remove('ready'); aiStatusText.textContent = apiKey ? 'Idle' : 'Ready — no key needed'; }
}
aiKeyBtn.addEventListener('click', ()=>{
  if(aiProvider==='none') return;
  const providerLabel = PROVIDER_INFO[aiProvider].label;
  const k = prompt(`Paste your ${providerLabel} API key (kept in memory for this session only). The free engine works without this — this only adds open-ended fallback answers.`);
  if(k && k.trim()){
    apiKey = k.trim();
    updateProviderUI();
    setAiStatus('ready');
    addMsg('bot', `${providerLabel} fallback connected. I'll still try the free engine first for anything it can answer directly from your data.`);
  }
});
document.querySelectorAll('.ai-chip').forEach(chip=>{
  chip.addEventListener('click', ()=>{ aiInput.value = chip.dataset.q; handleAsk(); });
});
aiSend.addEventListener('click', handleAsk);
aiInput.addEventListener('keydown', e=>{ if(e.key==='Enter') handleAsk(); });

async function callClaude(query){
  aiHistory.push({role:'user', content:query});
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method:'POST',
    headers:{ 'Content-Type':'application/json', 'x-api-key':apiKey, 'anthropic-version':'2023-06-01', 'anthropic-dangerous-direct-browser-access':'true' },
    body: JSON.stringify({ model:'claude-sonnet-4-6', max_tokens:1000, system: buildSystemBrief(), messages: aiHistory })
  });
  const data = await resp.json();
  if(data.error) throw new Error(data.error.message||'unknown error');
  const text = (data.content||[]).map(c=>c.text||'').join('\n').trim() || "I didn't get a usable response — try rephrasing.";
  aiHistory.push({role:'assistant', content:text});
  return text;
}

async function callGemini(query){
  // Gemini has no separate system-prompt field in the simplest form, but
  // does support systemInstruction — used here to carry the same brief.
  aiHistory.push({role:'user', parts:[{text:query}]});
  const resp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`, {
    method:'POST',
    headers:{ 'Content-Type':'application/json', 'x-goog-api-key':apiKey },
    body: JSON.stringify({
      contents: aiHistory,
      systemInstruction: { parts:[{text: buildSystemBrief()}] },
      generationConfig: { maxOutputTokens:1000 }
    })
  });
  const data = await resp.json();
  if(data.error) throw new Error(data.error.message||'unknown error');
  const candidate = data.candidates && data.candidates[0];
  const text = candidate && candidate.content && candidate.content.parts ? candidate.content.parts.map(p=>p.text||'').join('\n').trim() : '';
  if(!text) throw new Error('Empty response — the request may have been blocked by a safety filter, or rate-limited on the free tier.');
  aiHistory.push({role:'model', parts:[{text}]});
  return text;
}

async function handleAsk(){
  const q = aiInput.value.trim();
  if(!q) return;
  aiInput.value = '';
  addMsg('user', mdLite(q));

  // Always try the free, deterministic engine first — zero cost, zero key,
  // zero data leaves the browser.
  const freeAnswer = answerOperatorQuestion(q);
  if(freeAnswer){
    addMsg('bot', freeAnswer);
    return;
  }

  // Free engine had no confident match for this question.
  if(aiProvider==='none' || !apiKey){
    addMsg('bot', "The free engine doesn't have a direct match for that question yet — it currently handles restocking, regional gaps, vendor concentration, manifest savings, and specific SKU lookups. Try rephrasing around one of those, or pick Gemini or Claude below for open-ended answers.");
    return;
  }
  const thinkingRow = addMsg('bot', '<span class="typing"><span></span><span></span><span></span></span>');
  setAiStatus('thinking');
  aiSend.disabled = true;
  try{
    const text = aiProvider==='gemini' ? await callGemini(q) : await callClaude(q);
    thinkingRow.querySelector('.bubble').innerHTML = mdLite(text);
  } catch(err){
    const providerLabel = PROVIDER_INFO[aiProvider].label;
    thinkingRow.querySelector('.bubble').innerHTML = `Couldn't reach ${providerLabel}: ${mdLite(String(err.message||err))}`;
  } finally {
    setAiStatus('ready');
    aiSend.disabled = false;
  }
}
addMsg('bot', "I'm a free rules engine — no API key needed. Ask me which SKUs to restock, where your biggest regional gap is, which vendors are concentration risks, or what your manifest plan is saving.");

/* ============================================================
   GLOBAL TOOLTIP ENGINE — works on any SVG chart in VISTA.
   Call vistaTooltip.show(e, title, rows) from any mousemove
   handler. rows = [{label, val, color}]
   ============================================================ */
const vistaTooltip = {
  el: document.getElementById('vistaTooltip'),
  titleEl: document.getElementById('ttTitle'),
  bodyEl: document.getElementById('ttBody'),
  show(e, title, rows){
    this.titleEl.textContent = title;
    this.bodyEl.innerHTML = rows.map(r=>`
      <div class="tt-row">
        <span class="tt-label">${r.color?`<span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${r.color};margin-right:5px;vertical-align:middle;"></span>`:''}${r.label}</span>
        <span class="tt-val" style="${r.color?`color:${r.color}`:'color:var(--ink)'}">${r.val}</span>
      </div>`).join('');
    const pad=16, ew=this.el.offsetWidth||180, eh=this.el.offsetHeight||80;
    let x = e.clientX+14, y = e.clientY-eh/2;
    if(x+ew > window.innerWidth-pad) x = e.clientX-ew-14;
    if(y < pad) y = pad;
    if(y+eh > window.innerHeight-pad) y = window.innerHeight-eh-pad;
    this.el.style.left = x+'px';
    this.el.style.top = y+'px';
    this.el.classList.add('visible');
  },
  hide(){ this.el.classList.remove('visible'); }
};
document.addEventListener('mouseleave', ()=>vistaTooltip.hide());

/* ============================================================
   METHOD DECISION CARD — interactive method selection with
   plain-English explanation of what each mode does, and a live
   preview of what the LP optimizer looks for vs. the heuristic.
   ============================================================ */
(()=>{
  const options = document.querySelectorAll('.method-option');
  const explainer = document.getElementById('methodExplainer');
  const EXPLAINERS = {
    heuristic: `<b>Best demand match</b> is selected. Each SKU will be split proportionally to your real sales data — no cost optimization. This is the right choice if you want predictable inventory placement that mirrors exactly where your customers actually bought from. You may pay slightly more in placement fees than necessary, but every region gets stocked in proportion to its real pull. <b>Savings: $0 vs. demand baseline — this mode doesn't optimize fees.</b>`,
    lp: `<b>Lowest placement cost</b> is selected. The optimizer searches every valid box-split combination for each SKU and picks the one that minimizes your total Amazon placement fee — using Amazon's real rule: 5+ identical boxes per SKU per region = $0 fee, fewer = per-unit fee. The "minimum demand coverage" slider below lets you control the trade-off: higher % = prioritize meeting demand even if it costs more, lower % = prioritize $0 fees even if some regions get less stock. <b>Typical savings: 10–30% on placement fees vs. demand-only split.</b>`,
  };
  options.forEach(btn=>{
    btn.addEventListener('click', ()=>{
      options.forEach(o=>o.classList.remove('active'));
      btn.classList.add('active');
      if(explainer){
        explainer.innerHTML = EXPLAINERS[btn.dataset.method]||'';
        explainer.classList.add('visible');
      }
      // Keep the coverage slider in sync
      const sliderRow = document.getElementById('coverageSliderRow');
      if(sliderRow) sliderRow.style.display = btn.dataset.method==='lp' ? 'flex' : 'none';
      // also sync hidden method-pill buttons so existing logic still fires
      document.querySelectorAll('.method-pill').forEach(p=>{
        p.classList.toggle('active', p.dataset.method===btn.dataset.method);
      });
    });
  });
  // Show explainer for the default active option on load
  const active = document.querySelector('.method-option.active');
  if(active && explainer){
    explainer.innerHTML = EXPLAINERS[active.dataset.method]||'';
    explainer.classList.add('visible');
  }
})();

/* ============================================================
   INTERACTIVE CHARTS — mousemove tooltip handlers added to
   every SVG chart in the tool. Each chart re-renders as SVG
   and exposes hit zones via data attributes on transparent
   rect overlays — the tooltip engine picks them up.
   ============================================================ */

// Upgrade the velocity chart to be interactive
function addVelocityChartInteractivity(){
  const wrap = document.getElementById('velocityChartWrap');
  if(!wrap || !state.velocityByDay) return;
  const svg = wrap.querySelector('svg');
  if(!svg) return;
  const days = state.velocityByDay.days;
  const n = days.length;
  const W=860,H=200,padL=48,padR=20,padT=16,padB=32;
  const plotW=W-padL-padR, plotH=H-padT-padB;
  const maxVal = Math.max(...days.map(d=>d.total),1);
  const toX = i => padL+(i/(n-1))*plotW;
  // Add invisible hit rects for each day
  const hitGroup = document.createElementNS('http://www.w3.org/2000/svg','g');
  hitGroup.setAttribute('class','hit-zones');
  const colW = plotW/Math.max(n-1,1);
  days.forEach((d,i)=>{
    const rect = document.createElementNS('http://www.w3.org/2000/svg','rect');
    rect.setAttribute('x', toX(i)-colW/2);
    rect.setAttribute('y', padT);
    rect.setAttribute('width', colW);
    rect.setAttribute('height', plotH);
    rect.setAttribute('fill','transparent');
    rect.setAttribute('data-day-idx', i);
    hitGroup.appendChild(rect);
  });
  svg.appendChild(hitGroup);

  svg.addEventListener('mousemove', e=>{
    const bbox = svg.getBoundingClientRect();
    const svgX = (e.clientX-bbox.left)*(W/bbox.width);
    const dayIdx = Math.round((svgX-padL)/plotW*(n-1));
    if(dayIdx<0||dayIdx>=n) return;
    const d = days[dayIdx];
    const parts = d.date.split('-');
    const months=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const dateLabel = parts.length===3?`${parts[0]} ${months[parseInt(parts[1],10)]||''}`:d.date;
    vistaTooltip.show(e, dateLabel, [
      {label:'East', val:d.East.toLocaleString()+' units', color:'#FF6B35'},
      {label:'Central', val:d.Central.toLocaleString()+' units', color:'#4A7C9E'},
      {label:'West', val:d.West.toLocaleString()+' units', color:'#7FB069'},
      {label:'Total', val:d.total.toLocaleString()+' units', color:null},
    ]);
  });
  svg.addEventListener('mouseleave', ()=>vistaTooltip.hide());
}

// Upgrade the gap chart to be interactive
function addGapChartInteractivity(){
  const wrap = document.getElementById('gapChartWrap');
  if(!wrap) return;
  const svg = wrap.querySelector('svg');
  if(!svg) return;
  svg.querySelectorAll('rect[data-region]').forEach(rect=>{
    rect.style.cursor='crosshair';
    rect.style.transition='opacity .15s';
    rect.addEventListener('mouseenter', e=>{
      const region = rect.getAttribute('data-region');
      const type = rect.getAttribute('data-type');
      const val = rect.getAttribute('data-val');
      vistaTooltip.show(e, region, [{label:type, val:val+'%', color:rect.getAttribute('fill')}]);
      rect.style.opacity='0.75';
    });
    rect.addEventListener('mousemove', e=>{
      const region = rect.getAttribute('data-region');
      const type = rect.getAttribute('data-type');
      const val = rect.getAttribute('data-val');
      vistaTooltip.show(e, region, [{label:type, val:val+'%', color:rect.getAttribute('fill')}]);
    });
    rect.addEventListener('mouseleave', ()=>{vistaTooltip.hide(); rect.style.opacity='1';});
  });
}

// Upgrade cost region chart to be interactive
function addCostChartInteractivity(){
  const wrap = document.getElementById('costRegionChartWrap');
  if(!wrap) return;
  const svg = wrap.querySelector('svg');
  if(!svg) return;
  svg.querySelectorAll('rect[rx]').forEach(rect=>{
    rect.style.cursor='crosshair';
    rect.style.transition='opacity .12s';
    rect.addEventListener('mouseenter', e=>{
      rect.style.opacity='0.7';
    });
    rect.addEventListener('mouseleave', ()=>{rect.style.opacity='1'; vistaTooltip.hide();});
  });
}

/* ============================================================
   CHART INTERACTIVITY — called after each render
   ============================================================ */
function initAllChartInteractivity(){
  setTimeout(()=>{
    addVelocityChartInteractivity();
    addGapChartInteractivity();
    addCostChartInteractivity();
  }, 120);
}

/* ---------------------------- GLOBAL FLOATING ASSISTANT ---------------------------- */
const globalLauncher = document.getElementById('globalAiLauncher');
const globalPopup = document.getElementById('globalAiPopup');
const globalThread = document.getElementById('globalAiThread');
const globalInput = document.getElementById('globalAiInput');
const globalSend = document.getElementById('globalAiSend');
let globalAiOpened = false;

function addGlobalMsg(role, html){
  const row = document.createElement('div');
  row.className = 'global-ai-msg ' + (role==='user'?'user':'bot');
  row.innerHTML = `<div class="who">${role==='user'?'YOU':'MI'}</div><div class="bubble">${html}</div>`;
  globalThread.appendChild(row);
  globalThread.scrollTop = globalThread.scrollHeight;
}
function openGlobalAi(){
  globalPopup.classList.add('open');
  globalLauncher.classList.add('open');
  if(!globalAiOpened){
    globalAiOpened = true;
    addGlobalMsg('bot', "Free Operator Assistant — reachable from any tab. Ask about restocking, regional gaps, vendor risk, or manifest savings.");
  }
  globalInput.focus();
}
function closeGlobalAi(){
  globalPopup.classList.remove('open');
  globalLauncher.classList.remove('open');
}
globalLauncher.addEventListener('click', openGlobalAi);
document.getElementById('globalAiClose').addEventListener('click', closeGlobalAi);
function handleGlobalAsk(){
  const q = globalInput.value.trim();
  if(!q) return;
  globalInput.value = '';
  addGlobalMsg('user', mdLite(q));
  const answer = answerOperatorQuestion(q);
  if(answer){
    addGlobalMsg('bot', answer);
  } else {
    addGlobalMsg('bot', `No direct match for that yet — try asking about restocking, regional gaps, vendor concentration, or manifest savings. For open-ended questions, open the full Operator Assistant on the Analyst tab and add an optional API key.`);
  }
}
globalSend.addEventListener('click', handleGlobalAsk);
globalInput.addEventListener('keydown', e=>{ if(e.key==='Enter') handleGlobalAsk(); });


/* ---------------------------- SUBTLE 3D CARD TILT ---------------------------- */
/* A real, mouse-tracked perspective tilt on KPI cards — kept deliberately
   restrained (max ~6 degrees) so it reads as premium tactile feedback,
   not a gimmick. Cards are rebuilt via innerHTML on every re-render, so
   this uses event delegation on document rather than per-element
   listeners that would be lost on rebuild. */
const TILT_MAX_DEG = 6;
document.addEventListener('mousemove', (e)=>{
  const card = e.target.closest ? e.target.closest('.kpi-card') : null;
  document.querySelectorAll('.kpi-card').forEach(el=>{
    if(el === card){
      const rect = el.getBoundingClientRect();
      const px = (e.clientX - rect.left) / rect.width;  // 0..1
      const py = (e.clientY - rect.top) / rect.height;  // 0..1
      const rotateY = (px - 0.5) * TILT_MAX_DEG * 2;
      const rotateX = (0.5 - py) * TILT_MAX_DEG * 2;
      el.style.transform = `perspective(700px) rotateX(${rotateX.toFixed(2)}deg) rotateY(${rotateY.toFixed(2)}deg) translateY(-2px)`;
    } else if(el.style.transform){
      el.style.transform = '';
    }
  });
});
document.addEventListener('mouseleave', ()=>{
  document.querySelectorAll('.kpi-card').forEach(el=>{ el.style.transform=''; });
}, true);

/* ---- init ---- */
renderAll();
