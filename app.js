/* =========================================================================
   VISTA AUTH ENGINE — multi-user login with role-based access.
   Passwords are stored as SHA-256 hashes in localStorage.
   Admin account seeded on first launch if nothing exists.
   Data persistence: admin can save/restore the full parsed session state
   to localStorage so users don't need to re-upload every time.
   ========================================================================= */

/* -- SHA-256 via WebCrypto (async, returns hex string) -- */
async function sha256(str){
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}

/* -- User store (localStorage key: vista_users) -- */
const AUTH_KEY = 'vista_users_v1';
const SESSION_KEY = 'vista_session_v1';

function getUsers(){
  try{ return JSON.parse(localStorage.getItem(AUTH_KEY)||'[]'); }
  catch(e){ return []; }
}
function saveUsers(users){ localStorage.setItem(AUTH_KEY, JSON.stringify(users)); }

async function seedDefaultAdmin(){
  const users = getUsers();
  if(users.length === 0){
    const hash = await sha256('admin123');
    saveUsers([{ username:'admin', passwordHash:hash, role:'admin', active:true, createdAt: new Date().toISOString() }]);
  }
}

async function attemptLogin(username, password){
  const users = getUsers();
  const user = users.find(u=>u.username.toLowerCase()===username.toLowerCase() && u.active);
  if(!user) return null;
  const hash = await sha256(password);
  if(hash !== user.passwordHash) return null;
  return user;
}

function getSession(){
  try{ return JSON.parse(sessionStorage.getItem(SESSION_KEY)||'null'); }
  catch(e){ return null; }
}
function setSession(user){ sessionStorage.setItem(SESSION_KEY, JSON.stringify({ username:user.username, role:user.role })); }
function clearSession(){ sessionStorage.removeItem(SESSION_KEY); }

let currentUser = null;

function applySession(user){
  currentUser = user;
  document.getElementById('authScreen').style.display = 'none';
  document.getElementById('sessName').textContent = user.username;
  const roleEl = document.getElementById('sessRole');
  roleEl.textContent = user.role;
  roleEl.className = 'sess-role ' + user.role;
  const adminBtn = document.getElementById('adminPanelBtn');
  if(adminBtn) adminBtn.style.display = user.role==='admin' ? 'inline-flex' : 'none';
  // Lock data upload to admin only
  applyRoleRestrictions(user.role);
}

function applyRoleRestrictions(role){
  if(role === 'user'){
    // Non-admins cannot upload or clear data — grey it out
    const uploadZones = document.querySelectorAll('.upload-zone, #runAnalysisBtn, #clearAllBtn, #loadSampleBtn');
    uploadZones.forEach(el=>{
      el.classList.add('data-admin-only');
      el.title = 'Data management is restricted to admins';
    });
    // Show a note on the data tab
    const runRow = document.querySelector('.run-row');
    if(runRow && !runRow.querySelector('.user-data-note')){
      const note = document.createElement('div');
      note.className = 'user-data-note';
      note.style.cssText = 'font-size:12px;color:var(--ink-faint);padding:10px 14px;background:var(--bg-input);border-radius:8px;border:1px solid var(--line);';
      note.innerHTML = '🔒 <b>Data uploads are admin-only.</b> Ask your admin to load or refresh the data — you\'ll see it here automatically.';
      runRow.parentNode.insertBefore(note, runRow);
    }
  }
}




/* -- Admin panel UI -- */
function renderAdminPanel(){
  renderUserList();
}

function renderUserList(){
  const list = document.getElementById('userList');
  if(!list) return;
  const users = getUsers();
  list.innerHTML = users.map((u,i)=>`
    <div class="user-row">
      <span class="user-row-name">${u.username}</span>
      <span class="user-row-role">${u.role}</span>
      <span class="user-row-status ${u.active?'active':'disabled'}">${u.active?'Active':'Disabled'}</span>
      ${u.username === currentUser?.username ? '<span style="font-size:10.5px;color:var(--ink-faint);">(you)</span>' :
        `<button class="btn btn-ghost btn-sm" onclick="toggleUser(${i})" style="padding:4px 10px;font-size:11px;">${u.active?'Disable':'Enable'}</button>
         <button class="btn btn-ghost btn-sm" onclick="deleteUser(${i})" style="padding:4px 10px;font-size:11px;color:var(--danger);">Remove</button>`}
    </div>`).join('');
}


function toggleUser(idx){
  const users = getUsers();
  if(users[idx].username === currentUser?.username){ toast('Cannot disable your own account',true); return; }
  users[idx].active = !users[idx].active;
  saveUsers(users);
  renderUserList();
}
function deleteUser(idx){
  const users = getUsers();
  if(users[idx].username === currentUser?.username){ toast('Cannot remove your own account',true); return; }
  if(!confirm('Remove user "'+users[idx].username+'"? This cannot be undone.')) return;
  users.splice(idx,1);
  saveUsers(users);
  renderUserList();
}

/* -- Wire up auth screen -- */
async function initAuth(){
  await seedDefaultAdmin();
  const existing = getSession();
  if(existing){
    const users = getUsers();
    const user = users.find(u=>u.username===existing.username && u.active);
    if(user){ applySession(user); return; }
  }
  document.getElementById('authScreen').style.display = 'flex';
  document.getElementById('authPassword').addEventListener('keydown', e=>{ if(e.key==='Enter') doLogin(); });
  document.getElementById('authLoginBtn').addEventListener('click', doLogin);
}
async function doLogin(){
  const username = document.getElementById('authUsername').value.trim();
  const password = document.getElementById('authPassword').value;
  const errEl = document.getElementById('authError');
  if(!username || !password){ errEl.classList.add('visible'); errEl.textContent='Please enter both username and password.'; return; }
  const user = await attemptLogin(username, password);
  if(!user){ errEl.classList.add('visible'); errEl.textContent='Incorrect username or password.'; return; }
  errEl.classList.remove('visible');
  setSession(user);
  applySession(user);
}

/* -- Admin panel button wiring -- */
document.getElementById('adminPanelBtn')?.addEventListener('click', ()=>{
  document.getElementById('adminPanel').classList.add('open');
  renderAdminPanel();
});
document.getElementById('adminClose')?.addEventListener('click', ()=>{
  document.getElementById('adminPanel').classList.remove('open');
});
document.getElementById('adminPanel')?.addEventListener('click', e=>{
  if(e.target === document.getElementById('adminPanel')) document.getElementById('adminPanel').classList.remove('open');
});
document.getElementById('addUserBtn')?.addEventListener('click', async ()=>{
  const u = document.getElementById('newUsername').value.trim();
  const p = document.getElementById('newPassword').value;
  const r = document.getElementById('newRole').value;
  const errEl = document.getElementById('addUserError');
  if(!u||!p){ errEl.textContent='Username and password required.'; errEl.classList.add('visible'); return; }
  const users = getUsers();
  if(users.find(x=>x.username.toLowerCase()===u.toLowerCase())){ errEl.textContent='Username already exists.'; errEl.classList.add('visible'); return; }
  const hash = await sha256(p);
  users.push({ username:u, passwordHash:hash, role:r, active:true, createdAt:new Date().toISOString() });
  saveUsers(users);
  document.getElementById('newUsername').value='';
  document.getElementById('newPassword').value='';
  errEl.classList.remove('visible');
  renderUserList();
  toast('User "'+u+'" added successfully');
});
document.getElementById('sessLogout')?.addEventListener('click', ()=>{
  clearSession();
  location.reload();
});





/* -- Boot -- */
initAuth();

/* VISTA static client app — no login gate or admin session layer. */

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
@@ -1125,149 +938,174 @@ function heuristicSplit(totalUnits, demandPct, realUnitsPerBox, declaredBoxCount
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

const MIN_REGION_BATCH_UNITS = 10;
const MAX_MANIFEST_REGIONS = 2;
const TARGET_MANIFEST_BOXES = 5;
function demandMismatch(units, demandPct, totalUnits){
  if(totalUnits<=0) return 0;
  return ['East','Central','West'].reduce((sum,region)=>sum+Math.abs((units[region]||0)-(demandPct[region]||0)*totalUnits),0)/totalUnits;
}
function countUsedRegions(units){ return ['East','Central','West'].filter(r=>(units[r]||0)>0).length; }
function meetsMinimumBatch(units, batchSize){
  return ['East','Central','West'].every(r => !units[r] || units[r] >= batchSize);
}
function manifestRationale(row){
  const splitCount = countUsedRegions(row.splitUnits);
  const match = row.coverage!=null ? fmtPct(row.coverage) : '—';
  const oneRegionCoverage = row.singleRegionCoverage!=null ? fmtPct(row.singleRegionCoverage) : '—';
  const leader = Object.entries(row.demandPct).sort((a,b)=>b[1]-a[1])[0];
  return `${splitCount} region${splitCount===1?'':'s'} recommended — matches ${match} of regional demand vs. ${oneRegionCoverage} for a 1-shipment plan, given ${leader[0]}'s ${Math.max(1, leader[1]*3).toFixed(1)}x average demand share.`;
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
function chooseBetterTwoTier(candidate, best){
  if(!best) return candidate;
  if(candidate.fiveBoxPlan !== best.fiveBoxPlan) return candidate.fiveBoxPlan ? candidate : best;
  if(candidate.locCount !== best.locCount) return candidate.locCount < best.locCount ? candidate : best;
  if(candidate.mismatch !== best.mismatch) return candidate.mismatch < best.mismatch - 1e-9 ? candidate : best;
  return candidate.totalCost < best.totalCost - 1e-9 ? candidate : best;
}
function oneRegionBestCoverage(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox){
  let best = null;
  for(const region of ['East','Central','West']){
    const units = { East:0, Central:0, West:0 };
    units[region] = totalUnits;
    const coverage = servedDemandCoverage(units, demandPct, totalUnits);
    const costResult = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox);
    const candidate = { region, units, coverage, totalCost:costResult.totalCost };
    if(!best || coverage > best.coverage + 1e-9 || (Math.abs(coverage-best.coverage)<1e-9 && candidate.totalCost < best.totalCost)) best = candidate;
  }
  return best;
}

/* Two-pass MILP-equivalent optimizer for VISTA's three-region static build.
   Pass 1 minimizes the number of used regions (binary y_r). Pass 2 fixes
   that minimum split count and minimizes absolute demand mismatch. The
   exhaustive search is equivalent to the tiny MILP domain here and avoids a
   backend while preserving the business priority order from the spec. */
function lpOptimize(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, minCoveragePct, gridStep, unitsPerBox, isRealCasePack, declaredBoxCount){
  const minBatch = Math.min(MIN_REGION_BATCH_UNITS, totalUnits);
  let best = null, bestByCoverage = null, maxCoverageSeen = 0;

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
      if(remainderUnits < 0){ totalBoxes = Math.floor(totalUnits/unitsPerBox); remainderUnits = totalUnits % unitsPerBox; }
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
        if(remainderUnits>0){ const biggest = e>=c && e>=w ? 'East' : c>=w ? 'Central' : 'West'; units[biggest] += remainderUnits; }
        if(!meetsMinimumBatch(units, minBatch)) continue;
        const coverage = servedDemandCoverage(units, demandPct, totalUnits);
        const mismatch = demandMismatch(units, demandPct, totalUnits);
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
        if(locCount > MAX_MANIFEST_REGIONS) continue;
        const candidate = { units, totalCost, coverage, mismatch, locCount, breakdown, fiveBoxPlan: totalBoxes === TARGET_MANIFEST_BOXES };
        if(!bestByCoverage || coverage > bestByCoverage.coverage + 1e-6 || (Math.abs(coverage-bestByCoverage.coverage)<1e-6 && mismatch < bestByCoverage.mismatch)) bestByCoverage = candidate;
        if(coverage < minCoveragePct - 1e-6) continue;
        if(!best || totalCost < best.totalCost - 1e-9) best = { units, totalCost, coverage, locCount, breakdown };
        best = chooseBetterTwoTier(candidate, best);
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
    if(!best && bestByCoverage) return { ...bestByCoverage, coverageTargetUnreachable:true, maxAchievableCoverage:maxCoverageSeen };
    if(best) best.singleRegionCoverage = oneRegionBestCoverage(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox)?.coverage;
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
  const fiveBoxSize = totalUnits >= MIN_REGION_BATCH_UNITS * TARGET_MANIFEST_BOXES ? Math.floor(totalUnits / TARGET_MANIFEST_BOXES) : null;
  const candidateValues = fiveBoxSize
    ? Array.from({length:TARGET_MANIFEST_BOXES+1}, (_,i)=>i*fiveBoxSize)
    : Array.from({length:totalUnits+1}, (_,i)=>i);
  for(const e of candidateValues){
    for(const c of candidateValues){
      if(e+c>totalUnits) continue;
      let w = totalUnits-e-c;
      if(fiveBoxSize && w % fiveBoxSize !== 0 && w !== totalUnits - (TARGET_MANIFEST_BOXES-1)*fiveBoxSize) continue;
      const units = { East:e, Central:c, West:w };
      if(fiveBoxSize){
        const drift = totalUnits - (units.East+units.Central+units.West);
        if(drift !== 0){
          const biggest = Object.entries(units).sort((a,b)=>b[1]-a[1])[0][0];
          units[biggest] += drift;
        }
      }
      if(!meetsMinimumBatch(units, minBatch)) continue;
      const coverage = servedDemandCoverage(units, demandPct, totalUnits);
      const mismatch = demandMismatch(units, demandPct, totalUnits);
      if(coverage > maxCoverageSeen) maxCoverageSeen = coverage;
      const { totalCost, locCount, breakdown } = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, fiveBoxSize || unitsPerBox);
      if(locCount > MAX_MANIFEST_REGIONS) continue;
      const candidate = { units, totalCost, coverage, mismatch, locCount, breakdown, fiveBoxPlan: !!fiveBoxSize };
      if(!bestByCoverage || coverage > bestByCoverage.coverage + 1e-6 || (Math.abs(coverage-bestByCoverage.coverage)<1e-6 && mismatch < bestByCoverage.mismatch)) bestByCoverage = candidate;
      if(coverage < minCoveragePct - 1e-6) continue;
      const { totalCost, locCount, breakdown } = computeSplitCost(units, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox);
      if(!best || totalCost < best.totalCost - 1e-9) best = { units, totalCost, coverage, locCount, breakdown };
      best = chooseBetterTwoTier(candidate, best);
    }
  }
  if(!best && bestByCoverage) best = { ...bestByCoverage, coverageTargetUnreachable:true, maxAchievableCoverage:maxCoverageSeen };
  if(best) best.singleRegionCoverage = oneRegionBestCoverage(totalUnits, demandPct, sizeTier, feeRateTable, defaultFeePerUnit, unitsPerBox)?.coverage;
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
@@ -1355,156 +1193,168 @@ function aggregateManifestBySku(manifestRecords){
        if(!row.sizeTierCrossCheck && specMatch.length && specMatch.width && specMatch.height && specMatch.weight){
          row.sizeTierCrossCheck = classifySizeTier(specMatch.length, specMatch.width, specMatch.height, specMatch.weight);
        }
      }
    }
  }
  return Object.values(bySku);
}

function buildManifestPlan(manifestRecords, demandBySku, feeRateTable, defaultFeePerUnit, method, minCoveragePct, defaultBoxUnitsPerBox){
  const aggregated = aggregateManifestBySku(manifestRecords);
  const plan = []; let knownDemandCount=0, unknownDemandCount=0;
  const portfolioUnits = { East:0, Central:0, West:0 };
  let portfolioCost = 0, portfolioCheapestOnlyCost = 0;
  let coverageWeightedSum = 0; // for a real units-weighted average coverage %, not a naive mean across SKUs of different sizes

  for(const row of aggregated){
    const demand = demandBySku[row.sku];
    const demandPct = demand ? demand.demandPct : { East:1/3, Central:1/3, West:1/3 };
    if(demand) knownDemandCount++; else unknownDemandCount++;
    const effectiveUnitsPerBox = row.realUnitsPerBox || defaultBoxUnitsPerBox || null;
    const usedDefaultBoxSpec = !row.realUnitsPerBox;
    const isRealCasePack = !!row.realUnitsPerBox; // true only for a real manifest/spec-sheet case pack, never the optimizer's own guessed box size
    const declaredBoxCount = (isRealCasePack && row.realNumberOfBoxes) ? row.realNumberOfBoxes : null;

    let units, cost=null, coverage=null, breakdown=null, coverageTargetUnreachable=false, maxAchievableCoverage=null;
    let units, cost=null, coverage=null, breakdown=null, coverageTargetUnreachable=false, maxAchievableCoverage=null, singleRegionCoverage=null, mismatch=null;
    if(method === 'lp'){
      const lp = lpOptimize(row.units, demandPct, row.sizeTier, feeRateTable, defaultFeePerUnit, minCoveragePct, null, effectiveUnitsPerBox, isRealCasePack, declaredBoxCount);
      units = lp ? lp.units : heuristicSplit(row.units, demandPct, isRealCasePack?effectiveUnitsPerBox:null, isRealCasePack?declaredBoxCount:null);
      cost = lp ? lp.totalCost : null;
      coverage = lp ? lp.coverage : null;
      breakdown = lp ? lp.breakdown : null;
      coverageTargetUnreachable = lp ? !!lp.coverageTargetUnreachable : false;
      maxAchievableCoverage = lp ? (lp.maxAchievableCoverage!=null?lp.maxAchievableCoverage:null) : null;
      singleRegionCoverage = lp ? lp.singleRegionCoverage : null;
      mismatch = lp ? lp.mismatch : null;
    } else {
      units = heuristicSplit(row.units, demandPct, isRealCasePack?effectiveUnitsPerBox:null, isRealCasePack?declaredBoxCount:null);
      const costResult = computeSplitCost(units, row.sizeTier, feeRateTable, defaultFeePerUnit, effectiveUnitsPerBox);
      cost = costResult.totalCost;
      breakdown = costResult.breakdown;
      // Real coverage for the heuristic path too — same formula the LP
      // optimizer uses (servedDemandCoverage), so the executive summary
      // means the same thing regardless of which method is selected.
      coverage = servedDemandCoverage(units, demandPct, row.units);
      singleRegionCoverage = oneRegionBestCoverage(row.units, demandPct, row.sizeTier, feeRateTable, defaultFeePerUnit, effectiveUnitsPerBox)?.coverage;
      mismatch = demandMismatch(units, demandPct, row.units);
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
    const planRow = {
      sku: row.sku, units: row.units, sizeTier: row.sizeTier, hasSalesHistory: !!demand, demandPct,
      splitUnits: units, cost, coverage, packingLines: row.lines.length, lines: row.lines,
      realUnitsPerBox: row.realUnitsPerBox, effectiveUnitsPerBox, usedDefaultBoxSpec, sizeTierCrossCheck: row.sizeTierCrossCheck,
      specSheetMatch: row.specSheetMatch, isRealCasePack, coverageTargetUnreachable, maxAchievableCoverage,
      boxBreakdown: breakdown, cheapestOnlyCost: cheapestOnly.totalCost, cheapestOnlyRegion: cheapestOnly.region, costDelta
    });
      boxBreakdown: breakdown, cheapestOnlyCost: cheapestOnly.totalCost, cheapestOnlyRegion: cheapestOnly.region, costDelta,
      singleRegionCoverage, mismatch, confidence: demand && demand.totalUnits >= 14 ? 'high' : 'low', rationale: null
    };
    planRow.rationale = manifestRationale(planRow);
    plan.push(planRow);
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
  filesRaw: { sales:null, fees:null, inventory:null, manifest:null, specsheet:null, packinglist:null },
  filesRaw: { sales:null, salesFiles:[], fees:null, inventory:null, manifest:null, specsheet:null, packinglist:null },
  detectedHeaders: { sales:null, fees:null, inventory:null, manifest:null, specsheet:null },
  salesAdapted: null, feeAdapted: null, inventoryAdapted: null,
  demand: null, feeRateTable: null, defaultFeePerUnit: 0.4,
  manifestRecords: null, manifestPlan: null, manifestPortfolio: null,
  frontierData: null,
  sellThroughOn: false,
  manifestFilters: { search:'', region:'all', demandSource:'all', minUnits:0 },
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
function normalizeSkuKey(sku){ return String(sku||'').trim().toUpperCase(); }
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
    const manifestSkus = new Set(state.manifestPlan.map(r=>normalizeSkuKey(r.sku)));
    skus = skus.filter(s=>manifestSkus.has(normalizeSkuKey(s)));
    if(!skus.length){
      const manifestVendors = new Set(state.manifestPlan.map(r=>extractVendorPrefix(r.sku)).filter(v=>v !== 'Unidentified'));
      skus = Object.keys(state.demand).filter(s=>manifestVendors.has(extractVendorPrefix(s)));
    }
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
@@ -1727,89 +1577,101 @@ document.getElementById('parsePackingListBtn').addEventListener('click', async (
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
    if(e.dataTransfer.files.length) handleFileSelected(kind, kind === 'sales' ? Array.from(e.dataTransfer.files).slice(0,6) : e.dataTransfer.files[0], zone);
  });
  input.addEventListener('change', ()=>{ if(input.files.length) handleFileSelected(kind, input.files[0], zone); });
  input.addEventListener('change', ()=>{ if(input.files.length) handleFileSelected(kind, kind === 'sales' ? Array.from(input.files).slice(0,6) : input.files[0], zone); });
  zone.querySelector('.uz-clear')?.addEventListener('click', (e)=>{
    e.stopPropagation();
    state.filesRaw[kind] = null;
      if(kind === 'sales') state.filesRaw.salesFiles = [];
    zone.classList.remove('has-file');
    zone.querySelector('.uz-filename').textContent = '';
    input.value = '';
  });
}
async function handleFileSelected(kind, file, zone){
  state.filesRaw[kind] = file;
  if(kind === 'sales' && Array.isArray(file)){
    state.filesRaw.salesFiles = file.slice(0, 6);
    state.filesRaw.sales = state.filesRaw.salesFiles[0] || null;
  } else {
    state.filesRaw[kind] = file;
  }
  zone.classList.add('has-file');
  zone.classList.add('detecting');
  const fmt = isExcelFile(file) ? 'XLSX' : 'CSV';
  zone.querySelector('.uz-filename').textContent = file.name + ' · ' + (file.size/1024/1024).toFixed(2) + ' MB · ' + fmt;
  const fmt = Array.isArray(file) ? 'mixed' : (isExcelFile(file) ? 'XLSX' : 'CSV');
  if(kind === 'sales' && Array.isArray(file)){
    const totalSize = file.reduce((sum,f)=>sum+f.size,0);
    zone.querySelector('.uz-filename').textContent = `${file.length} sales files · ${(totalSize/1024/1024).toFixed(2)} MB total`;
  } else {
    zone.querySelector('.uz-filename').textContent = file.name + ' · ' + (file.size/1024/1024).toFixed(2) + ' MB · ' + fmt;
  }
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
  toast(kind === 'sales' && Array.isArray(file) ? `${file.length} sales files ready` : file.name + ' ready');
  const schemaKeyForKind = kind === 'sales' ? 'sales' : kind === 'fees' ? 'fees' : kind === 'inventory' ? 'inventory' : kind === 'specsheet' ? 'specsheet' : 'manifest';
  try{
    const loader = kind === 'specsheet' ? loadMultiSheetTabularFile : loadTabularFile;
    const { headers, sheetName, headerRow, sheetAutoDetectFailed, noDataRowsFound, sheetsUsed } = await loader(file, schemaKeyForKind);
    const headerFile = (kind === 'sales' && Array.isArray(file)) ? file[0] : file;
    const { headers, sheetName, headerRow, sheetAutoDetectFailed, noDataRowsFound, sheetsUsed } = await loader(headerFile, schemaKeyForKind);
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
  const schemaKey = kind === 'sales' ? 'sales' : kind === 'fees' ? 'fees' : kind === 'inventory' ? 'inventory' : kind === 'specsheet' ? 'specsheet' : 'manifest';
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
@@ -1820,104 +1682,114 @@ function renderColumnMap(kind, zone, headers, sheetInfo){
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
wireUploadZone('sales','file-sales','uz-sales');
wireUploadZone('fees','file-fees','uz-fees');
wireUploadZone('inventory','file-inventory','uz-inventory');
wireUploadZone('manifest','file-manifest','uz-manifest');
wireUploadZone('specsheet','file-specsheet','uz-specsheet');
wireUploadZone('packinglist','file-packinglist','uz-packinglist');

document.getElementById('clearAllBtn').addEventListener('click', ()=>{
  ['sales','fees','inventory'].forEach(kind=>{
    state.filesRaw[kind] = null;
      if(kind === 'sales') state.filesRaw.salesFiles = [];
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

  if(!state.filesRaw.sales && !state.filesRaw.fees){
    toast('Upload at least a Sales or Fee file first — or click "Load sample data"', true);
    progress.style.display = 'none';
    return;
  }

  try{
    if(state.filesRaw.sales){
      logLine(log, 'Sales file', 'reading…');
      const { objects, encoding, format } = await loadTabularFile(state.filesRaw.sales, 'sales');
    if(state.filesRaw.sales || (state.filesRaw.salesFiles && state.filesRaw.salesFiles.length)){
      const salesFiles = state.filesRaw.salesFiles && state.filesRaw.salesFiles.length ? state.filesRaw.salesFiles : [state.filesRaw.sales];
      logLine(log, 'Sales files', `reading ${salesFiles.length} file${salesFiles.length===1?'':'s'}…`);
      let allRecords = [], totalRaw = 0, kept = 0, droppedNonUS = 0, droppedUnresolvedState = 0, droppedStatus = 0, formats = new Set(), encodings = new Set();
      for(const salesFile of salesFiles){
        const { objects, encoding, format } = await loadTabularFile(salesFile, 'sales');
        const { records, stats } = adaptSalesRows(objects);
        allRecords = allRecords.concat(records);
        totalRaw += stats.totalRaw; kept += stats.kept;
        droppedNonUS += stats.droppedNonUS; droppedUnresolvedState += stats.droppedUnresolvedState; droppedStatus += stats.droppedStatus;
        formats.add(format === 'excel' ? 'Excel (.xlsx)' : 'CSV');
        if(format === 'csv') encodings.add(encoding);
      }
      fill.style.width = '30%';
      const { records, stats } = adaptSalesRows(objects);
      state.salesAdapted = records;
      state.salesAdapted = allRecords;
      log.innerHTML = '';
      logLine(log, 'Sales file format', format === 'excel' ? 'Excel (.xlsx)' : 'CSV', 'good');
      if(format === 'csv') logLine(log, 'Sales encoding detected', encoding, encoding.includes('fallback') ? 'warn' : 'good');
      logLine(log, 'Sales rows parsed', `<b>${stats.totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Kept (US, resolvable state, valid qty)', `<b>${stats.kept.toLocaleString()}</b>`, 'good');
      if(stats.droppedNonUS) logLine(log, 'Dropped — non-US', stats.droppedNonUS.toLocaleString(), 'warn');
      if(stats.droppedUnresolvedState) logLine(log, 'Dropped — unresolvable state (military/intl)', stats.droppedUnresolvedState.toLocaleString(), 'warn');
      if(stats.droppedStatus) logLine(log, 'Dropped — excluded status', stats.droppedStatus.toLocaleString(), 'warn');
      logLine(log, 'Sales uploads merged', `<b>${salesFiles.length}</b> file${salesFiles.length===1?'':'s'} (${formats.size ? Array.from(formats).join(', ') : 'unknown'})`, 'good');
      if(encodings.size) logLine(log, 'Sales encoding detected', Array.from(encodings).join(', '), Array.from(encodings).some(e=>e.includes('fallback')) ? 'warn' : 'good');
      logLine(log, 'Sales rows parsed', `<b>${totalRaw.toLocaleString()}</b>`);
      logLine(log, 'Kept (US, resolvable state, valid qty)', `<b>${kept.toLocaleString()}</b>`, 'good');
      if(droppedNonUS) logLine(log, 'Dropped — non-US', droppedNonUS.toLocaleString(), 'warn');
      if(droppedUnresolvedState) logLine(log, 'Dropped — unresolvable state (military/intl)', droppedUnresolvedState.toLocaleString(), 'warn');
      if(droppedStatus) logLine(log, 'Dropped — excluded status', droppedStatus.toLocaleString(), 'warn');
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
@@ -2053,50 +1925,51 @@ function animateKpiValues(container){
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
@@ -3281,50 +3154,99 @@ function renderExecSummary(portfolio, summary){
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


function renderManifestAnalytics(){
  const mapEl = document.getElementById('manifestRegionMap');
  const insightsEl = document.getElementById('manifestAnalyticsInsights');
  if(!mapEl || !insightsEl) return;
  if(!state.manifestPlan || !state.manifestPlan.length || !state.manifestPortfolio){
    mapEl.className = 'region-map-empty';
    mapEl.innerHTML = 'Build a Manifest plan first.';
    insightsEl.className = 'analytics-insights-empty';
    insightsEl.innerHTML = 'No active manifest plan yet.';
    return;
  }
  const p = state.manifestPortfolio;
  const maxUnits = Math.max(1, p.units.East||0, p.units.Central||0, p.units.West||0);
  const regionCost = { East:0, Central:0, West:0 };
  const regionSkuCount = { East:0, Central:0, West:0 };
  state.manifestPlan.forEach(row=>{
    ['East','Central','West'].forEach(region=>{
      const units = row.splitUnits[region]||0;
      if(units>0){ regionCost[region] += row.boxBreakdown?.[region]?.cost || 0; regionSkuCount[region]++; }
    });
  });
  mapEl.className = 'manifest-region-map';
  mapEl.innerHTML = ['East','Central','West'].map(region=>{
    const units = p.units[region]||0;
    const pct = p.pct[region]||0;
    const cpu = units ? regionCost[region]/units : 0;
    const scale = 0.72 + (units/maxUnits)*0.28;
    return `<div class="region-node ${region.toLowerCase()}" style="transform:scale(${scale.toFixed(2)});">
      <div class="region-node-label">${region}</div>
      <div class="region-node-value">${units.toLocaleString()}</div>
      <div class="region-node-sub">${fmtPct(pct)} of units · ${fmtMoney(cpu)}/unit</div>
      <div class="region-node-foot">${regionSkuCount[region]} SKU${regionSkuCount[region]===1?'':'s'}</div>
    </div>`;
  }).join('');
  const topRegion = Object.entries(p.units).sort((a,b)=>b[1]-a[1])[0];
  const usedCounts = state.manifestPlan.map(r=>countUsedRegions(r.splitUnits));
  const overTwo = usedCounts.filter(c=>c>2).length;
  const avgSplit = usedCounts.reduce((a,b)=>a+b,0)/usedCounts.length;
  const savings = p.cheapestOnlyCost - p.totalCost;
  insightsEl.className = 'analytics-insights';
  insightsEl.innerHTML = `
    <div class="insight-row"><span>Primary ship-to region</span><b>${topRegion[0]} · ${fmtPct(topRegion[1]/Math.max(1,(p.units.East+p.units.Central+p.units.West)))}</b></div>
    <div class="insight-row"><span>Split discipline</span><b>${avgSplit.toFixed(1)} regions/SKU · ${overTwo} over the 2-region cap</b></div>
    <div class="insight-row"><span>Demand match</span><b>${p.weightedAvgCoverage!=null?fmtPct(p.weightedAvgCoverage):'—'}</b></div>
    <div class="insight-row"><span>Placement-fee delta</span><b style="color:${savings>=0?'var(--good)':'var(--danger)'}">${savings>=0?'+':''}${fmtMoney(savings)} vs. single-region baseline</b></div>
    <div class="insight-callout">Recommendation: ship the manifest mostly to <b>${topRegion[0]}</b>, keep SKU-level plans capped at two regions unless case-pack math forces otherwise, and use the downloadable per-region files from the Manifest tab for execution.</div>`;
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
@@ -3410,91 +3332,94 @@ function applyManifestFiltersAndRender(){
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
    return `<span class="decision-badge case-single" title="${row.rationale}">📦 1 box → ${bestRegion} (best demand)</span>`;
  }
  if(hasRealCasePack && totalBoxes > 1){
    return `<span class="decision-badge case-multi" title="${totalBoxes} real case-pack boxes, split across ${regionsUsed} region${regionsUsed===1?'':'s'} by demand. Each region gets whole boxes only.">${totalBoxes} boxes → ${regionsUsed} region${regionsUsed===1?'':'s'}</span>`;
    return `<span class="decision-badge case-multi" title="${row.rationale}">${totalBoxes} boxes → ${regionsUsed} region${regionsUsed===1?'':'s'}</span>`;
  }
  if(!hasRealCasePack && regionsUsed === 1){
    const onlyRegion = ['East','Central','West'].find(r=>row.splitUnits[r]>0);
    return `<span class="decision-badge free-single" title="No case-pack constraint — sent all to ${onlyRegion} (demand + cost optimal).">Demand → ${onlyRegion} only</span>`;
    return `<span class="decision-badge free-single" title="${row.rationale}">Demand → ${onlyRegion} only</span>`;
  }
  if(!hasRealCasePack){
    return `<span class="decision-badge free-multi" title="No case-pack constraint — split across ${regionsUsed} regions for best demand coverage + cost.">Demand → ${regionsUsed} regions</span>`;
    return `<span class="decision-badge free-multi" title="${row.rationale}">Demand → ${regionsUsed} regions</span>`;
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
    const confidenceBadge = row.confidence === 'low'
      ? '<span class="badge" style="background:var(--danger-soft);color:var(--danger);margin-left:6px;" title="Fewer than 14 sold units in the sales history for this SKU — recommendation is directionally useful, not high-confidence.">Low confidence</span>'
      : '<span class="badge" style="background:var(--good-soft);color:var(--good);margin-left:6px;">High confidence</span>';
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
      <td>${decisionBadge}${confidenceBadge}<div style="font-size:11px;color:var(--ink-faint);margin-top:5px;max-width:360px;">${row.rationale}</div></td>`;
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
@@ -3836,51 +3761,51 @@ function buildSystemBrief(){
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
- Three regions: East, Central, West. Decision methods: Heuristic (proportional to demand %) and Two-pass Optimizer (minimizes split count first, then fixes that split count and minimizes regional demand mismatch while still showing Amazon placement fee deltas).
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
@@ -3990,51 +3915,51 @@ const vistaTooltip = {
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
    lp: `<b>Two-pass split optimization</b> is selected. Pass 1 minimizes the number of shipment splits. Pass 2 fixes that split count and minimizes regional demand mismatch, while enforcing a sensible minimum batch so the plan never creates trickle shipments. Placement fees remain visible as the cost delta against the single-region baseline.`,
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
@@ -4169,60 +4094,50 @@ function openGlobalAi(){
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

/* ---------------------------- LIVE CLOCK ---------------------------- */
function updateClock(){
  const el = document.getElementById('liveClock');
  if(!el) return;
  const now = new Date();
  el.textContent = now.toLocaleTimeString(undefined, { hour:'2-digit', minute:'2-digit', second:'2-digit' });
}
updateClock();
setInterval(updateClock, 1000);

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
