// MM Returns Planner (client-side)
// Requires SheetJS (xlsx.full.min.js)

const $ = (id) => document.getElementById(id);

const hours = [
  "09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00"
];

// Default arrivals curve (from your Excel model)
const defaultPct = [0.05,0.12,0.18,0.20,0.18,0.15,0.07,0.05];
// Default hour-by-hour utilisation curve (from SHIFT_STRUCTURE)
const defaultUtil = [0.30,0.95,1.00,1.00,1.00,1.00,0.80,1.00];

let routes = []; // [{route, parcels}]
let pct = [...defaultPct];
let util = [...defaultUtil];
let actualBags = Array(hours.length).fill(null); // editable per hour; null => use planned


async function ensureXLSX(){
  try{
    const res = await (window.__xlsxReady || Promise.resolve({ok: !!window.XLSX, url:null}));
    if(res && res.ok && window.XLSX) return true;
  }catch(e){}
  return !!window.XLSX;
}

function setImportStatus(msg, kind){
  const el = $("importStatus");
  el.textContent = msg;
  el.className = "warn" + (kind ? (" " + kind) : "");
}

function parseRoutesFromCSV(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim().length);
  if(!lines.length) return [];
  // Find header line with required columns
  let headerLineIdx = -1, routeIdx=-1, parcelsIdx=-1;
  for(let i=0;i<Math.min(lines.length,50);i++){
    const cols = lines[i].split(",").map(s=>s.trim().replace(/^"|"$/g,""));
    const lower = cols.map(c=>c.toLowerCase());
    const r = lower.indexOf("route");
    let p = lower.indexOf("parcels_returned");
    if(p===-1) p = lower.indexOf("parcels returned");
    if(p===-1) p = lower.findIndex(v=>v.includes("parcels") && v.includes("return"));
    if(r!==-1 && p!==-1){
      headerLineIdx=i; routeIdx=r; parcelsIdx=p; break;
    }
  }
  if(headerLineIdx===-1) throw new Error("CSV missing headers: Route + Parcels_Returned.");
  const out=[];
  for(let i=headerLineIdx+1;i<lines.length;i++){
    const cols = lines[i].split(",").map(s=>s.trim().replace(/^"|"$/g,""));
    const route = cols[routeIdx];
    const parcels = n0(cols[parcelsIdx]);
    if(!route && parcels<=0) continue;
    if(parcels<=0) continue;
    out.push({route, parcels});
  }
  return out;
}


function nowPill(){
  const d = new Date();
  const opts = {weekday:"short", day:"2-digit", month:"short", year:"numeric"};
  $("pillDate").textContent = d.toLocaleDateString(undefined, opts).toUpperCase();
}

function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }
function n0(x){ const v = Number(x); return Number.isFinite(v) ? v : 0; }
function ceil(n){ return Math.ceil(n); }
function round(n){ return Math.round(n); }

function sum(arr){ return arr.reduce((a,b)=>a+n0(b),0); }

function compute(){
  const totalParcels = n0($("inpTotalParcels").value);
  const vintedParcels = n0($("inpVintedParcels").value);
  const whiteParcels = n0($("inpWhiteParcels").value);
  const ppb = Math.max(1, n0($("inpParcelsPerBag").value));
  const windowHrs = Math.max(1, n0($("inpWindowHours").value));
  const utilDefault = clamp(n0($("inpUtilDefault").value), 0, 1);

  const stageSec = Math.max(0, n0($("inpStageSec").value));
  const segSec = Math.max(0, n0($("inpSegSec").value));
  const intakeSec = Math.max(0, n0($("inpIntakeSec").value));

  const peakMins = Math.max(30, n0($("inpPeakMins").value));
  const peakTopN = Math.max(1, Math.floor(n0($("inpPeakTopN").value)));
  const peakFactor = Math.max(0.5, n0($("inpPeakFactor").value));

  // Shares
  const vShare = totalParcels > 0 ? clamp(vintedParcels / totalParcels, 0, 1) : 0;
  const wShare = totalParcels > 0 ? clamp(whiteParcels / totalParcels, 0, 1) : 0;

  // Bags (daily)
  const totalBags = ceil(totalParcels / ppb);
  $("kpiBags").textContent = totalBags.toLocaleString();

  $("kpiAvgBagsHr").textContent = (totalBags / windowHrs).toFixed(0);

  // Profile % validation
  const pctSum = sum(pct);
  const pctOk = Math.abs(pctSum - 1) <= 0.0001;
  const pctStatus = $("pctStatus");
  pctStatus.textContent = pctOk ? "OK (100%)" : `FIX: sum is ${(pctSum*100).toFixed(1)}% (must be 100%)`;
  pctStatus.className = "warn " + (pctOk ? "good" : "bad");

  // Hourly plan
  const hourly = hours.map((h, i) => {
    const pct_i = n0(pct[i]);
    const parcels = round(totalParcels * pct_i);
    const bags = ceil(parcels / ppb);

    // utilisation per hour (fallback to default utilisation)
    const u = util[i] == null ? utilDefault : clamp(n0(util[i]), 0.05, 1);

    // required heads per step
    const stageHeads = ceil((((bags * stageSec) / 3600) / u));
    const segHeads = ceil((((bags * segSec) / 3600) / u));
    // intake: white parcels/hr approximated as bags * ppb * wShare (same as Excel model)
    const whiteParcelsHr = (bags * ppb * wShare);
    const intakeHeads = ceil((((whiteParcelsHr * intakeSec) / 3600) / u));
    const totalHeads = stageHeads + segHeads + intakeHeads;

    return {hour:h, pct:pct_i, util:u, parcels, bags, stageHeads, segHeads, intakeHeads, totalHeads};
  });

  // Peak hour / heads (from hourly plan)
  let peakRow = hourly[0] || {totalHeads:0, hour:"—", stageHeads:0, segHeads:0, intakeHeads:0};
  for(const row of hourly){
    if(row.totalHeads > peakRow.totalHeads){
      peakRow = row;
    }
  }
  $("kpiPeakHeads").textContent = Number.isFinite(peakRow.totalHeads) ? peakRow.totalHeads.toLocaleString() : "0";
  $("kpiPeakHour").textContent = peakRow.hour;

  // Peak safety check uses the peak hour breakdown (NOT the top-N stress test)
  const reqStage = peakRow.stageHeads;
  const reqSeg = peakRow.segHeads;
  const reqIntake = peakRow.intakeHeads;
  $("reqStage").textContent = reqStage.toLocaleString();
  $("reqSeg").textContent = reqSeg.toLocaleString();
  $("reqIntake").textContent = reqIntake.toLocaleString();

  // Optional: keep stress-test numbers for reference in import status if routes exist
  // (Top-N routes within peak window, buffered)
  const topRoutes = [...routes].sort((a,b)=>b.parcels-a.parcels).slice(0, peakTopN);
  const stressParcels = sum(topRoutes.map(r=>r.parcels)) * peakFactor;
  const stressBags = ceil(stressParcels / ppb);
  const stressHours = peakMins / 60;
  const stressBagsHr = stressHours > 0 ? (stressBags / stressHours) : 0;

  // Planned
  const planStage = n0($("planStage").value);
  const planSeg = n0($("planSeg").value);
  const planIntake = n0($("planIntake").value);
  const planFloat = n0($("planFloat").value);

  // Stream gaps/status
  const gapStage = planStage - reqStage;
  const gapSeg = planSeg - reqSeg;
  const gapIntake = planIntake - reqIntake;

  $("gapStage").textContent = gapStage.toLocaleString();
  $("gapSeg").textContent = gapSeg.toLocaleString();
  $("gapIntake").textContent = gapIntake.toLocaleString();

  const stStage = planStage >= reqStage ? "PASS" : "FAIL";
  const stSeg = planSeg >= reqSeg ? "PASS" : "FAIL";
  const stIntake = planIntake >= reqIntake ? "PASS" : "FAIL";
  $("statusStage").textContent = stStage;
  $("statusSeg").textContent = stSeg;
  $("statusIntake").textContent = stIntake;

  $("statusStage").style.color = stStage==="PASS" ? "var(--good)" : "var(--bad)";
  $("statusSeg").style.color = stSeg==="PASS" ? "var(--good)" : "var(--bad)";
  $("statusIntake").style.color = stIntake==="PASS" ? "var(--good)" : "var(--bad)";

  const totReq = reqStage + reqSeg + reqIntake;
  const totPlan = planStage + planSeg + planIntake + planFloat;
  $("totReq").textContent = totReq.toLocaleString();
  $("totPlan").textContent = totPlan.toLocaleString();

  const overall = totPlan >= totReq ? "PASS" : "FAIL";
  $("totStatus").textContent = overall;
  $("totStatus").style.color = overall==="PASS" ? "var(--good)" : "var(--bad)";
  $("kpiPeakStatus").textContent = overall;
  $("kpiPeakStatus").style.color = overall==="PASS" ? "var(--good)" : "var(--bad)";

  // Progress bar (planned vs actual volume progress)
  const plannedTotalBags = sum(hourly.map(r=>r.bags));
  const actualTotalBags = hourly.reduce((acc, r, i)=>{
    const a = actualBags[i] == null ? r.bags : Math.max(0, Math.floor(n0(actualBags[i])));
    return acc + a;
  }, 0);
  const pctAtt = plannedTotalBags > 0 ? clamp((actualTotalBags / plannedTotalBags) * 100, 0, 180) : 0;
  $("barFill").style.width = clamp(pctAtt, 0, 100) + "%";

  // Render hourly table (planned vs actual)
  const tbody = $("hourlyTable").querySelector("tbody");
  tbody.innerHTML = "";
  let cumPlanned = 0;
  let cumActual = 0;
  for(const r of hourly){
    const i = hours.indexOf(r.hour);
    const plannedB = r.bags;
    const actualB = actualBags[i] == null ? plannedB : Math.max(0, Math.floor(n0(actualBags[i])));
    cumPlanned += plannedB;
    cumActual += actualB;
    const cumGap = cumActual - cumPlanned;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.hour}</td>
      <td class="right">${plannedB.toLocaleString()}</td>
      <td class="right"><input class="inp green slim" data-actual-hour="${i}" type="number" min="0" step="1" value="${actualB.toLocaleString()}"></td>
      <td class="right">${r.util.toFixed(2)}</td>
      <td class="right">${r.stageHeads.toLocaleString()}</td>
      <td class="right">${r.segHeads.toLocaleString()}</td>
      <td class="right">${r.intakeHeads.toLocaleString()}</td>
      <td class="right">${r.totalHeads.toLocaleString()}</td>
      <td class="right">${cumGap.toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  }

  // Wire actual inputs
  tbody.querySelectorAll("input[data-actual-hour]").forEach(inp=>{
    inp.addEventListener("input", (e)=>{
      const idx = Number(e.target.getAttribute("data-actual-hour"));
      actualBags[idx] = n0(e.target.value);
      compute();
    });
  });

  // Render backlog / queue view (Queue_SIM)
  renderQueue(hourly, {
    planStage, planSeg, planIntake,
    stageSec, segSec, intakeSec,
    ppb, wShare
  });

  // Render arrivals profile table
  renderProfile();

  // Routes table (top 10)
  renderRoutesTable();

  // Import status
  const imp = $("importStatus");
  if(routes.length){
    const totalFromRoutes = sum(routes.map(r=>r.parcels));
    const stressNote = topRoutes.length
      ? ` Stress test (Top-${peakTopN}, ${peakMins}m, x${peakFactor.toFixed(2)}): ${round(stressBagsHr).toLocaleString()} bags/hr.`
      : "";
    imp.textContent = `Imported ${routes.length} routes • Total parcels from ROUTES_RAW = ${totalFromRoutes.toLocaleString()}.` + stressNote;
    imp.className = "warn good";
  } else {
    imp.textContent = "No file imported. Using manual totals.";
    imp.className = "warn";
  }
}

function renderQueue(hourly, cfg){
  const {
    planStage, planSeg, planIntake,
    stageSec, segSec, intakeSec,
    ppb, wShare
  } = cfg;

  const rateStage = stageSec > 0 ? (3600 / stageSec) : 0;
  const rateSeg = segSec > 0 ? (3600 / segSec) : 0;
  const rateIntake = intakeSec > 0 ? (3600 / intakeSec) : 0;

  let bagBacklog = 0;
  let whiteBacklog = 0;
  const tbody = $("queueTable").querySelector("tbody");
  tbody.innerHTML = "";

  for(let i=0;i<hourly.length;i++){
    const r = hourly[i];
    const arrivalsBags = r.bags;

    // Capacity scales with utilisation
    const stageCap = Math.floor(planStage * rateStage * r.util);
    const segCap = Math.floor(planSeg * rateSeg * r.util);
    const bagCap = Math.min(stageCap || 0, segCap || 0);

    bagBacklog = Math.max(0, bagBacklog + arrivalsBags - bagCap);

    const whiteDemand = Math.round(arrivalsBags * ppb * wShare);
    const intakeCap = Math.floor(planIntake * rateIntake * r.util);
    whiteBacklog = Math.max(0, whiteBacklog + whiteDemand - intakeCap);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.hour}</td>
      <td class="right">${arrivalsBags.toLocaleString()}</td>
      <td class="right">${stageCap.toLocaleString()}</td>
      <td class="right">${segCap.toLocaleString()}</td>
      <td class="right">${bagBacklog.toLocaleString()}</td>
      <td class="right">${whiteDemand.toLocaleString()}</td>
      <td class="right">${intakeCap.toLocaleString()}</td>
      <td class="right">${whiteBacklog.toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  }
}

function renderProfile(){
  const wrap = $("profileTable");
  wrap.innerHTML = "";
  // header row
  const hdr = document.createElement("div");
  hdr.className = "pRow";
  hdr.innerHTML = `<div class="muted">Hour</div><div class="muted">%</div><div class="muted">Util</div>`;
  wrap.appendChild(hdr);

  hours.forEach((h,i)=>{
    const row = document.createElement("div");
    row.className = "pRow";
    const pctInp = document.createElement("input");
    pctInp.className = "inp green slim";
    pctInp.type = "number";
    pctInp.min = "0";
    pctInp.max = "1";
    pctInp.step = "0.01";
    pctInp.value = pct[i].toFixed(2);
    pctInp.addEventListener("input", () => {
      pct[i] = clamp(n0(pctInp.value), 0, 1);
      compute();
    });

    const utilInp = document.createElement("input");
    utilInp.className = "inp green slim";
    utilInp.type = "number";
    utilInp.min = "0.05";
    utilInp.max = "1";
    utilInp.step = "0.05";
    utilInp.value = util[i].toFixed(2);
    utilInp.addEventListener("input", () => {
      util[i] = clamp(n0(utilInp.value), 0.05, 1);
      compute();
    });

    row.innerHTML = `<div>${h}</div>`;
    const pctCell = document.createElement("div"); pctCell.appendChild(pctInp);
    const utilCell = document.createElement("div"); utilCell.appendChild(utilInp);
    row.appendChild(pctCell);
    row.appendChild(utilCell);
    wrap.appendChild(row);
  });
}

function renderRoutesTable(){
  const tbody = $("routesTable").querySelector("tbody");
  tbody.innerHTML = "";
  const top10 = [...routes].sort((a,b)=>b.parcels-a.parcels).slice(0,10);
  if(!top10.length){
    $("routesFoot").textContent = "Import a file to see live route ranking.";
    return;
  }
  top10.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="right">${idx+1}</td>
      <td>${String(r.route ?? "")}</td>
      <td class="right">${n0(r.parcels).toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  });
  $("routesFoot").textContent = "Peak stress test uses Top‑N routes (configurable).";
}

function parseRoutesFromWorkbook(wb){
  // Find sheet name case-insensitive
  const sheetName = wb.SheetNames.find(n => n.toLowerCase() === "routes_raw") || wb.SheetNames.find(n => n.toLowerCase().includes("routes_raw"));
  if(!sheetName) throw new Error("Sheet ROUTES_RAW not found.");

  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});

  // Find header row containing 'route' and 'parcels' columns
  let headerIdx = -1;
  let routeCol = -1;
  let parcelsCol = -1;

  for(let i=0;i<Math.min(rows.length,200);i++){
    const row = rows[i].map(v => (typeof v === "string" ? v.trim() : v));
    const lower = row.map(v => (typeof v === "string" ? v.toLowerCase() : ""));
    const rIdx = lower.findIndex(v => v === "route");
    let pIdx = lower.findIndex(v => v === "parcels_returned");
    if(pIdx === -1) pIdx = lower.findIndex(v => v === "parcels returned");
    if(pIdx === -1) pIdx = lower.findIndex(v => v.includes("parcels") && v.includes("return"));

    if(rIdx !== -1 && pIdx !== -1){
      headerIdx = i;
      routeCol = rIdx;
      parcelsCol = pIdx;
      break;
    }
  }
  if(headerIdx === -1) throw new Error("Could not find headers: Route + Parcels_Returned.");

  const out = [];
  for(let i=headerIdx+1;i<rows.length;i++){
    const row = rows[i];
    const route = row[routeCol];
    const parcels = n0(row[parcelsCol]);
    if(route == null && (!parcels || parcels===0)) continue;
    if(parcels <= 0) continue;
    out.push({route, parcels});
  }
  return out;
}

async function handleFile(file){
  const name = (file.name || "").toLowerCase();

  // CSV fallback (no XLSX dependency)
  if(name.endsWith(".csv")){
    const text = await file.text();
    routes = parseRoutesFromCSV(text);
    const totalFromRoutes = sum(routes.map(r=>r.parcels));
    $("inpTotalParcels").value = Math.round(totalFromRoutes);
    // ROUTES_RAW provides only TOTAL parcels. Keep WHITE as user input and set VINTED = total - white.
    const whiteNow = Math.max(0, n0($("inpWhiteParcels").value));
    if(whiteNow > totalFromRoutes){
      $("inpWhiteParcels").value = Math.round(totalFromRoutes);
      $("inpVintedParcels").value = 0;
    } else {
      $("inpVintedParcels").value = Math.max(0, Math.round(totalFromRoutes - whiteNow));
    }
    setImportStatus(`Imported ${routes.length} routes from CSV • Total parcels = ${totalFromRoutes.toLocaleString()}.`, "good");
    compute();
    return;
  }

  // Excel requires XLSX
  const ok = await ensureXLSX();
  if(!ok){
    routes = [];
    setImportStatus("Import blocked: Excel parser (XLSX) failed to load. Try a different network, host on GitHub Pages, or export ROUTES_RAW as CSV and import the CSV.", "bad");
    compute();
    return;
  }

  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, {type:"array"});
  // helpful debug: list sheet names in status if parsing fails
  try{
    routes = parseRoutesFromWorkbook(wb);
  }catch(e){
    const sheets = (wb.SheetNames || []).join(", ");
    throw new Error(e.message + (sheets ? ` (Sheets found: ${sheets})` : ""));
  }

  const totalFromRoutes = sum(routes.map(r=>r.parcels));
  $("inpTotalParcels").value = Math.round(totalFromRoutes);
  // ROUTES_RAW provides only TOTAL parcels. Keep WHITE as user input and set VINTED = total - white.
  const whiteNow = Math.max(0, n0($("inpWhiteParcels").value));
  if(whiteNow > totalFromRoutes){
    $("inpWhiteParcels").value = Math.round(totalFromRoutes);
    $("inpVintedParcels").value = 0;
  } else {
    $("inpVintedParcels").value = Math.max(0, Math.round(totalFromRoutes - whiteNow));
  }
  setImportStatus(`Imported ${routes.length} routes • Total parcels from ROUTES_RAW = ${totalFromRoutes.toLocaleString()}.`, "good");
  compute();
}

function bind(){
  nowPill();

  const inputs = document.querySelectorAll("input");
  inputs.forEach(inp => inp.addEventListener("input", compute));

  // Hard refresh (matches typical dashboard expectation).
  // Browsers won't let us re-open the last imported file programmatically.
  $("btnRefresh").addEventListener("click", ()=>location.reload());

  $("fileInput").addEventListener("change", (e)=>{
    const f = e.target.files?.[0];
    if(!f) return;
    handleFile(f).catch(err=>{
      routes = [];
      setImportStatus("Import failed: " + err.message, "bad");
      compute();
    });
  });
}

bind();
compute();
