/* global XLSX */

const SHEET_NAME = "MasterAppointmentsWithInsurance"; // your exports use this
const desiredCols = ["Patient","Chart #","Time","Provider Profile","Appt Type","Carrier","CoPay","Pat Bal"];

const el = (id) => document.getElementById(id);

const copayFile = el("copayFile");
const ms1File = el("ms1File");
const ms2File = el("ms2File");
const runBtn = el("runBtn");
const clearBtn = el("clearBtn");
const statusEl = el("status");
const totalsEl = el("totals");
const strictChartOnlyEl = el("strictChartOnly");
const downloadCsvBtn = el("downloadCsvBtn");

const table = el("resultTable");
const thead = table.querySelector("thead");
const tbody = table.querySelector("tbody");

let lastMissingRows = [];
let lastHeaders = [];

function setStatus(msg, tone="info"){
  statusEl.textContent = msg;
  statusEl.style.color = tone === "error" ? "var(--danger)" : tone === "ok" ? "var(--ok)" : "var(--muted)";
}

function filesReady(){
  return copayFile.files.length && ms1File.files.length && ms2File.files.length;
}

function enableButtons(){
  const ready = filesReady();
  runBtn.disabled = !ready;
  clearBtn.disabled = !ready && !lastMissingRows.length;
}

[copayFile, ms1File, ms2File].forEach(inp => inp.addEventListener("change", () => {
  enableButtons();
  setStatus(filesReady() ? "Files loaded. Click Compare." : "Upload all three files.");
}));

clearBtn.addEventListener("click", () => {
  copayFile.value = "";
  ms1File.value = "";
  ms2File.value = "";
  lastMissingRows = [];
  lastHeaders = [];
  renderTotals(null);
  renderTable([], []);
  downloadCsvBtn.disabled = true;
  enableButtons();
  setStatus("Cleared.");
});

downloadCsvBtn.addEventListener("click", () => {
  if (!lastMissingRows.length) return;
  const csv = toCSV(lastHeaders, lastMissingRows);
  downloadBlob(csv, `late_adds_${todayStamp()}.csv`, "text/csv;charset=utf-8;");
});

runBtn.addEventListener("click", async () => {
  try{
    setStatus("Reading spreadsheets...");
    runBtn.disabled = true;

    const strictChartOnly = strictChartOnlyEl.checked;

    const copay = await readReport(copayFile.files[0]);
    const ms1 = await readReport(ms1File.files[0]);
    const ms2 = await readReport(ms2File.files[0]);

    setStatus("Comparing...");

    const msKeys = new Set([...makeKeys(ms1, strictChartOnly), ...makeKeys(ms2, strictChartOnly)].filter(Boolean));

    const copayKeys = makeKeys(copay, strictChartOnly);

    const missing = [];
    for (let i=0; i<copay.rows.length; i++){
      const key = copayKeys[i];
      if (!key) continue;
      if (!msKeys.has(key)){
        missing.push(copay.rows[i]);
      }
    }

    // de-dupe by Patient + Chart # if possible
    const deduped = dedupeRows(missing);

    // prepare headers and render
    const headers = pickColumns(copay.headers, desiredCols);

    lastMissingRows = deduped.map(r => projectRow(r, headers));
    lastHeaders = headers;

    renderTotals({
      copayRows: copay.rows.length,
      ms1Rows: ms1.rows.length,
      ms2Rows: ms2.rows.length,
      uniqueCopayPatients: uniqueCount(copay.rows, strictChartOnly),
      uniqueMsPatients: uniqueCount([...ms1.rows, ...ms2.rows], strictChartOnly),
      missingCount: lastMissingRows.length
    });

    renderTable(headers, lastMissingRows);
    downloadCsvBtn.disabled = lastMissingRows.length === 0;

    setStatus(`Done. Late adds found: ${lastMissingRows.length}`, "ok");
  }catch(err){
    console.error(err);
    setStatus(err?.message || "Something went wrong.", "error");
  }finally{
    runBtn.disabled = !filesReady();
    enableButtons();
  }
});

function todayStamp(){
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}

function toCSV(headers, rows){
  const esc = (v) => {
    const s = (v ?? "").toString();
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
    return s;
  };
  const lines = [];
  lines.push(headers.map(esc).join(","));
  for (const r of rows){
    lines.push(headers.map(h => esc(r[h])).join(","));
  }
  return lines.join("\n");
}

function downloadBlob(content, filename, mime){
  const blob = new Blob([content], {type: mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/**
 * Read report:
 * - Finds the header row by scanning first ~80 rows for 'Patient' and 'Chart'
 * - Returns { headers: [...], rows: [ {col:value,...}, ... ] }
 */
async function readReport(file){
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, {type:"array"});
  const ws = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error(`No sheets found in ${file.name}`);

  // get raw rows as arrays
  const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});

  const hdrIndex = detectHeaderIndex(raw);
  if (hdrIndex === -1) throw new Error(`Could not detect header row in ${file.name}`);

  const headers = raw[hdrIndex].map(h => (h ?? "").toString().trim());
  const rows = [];

  for (let r = hdrIndex+1; r < raw.length; r++){
    const rowArr = raw[r];
    if (!rowArr || rowArr.every(v => v === null || v === "")) continue;

    const obj = {};
    for (let c=0; c<headers.length; c++){
      const key = headers[c];
      if (!key) continue;
      obj[key] = rowArr[c] ?? null;
    }
    // require Patient column if present
    if (obj["Patient"] == null || obj["Patient"] === "") continue;
    rows.push(obj);
  }

  return { headers, rows };
}

function detectHeaderIndex(raw){
  const look = Math.min(raw.length, 90);
  for (let i=0; i<look; i++){
    const row = (raw[i] || []).map(v => (v ?? "").toString().trim().toLowerCase());
    const hasPatient = row.includes("patient");
    const hasChart = row.some(x => x.includes("chart"));
    const hasTime = row.some(x => x.includes("time"));
    if (hasPatient && hasChart && hasTime) return i;
  }
  for (let i=0; i<look; i++){
    const row = (raw[i] || []).map(v => (v ?? "").toString().trim().toLowerCase());
    const hasPatient = row.includes("patient");
    const hasChart = row.some(x => x.includes("chart"));
    if (hasPatient && hasChart) return i;
  }
  return -1;
}

function normChart(v){
  if (v == null) return null;
  const s = String(v).trim().replace(/\s+/g,"");
  if (!s) return null;
  // handle numeric-like charts (e.g. 163794.0)
  const n = Number(s);
  if (!Number.isNaN(n)) return String(Math.trunc(n));
  return s.replace(/\.0$/,"");
}

function normName(v){
  if (v == null) return null;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/^\*/,"").trim();     // remove leading '*'
  s = s.replace(/\s+/g," ");
  return s.toUpperCase();
}

/**
 * Key rules:
 * - Default: Chart # if present, else name fallback
 * - Strict: Chart # only (name ignored)
 */
function makeKeys(report, strictChartOnly){
  return report.rows.map(r => {
    const chart = normChart(r["Chart #"]);
    if (chart) return `C:${chart}`;
    if (strictChartOnly) return null;
    const name = normName(r["Patient"]);
    return name ? `N:${name}` : null;
  });
}

function dedupeRows(rows){
  const seen = new Set();
  const out = [];
  for (const r of rows){
    const name = normName(r["Patient"]) || "";
    const chart = normChart(r["Chart #"]) || "";
    const key = `${name}__${chart}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(r);
  }
  return out;
}

function pickColumns(allHeaders, preferred){
  const set = new Set(allHeaders);
  const cols = preferred.filter(c => set.has(c));
  // if none matched, just use all headers (rare)
  return cols.length ? cols : allHeaders.filter(Boolean);
}

function projectRow(row, headers){
  const out = {};
  for (const h of headers) out[h] = row[h] ?? "";
  return out;
}

function renderTotals(t){
  totalsEl.innerHTML = "";
  if (!t) return;

  const items = [
    ["Copay rows", t.copayRows],
    ["MS1 rows", t.ms1Rows],
    ["MS2 rows", t.ms2Rows],
    ["Unique patients in Copay", t.uniqueCopayPatients],
    ["Unique patients in (MS1 ∪ MS2)", t.uniqueMsPatients],
    ["Late adds (Copay not in MS1/MS2)", t.missingCount],
  ];

  for (const [label, value] of items){
    const div = document.createElement("div");
    div.className = "totalCard";
    div.innerHTML = `
      <div class="totalLabel">${label}</div>
      <div class="totalValue">${value}</div>
    `;
    totalsEl.appendChild(div);
  }
}

function renderTable(headers, rows){
  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!headers.length){
    thead.innerHTML = `<tr><th>No results</th></tr>`;
    return;
  }

  const trh = document.createElement("tr");
  for (const h of headers){
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  if (!rows.length){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = headers.length;
    td.textContent = "No late adds found.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  for (const r of rows){
    const tr = document.createElement("tr");
    for (const h of headers){
      const td = document.createElement("td");
      td.textContent = r[h] ?? "";
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}

function uniqueCount(rows, strictChartOnly){
  const set = new Set();
  for (const r of rows){
    const chart = normChart(r["Chart #"]);
    if (chart) { set.add(`C:${chart}`); continue; }
    if (strictChartOnly) continue;
    const name = normName(r["Patient"]);
    if (name) set.add(`N:${name}`);
  }
  return set.size;
}

// init
setStatus("Upload all three files.");
enableButtons();
