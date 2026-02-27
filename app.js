/* global XLSX */

const PREFERRED_SHEET = "MasterAppointmentsWithInsurance";

// what we want to show (in this order) when those headers exist
const desiredCols = ["Time","Patient","Chart #","Provider Profile","Appt Type","Carrier","CoPay","Pat Bal"];

const el = (id) => document.getElementById(id);

const revisedFile = el("revisedFile");
const reprintFile = el("reprintFile");
const runBtn = el("runBtn");
const clearBtn = el("clearBtn");
const statusEl = el("status");
const totalsEl = el("totals");

const strictChartOnlyEl = el("strictChartOnly");
const alsoShowReverseEl = el("alsoShowReverse");

const lateAddsTable = el("lateAddsTable");
const lateAddsThead = lateAddsTable.querySelector("thead");
const lateAddsTbody = lateAddsTable.querySelector("tbody");

const reverseCard = el("reverseCard");
const reverseTable = el("reverseTable");
const reverseThead = reverseTable.querySelector("thead");
const reverseTbody = reverseTable.querySelector("tbody");

const downloadLateAddsCsvBtn = el("downloadLateAddsCsvBtn");
const downloadReverseCsvBtn = el("downloadReverseCsvBtn");

let lastLateAdds = [];
let lastReverse = [];
let lastHeaders = [];

function setStatus(msg, tone="info"){
  statusEl.textContent = msg;
  statusEl.style.color = tone === "error" ? "var(--danger)" : tone === "ok" ? "var(--ok)" : "var(--muted)";
}

function filesReady(){
  return revisedFile.files.length && reprintFile.files.length;
}

function enableButtons(){
  const ready = filesReady();
  runBtn.disabled = !ready;
  clearBtn.disabled = !ready && !lastLateAdds.length && !lastReverse.length;
}

[revisedFile, reprintFile].forEach(inp => inp.addEventListener("change", () => {
  enableButtons();
  setStatus(filesReady() ? "Files loaded. Click Compare." : "Upload both files.");
}));

clearBtn.addEventListener("click", () => {
  revisedFile.value = "";
  reprintFile.value = "";
  lastLateAdds = [];
  lastReverse = [];
  lastHeaders = [];
  renderTotals(null);
  renderTable(lateAddsThead, lateAddsTbody, [], []);
  renderTable(reverseThead, reverseTbody, [], []);
  reverseCard.style.display = "none";
  downloadLateAddsCsvBtn.disabled = true;
  downloadReverseCsvBtn.disabled = true;
  enableButtons();
  setStatus("Cleared.");
});

downloadLateAddsCsvBtn.addEventListener("click", () => {
  if (!lastLateAdds.length) return;
  downloadBlob(toCSV(lastHeaders, lastLateAdds), `late_adds_${todayStamp()}.csv`, "text/csv;charset=utf-8;");
});

downloadReverseCsvBtn.addEventListener("click", () => {
  if (!lastReverse.length) return;
  downloadBlob(toCSV(lastHeaders, lastReverse), `revised_not_in_reprint_${todayStamp()}.csv`, "text/csv;charset=utf-8;");
});

runBtn.addEventListener("click", async () => {
  try{
    setStatus("Reading spreadsheets...");
    runBtn.disabled = true;

    const strictChartOnly = strictChartOnlyEl.checked;
    const showReverse = alsoShowReverseEl.checked;

    const revised = await readReport(revisedFile.files[0]);
    const reprint = await readReport(reprintFile.files[0]);

    setStatus("Comparing...");

    // build key sets
    const revisedKeys = new Set(revised.rows.map(r => makeKey(r, revised.idx, strictChartOnly)).filter(Boolean));
    const reprintKeys = new Set(reprint.rows.map(r => makeKey(r, reprint.idx, strictChartOnly)).filter(Boolean));

    // late adds: Reprint - Revised
    const lateAddsRaw = [];
    for (const row of reprint.rows){
      const key = makeKey(row, reprint.idx, strictChartOnly);
      if (key && !revisedKeys.has(key)) lateAddsRaw.push(row);
    }

    // reverse: Revised - Reprint (optional)
    const reverseRaw = [];
    if (showReverse){
      for (const row of revised.rows){
        const key = makeKey(row, revised.idx, strictChartOnly);
        if (key && !reprintKeys.has(key)) reverseRaw.push(row);
      }
    }

    // choose columns to render (based on the reprint file, since that’s the “current schedule”)
    const headersToShow = chooseHeadersToShow(reprint.headers, reprint.headerRow, desiredCols);
    lastHeaders = headersToShow;

    // dedupe and project to objects for display/export
    lastLateAdds = projectRows(dedupeRows(lateAddsRaw, reprint.idx), headersToShow, reprint.idx);
    lastReverse = projectRows(dedupeRows(reverseRaw, revised.idx), headersToShow, revised.idx);

    renderTotals({
      revisedRows: revised.rows.length,
      reprintRows: reprint.rows.length,
      revisedUnique: uniqueCount(revised.rows, revised.idx, strictChartOnly),
      reprintUnique: uniqueCount(reprint.rows, reprint.idx, strictChartOnly),
      lateAdds: lastLateAdds.length,
      reverse: lastReverse.length
    });

    renderTable(lateAddsThead, lateAddsTbody, headersToShow, lastLateAdds);
    downloadLateAddsCsvBtn.disabled = lastLateAdds.length === 0;

    if (showReverse){
      reverseCard.style.display = "";
      renderTable(reverseThead, reverseTbody, headersToShow, lastReverse);
      downloadReverseCsvBtn.disabled = lastReverse.length === 0;
    } else {
      reverseCard.style.display = "none";
      renderTable(reverseThead, reverseTbody, [], []);
      downloadReverseCsvBtn.disabled = true;
    }

    setStatus(`Done. Late adds found: ${lastLateAdds.length}`, "ok");
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
 * Reads your export safely without shifting columns:
 * - Finds header row
 * - Keeps header positions intact (no filtering)
 * - Stores data rows as arrays aligned to the header row
 * - Builds an index map (Patient/Chart#/etc) for comparisons
 */
async function readReport(file){
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, {type:"array"});

  const sheetName = wb.Sheets[PREFERRED_SHEET] ? PREFERRED_SHEET : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`No sheets found in ${file.name}`);

  const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});
  const hdrIndex = detectHeaderIndex(raw);
  if (hdrIndex === -1) throw new Error(`Could not detect header row in ${file.name}`);

  // keep positional headers (including blanks!)
  const headerRow = raw[hdrIndex] || [];
  const headers = headerRow.map(h => (h ?? "").toString().trim());

  // map normalized header name -> column index (first occurrence)
  const idx = buildIndexMap(headers);

  // read rows aligned to header length (no shifting)
  const rows = [];
  for (let r = hdrIndex + 1; r < raw.length; r++){
    const arr = raw[r];
    if (!arr || arr.every(v => v === null || v === "")) continue;

    // must have patient value in the Patient column (by index)
    const p = safeCell(arr, idx.patient);
    if (p == null || String(p).trim() === "") continue;

    rows.push(arr);
  }

  return { sheetName, headers, headerRow, idx, rows };
}

function detectHeaderIndex(raw){
  const look = Math.min(raw.length, 120);
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

function buildIndexMap(headers){
  const norm = (s) => String(s || "").trim().toLowerCase();

  const find = (target) => {
    const t = norm(target);
    for (let i=0; i<headers.length; i++){
      if (norm(headers[i]) === t) return i;
    }
    return -1;
  };

  const patient = find("Patient");
  const chart = find("Chart #");
  const time = find("Time");

  return {
    patient: patient === -1 ? null : patient,
    chart: chart === -1 ? null : chart,
    time: time === -1 ? null : time,
    // you can add more if you ever want index-based operations
  };
}

function safeCell(arr, idx){
  if (idx == null) return null;
  return idx < arr.length ? arr[idx] : null;
}

function normChart(v){
  if (v == null) return null;
  const s = String(v).trim().replace(/\s+/g,"");
  if (!s) return null;
  const n = Number(s);
  if (!Number.isNaN(n)) return String(Math.trunc(n));
  return s.replace(/\.0$/,"");
}

function normName(v){
  if (v == null) return null;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/^\*/,"").trim();
  s = s.replace(/\s+/g," ");
  return s.toUpperCase();
}

function makeKey(rowArr, idx, strictChartOnly){
  const chartVal = normChart(safeCell(rowArr, idx.chart));
  if (chartVal) return `C:${chartVal}`;

  if (strictChartOnly) return null;

  const nameVal = normName(safeCell(rowArr, idx.patient));
  return nameVal ? `N:${nameVal}` : null;
}

function dedupeRows(rowsArr, idx){
  const seen = new Set();
  const out = [];
  for (const row of rowsArr){
    const name = normName(safeCell(row, idx.patient)) || "";
    const chart = normChart(safeCell(row, idx.chart)) || "";
    const key = `${name}__${chart}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(row);
  }
  return out;
}

function chooseHeadersToShow(headers, headerRow, preferred){
  // Build a set of available (non-empty) headers
  const available = new Set(headers.filter(h => String(h).trim() !== ""));

  // If our preferred columns exist, use them in that order
  const picked = preferred.filter(h => available.has(h));
  if (picked.length) return picked;

  // fallback: show all non-empty headers in their original order
  return headers.filter(h => String(h).trim() !== "");
}

function projectRows(rowsArr, headersToShow, idx){
  // We need to map header text -> actual column index (by scanning the header row)
  // This preserves correct alignment even if there are blank header cells.
  // If a header appears twice, we take the first match (consistent with index map behavior).
  const colIndex = {};
  const headerLower = (s) => String(s || "").trim().toLowerCase();

  return rowsArr.map(arr => {
    const obj = {};
    for (const h of headersToShow){
      // lazy lookup: find column index for this header name
      if (colIndex[h] == null){
        colIndex[h] = findHeaderIndexByName(h, idx, headersToShow); // placeholder; overwritten below
      }
      obj[h] = ""; // filled below
    }
    // Fill values by scanning the report headers each time (robust + still fast at these sizes)
    // We’ll do it in a way that never shifts columns:
    for (const h of headersToShow){
      const col = findHeaderIndexByName(h, idx, null); // uses exact match scan
      obj[h] = (col == null) ? "" : (arr[col] ?? "");
    }
    return obj;
  });
}

function findHeaderIndexByName(name, idx, _unused){
  // We don't have the full headers array here, so we rely on the known ones via idx when possible.
  // For the standard columns we care about, use idx (fast + consistent).
  const n = String(name || "").trim().toLowerCase();
  if (n === "patient") return idx.patient;
  if (n === "chart #") return idx.chart;
  if (n === "time") return idx.time;

  // If you later add more columns and want them index-based, expand idx in buildIndexMap().
  return null;
}

function uniqueCount(rowsArr, idx, strictChartOnly){
  const set = new Set();
  for (const r of rowsArr){
    const chart = normChart(safeCell(r, idx.chart));
    if (chart) { set.add(`C:${chart}`); continue; }
    if (strictChartOnly) continue;
    const name = normName(safeCell(r, idx.patient));
    if (name) set.add(`N:${name}`);
  }
  return set.size;
}

function renderTotals(t){
  totalsEl.innerHTML = "";
  if (!t) return;

  const items = [
    ["Revised rows", t.revisedRows],
    ["Reprint rows", t.reprintRows],
    ["Unique patients in Revised", t.revisedUnique],
    ["Unique patients in Reprint", t.reprintUnique],
    ["Late adds (Reprint not in Revised)", t.lateAdds],
    ["Revised not in Reprint (optional)", t.reverse],
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

function renderTable(headEl, bodyEl, headers, rows){
  headEl.innerHTML = "";
  bodyEl.innerHTML = "";

  if (!headers.length){
    headEl.innerHTML = `<tr><th>No results</th></tr>`;
    return;
  }

  const trh = document.createElement("tr");
  for (const h of headers){
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  }
  headEl.appendChild(trh);

  if (!rows.length){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = headers.length;
    td.textContent = "No rows found.";
    tr.appendChild(td);
    bodyEl.appendChild(tr);
    return;
  }

  for (const r of rows){
    const tr = document.createElement("tr");
    for (const h of headers){
      const td = document.createElement("td");
      td.textContent = r[h] ?? "";
      tr.appendChild(td);
    }
    bodyEl.appendChild(tr);
  }
}

// init
setStatus("Upload Revised and Reprint.");
enableButtons();
