/* global XLSX */

const PREFERRED_SHEET = "MasterAppointmentsWithInsurance";
const desiredCols = ["Patient","Chart #","Time","Provider Profile","Appt Type","Carrier","CoPay","Pat Bal"];

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
  const csv = toCSV(lastHeaders, lastLateAdds);
  downloadBlob(csv, `late_adds_${todayStamp()}.csv`, "text/csv;charset=utf-8;");
});

downloadReverseCsvBtn.addEventListener("click", () => {
  if (!lastReverse.length) return;
  const csv = toCSV(lastHeaders, lastReverse);
  downloadBlob(csv, `revised_not_in_reprint_${todayStamp()}.csv`, "text/csv;charset=utf-8;");
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

    // Key sets
    const revisedKeys = new Set(makeKeys(revised.rows, strictChartOnly).filter(Boolean));
    const reprintKeys = new Set(makeKeys(reprint.rows, strictChartOnly).filter(Boolean));

    // Reprint - Revised = late adds
    const lateAddsRaw = [];
    for (const row of reprint.rows){
      const key = makeKey(row, strictChartOnly);
      if (key && !revisedKeys.has(key)) lateAddsRaw.push(row);
    }

    // Revised - Reprint (optional)
    const reverseRaw = [];
    if (showReverse){
      for (const row of revised.rows){
        const key = makeKey(row, strictChartOnly);
        if (key && !reprintKeys.has(key)) reverseRaw.push(row);
      }
    }

    // Columns
    const headers = pickColumns(reprint.headers, desiredCols);
    lastHeaders = headers;

    // Dedup + project
    lastLateAdds = projectRows(dedupeRows(lateAddsRaw), headers);
    lastReverse = projectRows(dedupeRows(reverseRaw), headers);

    renderTotals({
      revisedRows: revised.rows.length,
      reprintRows: reprint.rows.length,
      revisedUnique: uniqueCount(revised.rows, strictChartOnly),
      reprintUnique: uniqueCount(reprint.rows, strictChartOnly),
      lateAdds: lastLateAdds.length,
      reverse: lastReverse.length
    });

    renderTable(lateAddsThead, lateAddsTbody, headers, lastLateAdds);
    downloadLateAddsCsvBtn.disabled = lastLateAdds.length === 0;

    if (showReverse){
      reverseCard.style.display = "";
      renderTable(reverseThead, reverseTbody, headers, lastReverse);
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

/** Read report with header-row auto detect */
async function readReport(file){
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, {type:"array"});

  // prefer expected sheet; fallback to first sheet
  const sheetName = wb.Sheets[PREFERRED_SHEET] ? PREFERRED_SHEET : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`No sheets found in ${file.name}`);

  // raw as arrays
  const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});
  const hdrIndex = detectHeaderIndex(raw);
  if (hdrIndex === -1) throw new Error(`Could not detect header row in ${file.name}`);

  const headers = raw[hdrIndex].map(h => (h ?? "").toString().trim()).filter(Boolean);
  const rows = [];

  for (let r = hdrIndex+1; r < raw.length; r++){
    const rowArr = raw[r];
    if (!rowArr || rowArr.every(v => v === null || v === "")) continue;

    const obj = {};
    for (let c=0; c<headers.length; c++){
      const key = headers[c];
      obj[key] = rowArr[c] ?? null;
    }
    if (obj["Patient"] == null || obj["Patient"] === "") continue;
    rows.push(obj);
  }

  return { sheetName, headers, rows };
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

function makeKey(row, strictChartOnly){
  const chart = normChart(row["Chart #"]);
  if (chart) return `C:${chart}`;
  if (strictChartOnly) return null;
  const name = normName(row["Patient"]);
  return name ? `N:${name}` : null;
}

function makeKeys(rows, strictChartOnly){
  return rows.map(r => makeKey(r, strictChartOnly));
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
  return cols.length ? cols : allHeaders.filter(Boolean);
}

function projectRows(rows, headers){
  return rows.map(row => {
    const out = {};
    for (const h of headers) out[h] = row[h] ?? "";
    return out;
  });
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
setStatus("Upload Revised and Reprint.");
enableButtons();
