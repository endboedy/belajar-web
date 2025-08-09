// script.js - merge & filter logic menggunakan SheetJS
let dataIW39 = [], data1 = [], data2 = [], dataSUM57 = [], dataPlanning = [];
let mergedData = [];

//=================================================
// Utilities - normalize keys, find matches robustly
//=================================================
function normalizeKey(k){
  if(k===undefined || k===null) return "";
  return String(k).toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}
function normalizeRows(rows){
  return rows.map(r=>{
    const obj = {};
    Object.keys(r||{}).forEach(k=>{
      obj[normalizeKey(k)] = r[k];
    });
    return obj;
  });
}
function findRowByKey(rowsNorm, keyName, value){
  if(!value) return undefined;
  const nv = String(value).trim();
  if(nv==="") return undefined;
  return rowsNorm.find(r=>{
    // try several possible normalized key names
    if(r[normalizeKey(keyName)] && String(r[normalizeKey(keyName)]) === nv) return true;
    // fallback: check any field that equals value
    return Object.values(r).some(v => String(v) === nv);
  });
}
function findFirstKeyContains(rowNorm, patterns){
  // patterns: array of substrings to search in normalized keys
  const keys = Object.keys(rowNorm||{});
  for(const p of patterns){
    for(const k of keys){
      if(k.includes(p)) return k;
    }
  }
  return null;
}
function asNumber(v){
  if (v === undefined || v === null || v === "") return 0;
  const n = Number(String(v).toString().replace(/[^0-9.\-]/g,'')); // remove currency/commas
  return isNaN(n) ? 0 : n;
}

//=================================================
// File read helpers
//=================================================
function readExcelFile(file) {
  return new Promise((resolve) => {
    if(!file) return resolve([]);
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, {defval: ""});
        resolve(json);
      }catch(err){
        console.error("read error", err);
        resolve([]);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

//=================================================
// UI wiring
//=================================================
document.getElementById('btnUpload').addEventListener('click', ()=> showPage('pageUpload'));
document.getElementById('btnLembar').addEventListener('click', ()=> showPage('pageLembar'));
document.getElementById('btnSummary').addEventListener('click', ()=> showPage('pageSummary'));
document.getElementById('btnDownload').addEventListener('click', downloadExcel);
document.getElementById('btnProcess').addEventListener('click', processData);

// default page
showPage('pageUpload');

//=================================================
// Main: processData -> mergeData -> render
//=================================================
async function processData(){
  document.getElementById('uploadMessage').textContent = "Memproses... tunggu sebentar.";
  // read files (if missing, returns [])
  dataIW39 = await readExcelFile(document.getElementById('fileIW39').files[0]);
  data1 = await readExcelFile(document.getElementById('fileData1').files[0]);
  data2 = await readExcelFile(document.getElementById('fileData2').files[0]);
  dataSUM57 = await readExcelFile(document.getElementById('fileSUM57').files[0]);
  dataPlanning = await readExcelFile(document.getElementById('filePlanning').files[0]);

  // normalize
  const iwNorm = normalizeRows(dataIW39);
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(dataSUM57);
  const planNorm = normalizeRows(dataPlanning);

  // save normalized sets for lookups in closure
  window.__norm = {iw:iwNorm, d1:d1Norm, d2:d2Norm, sum:sumNorm, plan:planNorm};

  mergeData(iwNorm, d1Norm, d2Norm, sumNorm, planNorm);

  document.getElementById('uploadMessage').textContent = `Selesai. Baris diproses: ${mergedData.length}. Lihat di Lembar Kerja.`;
  showPage('pageLembar');
}

//=================================================
// Merge logic (sesuai rules)
//=================================================
function mergeData(iw, d1, d2, sum, plan){
  const month = document.getElementById('monthInput').value || "";
  const remanInput = String(document.getElementById('remanInput').value || "").trim();
  mergedData = [];

  // if no IW39 rows, set empty and render header
  if(!iw || iw.length===0){
    mergedData = [];
    renderTable([]);
    renderSummary();
    return;
  }

  // For each IW39 row -> create merged row
  iw.forEach(orig=>{
    // orig: normalized keys object (keys normalized)
    // But we also want original values: they are in dataIW39 list at same index.
    // Find corresponding original row by matching some field (like order or mat)
    // Simpler: use normalized orig values directly.
    const get = (k) => orig[normalizeKey(k)] ?? "";

    const Room = orig['room'] ?? orig['location'] ?? "";
    const Order = orig['order'] ?? orig['orderno'] ?? orig['noorder'] ?? "";
    const OrderType = orig['ordertype'] ?? orig['ordertype'] ?? orig['type'] ?? "";
    const DescriptionRaw = orig['description'] ?? orig['desc'] ?? "";
    const CreatedOn = orig['createdon'] ?? orig['createddate'] ?? orig['created'] ?? "";
    const UserStatus = orig['userstatus'] ?? orig['status'] ?? "";
    const MAT = orig['mat'] ?? orig['material'] ?? "";

    // find matching rows in other sheets
    const d1row = d1.find(r => {
      if(!r) return false;
      if(r['mat'] && String(r['mat']) === String(MAT)) return true;
      // fallback: check any value equals MAT
      return Object.values(r).some(v => String(v) === String(MAT));
    }) || {};

    const d2row = d2.find(r => {
      if(!r) return false;
      if(r['mat'] && String(r['mat']) === String(MAT)) return true;
      return Object.values(r).some(v => String(v) === String(MAT));
    }) || {};

    const sumrow = sum.find(r => {
      if(!r) return false;
      if(r['order'] && String(r['order']) === String(Order)) return true;
      return Object.values(r).some(v => String(v) === String(Order));
    }) || {};

    const planrow = plan.find(r => {
      if(!r) return false;
      if(r['order'] && String(r['order']) === String(Order)) return true;
      return Object.values(r).some(v => String(v) === String(Order));
    }) || {};

    // Description rule
    let Description = "";
    if(String(DescriptionRaw).toUpperCase().startsWith("JR")) {
      Description = "JR";
    } else {
      // try lookup Description from Data1
      Description = d1row['description'] || d1row['desc'] || d1row['keterangan'] || "";
    }

    // Section & CPH
    const Section = d1row['section'] || d1row['dept'] || d1row['location'] || "";
    const CPH = d2row['cph'] || d2row['costperhour'] || d2row['cphvalue'] || d2row['cph'] || "";

    // Status Part & Aging
    const StatusPart = sumrow['statuspart'] || sumrow['status'] || sumrow['partcomplete'] || "";
    const Aging = sumrow['aging'] || sumrow['age'] || "";

    // Planning & Status AMT
    const Planning = planrow['planning'] || planrow['description'] || "";
    const StatusAMT = planrow['statusamt'] || planrow['status'] || planrow['statusamt'] || "";

    // Cost calc: try find plan & actual fields in orig
    const planKey = findFirstKeyContains(orig, ['plan','totalsum','totalplan']);
    const actualKey = findFirstKeyContains(orig, ['actual','totalactual','totalsum(actual)','totalactual']);
    const planVal = planKey ? asNumber(orig[planKey]) : asNumber(orig['totalsumplan'] || orig['plan'] || 0);
    const actualVal = actualKey ? asNumber(orig[actualKey]) : asNumber(orig['totalactual'] || 0);

    let cost = "-";
    const rawCost = (planVal - actualVal) / 16500;
    if (!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) {
      cost = Number(parseFloat(rawCost).toFixed(2));
    } else {
      cost = "-";
    }

    // Include & Exclude
    let Include = "-";
    if(cost === "-"){
      Include = "-";
    } else {
      if(String(remanInput).toLowerCase().includes("reman")){
        Include = Number((cost * 0.25).toFixed(2));
      } else {
        Include = Number(Number(cost).toFixed(2));
      }
    }
    let Exclude = "-";
    if(String(OrderType).toUpperCase() === "PM38") Exclude = "-";
    else Exclude = Include;

    // Build row with exact column order required
    const rowOut = {
      "Room": Room || "",
      "Order Type": OrderType || "",
      "Order": Order || "",
      "Description": Description || "",
      "Created On": CreatedOn || "",
      "User Status": UserStatus || "",
      "MAT": MAT || "",
      "CPH": CPH || "",
      "Section": Section || "",
      "Status Part": StatusPart || "",
      "Aging": Aging || "",
      "Month": month || "",
      "Cost": cost,
      "Reman": remanInput || "",
      "Include": Include,
      "Exclude": Exclude,
      "Planning": Planning || "",
      "Status AMT": StatusAMT || ""
    };

    mergedData.push(rowOut);
  });

  renderTable(mergedData);
  renderSummary();
}

//=================================================
// Render table & filter
//=================================================
const DISPLAY_COLUMNS = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length === 0){
    // render header only
    let html = "<div class='table-wrap'><table><thead><tr>";
    DISPLAY_COLUMNS.forEach(c => html += `<th>${c}</th>`);
    html += "</tr></thead><tbody><tr><td colspan='"+DISPLAY_COLUMNS.length+"' style='text-align:center;padding:18px;color:#666'>Tidak ada data. Silakan upload file dan klik Proses Data.</td></tr></tbody></table></div>`;
    container.innerHTML = html;
    return;
  }
  // build rows
  let html = "<div class='table-wrap'><table><thead><tr>";
  DISPLAY_COLUMNS.forEach(c => html += `<th>${c}</th>`);
  html += "</tr></thead><tbody>";
  data.forEach(row => {
    html += "<tr>";
    DISPLAY_COLUMNS.forEach(col => {
      let v = row[col];
      if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      html += `<td>${v !== undefined && v !== null ? v : ""}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table></div>";
  container.innerHTML = html;
}

//=================================================
// Filtering
//=================================================
function filterTable(){
  const roomVal = (document.getElementById('filterRoom').value || "").toString().toLowerCase();
  const orderVal = (document.getElementById('filterOrder').value || "").toString().toLowerCase();
  const matVal = (document.getElementById('filterMAT').value || "").toString().toLowerCase();
  const sectionVal = (document.getElementById('filterSection').value || "").toString().toLowerCase();
  const cphVal = (document.getElementById('filterCPH').value || "").toString().toLowerCase();

  const filtered = mergedData.filter(r=>{
    return (String(r["Room"]||"").toLowerCase().includes(roomVal)) &&
           (String(r["Order"]||"").toLowerCase().includes(orderVal)) &&
           (String(r["MAT"]||"").toLowerCase().includes(matVal)) &&
           (String(r["Section"]||"").toLowerCase().includes(sectionVal)) &&
           (String(r["CPH"]||"").toLowerCase().includes(cphVal));
  });

  renderTable(filtered);
}

//=================================================
// Summary
//=================================================
function renderSummary(){
  const totalCost = mergedData.reduce((s,r)=> s + (typeof r.Cost === 'number' ? r.Cost : 0), 0);
  const totalInclude = mergedData.reduce((s,r)=> s + (typeof r.Include === 'number' ? r.Include : 0), 0);
  const totalExclude = mergedData.reduce((s,r)=> s + (typeof r.Exclude === 'number' ? r.Exclude : 0), 0);

  const el = document.getElementById('summaryContainer');
  el.innerHTML = `<div>
    <p><strong>Total baris:</strong> ${mergedData.length}</p>
    <p><strong>Total Cost:</strong> ${totalCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
    <p><strong>Total Include:</strong> ${totalInclude.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
    <p><strong>Total Exclude:</strong> ${totalExclude.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
  </div>`;
}

//=================================================
// Download hasil ke Excel
//=================================================
function downloadExcel(){
  if(!mergedData || mergedData.length===0){
    alert("Tidak ada data untuk didownload. Proses data terlebih dahulu.");
    return;
  }
  // ensure columns order
  const ws = XLSX.utils.json_to_sheet(mergedData, {header: DISPLAY_COLUMNS});
  // add header row exactly as DISPLAY_COLUMNS
  XLSX.utils.sheet_add_aoa(ws, [DISPLAY_COLUMNS], {origin: "A1"});
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hasil_Merge");
  XLSX.writeFile(wb, "Hasil_Merge.xlsx");
}

//=================================================
// Simple page switcher
//=================================================
function showPage(id){
  document.querySelectorAll('.page').forEach(p=>p.style.display='none');
  document.getElementById(id).style.display = 'block';
}
