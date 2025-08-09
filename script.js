// script.js
// Handles: reading excel files, lookup logic, add multi orders, save/load localStorage, edit/delete, download

// Data stores
let iwData = [];      // IW39 rows (array of objects)
let data1 = [];       // Data1 rows
let data2 = [];       // Data2 rows
let sum57 = [];       // SUM57 rows
let planning = [];    // Planning rows
let budget = [];      // Budget rows (unused currently)
let merged = [];      // hasil lembar kerja (array of objects)

// Helpers: normalize keys to simple form
function normalizeKey(k){
  return String(k||"").toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}
function normalizeRows(rows){
  return (rows||[]).map(r=>{
    const o={};
    Object.keys(r||{}).forEach(k=>{
      o[normalizeKey(k)] = r[k];
    });
    return o;
  });
}
function asNumber(v){
  if(v===undefined || v===null || v==="" ) return 0;
  const num = Number(String(v).toString().replace(/[^0-9.\-]/g,''));
  return isNaN(num) ? 0 : num;
}

// Read file as json rows (sheet 0)
function readExcelFile(file){
  return new Promise((resolve)=>{
    if(!file) return resolve([]);
    const reader = new FileReader();
    reader.onload = e=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data,{type:'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
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

// UI wiring
document.getElementById('navUpload').addEventListener('click', ()=> showPage('pageUpload'));
document.getElementById('navLembar').addEventListener('click', ()=> showPage('pageLembar'));
document.getElementById('navSummary').addEventListener('click', ()=> showPage('pageSummary'));
document.getElementById('navDownload').addEventListener('click', ()=> showPage('pageDownload'));

document.getElementById('btnLoad').addEventListener('click', async ()=>{
  const fIW = document.getElementById('fileIW39').files[0];
  const fSUM = document.getElementById('fileSUM57').files[0];
  const fPlan = document.getElementById('filePlanning').files[0];
  const fBud = document.getElementById('fileBudget').files[0];
  const fD1 = document.getElementById('fileData1').files[0];
  const fD2 = document.getElementById('fileData2').files[0];

  document.getElementById('loadMsg').textContent = "Loading files...";
  iwData = await readExcelFile(fIW);
  sum57 = await readExcelFile(fSUM);
  planning = await readExcelFile(fPlan);
  budget = await readExcelFile(fBud);
  data1 = await readExcelFile(fD1);
  data2 = await readExcelFile(fD2);

  // normalize sets for easier lookup
  iwNorm = normalizeRows(iwData);
  d1Norm = normalizeRows(data1);
  d2Norm = normalizeRows(data2);
  sumNorm = normalizeRows(sum57);
  planNorm = normalizeRows(planning);

  // keep normalized copies globally (for debugging)
  window.__nd = {iw:iwNorm,d1:d1Norm,d2:d2Norm,sum:sumNorm,plan:planNorm};

  document.getElementById('loadMsg').textContent = `Selesai load. IW39 baris: ${iwData.length}`;
});

// Add orders button
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim()===""){ document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan."; return; }
  if(!iwData || iwData.length===0){ document.getElementById('lmMsg').textContent = "Belum load IW39. Silakan upload & Load Files terlebih dahulu."; return; }

  // parse orders: split by newline or comma
  const arr = raw.split(/\r?\n|,/).map(s=>s.trim()).filter(s=>s!=="");
  const month = document.getElementById('inputMonth').value || "";
  const remanVal = document.getElementById('inputReman').value || "";

  // build for each order: lookup in iwData (match any field equals order)
  // We'll use normalized maps to find row in iwNorm
  const iwNorm = normalizeRows(iwData);
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);

  arr.forEach(orderKey=>{
    // find row in iwNorm: search any field equals orderKey (string compare)
    const iwRow = iwNorm.find(r => Object.values(r).some(v => String(v).trim() === String(orderKey).trim()));
    if(!iwRow){
      // no found -> push a minimal row with Order = orderKey and empty lookups
      merged.push(buildMergedRow({}, orderKey, month, remanVal, d1Norm, d2Norm, sumNorm, planNorm));
    } else {
      merged.push(buildMergedRow(iwRow, iwRow['order'] || orderKey, month, remanVal, d1Norm, d2Norm, sumNorm, planNorm));
    }
  });

  // clear input area
  document.getElementById('inputOrders').value = "";
  document.getElementById('lmMsg').textContent = `Berhasil menambahkan ${arr.length} order.`;
  renderTable(merged);
});

// Build merged row from normalized iw row (origNormalized), orderKey, month, reman
function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
  // origNorm has normalized keys (lowercase, no spaces)
  // helper to get original-like fields
  const get = (o, names) => {
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]) !== "") return o[n];
    }
    return "";
  };

  // Read IW fields
  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  const DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOn = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  // lookups: Data1 (section), Data2 (cph), SUM57 (status part & aging), Planning (planning,status amt)
  // find by matching MAT for data1/data2; find by Order for sum and plan
  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const d2row = (d2Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};

  // Description rule (if IW39 desc starts with JR -> Description = JR)
  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(d1row, ['description','desc','keterangan','shortdesc']) || "";

  // CPH rule: if Description starts with JR => CPH = "JR"; else lookup in data2 by MAT, fallback data1
  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "";

  const Section = get(d1row, ['section','dept','deptcode','department']);
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";
  const Aging = get(sumrow, ['aging','age']) || "";
  const Planning = get(planrow, ['planning','eventstart','description']) || "";
  const StatusAMT = get(planrow, ['statusamt','status']) || "";

  // Cost calculation: from origNorm find keys containing 'plan' and 'actual'
  const planKey = Object.keys(origNorm||{}).find(k => k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k => k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal) / 16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(2));

  // Include & Exclude
  let Include = "-";
  if(typeof Cost === "number") Include = (String(remanVal).toLowerCase().includes("reman") ? Number((Cost * 0.25).toFixed(2)) : Number(Cost.toFixed(2)));
  let Exclude = (String(OrderType).toUpperCase() === "PM38") ? "-" : Include;

  return {
    "Room": Room||"",
    "Order Type": OrderType||"",
    "Order": Order||"",
    "Description": Description||"",
    "Created On": CreatedOn||"",
    "User Status": UserStatus||"",
    "MAT": MAT||"",
    "CPH": CPH||"",
    "Section": Section||"",
    "Status Part": StatusPart||"",
    "Aging": Aging||"",
    "Month": monthVal||"",
    "Cost": Cost,
    "Reman": remanVal||"",
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning||"",
    "Status AMT": StatusAMT||""
  };
}

// Render table with action column (edit/delete)
const DISPLAY = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT","Action"];

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }

  let html = "<table><thead><tr>";
  DISPLAY.forEach(col => {
    html += `<th>${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    // all columns except Action
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col => {
      let cls = (col==="Description") ? "col-desc" : "";
      let v = row[col];
      if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      // Month and Reman should be editable inline via Edit action, but show as text here
      html += `<td class="${cls}">${v !== undefined && v !== null ? v : ""}</td>`;
    });
    // Action buttons
    html += `<td>
      <button class="action-btn small" onclick="editRow(${idx})">Edit</button>
      <button class="action-btn small" onclick="deleteRow(${idx})">Delete</button>
    </td>`;
    html += "</tr>";
  });

  html += "</tbody></table>";
  container.innerHTML = html;
}

// Edit row: convert row to inline editable (Month & Reman editable)
function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  // build table but replace the edited row with inputs
  let html = "<table><thead><tr>";
  DISPLAY.forEach(col => html += `<th>${col}</th>`);
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col=>{
      let cls = (col==="Description") ? "col-desc" : "";
      if(i === idx){
        if(col === "Month"){
          html += `<td class="${cls}">
            <select id="edit_month" >
              <option value="">--Pilih--</option>
              <option>Jan</option><option>Feb</option><option>Mar</option>
              <option>Apr</option><option>Mei</option><option>Jun</option>
              <option>Jul</option><option>Agu</option><option>Sep</option>
              <option>Okt</option><option>Nov</option><option>Des</option>
            </select>
          </td>`;
        } else if(col === "Reman"){
          html += `<td class="${cls}"><input id="edit_reman" type="text" value="${r.Reman||""}" /></td>`;
        } else {
          let v = r[col];
          if(typeof v === "number") v = v.toLocaleString();
          html += `<td class="${cls}">${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

    // action column
    if(i === idx){
      html += `<td>
        <button class="action-btn small" onclick="saveEdit(${i})">Save</button>
        <button class="action-btn small" onclick="cancelEdit()">Cancel</button>
      </td>`;
    } else {
      html += `<td>
        <button class="action-btn small" onclick="editRow(${i})">Edit</button>
        <button class="action-btn small" onclick="deleteRow(${i})">Delete</button>
      </td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table>";
  container.innerHTML = html;

  // set current values into selects/inputs
  const sel = document.getElementById('edit_month');
  if(sel) sel.value = row.Month || "";
  const rin = document.getElementById('edit_reman');
  if(rin) rin.value = row.Reman || "";
}

function saveEdit(idx){
  const month = document.getElementById('edit_month') ? document.getElementById('edit_month').value : "";
  const reman = document.getElementById('edit_reman') ? document.getElementById('edit_reman').value : "";
  if(merged[idx]){
    merged[idx].Month = month;
    merged[idx].Reman = reman;
    // recompute Include & Exclude if needed
    const cost = merged[idx].Cost;
    if(typeof cost === "number"){
      merged[idx].Include = String(reman).toLowerCase().includes("reman") ? Number((cost * 0.25).toFixed(2)) : Number(cost.toFixed(2));
    } else {
      merged[idx].Include = "-";
    }
    merged[idx].Exclude = String(merged[idx]["Order Type"]).toUpperCase() === "PM38" ? "-" : merged[idx].Include;
  }
  renderTable(merged);
  document.getElementById('lmMsg').textContent = "Perubahan disimpan (sementara). Jangan lupa klik Save Lembar Kerja untuk persist.";
}

function cancelEdit(){
  renderTable(merged);
}

// delete row
function deleteRow(idx){
  if(!confirm("Hapus baris ini?")) return;
  merged.splice(idx,1);
  renderTable(merged);
  document.getElementById('lmMsg').textContent = "Baris dihapus (sementara). Klik Save Lembar Kerja untuk persist.";
}

// Save / Load Lembar Kerja to localStorage
document.getElementById('btnSave').addEventListener('click', ()=>{
  localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
  document.getElementById('lmMsg').textContent = "Lembar Kerja tersimpan di localStorage.";
});

// load on start
function loadSaved(){
  const raw = localStorage.getItem('ndarboe_merged');
  if(raw){
    try{
      merged = JSON.parse(raw);
      renderTable(merged);
    }catch(e){
      console.error("parse saved",e);
    }
  }
}
loadSaved();

// Summary page logic
document.getElementById('summaryMonth').addEventListener('change', ()=>{
  const month = document.getElementById('summaryMonth').value;
  if(!month){ document.getElementById('summaryResult').innerHTML = "Pilih bulan untuk melihat ringkasan."; return; }
  const rows = (merged||[]).filter(r => String(r.Month||"").toLowerCase() === String(month).toLowerCase());
  if(!rows || rows.length===0){
    document.getElementById('summaryResult').innerHTML = `<div class="hint">Tidak ada data untuk bulan ${month}.</div>`;
    return;
  }
  // compute totals
  const totalCost = rows.reduce((s,r)=> s + (typeof r.Cost === 'number' ? r.Cost : 0), 0);
  const totalInclude = rows.reduce((s,r)=> s + (typeof r.Include === 'number' ? r.Include : 0), 0);
  const totalExclude = rows.reduce((s,r)=> s + (typeof r.Exclude === 'number' ? r.Exclude : 0), 0);

  let html = `<div><p><strong>Baris:</strong> ${rows.length}</p>
    <p><strong>Total Cost:</strong> ${totalCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
    <p><strong>Total Include:</strong> ${totalInclude.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
    <p><strong>Total Exclude:</strong> ${totalExclude.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p></div>`;

  document.getElementById('summaryResult').innerHTML = html;
});

// Download excel
document.getElementById('btnDownloadExcel').addEventListener('click', ()=>{
  if(!merged || merged.length===0){ alert("Tidak ada data untuk di-download."); return; }
  const header = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
  const ws = XLSX.utils.json_to_sheet(merged, {header: header});
  XLSX.utils.sheet_add_aoa(ws, [header], {origin:"A1"});
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hasil_Merge");
  XLSX.writeFile(wb, "Hasil_Merge.xlsx");
});

// Show Page helper
function showPage(id){
  document.querySelectorAll('.page').forEach(p=> p.style.display = 'none');
  document.getElementById(id).style.display = 'block';
}

// also nav buttons outside
document.getElementById('navUpload').addEventListener('click', ()=> showPage('pageUpload'));
document.getElementById('navLembar').addEventListener('click', ()=> showPage('pageLembar'));
document.getElementById('navSummary').addEventListener('click', ()=> showPage('pageSummary'));
document.getElementById('navDownload').addEventListener('click', ()=> showPage('pageDownload'));

// initial page
showPage('pageUpload');
