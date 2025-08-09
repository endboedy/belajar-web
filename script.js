// script.js

// Data stores (global)
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
  if(v===undefined || v===null || v==="") return 0;
  const num = Number(String(v).toString().replace(/[^0-9.\-]/g,''));
  return isNaN(num) ? 0 : num;
}
// Format date "dd-mmm-yyyy"
function formatDate(d){
  if(!d) return "";
  let dateObj;
  if(d instanceof Date) dateObj = d;
  else dateObj = new Date(d);
  if(isNaN(dateObj.getTime())) return "";
  const day = dateObj.getDate().toString().padStart(2,'0');
  const monthNames = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const mon = monthNames[dateObj.getMonth()] || "";
  const year = dateObj.getFullYear();
  return `${day}-${mon}-${year}`;
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

// Show page helper
function showPage(pageId) {
  document.querySelectorAll('.page').forEach(p => p.style.display = 'none');
  const page = document.getElementById(pageId);
  if (page) page.style.display = 'block';
}

// Nav menu buttons
document.getElementById('navUpload').addEventListener('click', () => showPage('pageUpload'));
document.getElementById('navLembar').addEventListener('click', () => showPage('pageLembar'));
document.getElementById('navSummary').addEventListener('click', () => showPage('pageSummary'));
document.getElementById('navDownload').addEventListener('click', () => showPage('pageDownload'));

// Load Files button
document.getElementById('btnLoad').addEventListener('click', async () => {
  const loadMsg = document.getElementById('loadMsg');
  if (loadMsg) loadMsg.textContent = "Loading files...";

  const fIW = document.getElementById('fileIW39').files[0];
  const fSUM = document.getElementById('fileSUM57').files[0];
  const fPlan = document.getElementById('filePlanning').files[0];
  const fBud = document.getElementById('fileBudget').files[0];
  const fD1 = document.getElementById('fileData1').files[0];
  const fD2 = document.getElementById('fileData2').files[0];

  iwData = fIW ? await readExcelFile(fIW) : [];
  sum57 = fSUM ? await readExcelFile(fSUM) : [];
  planning = fPlan ? await readExcelFile(fPlan) : [];
  budget = fBud ? await readExcelFile(fBud) : [];
  data1 = fD1 ? await readExcelFile(fD1) : [];
  data2 = fD2 ? await readExcelFile(fD2) : [];

  if (loadMsg) loadMsg.textContent = "Files loaded.";
});

// Add Orders button - tanpa input Month & Reman di luar
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim() === ""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  if(!iwData || iwData.length === 0){
    document.getElementById('lmMsg').textContent = "Belum load IW39. Silakan upload & Load Files terlebih dahulu.";
    return;
  }

  const iwNorm = normalizeRows(iwData);
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);

  const arr = raw.split(/\r?\n|,/).map(s=>s.trim()).filter(s=>s!=="");
  arr.forEach(orderKey=>{
    const iwRow = iwNorm.find(r => Object.values(r).some(v => String(v).trim() === String(orderKey).trim()));
    if(!iwRow){
      merged.push(buildMergedRow({}, orderKey, "", "", d1Norm, d2Norm, sumNorm, planNorm));
    } else {
      merged.push(buildMergedRow(iwRow, iwRow['order'] || orderKey, "", "", d1Norm, d2Norm, sumNorm, planNorm));
    }
  });

  document.getElementById('inputOrders').value = "";
  document.getElementById('lmMsg').textContent = `Berhasil menambahkan ${arr.length} order.`;
  renderTable(merged);
});

// Build merged row from normalized iw row (origNormalized), orderKey, month, reman
function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
  const get = (o, names) => {
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]) !== "") return o[n];
    }
    return "";
  };

  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  const DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const d2row = (d2Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};

  // Description rule
  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(d1row, ['description','desc','keterangan','shortdesc']) || "";

  // CPH rule
  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "";

  const Section = get(d1row, ['section','dept','deptcode','department']);
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";
  let AgingRaw = get(sumrow, ['aging','age']) || "";

  // Format Aging tanpa angka di belakang koma
  let Aging = "";
  if(typeof AgingRaw === "number") Aging = Math.round(AgingRaw).toString();
  else Aging = String(AgingRaw).split('.')[0]; // fallback

  // Format tanggal
  const CreatedOn = formatDate(CreatedOnRaw);
  const Planning = formatDate(get(planrow, ['planning','eventstart','description','date']));

  // Status AMT
  const StatusAMT = get(planrow, ['statusamt','status']) || "";

  // Cost calculation from origNorm keys containing 'plan' and 'actual'
  const planKey = Object.keys(origNorm||{}).find(k => k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k => k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal) / 16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(2));

  // Include & Exclude calculation
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
    "Month": monthVal || "",
    "Cost": Cost,
    "Reman": remanVal || "",
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning||"",
    "Status AMT": StatusAMT||""
  };
}

// Display columns including Month & Reman
const DISPLAY = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT","Action"];

// Render table with edit/delete and inline editing for Month & Reman
function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }

  // Maksimalkan lebar dengan scroll horizontal
  container.style.overflowX = 'auto';
  container.style.width = '100%';

  let html = "<table><thead><tr>";
  DISPLAY.forEach(col => {
    html += `<th>${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col => {
      let cls = (col==="Description") ? "col-desc" : "";
      let v = row[col];

      if(col === "Order"){
        // Hilangkan angka di belakang koma (tanda titik), tampilkan saja tanpa desimal
        if(typeof v === "number") v = Math.floor(v).toString();
        else if(typeof v === "string") v = v.split('.')[0];
      } else if(col === "Created On" || col === "Planning"){
        // Sudah diformat tanggal (dd-mmm-yyyy)
      } else if(col === "Aging"){
        // Sudah dibulatkan
      } else if(typeof v === "number") {
        v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      }

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

// Edit row inline for Month & Reman
function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');

  container.style.overflowX = 'auto';
  container.style.width = '100%';

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
          if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
          html += `<td class="${cls}">${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

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
    // Recompute Include & Exclude if needed
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

function deleteRow(idx){
  if(!confirm("Hapus baris ini?")) return;
  merged.splice(idx,1);
  renderTable(merged);
  document.getElementById('lmMsg').textContent = "Baris dihapus (sementara). Klik Save Lembar Kerja untuk persist.";
}

// Save / Load Lembar Kerja to localStorage
document.addEventListener('DOMContentLoaded', () => {
  // Save / Load Lembar Kerja to localStorage
  const btnSave = document.getElementById('btnSave');
  if (btnSave) {
    btnSave.addEventListener('click', () => {
      localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Lembar Kerja tersimpan di localStorage.";
    });
  }

  function loadSaved() {
    const raw = localStorage.getItem('ndarboe_merged');
    if(raw){
      try{
        merged = JSON.parse(raw);
        renderTable(merged);
        const lmMsg = document.getElementById('lmMsg');
        if(lmMsg) lmMsg.textContent = "Lembar Kerja dimuat dari localStorage.";
      }catch(e){
        console.error(e);
        const lmMsg = document.getElementById('lmMsg');
        if(lmMsg) lmMsg.textContent = "Gagal memuat Lembar Kerja dari localStorage.";
      }
    }
  }
  loadSaved();

  // Clear Lembar Kerja
  const btnClear = document.getElementById('btnClear');
  if(btnClear){
    btnClear.addEventListener('click', () => {
      if(!confirm("Yakin ingin hapus semua Lembar Kerja?")) return;
      merged = [];
      renderTable(merged);
      localStorage.removeItem('ndarboe_merged');
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Lembar Kerja dihapus.";
    });
  }

  // Show hint message clear on inputOrders focus
  const inputOrders = document.getElementById('inputOrders');
  if(inputOrders){
    inputOrders.addEventListener('focus', () => {
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "";
    });
  }
});

// Maksimalkan section#id container lebar penuh
const sectionLembar = document.getElementById('pageLembar');
if(sectionLembar){
  sectionLembar.style.width = '100%';
  sectionLembar.style.maxWidth = '100vw';
  sectionLembar.style.overflowX = 'auto';
  sectionLembar.style.boxSizing = 'border-box';
}

// CSS style inject untuk tabel
const style = document.createElement('style');
style.textContent = `
  table {
    border-collapse: collapse;
    width: 100%;
    max-width: 100vw;
    table-layout: fixed;
    font-family: Arial, sans-serif;
    font-size: 14px;
  }
  th, td {
    border: 1px solid #666;
    padding: 8px 10px;
    word-wrap: break-word;
    white-space: normal;
    text-align: center;
  }
  th {
    background-color: #f0f0f0;
  }
  .col-desc {
    text-align: left;
    max-width: 200px;
    white-space: normal;
  }
  .action-btn {
    margin: 2px 4px;
    padding: 3px 6px;
    font-size: 12px;
    cursor: pointer;
  }
  .small {
    font-size: 11px;
    padding: 2px 5px;
  }
  #tableContainer {
    overflow-x: auto;
  }
  #inputOrders {
    width: 100%;
    min-height: 60px;
    font-family: monospace;
    font-size: 14px;
  }
`;
document.head.appendChild(style);
