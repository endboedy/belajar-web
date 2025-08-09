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

// Fungsi ganti halaman
function showPage(pageId) {
  document.querySelectorAll('.page').forEach(p => p.style.display = 'none');
  const page = document.getElementById(pageId);
  if (page) page.style.display = 'block';
}

// Pasang event click di menu utama
document.getElementById('navUpload').addEventListener('click', () => showPage('pageUpload'));
document.getElementById('navLembar').addEventListener('click', () => showPage('pageLembar'));
document.getElementById('navSummary').addEventListener('click', () => showPage('pageSummary'));
document.getElementById('navDownload').addEventListener('click', () => showPage('pageDownload'));

// Event click tombol Load Files
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

  if (loadMsg) loadMsg.textContent = `Files loaded. IW39 baris: ${iwData.length}`;
});

// Tombol Add Orders
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim()===""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  if(!iwData || iwData.length===0){
    document.getElementById('lmMsg').textContent = "Belum load IW39. Silakan upload & Load Files terlebih dahulu.";
    return;
  }

  // parse orders: split by newline or comma
  const arr = raw.split(/\r?\n|,/).map(s=>s.trim()).filter(s=>s!=="");

  // Build normalized rows for lookup
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
      merged.push(buildMergedRow({}, orderKey, "", "", d1Norm, d2Norm, sumNorm, planNorm));
    } else {
      merged.push(buildMergedRow(iwRow, iwRow['order'] || orderKey, "", "", d1Norm, d2Norm, sumNorm, planNorm));
    }
  });

  // clear input area
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

  // Dari IW39
  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  const DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  // Lookup Section dari Data1 by MAT
  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  // Lookup Status Part dari SUM57 by Order
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};
  // Lookup Planning by Order
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};

  // Description dari IW39 langsung
  const Description = DescriptionRaw || "";

  // Section dari Data1
  const Section = get(d1row, ['section','dept','deptcode','department']) || "";

  // Status Part dari SUM57
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";

  // Aging dari SUM57, hapus angka di belakang koma
  let Aging = get(sumrow, ['aging','age']) || "";
  if(typeof Aging === "number") Aging = Math.floor(Aging);
  else if(!isNaN(Number(Aging))) Aging = Math.floor(Number(Aging));

  // Format Created On jadi "dd-mmm-yyyy"
  let CreatedOn = "";
  if(CreatedOnRaw){
    const d = new Date(CreatedOnRaw);
    if(!isNaN(d)) CreatedOn = d.toLocaleDateString('en-GB', {day:'2-digit', month:'short', year:'numeric'}).replace(/ /g, '-');
    else CreatedOn = CreatedOnRaw;
  }

  // Format Planning jadi "dd-mmm-yyyy" jika berupa tanggal
  let Planning = "";
  const planDateRaw = get(planrow, ['planning','eventstart','description']);
  if(planDateRaw){
    const dp = new Date(planDateRaw);
    if(!isNaN(dp)) Planning = dp.toLocaleDateString('en-GB', {day:'2-digit', month:'short', year:'numeric'}).replace(/ /g, '-');
    else Planning = planDateRaw;
  }

  // Cost dan Reman sama seperti sebelumnya
  const planKey = Object.keys(origNorm||{}).find(k => k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k => k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal) / 16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(2));

  let Include = "-";
  if(typeof Cost === "number") Include = (String(remanVal).toLowerCase().includes("reman") ? Number((Cost * 0.25).toFixed(2)) : Number(Cost.toFixed(2)));
  let Exclude = (String(OrderType).toUpperCase() === "PM38") ? "-" : Include;

  return {
    "Room": Room||"",
    "Order Type": OrderType||"",
    "Order": Order ? String(Order) : "",
    "Description": Description||"",
    "Created On": CreatedOn||"",
    "User Status": UserStatus||"",
    "MAT": MAT||"",
    "CPH": get(d1row, ['cph','costperhour','cphvalue']) || "",
    "Section": Section||"",
    "Status Part": StatusPart||"",
    "Aging": Aging||"",
    "Cost": Cost,
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning||"",
    "Status AMT": get(planrow, ['statusamt','status']) || ""
  };
}

// Render table with action column (edit/delete)
const DISPLAY = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Cost","Include","Exclude","Planning","Status AMT","Action"];

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }

  let html = `<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;"><thead><tr>`;
  DISPLAY.forEach(col => {
    html += `<th style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    const cols = DISPLAY.filter(c => c !== "Action");
    cols.forEach(col => {
      let cls = (col === "Description") ? "col-desc" : "";
      let v = row[col];

      // Formatting khusus
      if(col === "Order"){
        v = String(v); // tampil utuh tanpa desimal
      } else if(col === "Aging"){
        if(typeof v === "number") v = Math.floor(v);
        else if(!isNaN(Number(v))) v = Math.floor(Number(v));
      } else if(typeof v === "number"){
        // angka lain tampil 2 desimal
        v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      }

      html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${v !== undefined && v !== null ? v : ""}</td>`;
    });

    // Action buttons
    html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
      <button class="action-btn small" onclick="editRow(${idx})">Edit</button>
      <button class="action-btn small" onclick="deleteRow(${idx})">Delete</button>
    </td>`;

    html += "</tr>";
  });

  html += "</tbody></table></div>";
  container.innerHTML = html;
}

// Edit row: convert row to inline editable (Month & Reman editable) - kita hapus edit Month & Reman sesuai request
function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  // build table but replace the edited row with inputs
  let html = "<table style='width:100%;border-collapse:collapse;'><thead><tr>";
  DISPLAY.forEach(col => html += `<th style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${col}</th>`);
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col=>{
      let cls = (col==="Description") ? "col-desc" : "";
      if(i === idx){
        // Karena Month dan Reman sudah dihilangkan, tampil semua read-only kecuali Reman dan Month tidak ada
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${v !== undefined && v !== null ? v : ""}</td>`;
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

    // action column
    if(i === idx){
      html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
        <button class="action-btn small" onclick="cancelEdit()">Cancel</button>
      </td>`;
    } else {
      html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
        <button class="action-btn small" onclick="editRow(${i})">Edit</button>
        <button class="action-btn small" onclick="deleteRow(${i})">Delete</button>
      </td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";
  container.innerHTML = html;
}

function saveEdit(idx){
  // Karena tidak ada field yg bisa diedit, saveEdit hanya batal saja
  cancelEdit();
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
  if(!merged || merged.length === 0){
    alert("Tidak ada data untuk diunduh.");
    return;
  }
  const wb = XLSX.utils.book_new();
  const ws_data = [];
  const cols = DISPLAY.filter(c => c !== "Action");
  ws_data.push(cols);
  merged.forEach(r=>{
    ws_data.push(cols.map(c=>r[c]));
  });
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "Lembar Kerja");
  XLSX.writeFile(wb, "LembarKerja_Export.xlsx");
});

// CSS styles
const style = document.createElement('style');
style.textContent = `
  table {
    width: 100%;
    border-collapse: collapse;
    font-family: Arial, sans-serif;
    font-size: 14px;
  }
  th, td {
    border: 1px solid #ccc;
    padding: 6px 10px;
    white-space: nowrap;
  }
  .col-desc {
    max-width: 300px;
    white-space: normal;
    word-wrap: break-word;
  }
  .action-btn.small {
    font-size: 12px;
    padding: 4px 8px;
    margin: 0 2px;
    cursor: pointer;
  }
  .hint {
    font-style: italic;
    color: #555;
  }
`;
document.head.appendChild(style);

// Init page awal: tampil halaman upload
showPage('pageUpload');
