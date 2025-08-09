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

// Format tanggal ke dd-mmm-yyyy (contoh: 12-Aug-2025)
function formatDateString(input){
  if(!input) return "";
  let d;
  if(input instanceof Date){
    d = input;
  } else {
    d = new Date(input);
    if(isNaN(d)) return input; // kalau bukan tanggal valid, return apa adanya
  }
  const day = d.getDate().toString().padStart(2,'0');
  const monthNames = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const month = monthNames[d.getMonth()] || "???";
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
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

  if (loadMsg) loadMsg.textContent = "Files loaded.";

  // Clear previous merged data on load new files
  merged = [];
  localStorage.removeItem('ndarboe_merged');
  renderTable(merged);

  document.getElementById('lmMsg').textContent = "Silakan input order dan Add Orders.";
});

// Tombol Add Orders
document.getElementById('btnAddOrders').addEventListener('click', () => {
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim()===""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  if(!iwData || iwData.length===0){
    document.getElementById('lmMsg').textContent = "Belum load IW39. Silakan upload & Load Files terlebih dahulu.";
    return;
  }

  const arr = raw.split(/\r?\n|,/).map(s=>s.trim()).filter(s=>s!=="");
  const remanVal = ""; // Karena dihilangkan input Month & Reman atas tabel

  // normalize source data
  const iwNorm = normalizeRows(iwData);
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);

  arr.forEach(orderKey=>{
    // find row in iwNorm: search any field equals orderKey (string compare)
    const iwRow = iwNorm.find(r => Object.values(r).some(v => String(v).trim() === String(orderKey).trim()));
    if(!iwRow){
      merged.push(buildMergedRow({}, orderKey, "", remanVal, d1Norm, d2Norm, sumNorm, planNorm));
    } else {
      merged.push(buildMergedRow(iwRow, iwRow['order'] || orderKey, "", remanVal, d1Norm, d2Norm, sumNorm, planNorm));
    }
  });

  document.getElementById('inputOrders').value = "";
  document.getElementById('lmMsg').textContent = `Berhasil menambahkan ${arr.length} order.`;

  // Simpan hasil merged ke localStorage
  localStorage.setItem('ndarboe_merged', JSON.stringify(merged));

  renderTable(merged);
});

// Build merged row from normalized iw row (origNormalized), orderKey, month, reman
function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
  const get = (o, names) => {
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]).trim() !== "") return o[n];
    }
    return null; // null kalau tidak ketemu
  };

  // IW fields
  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  let DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  // lookups
  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || null;
  const d2row = (d2Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || null;
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || null;
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || null;

  // Description rule
  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(d1row, ['description','desc','keterangan','shortdesc']) || "Tidak Ada";

  // CPH rule
  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "Tidak Ada";

  const Section = get(d1row, ['section','dept','deptcode','department']) || "Tidak Ada";
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "Tidak Ada";
  let Aging = get(sumrow, ['aging','age']);
  const PlanningRaw = get(planrow, ['planning','eventstart','description']);
  const StatusAMT = get(planrow, ['statusamt','status']) || "Tidak Ada";

  // Format tanggal Created On & Planning
  const CreatedOn = formatDateString(CreatedOnRaw) || "Tidak Ada";
  const Planning = formatDateString(PlanningRaw) || "Tidak Ada";

  // Cost calculation
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

  // Buat Aging bulatkan tanpa desimal atau jika null kasih "Tidak Ada"
  if(typeof Aging === "number") Aging = Math.floor(Aging);
  else if(!isNaN(Number(Aging))) Aging = Math.floor(Number(Aging));
  else Aging = "Tidak Ada";

  // Hilangkan koma/desimal di Order (ubah ke string agar tampil bersih)
  const OrderClean = String(Order);

  // UserStatus & Room & OrderType => Tampilkan "OK" kalau ada data, "Tidak Ada" kalau null/empty
  const RoomDisp = Room ? "OK" : "Tidak Ada";
  const OrderTypeDisp = OrderType ? "OK" : "Tidak Ada";
  const UserStatusDisp = UserStatus ? "OK" : "Tidak Ada";

  return {
    "Room": RoomDisp,
    "Order Type": OrderTypeDisp,
    "Order": OrderClean || "Tidak Ada",
    "Description": Description,
    "Created On": CreatedOn,
    "User Status": UserStatusDisp,
    "MAT": MAT || "Tidak Ada",
    "CPH": CPH,
    "Section": Section,
    "Status Part": StatusPart,
    "Aging": Aging,
    "Month": monthVal || "",
    "Cost": Cost,
    "Reman": remanVal || "",
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning,
    "Status AMT": StatusAMT
  };
}

// Render tabel dinamis dengan semua kolom lengkap
function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }

  // Ambil semua kolom unik dari data
  const allColsSet = new Set();
  data.forEach(row => {
    Object.keys(row).forEach(k => allColsSet.add(k));
  });
  const allCols = Array.from(allColsSet);

  // Pastikan kolom Action selalu terakhir
  if(!allCols.includes("Action")) allCols.push("Action");

  let html = `<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;"><thead><tr>`;
  allCols.forEach(col => {
    html += `<th style="border:1px solid #ccc;padding:8px;white-space:nowrap;background:#eee;">${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    allCols.forEach(col => {
      if(col === "Action"){
        html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
          <button class="action-btn small" onclick="editRow(${idx})">Edit</button>
          <button class="action-btn small" onclick="deleteRow(${idx})">Delete</button>
        </td>`;
      } else {
        let v = row[col];

        // Format khusus
        if(col === "Order"){
          v = String(v);
        } else if(col === "Aging"){
          if(typeof v === "number") v = Math.floor(v);
          else if(!isNaN(Number(v))) v = Math.floor(Number(v));
        } else if(typeof v === "number"){
          v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        }
        const cls = (col === "Description") ? "col-desc" : "";
        html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;white-space:nowrap;">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });
    html += "</tr>";
  });

  html += "</tbody></table></div>";
  container.innerHTML = html;
}

// Edit row (inline editable Month & Reman)
function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  // ambil semua kolom
  const allColsSet = new Set();
  merged.forEach(r => Object.keys(r).forEach(k => allColsSet.add(k)));
  const allCols = Array.from(allColsSet);
  if(!allCols.includes("Action")) allCols.push("Action");

  // bangun tabel baru dengan baris edit
  let html = `<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;"><thead><tr>`;
  allCols.forEach(col => {
    html += `<th style="border:1px solid #ccc;padding:8px;white-space:nowrap;background:#eee;">${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    allCols.forEach(col => {
      if(col === "Action"){
        if(i === idx){
          html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
            <button class="action-btn small" onclick="saveEdit(${i})">Save</button>
            <button class="action-btn small" onclick="cancelEdit()">Cancel</button>
          </td>`;
        } else {
          html += `<td style="border:1px solid #ccc;padding:8px;white-space:nowrap;">
            <button class="action-btn small" onclick="editRow(${i})">Edit</button>
            <button class="action-btn small" onclick="deleteRow(${i})">Delete</button>
          </td>`;
        }
        return;
      }
      let cls = (col === "Description") ? "col-desc" : "";

      if(i === idx){
        // field editable hanya Month dan Reman
        if(col === "Month"){
          html += `<td class="${cls}" style="border:1px solid #ccc;padding:4px;">
            <select id="edit_month" >
              <option value="">--Pilih--</option>
              <option>Jan</option><option>Feb</option><option>Mar</option>
              <option>Apr</option><option>Mei</option><option>Jun</option>
              <option>Jul</option><option>Agu</option><option>Sep</option>
              <option>Okt</option><option>Nov</option><option>Des</option>
            </select>
          </td>`;
        } else if(col === "Reman"){
          html += `<td class="${cls}" style="border:1px solid #ccc;padding:4px;">
            <input id="edit_reman" type="text" value="${r.Reman||""}" />
          </td>`;
        } else {
          let v = r[col];
          if(typeof v === "number") v = v.toLocaleString();
          html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;">${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}" style="border:1px solid #ccc;padding:8px;">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });
    html += "</tr>";
  });

  html += "</tbody></table></div>";
  container.innerHTML = html;

  // set current values
  const sel = document.getElementById('edit_month');
  if(sel) sel.value = row.Month || "";
  const rin = document.getElementById('edit_reman');
  if(rin) rin.value = row.Reman || "";
}

function saveEdit(idx){
  const month = document.getElementById('edit_month').value || "";
  const reman = document.getElementById('edit_reman').value || "";
  if(idx<0 || idx>=merged.length) return;

  merged[idx].Month = month;
  merged[idx].Reman = reman;

  // simpan ke localStorage
  localStorage.setItem('ndarboe_merged', JSON.stringify(merged));

  renderTable(merged);
  document.getElementById('lmMsg').textContent = `Row ke-${idx+1} disimpan.`;
}

function cancelEdit(){
  renderTable(merged);
}

// Delete row
function deleteRow(idx){
  if(idx<0 || idx>=merged.length) return;
  if(!confirm(`Hapus baris ke-${idx+1}?`)) return;
  merged.splice(idx,1);
  localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
  renderTable(merged);
  document.getElementById('lmMsg').textContent = `Baris ke-${idx+1} dihapus.`;
}

// Load merged data dari localStorage saat mulai
function loadMergedData(){
  const s = localStorage.getItem('ndarboe_merged');
  if(s){
    try{
      merged = JSON.parse(s);
      if(!Array.isArray(merged)) merged = [];
    }catch(e){
      merged = [];
    }
  }
}
loadMergedData();
renderTable(merged);

// Inisialisasi halaman awal: tampilkan pageUpload
showPage('pageUpload');
document.getElementById('lmMsg').textContent = "Silakan upload file dan input order lalu Add Orders.";

// Styles tambahan (bisa juga di CSS)
const styleEl = document.createElement('style');
styleEl.textContent = `
  .col-desc { max-width:300px; white-space:normal; }
  button.action-btn.small { margin:2px; padding:4px 8px; font-size:0.8em; }
  table { font-family: Arial, sans-serif; font-size: 13px; }
  th, td { text-align: left; }
  /* Scroll horizontal pada container */
  #tableContainer > div { overflow-x:auto; }
`;
document.head.appendChild(styleEl);
