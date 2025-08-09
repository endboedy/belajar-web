// script.js
// Handles: reading excel files, lookup logic, add multi orders, save/load localStorage, edit/delete, download

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

// Build merged row from normalized iw row (origNormalized), orderKey, month, reman
function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
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
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  // Format Created On "dd-mmm-yyyy"
  let CreatedOn = "";
  if(CreatedOnRaw){
    const d = new Date(CreatedOnRaw);
    if(!isNaN(d)) {
      CreatedOn = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    } else {
      CreatedOn = CreatedOnRaw;
    }
  }

  // lookups: Data1 (section), Data2 (cph), SUM57 (status part & aging), Planning (planning,status amt)
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
  
  // Format Aging no decimal
  let AgingRaw = get(sumrow, ['aging','age']);
  let Aging = "";
  if(AgingRaw !== "" && AgingRaw !== undefined && AgingRaw !== null){
    Aging = parseInt(Number(AgingRaw));
  }

  // Format Planning "dd-mmm-yyyy"
  let PlanningRaw = get(planrow, ['planning','eventstart','description']);
  let Planning = "";
  if(PlanningRaw){
    const d = new Date(PlanningRaw);
    if(!isNaN(d)) {
      Planning = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    } else {
      Planning = PlanningRaw;
    }
  }
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
  if(typeof Cost === "number"){
    Include = (String(remanVal).toLowerCase().includes("reman") ? Number((Cost * 0.25).toFixed(2)) : Number(Cost.toFixed(2)));
  }
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

  // Maksimalkan lebar container ke 100vw
  container.style.overflowX = "auto";
  container.style.width = "100vw";

  let html = "<table style='min-width: 1200px; border-collapse: collapse;'>";
  html += "<thead><tr>";
  DISPLAY.forEach(col => {
    html += `<th style="border:1px solid #ccc; padding:6px 8px; text-align:left;">${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col => {
      let cls = (col==="Description") ? "col-desc" : "";
      let v = row[col];
      if(col === "Order"){
        // hapus tanda titik dan angka dibelakang koma
        if(typeof v === "string") v = v.replace(/\./g, "").split(",")[0];
      }
      if(col === "Created On" || col === "Planning"){
        // tanggal sudah diformat di buildMergedRow
      }
      if(col === "Aging"){
        // hapus angka dibelakang koma sudah di buildMergedRow
      }
      if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      html += `<td class="${cls}" style="border:1px solid #ccc; padding:6px 8px;">${v !== undefined && v !== null ? v : ""}</td>`;
    });
    // Action buttons
    html += `<td style="border:1px solid #ccc; padding:6px 8px;">
      <button class="action-btn small" onclick="editRow(${idx})">Edit</button>
      <button class="action-btn small" onclick="deleteRow(${idx})">Delete</button>
    </td>`;
    html += "</tr>";
  });

  html += "</tbody></table>";
  container.innerHTML = html;
}

function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  let html = "<table style='min-width: 1200px; border-collapse: collapse;'><thead><tr>";
  DISPLAY.forEach(col => html += `<th style="border:1px solid #ccc; padding:6px 8px; text-align:left;">${col}</th>`);
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col=>{
      let cls = (col==="Description") ? "col-desc" : "";
      if(i === idx){
        if(col === "Month"){
          html += `<td class="${cls}" style="border:1px solid #ccc; padding:6px 8px;">
            <select id="edit_month" >
              <option value="">--Pilih--</option>
              <option>Jan</option><option>Feb</option><option>Mar</option>
              <option>Apr</option><option>Mei</option><option>Jun</option>
              <option>Jul</option><option>Agu</option><option>Sep</option>
              <option>Okt</option><option>Nov</option><option>Des</option>
            </select>
          </td>`;
        } else if(col === "Reman"){
          html += `<td class="${cls}" style="border:1px solid #ccc; padding:6px 8px;"><input id="edit_reman" type="text" value="${r.Reman||""}" /></td>`;
        } else {
          let v = r[col];
          if(typeof v === "number") v = v.toLocaleString();
          html += `<td class="${cls}" style="border:1px solid #ccc; padding:6px 8px;">${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}" style="border:1px solid #ccc; padding:6px 8px;">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

    if(i === idx){
      html += `<td style="border:1px solid #ccc; padding:6px 8px;">
        <button class="action-btn small" onclick="saveEdit(${i})">Save</button>
        <button class="action-btn small" onclick="cancelEdit()">Cancel</button>
      </td>`;
    } else {
      html += `<td style="border:1px solid #ccc; padding:6px 8px;">
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
  const lmMsg = document.getElementById('lmMsg');
  if(lmMsg) lmMsg.textContent = "Perubahan disimpan (sementara). Jangan lupa klik Save Lembar Kerja untuk persist.";
}

function cancelEdit(){
  renderTable(merged);
}

function deleteRow(idx){
  if(!confirm("Hapus baris ini?")) return;
  merged.splice(idx,1);
  renderTable(merged);
  const lmMsg = document.getElementById('lmMsg');
  if(lmMsg) lmMsg.textContent = "Baris dihapus (sementara). Klik Save Lembar Kerja untuk persist.";
}

document.addEventListener('DOMContentLoaded', () => {
  // Pasang event click di menu utama (cek dulu elemen agar tidak error)
  const navUpload = document.getElementById('navUpload');
  if(navUpload) navUpload.addEventListener('click', () => showPage('pageUpload'));

  const navLembar = document.getElementById('navLembar');
  if(navLembar) navLembar.addEventListener('click', () => showPage('pageLembar'));

  const navSummary = document.getElementById('navSummary');
  if(navSummary) navSummary.addEventListener('click', () => showPage('pageSummary'));

  const navDownload = document.getElementById('navDownload');
  if(navDownload) navDownload.addEventListener('click', () => showPage('pageDownload'));

  const btnLoad = document.getElementById('btnLoad');
  if(btnLoad) btnLoad.addEventListener('click', async () => {
    const loadMsg = document.getElementById('loadMsg');
    if(loadMsg) loadMsg.textContent = "Loading files...";

    const fIW = document.getElementById('fileIW39')?.files[0];
    const fSUM = document.getElementById('fileSUM57')?.files[0];
    const fPlan = document.getElementById('filePlanning')?.files[0];
    const fBud = document.getElementById('fileBudget')?.files[0];
    const fD1 = document.getElementById('fileData1')?.files[0];
    const fD2 = document.getElementById('fileData2')?.files[0];

    iwData = fIW ? await readExcelFile(fIW) : [];
    sum57 = fSUM ? await readExcelFile(fSUM) : [];
    planning = fPlan ? await readExcelFile(fPlan) : [];
    budget = fBud ? await readExcelFile(fBud) : [];
    data1 = fD1 ? await readExcelFile(fD1) : [];
    data2 = fD2 ? await readExcelFile(fD2) : [];

    const iwNorm = normalizeRows(iwData);
    const d1Norm = normalizeRows(data1);
    const d2Norm = normalizeRows(data2);
    const sumNorm = normalizeRows(sum57);
    const planNorm = normalizeRows(planning);

    window.__nd = {iw:iwNorm,d1:d1Norm,d2:d2Norm,sum:sumNorm,plan:planNorm};

    if(loadMsg) loadMsg.textContent = `Selesai load. IW39 baris: ${iwData.length}`;
  });

  const btnAddOrders = document.getElementById('btnAddOrders');
  if(btnAddOrders) btnAddOrders.addEventListener('click', ()=>{
    const raw = document.getElementById('inputOrders')?.value || "";
    if(raw.trim()===""){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Tidak ada order yang dimasukkan.";
      return;
    }
    if(!iwData || iwData.length===0){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "File IW39 belum di-load.";
      return;
    }
    const orders = raw.split(/[\r\n]+/).map(s=>s.trim()).filter(s=>s.length>0);

    const iwNorm = normalizeRows(iwData);
    const d1Norm = normalizeRows(data1);
    const d2Norm = normalizeRows(data2);
    const sumNorm = normalizeRows(sum57);
    const planNorm = normalizeRows(planning);

    orders.forEach(orderKey => {
      const row = iwNorm.find(r => {
        const v = r['order'] || r['orderno'] || r['no'] || "";
        return String(v).trim() === orderKey.trim();
      });
      if(row){
        // default Month & Reman empty at first
        const m = "";
        const rmn = "";
        const newRow = buildMergedRow(row, orderKey, m, rmn, d1Norm, d2Norm, sumNorm, planNorm);
        merged.push(newRow);
      }
    });
    renderTable(merged);
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = `${orders.length} order berhasil ditambahkan.`;
  });

  const btnSave = document.getElementById('btnSave');
  if(btnSave) btnSave.addEventListener('click', ()=>{
    localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "Lembar Kerja tersimpan di localStorage.";
  });

  function loadSaved(){
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

  const btnClear = document.getElementById('btnClear');
  if(btnClear) btnClear.addEventListener('click', ()=>{
    if(!confirm("Yakin ingin hapus semua Lembar Kerja?")) return;
    merged = [];
    renderTable(merged);
    localStorage.removeItem('ndarboe_merged');
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "Lembar Kerja dihapus.";
  });

  const inputOrders = document.getElementById('inputOrders');
  if(inputOrders) inputOrders.addEventListener('focus', ()=>{
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "";
  });

  // Download JSON as Excel file
  const btnDownload = document.getElementById('btnDownload');
  if(btnDownload) btnDownload.addEventListener('click', ()=>{
    if(!merged || merged.length===0){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Tidak ada data untuk diunduh.";
      return;
    }
    const ws = XLSX.utils.json_to_sheet(merged);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "LembarKerja");
    XLSX.writeFile(wb, "LembarKerja.xlsx");
  });
});

// Show / Hide pages
function showPage(pageId){
  document.querySelectorAll('.page').forEach(p => {
    if(p.id === pageId) p.style.display = "block";
    else p.style.display = "none";
  });
}

// Expose edit and delete globally so buttons in table can access
window.editRow = editRow;
window.deleteRow = deleteRow;
window.saveEdit = saveEdit;
window.cancelEdit = cancelEdit;
