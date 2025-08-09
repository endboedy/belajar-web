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

// Format date to dd-mmm-yyyy
function formatDateString(dateInput){
  if(!dateInput) return null;
  const d = new Date(dateInput);
  if(isNaN(d)) return null;
  const day = d.getDate().toString().padStart(2,'0');
  const monthNames = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const month = monthNames[d.getMonth()] || "???";
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

// Build merged row with lookup and formatting
function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
  const get = (o, names) => {
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]).trim() !== "") return o[n];
    }
    return null;
  };

  // IW fields
  const RoomVal = get(origNorm, ['room','location','area']);
  const OrderTypeVal = get(origNorm, ['ordertype','type']);
  const OrderValRaw = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  let DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatusVal = get(origNorm, ['userstatus','status']);
  const MATVal = get(origNorm, ['mat','material','materialcode']);

  // Lookup rows
  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MATVal).trim())) || null;
  const d2row = (d2Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MATVal).trim())) || null;
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(OrderValRaw).trim())) || null;
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(OrderValRaw).trim())) || null;

  // Description rule
  let DescriptionVal = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) DescriptionVal = "JR";
  else DescriptionVal = get(d1row, ['description','desc','keterangan','shortdesc']) || "Tidak Ada";

  // CPH rule
  let CPHVal = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPHVal = "JR";
  else CPHVal = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "Tidak Ada";

  const SectionVal = get(d1row, ['section','dept','deptcode','department']) || "Tidak Ada";
  const StatusPartVal = get(sumrow, ['statuspart','status','status_part']) || "Tidak Ada";

  let AgingVal = get(sumrow, ['aging','age']);
  if(typeof AgingVal === "number") AgingVal = Math.floor(AgingVal);
  else if(!isNaN(Number(AgingVal))) AgingVal = Math.floor(Number(AgingVal));
  else AgingVal = "Tidak Ada";

  const PlanningRawVal = get(planrow, ['planning','eventstart','description']);
  const StatusAMTVal = get(planrow, ['statusamt','status']) || "Tidak Ada";

  // Format tanggal
  const CreatedOnVal = formatDateString(CreatedOnRaw) || "Tidak Ada";
  const PlanningVal = formatDateString(PlanningRawVal) || "Tidak Ada";

  // Cost calculation
  const planKey = Object.keys(origNorm||{}).find(k => k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k => k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let CostVal = "-";
  const rawCost = (planVal - actualVal) / 16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) CostVal = Number(rawCost.toFixed(2));

  // Include & Exclude
  let IncludeVal = "-";
  if(typeof CostVal === "number") IncludeVal = (String(remanVal).toLowerCase().includes("reman") ? Number((CostVal * 0.25).toFixed(2)) : Number(CostVal.toFixed(2)));
  let ExcludeVal = (String(OrderTypeVal).toUpperCase() === "PM38") ? "-" : IncludeVal;

  // Order tanpa desimal (as string)
  const OrderVal = String(OrderValRaw).split(".")[0];

  // Room, Order Type, User Status tampil "OK" kalau ada, "Tidak Ada" kalau kosong
  const RoomDisp = RoomVal ? "OK" : "Tidak Ada";
  const OrderTypeDisp = OrderTypeVal ? "OK" : "Tidak Ada";
  const UserStatusDisp = UserStatusVal ? "OK" : "Tidak Ada";

  return {
    "Room": RoomDisp,
    "Order Type": OrderTypeDisp,
    "Order": OrderVal || "Tidak Ada",
    "Description": DescriptionVal,
    "Created On": CreatedOnVal,
    "User Status": UserStatusDisp,
    "MAT": MATVal || "Tidak Ada",
    "CPH": CPHVal,
    "Section": SectionVal,
    "Status Part": StatusPartVal,
    "Aging": AgingVal,
    "Month": monthVal || "",
    "Cost": CostVal,
    "Reman": remanVal || "",
    "Include": IncludeVal,
    "Exclude": ExcludeVal,
    "Planning": PlanningVal,
    "Status AMT": StatusAMTVal
  };
}

// Render table with action column (edit/delete)
const DISPLAY = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }

  let html = "<div style='overflow-x:auto;'><table style='width:100%;table-layout:auto; border-collapse: collapse;' border='1'>";
  html += "<thead><tr>";
  DISPLAY.forEach(col => {
    html += `<th style="min-width:100px;padding:6px 10px;text-align:center;background:#eee;">${col}</th>`;
  });
  html += "<th style='min-width:100px;padding:6px 10px;text-align:center;background:#eee;'>Action</th>";
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    DISPLAY.forEach(col => {
      let cls = (col === "Description") ? "col-desc" : "";
      let v = row[col];

      // Format number columns with 2 decimals if number (Cost, Include, Exclude)
      if(["Cost","Include","Exclude"].includes(col) && typeof v === "number") {
        v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      }

      // Format Aging & Order without decimals
      if(col === "Aging") {
        if(typeof v === "number") v = Math.floor(v);
      }

      if(col === "Order"){
        // pastikan string dan tidak ada desimal
        v = String(v).split(".")[0];
      }

      html += `<td class="${cls}" style="padding:6px 10px;white-space:nowrap;vertical-align:middle;text-align:center">${v !== undefined && v !== null ? v : ""}</td>`;
    });

    // Action buttons column
    html += `<td style="padding:6px 10px;text-align:center;white-space:nowrap;">
      <button class="action-btn small" onclick="editRow(${idx})">Edit</button>
      <button class="action-btn small" onclick="deleteRow(${idx})">Delete</button>
    </td>`;

    html += "</tr>";
  });

  html += "</tbody></table></div>";
  container.innerHTML = html;
}

// Edit row: convert row to inline editable (Month & Reman editable)
function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  // build table but replace the edited row with inputs
  let html = "<div style='overflow-x:auto;'><table style='width:100%;table-layout:auto; border-collapse: collapse;' border='1'><thead><tr>";
  DISPLAY.forEach(col => html += `<th style="min-width:100px;padding:6px 10px;text-align:center;background:#eee;">${col}</th>`);
  html += "<th style='min-width:100px;padding:6px 10px;text-align:center;background:#eee;'>Action</th>";
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    DISPLAY.forEach(col=>{
      let cls = (col==="Description") ? "col-desc" : "";
      if(i === idx){
        if(col === "Month"){
          html += `<td class="${cls}" style="padding:6px 10px;white-space:nowrap;vertical-align:middle;text-align:center">
            <select id="edit_month" >
              <option value="">--Pilih--</option>
              <option>Jan</option><option>Feb</option><option>Mar</option>
              <option>Apr</option><option>Mei</option><option>Jun</option>
              <option>Jul</option><option>Agu</option><option>Sep</option>
              <option>Okt</option><option>Nov</option><option>Des</option>
            </select>
          </td>`;
        } else if(col === "Reman"){
          html += `<td class="${cls}" style="padding:6px 10px;white-space:nowrap;vertical-align:middle;text-align:center"><input id="edit_reman" type="text" value="${r.Reman||""}" /></td>`;
        } else {
          let v = r[col];
          if(typeof v === "number") v = v.toLocaleString();
          html += `<td class="${cls}" style="padding:6px 10px;white-space:nowrap;vertical-align:middle;text-align:center">${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td class="${cls}" style="padding:6px 10px;white-space:nowrap;vertical-align:middle;text-align:center">${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

    // action column
    if(i === idx){
      html += `<td style="padding:6px 10px;text-align:center;white-space:nowrap;">
        <button class="action-btn small" onclick="saveEdit(${i})">Save</button>
        <button class="action-btn small" onclick="cancelEdit()">Cancel</button>
      </td>`;
    } else {
      html += `<td style="padding:6px 10px;text-align:center;white-space:nowrap;">
        <button class="action-btn small" onclick="editRow(${i})">Edit</button>
        <button class="action-btn small" onclick="deleteRow(${i})">Delete</button>
      </td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";
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

  iwData = normalizeRows(iwData);
  data1 = normalizeRows(data1);
  data2 = normalizeRows(data2);
  sum57 = normalizeRows(sum57);
  planning = normalizeRows(planning);
  budget = normalizeRows(budget);

  document.getElementById('loadMsg').textContent = "Files loaded. Ready to add orders.";
});

// Tombol Add Orders (multi orders dipisah koma atau newline)
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim()===""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  if(iwData.length === 0){
    document.getElementById('lmMsg').textContent = "Data IW39 belum dimuat. Load dulu file IW39.";
    return;
  }
  const ordersRaw = raw.split(/[\n,;]/).map(s=>s.trim()).filter(s=>s.length>0);
  const ordersSet = new Set(ordersRaw);

  const d1Norm = data1;
  const d2Norm = data2;
  const sumNorm = sum57;
  const planNorm = planning;

  ordersSet.forEach(o=>{
    // cari data IW39 order o
    const found = iwData.find(r=>{
      const ordr = r.order || r.orderno || r.no || r.noorder || "";
      return String(ordr).trim() === o;
    });
    if(!found){
      // tambahkan baris dengan minimal data
      merged.push({
        "Room":"Tidak Ada",
        "Order Type":"Tidak Ada",
        "Order": o,
        "Description":"Tidak Ada",
        "Created On":"Tidak Ada",
        "User Status":"Tidak Ada",
        "MAT":"Tidak Ada",
        "CPH":"Tidak Ada",
        "Section":"Tidak Ada",
        "Status Part":"Tidak Ada",
        "Aging":"Tidak Ada",
        "Month":"",
        "Cost":"-",
        "Reman":"",
        "Include":"-",
        "Exclude":"-",
        "Planning":"Tidak Ada",
        "Status AMT":"Tidak Ada"
      });
    } else {
      const normFound = {};
      Object.keys(found).forEach(k=>{
        normFound[normalizeKey(k)] = found[k];
      });

      // default Month & Reman kosong, bisa diedit nanti
      const rowBuilt = buildMergedRow(normFound, o, "", "", d1Norm, d2Norm, sumNorm, planNorm);
      merged.push(rowBuilt);
    }
  });

  renderTable(merged);
  document.getElementById('lmMsg').textContent = "Order(s) ditambahkan (sementara). Jangan lupa klik Save Lembar Kerja.";
});

// Clear all merged data
document.getElementById('btnClear').addEventListener('click', ()=>{
  if(!confirm("Hapus semua data di Lembar Kerja?")) return;
  merged = [];
  renderTable(merged);
  document.getElementById('lmMsg').textContent = "Lembar Kerja dibersihkan.";
});
