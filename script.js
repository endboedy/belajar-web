/* =========================
   script.js - full version
   - Upload IW39, SUM57, Planning, Budget, Data1, Data2
   - Build output table with lookups & calculations
   - Live freetext filters (Room, Order, CPH, MAT, Section)
   - Duplicate Order highlight (red / white)
   - Event delegation for Order click, Edit/Delete
   - Modal edit for Month & Reman (simple)
   - Save / Load to localStorage
   ========================= */

/* ========== GLOBAL DATA ========== */
let iwData = [];      // rows from IW39
let sum57 = [];       // rows from SUM57
let planning = [];    // rows from Planning
let budget = [];      // rows from Budget (unused here but stored)
let data1 = [];       // Data1 (for Section lookup)
let data2 = [];       // Data2 (for CPH lookup)
let merged = [];      // merged rows for Lembar Kerja

/* ========== UTIL HELPERS ========== */
function normalizeKey(k){
  return String(k||"").toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}
function normalizeRows(rows){
  return (rows||[]).map(r=>{
    const out = {};
    Object.keys(r||{}).forEach(k=>{
      out[normalizeKey(k)] = r[k];
    });
    return out;
  });
}
function asNumber(v){
  if(v===undefined || v===null || v==="") return 0;
  // strip non numeric except dot and minus
  const s = String(v).toString().replace(/[^0-9.\-]/g,'');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
// Format date from excel or JS Date to dd-Mmm-yyyy (Indo month short)
function formatDateExcel(dateInput){
  if(!dateInput && dateInput !== 0) return "";
  let d;
  if(typeof dateInput === "string"){
    d = new Date(dateInput);
    if(isNaN(d)){
      const n = Number(dateInput);
      if(!isNaN(n)) d = new Date(Date.UTC(1900,0,n-1));
      else return String(dateInput);
    }
  } else if(dateInput instanceof Date){
    d = dateInput;
  } else if(typeof dateInput === "number"){
    d = new Date(Date.UTC(1900,0,dateInput-1));
  } else {
    return String(dateInput);
  }
  const day = String(d.getUTCDate()).padStart(2,'0');
  const mnames = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const month = mnames[d.getUTCMonth()] || "";
  const year = d.getUTCFullYear();
  return `${day}-${month}-${year}`;
}

/* ========== FILE UPLOAD HANDLER ========== */
/*
  This function accepts multiple files at once. It will try to detect which file is which
  based on filename containing keywords: 'iw39', 'sum57', 'planning', 'budget', 'data1', 'data2'.
  If your filenames differ, tweak detection.
*/
function handleFilesFromInput(fileList){
  if(!fileList || fileList.length === 0) {
    showMsg("Tidak ada file dipilih.", "error");
    return;
  }
  const promises = [];
  for(const f of fileList){
    promises.push(readExcelToJson(f));
  }

  Promise.all(promises).then(results=>{
    // results: array of {name, sheets: {sheetName: rows...}}
    results.forEach(res=>{
      const lname = res.name.toLowerCase();
      // try to identify by filename
      if(lname.includes('iw39')) {
        // choose first sheet
        iwData = getFirstSheetRows(res);
      } else if(lname.includes('sum57')) {
        sum57 = getFirstSheetRows(res);
      } else if(lname.includes('planning')) {
        planning = getFirstSheetRows(res);
      } else if(lname.includes('budget')) {
        budget = getFirstSheetRows(res);
      } else if(lname.includes('data1')) {
        data1 = getFirstSheetRows(res);
      } else if(lname.includes('data2')) {
        data2 = getFirstSheetRows(res);
      } else {
        // unknown file: heuristics: check header names
        const first = getFirstSheetRows(res);
        const keys = Object.keys(first[0] || {}).map(k=>normalizeKey(k)).join(' ');
        if(keys.includes('order') && keys.includes('description')) iwData = first;
        else if(keys.includes('order') && keys.includes('statuspart')) sum57 = first;
        else if(keys.includes('eventstart') || keys.includes('statusamt')) planning = first;
        else if(keys.includes('room') && keys.includes('section')) data1 = first;
        else if(keys.includes('mat') && keys.includes('cph')) data2 = first;
        else {
          // unknown - ignore or keep in budget
          budget = first;
        }
      }
    });

    // Normalize: convert header keys to lowercase-keys map to ease lookup
    // We store both raw arrays and normalized forms inside functions when needed.

    showMsg("File berhasil dimuat. IW39 rows: " + (iwData.length || 0));
    // reset merged when new files loaded
    merged = [];
    renderTable(merged);
  }).catch(err=>{
    console.error(err);
    showMsg("Gagal membaca file: " + err.message, "error");
  });
}

// read single excel file to json per sheet
function readExcelToJson(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = e=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheets = {};
        wb.SheetNames.forEach(name=>{
          sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {defval:""});
        });
        resolve({name: file.name, sheets});
      }catch(er){
        reject(er);
      }
    };
    reader.onerror = er => reject(er);
    reader.readAsArrayBuffer(file);
  });
}
function getFirstSheetRows(res){ // res from readExcelToJson
  const firstName = Object.keys(res.sheets)[0];
  return res.sheets[firstName] || [];
}

/* ========== MESSAGE HUD ========== */
function showMsg(txt, type="info"){
  const el = document.getElementById('lmMsg');
  if(!el) {
    console.log(txt);
    return;
  }
  el.textContent = txt;
  el.style.color = (type==="error") ? "crimson" : "#333";
}

/* ========== BUILD MERGED ROW ========== */
/*
  Input: original iwRow (raw row object from IW39) OR an object with minimal keys.
  The function will use normalized lookup across other source arrays.
*/
function buildMergedRow(origRow, explicitOrder){
  // normalize origRow keys
  const origNormArr = normalizeRows([origRow]);
  const orig = origNormArr[0] || {};

  const getFrom = (names)=>{
    for(const n of names){
      if(n in orig && orig[n] !== "" && orig[n] !== null && orig[n] !== undefined) return orig[n];
    }
    return "";
  };

  // Basic fields from IW39
  const Room = getFrom(['room','location','area']) || "-";
  const OrderType = getFrom(['ordertype','type']) || "";
  const Order = explicitOrder ? String(explicitOrder) : (getFrom(['order','orderno','no','noorder']) || "");
  const DescriptionRaw = getFrom(['description','desc','keterangan','shortdesc']) || "";
  const CreatedOnRaw = getFrom(['createdon','created','date']) || "";
  const CreatedOn = formatDateExcel(CreatedOnRaw) || "";
  const UserStatus = getFrom(['userstatus','status']) || "";
  const MAT = getFrom(['mat','material','materialcode']) || "";

  // Normalize other data sources for lookup
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);
  const iwNorm = normalizeRows(iwData);

  // find d1 by room
  const d1rowByRoom = d1Norm.find(r => (r['room']||"").toString().trim().toUpperCase() === (Room||"").toString().trim().toUpperCase());
  // find d2 by MAT
  const d2rowByMAT = d2Norm.find(r => (r['mat']||"").toString().trim().toUpperCase() === (MAT||"").toString().trim().toUpperCase());
  // find sum57 by order
  const sumrow = sumNorm.find(r => (r['order']||"").toString().trim().toUpperCase() === (String(Order)||"").trim().toUpperCase());
  // planning row by order
  const planrow = planNorm.find(r => (r['order']||"").toString().trim().toUpperCase() === (String(Order)||"").trim().toUpperCase());
  // also attempt to find iw row by order to get TotalPlan / TotalActual
  const iwrow = iwNorm.find(r => (r['order']||"").toString().trim().toUpperCase() === (String(Order)||"").trim().toUpperCase());

  // CPH logic:
  // If Description starts with "JR" (2 letters = JR) then 'JR'
  let CPH = "";
  if(String(DescriptionRaw || "").toUpperCase().startsWith("JR")) {
    CPH = "JR";
  } else if(d2rowByMAT && (d2rowByMAT['cph'] || d2rowByMAT['costperhour'] || d2rowByMAT['cphvalue'])){
    CPH = d2rowByMAT['cph'] || d2rowByMAT['costperhour'] || d2rowByMAT['cphvalue'];
  } else {
    CPH = "";
  }

  // Status Part & Aging from SUM57
  const StatusPart = sumrow ? (sumrow['statuspart'] || sumrow['status'] || "") : "";
  const AgingRaw = sumrow ? (sumrow['aging'] || sumrow['age'] || "") : "";
  let Aging = "";
  if(typeof AgingRaw === "number") Aging = Math.round(AgingRaw).toString();
  else Aging = String(AgingRaw).split('.')[0] || "";

  // Planning & Status AMT
  const PlanningRaw = planrow ? (planrow['eventstart'] || planrow['planning'] || "") : "";
  const Planning = PlanningRaw ? formatDateExcel(PlanningRaw) : "";
  const StatusAMT = planrow ? (planrow['statusamt'] || "") : "";

  // Cost calculation: (IW39.TotalPlan - IW39.TotalActual)/16500, if < 0 => "-"
  let Cost = "-";
  if(iwrow){
    // try to find keys that might contain plan/actual
    const planKeys = Object.keys(iwrow).filter(k=>normalizeKey(k).includes('plan'));
    const actualKeys = Object.keys(iwrow).filter(k=>normalizeKey(k).includes('actual') || normalizeKey(k).includes('act'));
    const planVal = planKeys.length ? asNumber(iwrow[planKeys[0]]) : asNumber(iwrow['TotalPlan'] || iwrow['totalplan'] || 0);
    const actualVal = actualKeys.length ? asNumber(iwrow[actualKeys[0]]) : asNumber(iwrow['TotalActual'] || iwrow['totalactual'] || 0);
    const rawCost = (planVal - actualVal) / 16500;
    const finite = Number.isFinite(rawCost);
    if(finite && rawCost >= 0){
      // round to 1 decimal
      Cost = Number(rawCost.toFixed(1));
    } else {
      Cost = "-";
    }
  }

  // Reman default blank (manual input)
  const Reman = ""; // to be filled by user later
  // Include: if Reman contains "reman" (case-insensitive) => Cost*0.25 else Cost
  // We'll compute include/exclude at render time since Reman may be edited later.

  const Section = d1rowByRoom ? (d1rowByRoom['section'] || d1rowByRoom['dept'] || "") : "";
  const OrderTypeVal = OrderType || "";

  return {
    "Room": Room,
    "Order Type": OrderTypeVal,
    "Order": String(Order || "").split('.')[0],
    "Description": DescriptionRaw,
    "Created On": CreatedOn,
    "User Status": UserStatus,
    "MAT": MAT,
    "CPH": CPH,
    "Section": Section || "",
    "Status Part": StatusPart || "",
    "Aging": Aging || "",
    "Month": "",   // manual dropdown
    "Cost": Cost,
    "Reman": Reman,
    "Include": Cost === "-" ? "-" : Cost, // default include same as cost (if reman absent)
    "Exclude": OrderTypeVal.toUpperCase() === "PM38" ? "-" : (Cost === "-" ? "-" : Cost),
    "Planning": Planning,
    "Status AMT": StatusAMT || ""
  };
}

/* ========== ADD ORDERS FROM INPUT (textarea) ========== */
function addOrdersFromInput(){
  const raw = (document.getElementById('inputOrders') && document.getElementById('inputOrders').value) || "";
  if(raw.trim() === ""){
    showMsg("Tidak ada order yang dimasukkan.");
    return;
  }
  showMsg("");
  const parts = raw.split(/[\n,;]+/).map(s=>s.trim()).filter(s=>s!=="");
  const iwNorm = normalizeRows(iwData);
  for(const p of parts){
    const key = String(p).split('.')[0].trim().toUpperCase();
    // find iw row by order
    const found = iwNorm.find(r => (String(r['order']||"").trim().toUpperCase() === key));
    if(found){
      const mr = buildMergedRow(found, key);
      merged.push(mr);
    } else {
      // push minimal row with order only
      merged.push({
        "Room":"-",
        "Order Type":"-",
        "Order":key,
        "Description":"-",
        "Created On":"",
        "User Status":"",
        "MAT":"",
        "CPH":"",
        "Section":"",
        "Status Part":"",
        "Aging":"",
        "Month":"",
        "Cost":"-",
        "Reman":"",
        "Include":"-",
        "Exclude":"-",
        "Planning":"",
        "Status AMT":""
      });
    }
  }
  // After adding, render
  renderTable(merged);
  // clear input
  const inp = document.getElementById('inputOrders');
  if(inp) inp.value = "";
}

/* ========== RENDER TABLE ========== */
function findDuplicateOrdersCount(rows){
  const counts = {};
  rows.forEach(r=>{
    const k = (r.Order || "").toString().trim().toUpperCase();
    counts[k] = (counts[k]||0) + 1;
  });
  return counts;
}

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!container){
    console.error("Element #tableContainer not found in DOM.");
    return;
  }
  container.innerHTML = "";

  if(!data || data.length === 0){
    container.textContent = "Data Lembar Kerja kosong.";
    return;
  }

  const columns = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

  // FILTER BAR (create inputs)
  const filterDiv = document.createElement('div');
  filterDiv.style.display = "flex";
  filterDiv.style.gap = "8px";
  filterDiv.style.margin = "8px 0";
  filterDiv.style.flexWrap = "wrap";

  // create filters for: Room, Order, CPH, MAT, Section
  const makeFilter = (placeholder, key) => {
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = placeholder;
    inp.dataset.col = key;
    inp.style.padding = "6px 8px";
    inp.style.border = "1px solid #ccc";
    inp.style.borderRadius = "4px";
    inp.addEventListener('input', ()=> {
      applyExternalFilters();
    });
    filterDiv.appendChild(inp);
    return inp;
  };
  const fRoom = makeFilter('Filter Room', 'Room');
  const fOrder = makeFilter('Filter Order', 'Order');
  const fCPH = makeFilter('Filter CPH', 'CPH');
  const fMAT = makeFilter('Filter MAT', 'MAT');
  const fSection = makeFilter('Filter Section', 'Section');

  container.appendChild(filterDiv);

  // Table creation
  const table = document.createElement('table');
  table.style.width = "100%";
  table.style.borderCollapse = "collapse";
  table.className = "data-table";

  // thead with sticky header (CSS should have position:sticky for th)
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  columns.forEach(c=>{
    const th = document.createElement('th');
    th.textContent = c;
    th.style.padding = "8px";
    th.style.border = "1px solid #ddd";
    th.style.backgroundColor = "#cce5ff";
    th.style.position = "sticky";
    th.style.top = "0";
    th.style.zIndex = "3";
    trh.appendChild(th);
  });
  // action column
  const thAct = document.createElement('th');
  thAct.textContent = "Action";
  thAct.style.padding = "8px";
  thAct.style.border = "1px solid #ddd";
  thAct.style.backgroundColor = "#cce5ff";
  trh.appendChild(thAct);
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  // duplicate detection
  const dupCounts = findDuplicateOrdersCount(data);

  data.forEach((row, idx)=>{
    const tr = document.createElement('tr');
    const orderKey = (row["Order"] || "").toString().trim().toUpperCase();
    const isDup = dupCounts[orderKey] > 1;

    if(isDup){
      tr.classList.add('duplicate-order'); // CSS: red bg + white text
      // fallback inline style if CSS not set
      tr.style.backgroundColor = "#d9534f";
      tr.style.color = "white";
    }

    columns.forEach(col=>{
      const td = document.createElement('td');
      td.style.border = "1px solid #ddd";
      td.style.padding = "6px";
      td.style.whiteSpace = "nowrap";
      td.style.overflow = "hidden";
      td.style.textOverflow = "ellipsis";

      let val = row[col];

      // If column depends on lookup or calculation (ensure reflect latest values)
      if(col === "Description"){
        // prefer existing value else try lookup from iwData
        if(!val || val === "-" || val === "") {
          // lookup in iwData
          const iwNorm = normalizeRows(iwData);
          const found = iwNorm.find(r => (String(r['order']||"").trim().toUpperCase() === orderKey));
          val = found ? (found['description'] || found['desc'] || found['keterangan'] || "-") : "-";
        }
      }
      if(col === "CPH"){
        // if description starts with JR => JR, else lookup data2 by MAT
        if(String(row["Description"]||"").toUpperCase().startsWith("JR")) {
          val = "JR";
        } else {
          if(!val || val === "-"){
            const d2 = normalizeRows(data2).find(r => (String(r['mat']||"").trim().toUpperCase() === (String(row['MAT']||"").trim().toUpperCase())));
            if(d2) val = d2['cph'] || d2['costperhour'] || d2['cphvalue'] || val;
          }
        }
      }
      if(col === "Section"){
        if(!val || val === "-"){
          const d1 = normalizeRows(data1).find(r => (String(r['room']||"").trim().toUpperCase() === (String(row['Room']||"").trim().toUpperCase())));
          val = d1 ? (d1['section'] || d1['dept'] || "") : val;
        }
      }
      if(col === "Status Part" || col === "Aging"){
        if(!val || val === "-"){
          const s = normalizeRows(sum57).find(r => (String(r['order']||"").trim().toUpperCase() === orderKey));
          if(s){
            val = (col === "Status Part") ? (s['statuspart'] || s['status'] || "") : (s['aging'] || s['age'] || "");
          }
        }
      }
      if(col === "Planning" || col === "Status AMT"){
        if(!val || val === "-"){
          const p = normalizeRows(planning).find(r => (String(r['order']||"").trim().toUpperCase() === orderKey));
          if(p){
            if(col === "Planning"){
              val = p['eventstart'] ? formatDateExcel(p['eventstart']) : (p['planning'] ? formatDateExcel(p['planning']) : "");
            } else {
              val = p['statusamt'] || "";
            }
          }
        }
      }
      if(col === "Cost"){
        // if cost is '-' keep '-', else ensure numeric formatted with 1 decimal if number
        if(typeof val === "number") {
          val = Number(val.toFixed(1));
        } else if(val === "-" || val === "" || val === null) {
          // try compute from iwData row
          const iwNorm = normalizeRows(iwData);
          const iwrow = iwNorm.find(r => (String(r['order']||"").trim().toUpperCase() === orderKey));
          if(iwrow){
            const planKeys = Object.keys(iwrow).filter(k=>normalizeKey(k).includes('plan'));
            const actualKeys = Object.keys(iwrow).filter(k=>normalizeKey(k).includes('actual') || normalizeKey(k).includes('act'));
            const planVal = planKeys.length ? asNumber(iwrow[planKeys[0]]) : asNumber(iwrow['totalplan']||iwrow['totalplan']||0);
            const actualVal = actualKeys.length ? asNumber(iwrow[actualKeys[0]]) : asNumber(iwrow['totalactual']||iwrow['totalactual']||0);
            const rawCost = (planVal - actualVal) / 16500;
            if(Number.isFinite(rawCost) && rawCost >= 0){
              val = Number(rawCost.toFixed(1));
            } else {
              val = "-";
            }
          } else {
            val = "-";
          }
        }
      }
      if(col === "Include"){
        const rem = String(row["Reman"]||"").toLowerCase();
        const currentCost = (row["Cost"] === "-" ? "-" : row["Cost"]);
        if(currentCost === "-" || currentCost === "" || currentCost === null){
          val = "-";
        } else {
          const cnum = Number(currentCost);
          if(rem.includes("reman")){
            val = Number((cnum * 0.25).toFixed(1));
          } else {
            val = Number(cnum.toFixed(1));
          }
        }
      }
      if(col === "Exclude"){
        const orderType = String(row["Order Type"]||"").toUpperCase();
        if(orderType === "PM38") val = "-";
        else {
          // same as Include
          const inc = (row["Include"] === "-" ? "-" : row["Include"]);
          val = inc;
        }
      }

      // format numeric right align
      if(["Cost","Include","Exclude"].includes(col) && val !== "-" && val !== ""){
        td.style.textAlign = "right";
      }

      td.textContent = (val === undefined || val === null) ? "" : String(val);
      tr.appendChild(td);
    });

    // Action buttons cell
    const tdAct = document.createElement('td');
    tdAct.style.border = "1px solid #ddd";
    tdAct.style.padding = "6px";
    tdAct.style.whiteSpace = "nowrap";

    const btnEdit = document.createElement('button');
    btnEdit.textContent = "Edit";
    btnEdit.className = "action-btn small btn-edit";
    btnEdit.dataset.idx = idx;
    btnEdit.style.marginRight = "6px";

    const btnDelete = document.createElement('button');
    btnDelete.textContent = "Delete";
    btnDelete.className = "action-btn small btn-delete";
    btnDelete.dataset.idx = idx;
    btnDelete.style.backgroundColor = "#e74c3c";

    tdAct.appendChild(btnEdit);
    tdAct.appendChild(btnDelete);

    tr.appendChild(tdAct);
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  container.appendChild(table);
}

/* ========== APPLY EXTERNAL FILTERS (from created inputs above) ========== */
function applyExternalFilters(){
  // read filter inputs (if present)
  const fRoom = document.querySelector('input[placeholder="Filter Room"]');
  const fOrder = document.querySelector('input[placeholder="Filter Order"]');
  const fCPH = document.querySelector('input[placeholder="Filter CPH"]');
  const fMAT = document.querySelector('input[placeholder="Filter MAT"]');
  const fSection = document.querySelector('input[placeholder="Filter Section"]');

  const roomVal = fRoom ? fRoom.value.trim().toLowerCase() : "";
  const orderVal = fOrder ? fOrder.value.trim().toLowerCase() : "";
  const cphVal = fCPH ? fCPH.value.trim().toLowerCase() : "";
  const matVal = fMAT ? fMAT.value.trim().toLowerCase() : "";
  const secVal = fSection ? fSection.value.trim().toLowerCase() : "";

  // filter merged source - if merged empty, attempt to build from iwData? We'll filter merged (the Lembar Kerja)
  const source = merged.length ? merged : []; // if no merged, show nothing
  const filtered = source.filter(r => {
    return (!roomVal || (String(r.Room||r["Room"]||"").toLowerCase().includes(roomVal)))
        && (!orderVal || (String(r.Order||"").toLowerCase().includes(orderVal)))
        && (!cphVal || (String(r.CPH||"").toLowerCase().includes(cphVal)))
        && (!matVal || (String(r.MAT||"").toLowerCase().includes(matVal)))
        && (!secVal || (String(r.Section||"").toLowerCase().includes(secVal)));
  });

  renderTable(filtered);
}

/* ========== EVENT DELEGATION (Edit/Delete/Order Click) ========== */
document.addEventListener('click', function(e){
  // Order button (if created inside table as <button data-order=...>)
  if(e.target.matches('.btn-order')){
    const order = e.target.dataset.order;
    alert(`Order ${order} clicked (you can open detail panel here).`);
    return;
  }
  // Edit
  if(e.target.matches('.btn-edit')){
    const idx = Number(e.target.dataset.idx);
    openEditModal(idx);
    return;
  }
  // Delete
  if(e.target.matches('.btn-delete')){
    const idx = Number(e.target.dataset.idx);
    if(!isNaN(idx) && merged[idx]){
      if(confirm(`Hapus order ${merged[idx].Order}?`)){
        merged.splice(idx,1);
        renderTable(merged);
      }
    }
    return;
  }
});

/* ========== MODAL EDIT (Month & Reman) ========== */
function openEditModal(idx){
  const modal = document.getElementById('editModal');
  if(!modal){ alert("Modal edit tidak ditemukan."); return; }
  const row = merged[idx];
  if(!row) return;
  document.getElementById('editIdx').value = idx;
  document.getElementById('editMonth').value = row['Month'] || '';
  document.getElementById('editReman').value = row['Reman'] || '';
  modal.style.display = 'block';
}
function closeEditModal(){
  const modal = document.getElementById('editModal');
  if(modal) modal.style.display = 'none';
}
function saveEditModal(){
  const idx = Number(document.getElementById('editIdx').value);
  if(isNaN(idx) || !merged[idx]) return;
  merged[idx]['Month'] = document.getElementById('editMonth').value;
  merged[idx]['Reman'] = document.getElementById('editReman').value;
  // recalc include/exclude
  const cost = merged[idx]['Cost'];
  if(cost === "-" || cost === "" || cost === null){
    merged[idx]['Include'] = "-";
    merged[idx]['Exclude'] = "-";
  } else {
    const cnum = Number(cost);
    if(String(merged[idx]['Reman']||"").toLowerCase().includes('reman')){
      merged[idx]['Include'] = Number((cnum * 0.25).toFixed(1));
    } else {
      merged[idx]['Include'] = Number(cnum.toFixed(1));
    }
    merged[idx]['Exclude'] = (String(merged[idx]['Order Type']||"").toUpperCase()==="PM38") ? "-" : merged[idx]['Include'];
  }
  renderTable(merged);
  closeEditModal();
}

/* ========== SAVE / LOAD localStorage ========== */
function saveLembarKerja(){
  try{
    localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
    showMsg("Lembar Kerja tersimpan di localStorage.");
  }catch(err){
    showMsg("Gagal menyimpan: " + err.message, "error");
  }
}
function loadSaved(){
  const raw = localStorage.getItem('ndarboe_merged');
  if(raw){
    try{
      merged = JSON.parse(raw);
      renderTable(merged);
      showMsg("Lembar Kerja dimuat dari localStorage.");
    }catch(e){
      console.error(e);
      showMsg("Gagal memuat Lembar Kerja dari localStorage.", "error");
    }
  } else {
    showMsg("Tidak ada data tersimpan di localStorage.");
  }
}

/* ========== INIT: wire up DOM controls if exist ========== */
document.addEventListener('DOMContentLoaded', ()=>{
  // file input (#fileInput) - supports multiple files
  const fileInput = document.getElementById('fileInput');
  if(fileInput){
    fileInput.addEventListener('change', (ev)=>{
      handleFilesFromInput(ev.target.files);
    });
  }

  // btnAddOrders
  const btnAdd = document.getElementById('btnAddOrders');
  if(btnAdd){
    btnAdd.addEventListener('click', addOrdersFromInput);
  }

  // btnSaveLembar / Load
  const btnSave = document.getElementById('btnSaveLembar');
  if(btnSave) btnSave.addEventListener('click', saveLembarKerja);
  const btnLoad = document.getElementById('btnLoadLembar');
  if(btnLoad) btnLoad.addEventListener('click', loadSaved);

  // Modal buttons
  const btnClose = document.getElementById('btnCloseEdit');
  if(btnClose) btnClose.addEventListener('click', closeEditModal);
  const btnSaveEdit = document.getElementById('btnSaveEdit');
  if(btnSaveEdit) btnSaveEdit.addEventListener('click', saveEditModal);

  // quick attempt to apply filters if inputs exist on load
  applyExternalFilters();
});
