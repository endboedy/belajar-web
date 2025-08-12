// script.js - Ndarboe.net (FINAL for client-side merging & Lembar Kerja)
// Dependencies: XLSX (already included in index.html)

// ---------------------- Global stores ----------------------
let iw39Data = [];      // rows from IW39
let sum57Data = [];     // rows from SUM57
let planningData = [];  // rows from Planning
let data1Data = [];     // rows from Data1
let data2Data = [];     // rows from Data2
let budgetData = [];    // rows from Budget (optional)

let mergedData = [];    // result after mergeData()

// small UI state saved in localStorage (only tiny settings)
const UI_LS_KEY = "ndarboe_ui_state_v1";

// ---------------------- Utilities ----------------------
function getVal(row, candidates){
  if(!row) return undefined;
  for(const c of candidates){
    if(c in row && row[c] !== undefined && row[c] !== null && row[c] !== "") return row[c];
    // try case-insensitive match
    for(const k of Object.keys(row)){
      if(k.toLowerCase() === c.toLowerCase() && row[k] !== "" && row[k] !== null && row[k] !== undefined) {
        return row[k];
      }
    }
  }
  return undefined;
}
function safeNum(v){
  if(v === undefined || v === null || v === "" ) return NaN;
  // remove non-numeric except dot and minus
  const n = Number(String(v).replace(/[^0-9\.\-]/g,""));
  return isNaN(n)? NaN : n;
}
function downloadFile(filename, content, mime){
  const blob = new Blob([content], { type: mime || 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ---------------------- XLSX parse ----------------------
function parseFileToJson(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try{
        const bin = e.target.result;
        const wb = XLSX.read(bin, { type: 'binary' });
        const firstSheet = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheet];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(json);
      } catch(err){
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
}

// ---------------------- Merge Logic (core) ----------------------
function mergeData(){
  mergedData = [];

  if(!iw39Data || iw39Data.length === 0){
    alert("IW39 belum di-upload. Upload IW39 dulu lalu klik Refresh.");
    return;
  }

  // prepare lookup maps
  const sum57ByKey = new Map();
  sum57Data.forEach(r => {
    const k = (getVal(r, ["Order","Order No","Order_No","Key","No"]) || "").toString().trim();
    if(k) sum57ByKey.set(k, r);
  });

  const data1ByKey = new Map();
  data1Data.forEach(r => {
    const k = (getVal(r, ["Key","Order","Order No","MAT","Equipment"]) || "").toString().trim();
    if(k) data1ByKey.set(k, r);
  });

  const data2ByMat = new Map();
  data2Data.forEach(r => {
    const mk = (getVal(r, ["MAT","Mat","Material","Key"]) || "").toString().trim();
    if(mk) data2ByMat.set(mk, r);
  });

  const planningByOrder = new Map();
  planningData.forEach(r => {
    const ord = (getVal(r, ["Order","Order No","Order_No","Key"]) || "").toString().trim();
    if(ord) planningByOrder.set(ord, r);
  });

  // iterate IW39 rows and build merged rows
  iw39Data.forEach(row => {
    const order = (getVal(row, ["Order","Order No","Order_No","ORD","Key"]) || "").toString().trim();
    const room = getVal(row, ["Room","ROOM","Location","Lokasi"]) || "";
    const orderType = getVal(row, ["Order Type","OrderType","Type","Order_Type"]) || "";
    const description = getVal(row, ["Description","Desc","Keterangan"]) || "";
    const createdOn = getVal(row, ["Created On","CreatedOn","Created_Date","Tanggal"]) || "";
    const userStatus = getVal(row, ["User Status","UserStatus","Status User"]) || "";
    const mat = (getVal(row, ["MAT","Mat","Material"]) || "").toString().trim();

    // Cost uses TotalPlan & TotalActual
    const totalPlan = safeNum(getVal(row, ["TotalPlan","Total Plan","Plan","Total_Plan","TotalPlan."]));
    const totalActual = safeNum(getVal(row, ["TotalActual","Total Actual","Actual","Total_Actual"]));
    let cost;
    if(!isNaN(totalPlan) && !isNaN(totalActual)){
      const calc = (totalPlan - totalActual) / 16500;
      cost = (calc < 0) ? "-" : Number(Math.round((calc + Number.EPSILON) * 100) / 100).toFixed(2);
    } else {
      cost = "-";
    }

    // CPH lookup rule
    let cph = "";
    if(mat && mat.toUpperCase().startsWith("JR")) {
      cph = "JR";
    } else if(mat && data2ByMat.has(mat)){
      cph = getVal(data2ByMat.get(mat), ["CPH","Cph","cph","CPH Code","Code"]) || "";
    } else {
      cph = "";
    }

    // Section from Data1 by Order or MAT
    let section = "";
    if(order && data1ByKey.has(order)){
      section = getVal(data1ByKey.get(order), ["Section","Section Name","SectionName","SECTION"]) || "";
    } else if(mat && data1ByKey.has(mat)){
      section = getVal(data1ByKey.get(mat), ["Section","Section Name","SectionName","SECTION"]) || "";
    } else {
      // fallback: try to find a match where some column equals order
      for(const v of data1Data){
        const k = (getVal(v, ["Key","Order","MAT","Equipment"]) || "").toString().trim();
        if(k === order || k === mat){
          section = getVal(v, ["Section","Section Name","SectionName","SECTION"]) || "";
          break;
        }
      }
    }

    // Status Part & Aging from SUM57 (by Order or MAT)
    let statusPart = "";
    let aging = "";
    const sumRow = sum57ByKey.get(order) || sum57ByKey.get(mat);
    if(sumRow){
      statusPart = getVal(sumRow, ["Status Part","StatusPart","Part Status","Status"]) || "";
      aging = getVal(sumRow, ["Aging","Age"]) || "";
    }

    // Planning & Status AMT
    let planning = "";
    let statusAMT = "";
    const pRow = planningByOrder.get(order);
    if(pRow){
      planning = getVal(pRow, ["Event Start","Planning","Start","EventStart"]) || "";
      statusAMT = getVal(pRow, ["Status AMT","StatusAMT","AMT Status","Status"]) || "";
    }

    // Month & Reman defaults (may be empty)
    let month = getVal(row, ["Month"]) || "";
    let reman = getVal(row, ["Reman"]) || "";

    // Include calculation (depends on reman & cost)
    let include;
    if(cost === "-" || cost === undefined) {
      include = "-";
    } else {
      const costNum = Number(cost);
      if(String(reman).toLowerCase() === "reman") include = Number(Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2);
      else include = Number(costNum).toFixed(2);
    }

    // Exclude
    let exclude;
    if(String(orderType).trim().toUpperCase() === "PM38") exclude = "-";
    else exclude = include;

    const mergedRow = {
      Room: room,
      "Order Type": orderType,
      Order: order,
      Description: description,
      "Created On": createdOn,
      "User Status": userStatus,
      MAT: mat,
      CPH: cph,
      Section: section,
      "Status Part": statusPart,
      Aging: aging,
      Month: month,
      Cost: cost,
      Reman: reman,
      Include: include,
      Exclude: exclude,
      Planning: planning,
      "Status AMT": statusAMT,
      // Keep original IW39 totals in case user wants to export full data
      _IW39_totalPlan: isNaN(totalPlan) ? "" : totalPlan,
      _IW39_totalActual: isNaN(totalActual) ? "" : totalActual
    };

    mergedData.push(mergedRow);
  });

  // Re-apply any previously user-edits saved in small UI state (if present)
  // (We do not store full mergedData to localStorage to avoid quota)
  const uiStateRaw = localStorage.getItem(UI_LS_KEY);
  if(uiStateRaw){
    try {
      const ui = JSON.parse(uiStateRaw);
      if(ui && Array.isArray(ui.userEdits)){
        const editMap = new Map(ui.userEdits.map(r => [r.Order, r]));
        mergedData = mergedData.map(r => {
          const s = editMap.get(r.Order);
          if(s){
            if(s.Month !== undefined) r.Month = s.Month;
            if(s.Reman !== undefined) r.Reman = s.Reman;
            // recalc include/exclude
            if(r.Cost !== "-" && !isNaN(Number(r.Cost))){
              const costNum = Number(r.Cost);
              r.Include = (String(r.Reman).toLowerCase() === "reman") ? Number(Math.round((costNum*0.25+Number.EPSILON)*100)/100).toFixed(2) : Number(costNum).toFixed(2);
            } else {
              r.Include = "-";
            }
            r.Exclude = (String(r["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : r.Include;
          }
          return r;
        });
      }
    } catch(e){
      console.warn("UI state parse failed:", e);
    }
  }
}

// ---------------------- Render table ----------------------
function renderTable(data){
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";

  const rows = Array.isArray(data) ? data : mergedData;
  if(!rows || rows.length === 0){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 19;
    td.style.textAlign = "center";
    td.textContent = "Tidak ada data";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  rows.forEach((row, idx) => {
    const tr = document.createElement("tr");

    const addCell = (val) => {
      const td = document.createElement("td");
      td.textContent = (val === undefined || val === null) ? "" : val;
      return td;
    };

    tr.appendChild(addCell(row.Room));
    tr.appendChild(addCell(row["Order Type"]));
    tr.appendChild(addCell(row.Order));
    tr.appendChild(addCell(row.Description));
    tr.appendChild(addCell(row["Created On"]));
    tr.appendChild(addCell(row["User Status"]));
    tr.appendChild(addCell(row.MAT));
    tr.appendChild(addCell(row.CPH));
    tr.appendChild(addCell(row.Section));
    tr.appendChild(addCell(row["Status Part"]));
    tr.appendChild(addCell(row.Aging));

    // Month cell (editable)
    const tdMonth = document.createElement("td");
    tdMonth.textContent = row.Month || "";
    tdMonth.dataset.col = "Month";
    tr.appendChild(tdMonth);

    tr.appendChild(addCell(row.Cost));
    // Reman (editable)
    const tdReman = document.createElement("td");
    tdReman.textContent = row.Reman || "";
    tdReman.dataset.col = "Reman";
    tr.appendChild(tdReman);

    tr.appendChild(addCell(row.Include));
    tr.appendChild(addCell(row.Exclude));
    tr.appendChild(addCell(row.Planning));
    tr.appendChild(addCell(row["Status AMT"]));

    // Action cell
    const tdAction = document.createElement("td");
    const editBtn = document.createElement("button");
    editBtn.textContent = "Edit";
    editBtn.className = "action-btn edit-btn";
    editBtn.addEventListener("click", () => startEditRow(idx, tr));
    tdAction.appendChild(editBtn);

    const delBtn = document.createElement("button");
    delBtn.textContent = "Delete";
    delBtn.className = "action-btn delete-btn";
    delBtn.addEventListener("click", () => {
      if(confirm("Hapus baris order " + (row.Order || "") + " ?")){
        // remove from mergedData by matching Order (safe if table filtered)
        const globalIndex = mergedData.findIndex(r => r.Order === row.Order);
        if(globalIndex !== -1) mergedData.splice(globalIndex, 1);
        // also remove any saved UI edit for this order
        removeUserEdit(row.Order);
        renderTable(mergedData);
      }
    });
    tdAction.appendChild(delBtn);

    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

// ---------------------- Edit row inline ----------------------
function startEditRow(index, trElement){
  // index is index w.r.t. the currently shown set used in renderTable
  // but renderTable uses mergedData by default, so idx is global index in mergedData
  const row = mergedData[index];
  if(!row) return;

  const monthTd = trElement.querySelector('td[data-col="Month"]');
  const remanTd = trElement.querySelector('td[data-col="Reman"]');
  if(!monthTd || !remanTd) return;

  // create select for months
  const months = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const sel = document.createElement("select");
  months.forEach(m => {
    const o = document.createElement("option");
    o.value = m;
    o.text = m || "--";
    if(m === row.Month) o.selected = true;
    sel.appendChild(o);
  });
  monthTd.innerHTML = "";
  monthTd.appendChild(sel);

  const remInput = document.createElement("input");
  remInput.type = "text";
  remInput.value = row.Reman || "";
  remInput.style.width = "100%";
  remanTd.innerHTML = "";
  remanTd.appendChild(remInput);

  // action cell: replace with Save & Cancel
  const actionTd = trElement.querySelector('td:last-child');
  actionTd.innerHTML = "";

  const saveBtn = document.createElement("button");
  saveBtn.textContent = "Save";
  saveBtn.className = "action-btn save-btn";
  saveBtn.addEventListener("click", () => {
    // apply changes to mergedData
    row.Month = sel.value;
    row.Reman = remInput.value;

    // recalc Include & Exclude
    if(row.Cost !== "-" && !isNaN(Number(row.Cost))){
      const costNum = Number(row.Cost);
      row.Include = (String(row.Reman).toLowerCase() === "reman") ? Number(Math.round((costNum*0.25+Number.EPSILON)*100)/100).toFixed(2) : Number(costNum).toFixed(2);
    } else {
      row.Include = "-";
    }
    row.Exclude = (String(row["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : row.Include;

    // persist edit to small UI state
    saveUserEdit(row.Order, { Order: row.Order, Month: row.Month, Reman: row.Reman });

    renderTable(mergedData);
  });
  actionTd.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "action-btn cancel-btn";
  cancelBtn.addEventListener("click", () => {
    renderTable(mergedData);
  });
  actionTd.appendChild(cancelBtn);
}

// ---------------------- User edit persistence (small) ----------------------
function saveUserEdit(orderKey, editObj){
  // only keep small array of edits in localStorage to reapply after merge (not full dataset)
  let uiState = { userEdits: [] };
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if(raw) uiState = JSON.parse(raw);
  } catch(e){}
  // remove existing for same key
  uiState.userEdits = uiState.userEdits.filter(r => r.Order !== orderKey);
  uiState.userEdits.push(editObj);
  try {
    localStorage.setItem(UI_LS_KEY, JSON.stringify(uiState));
  } catch(e){
    console.warn("Gagal menyimpan UI state:", e);
  }
}
function removeUserEdit(orderKey){
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if(!raw) return;
    const uiState = JSON.parse(raw);
    uiState.userEdits = uiState.userEdits.filter(r => r.Order !== orderKey);
    localStorage.setItem(UI_LS_KEY, JSON.stringify(uiState));
  } catch(e){}
}

// ---------------------- Filter / Reset ----------------------
function filterData(){
  let filtered = mergedData.slice();

  const room = document.getElementById("filter-room").value.trim().toLowerCase();
  const order = document.getElementById("filter-order").value.trim().toLowerCase();
  const cph = document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat = document.getElementById("filter-mat").value.trim().toLowerCase();
  const section = document.getElementById("filter-section").value.trim().toLowerCase();

  if(room) filtered = filtered.filter(d => (d.Room || "").toString().toLowerCase().includes(room));
  if(order) filtered = filtered.filter(d => (d.Order || "").toString().toLowerCase().includes(order));
  if(cph) filtered = filtered.filter(d => (d.CPH || "").toString().toLowerCase().includes(cph));
  if(mat) filtered = filtered.filter(d => (d.MAT || "").toString().toLowerCase().includes(mat));
  if(section) filtered = filtered.filter(d => (d.Section || "").toString().toLowerCase().includes(section));

  renderTable(filtered);
}
function resetFilter(){
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section"].forEach(id => {
    const el = document.getElementById(id);
    if(el) el.value = "";
  });
  renderTable(mergedData);
}

// ---------------------- Add Orders manual ----------------------
function addOrders(){
  const input = document.getElementById("add-order-input").value.trim();
  const status = document.getElementById("add-order-status");
  if(!input){
    status.textContent = "Masukkan order dulu ya bro!";
    status.style.color = "red";
    return;
  }
  const orders = input.split(/[\s,]+/).filter(o => o.length>0);
  let added = 0;
  orders.forEach(o => {
    if(mergedData.find(r => r.Order === o)) return; // skip duplicates
    // try find full IW39 row
    const iw = iw39Data.find(r => {
      const val = (getVal(r, ["Order","Order No","Order_No","Key"]) || "").toString().trim();
      return val === o;
    });
    if(iw){
      // build merged row for this single IW row using same logic (re-using mergeData approach)
      const tempIw39 = [iw];
      // quick single-row merge function - reuse code path: easier to just call mergeData after adding nothing
      // For simplicity we'll push minimal row (user can click Refresh to rebuild full mergedData)
      mergedData.push({
        Room: getVal(iw, ["Room","ROOM","Location"]) || "",
        "Order Type": getVal(iw, ["Order Type","OrderType"]) || "",
        Order: (getVal(iw, ["Order","Order No","Key"]) || "").toString().trim(),
        Description: getVal(iw, ["Description","Desc"]) || "",
        "Created On": getVal(iw, ["Created On","CreatedOn"]) || "",
        "User Status": getVal(iw, ["User Status","UserStatus"]) || "",
        MAT: (getVal(iw, ["MAT","Mat","Material"]) || "").toString().trim(),
        CPH: "",
        Section: "",
        "Status Part": "",
        Aging: "",
        Month: "",
        Cost: "-",
        Reman: "",
        Include: "-",
        Exclude: "-",
        Planning: "",
        "Status AMT": ""
      });
      added++;
    } else {
      mergedData.push({
        Room: "",
        "Order Type": "",
        Order: o,
        Description: "",
        "Created On": "",
        "User Status": "",
        MAT: "",
        CPH: "",
        Section: "",
        "Status Part": "",
        Aging: "",
        Month: "",
        Cost: "-",
        Reman: "",
        Include: "-",
        Exclude: "-",
        Planning: "",
        "Status AMT": ""
      });
      added++;
    }
  });

  document.getElementById("add-order-status").textContent = `${added} order berhasil ditambahkan.`;
  document.getElementById("add-order-status").style.color = "green";
  document.getElementById("add-order-input").value = "";
  renderTable(mergedData);
}

// ---------------------- Export mergedData to Excel ----------------------
function exportMergedToExcel(){
  if(!mergedData || mergedData.length === 0){
    alert("Tidak ada data untuk diexport.");
    return;
  }
  // Prepare sheet rows (remove _IW39 keys)
  const exportRows = mergedData.map(r => {
    const o = Object.assign({}, r);
    delete o._IW39_totalPlan;
    delete o._IW39_totalActual;
    return o;
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Merged");
  const wbout = XLSX.write(wb, { bookType:'xlsx', type:'binary' });

  // binary string to ArrayBuffer
  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  downloadFile("merged_ndarboe.xlsx", s2ab(wbout), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// ---------------------- Import/Export JSON (backup) ----------------------
function downloadJSONBackup(){
  const payload = { iw39Data, sum57Data, planningData, data1Data, data2Data, budgetData, mergedData, timestamp: new Date().toISOString() };
  downloadFile("ndarboe_backup.json", JSON.stringify(payload, null, 2), "application/json");
}
function loadJSONBackupFile(file){
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const obj = JSON.parse(e.target.result);
      iw39Data = obj.iw39Data || [];
      sum57Data = obj.sum57Data || [];
      planningData = obj.planningData || [];
      data1Data = obj.data1Data || [];
      data2Data = obj.data2Data || [];
      budgetData = obj.budgetData || [];
      mergedData = obj.mergedData || [];
      renderTable(mergedData);
      alert("Backup JSON dimuat.");
    } catch(err){
      alert("Gagal memuat JSON: " + err.message);
    }
  };
  reader.readAsText(file);
}

// ---------------------- UI wiring ----------------------
function wireUp(){
  // Sidebar menu switching
  document.querySelectorAll(".menu-item").forEach(item => {
    item.addEventListener("click", () => {
      document.querySelectorAll(".menu-item").forEach(i => i.classList.remove("active"));
      document.querySelectorAll(".content-section").forEach(s => s.classList.remove("active"));
      item.classList.add("active");
      const menu = item.getAttribute("data-menu");
      const sec = document.getElementById(menu);
      if(sec) sec.classList.add("active");
    });
  });

  // Upload button
  document.getElementById("upload-btn").addEventListener("click", async () => {
    const sel = document.getElementById("file-select").value;
    const f = document.getElementById("file-input").files[0];
    if(!f){ alert("Pilih file terlebih dahulu"); return; }
    document.getElementById("upload-status").textContent = `Parsing ${f.name} ...`;
    try {
      const json = await parseFileToJson(f);
      switch(sel){
        case "IW39": iw39Data = json; break;
        case "SUM57": sum57Data = json; break;
        case "Planning": planningData = json; break;
        case "Data1": data1Data = json; break;
        case "Data2": data2Data = json; break;
        case "Budget": budgetData = json; break;
      }
      document.getElementById("upload-status").textContent = `${sel} loaded (${json.length} rows)`;
      document.getElementById("file-input").value = "";
    } catch(err){
      console.error(err);
      alert("Gagal parsing file: " + err.message);
    }
  });

  // Clear files (memory)
  document.getElementById("clear-files-btn").addEventListener("click", () => {
    if(!confirm("Clear semua data yang sudah di-upload di memory?")) return;
    iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[];
    mergedData=[];
    document.getElementById("upload-status").textContent = "Data cleared";
    renderTable([]);
  });

  // Refresh (merge)
  document.getElementById("refresh-btn").addEventListener("click", () => {
    if(!iw39Data || iw39Data.length === 0){
      alert("Upload IW39 dulu sebelum Refresh.");
      return;
    }
    mergeData();
    renderTable(mergedData);
    document.getElementById("add-order-status").textContent = "";
  });

  // Add orders
  document.getElementById("add-order-btn").addEventListener("click", addOrders);

  // Filter / Reset
  document.getElementById("filter-btn").addEventListener("click", filterData);
  document.getElementById("reset-btn").addEventListener("click", resetFilter);

  // Save (export) merged to Excel
  document.getElementById("save-btn").addEventListener("click", () => {
    exportMergedToExcel();
  });

  // Load (backup JSON) - open file input
  document.getElementById("load-btn").addEventListener("click", () => {
    // open native file selector
    const inpf = document.createElement("input");
    inpf.type = "file";
    inpf.accept = ".json";
    inpf.addEventListener("change", (e) => {
      const f = e.target.files[0];
      if(f) loadJSONBackupFile(f);
    });
    inpf.click();
  });

  // Export/Import backup JSON buttons (optional quick links)
  // If you want add explicit buttons for backup in UI, wire them here:
  // document.getElementById("export-json-btn").addEventListener("click", downloadJSONBackup);

  // Try to load UI state small settings
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if(raw){
      // no heavy data - maybe last active menu etc. (not required)
      const ui = JSON.parse(raw);
      // (optional) apply ui settings
    }
  } catch(e){}

  // initial render
  renderTable([]);
}

// ---------------------- Init ----------------------
window.addEventListener("DOMContentLoaded", () => {
  wireUp();
});
