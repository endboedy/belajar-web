/* script.js
 - Upload / parse Excel files (XLSX)
 - Merge IW39 + lookups (Data1, Data2, SUM57, Planning)
 - Render table, filter, edit Month/Reman, delete rows
*/

// --------- Global data stores ----------
let iw39Data = [];      // array of objects from IW39
let sum57Data = [];     // array from SUM57
let planningData = [];  // array from Planning
let data1Data = [];     // Data1
let data2Data = [];     // Data2
let budgetData = [];    // Budget (unused currently)

let mergedData = [];    // result of mergeData()
// localStorage keys
const LS_KEY = "ndarboe_lembar_data_v1";

// helpful utilities to read flexible column names
function getVal(row, candidates){
  if(!row) return undefined;
  for(const c of candidates){
    if(c in row && row[c] !== undefined && row[c] !== null) return row[c];
    // also try trimmed lowercase match
    for(const k of Object.keys(row)){
      if(k.toLowerCase() === c.toLowerCase()){
        return row[k];
      }
    }
  }
  return undefined;
}
function safeNum(v){
  if(v === undefined || v === null || v === "" ) return NaN;
  const n = Number(String(v).toString().replace(/[^0-9\.\-]/g,""));
  return isNaN(n)? NaN : n;
}

// ---------- XLSX parsing ----------
function parseFile(file, destKey){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try{
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(json);
      } catch(err){
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
}

// ---------- Merge logic ----------
function mergeData(){
  mergedData = [];

  if(!iw39Data || iw39Data.length === 0){
    alert("IW39 belum di-upload atau kosong. Upload IW39 dulu.");
    return;
  }

  // create lookup maps for faster search (keyed by Order or MAT depending)
  const sum57ByKey = new Map();
  sum57Data.forEach(r => {
    const k = (getVal(r, ["Order","Order No","Order_No","OrderID","Key"]) || "").toString().trim();
    if(k) sum57ByKey.set(k, r);
  });

  const data1ByKey = new Map();
  data1Data.forEach(r => {
    const k = (getVal(r, ["Key","Order","Order No","MAT","Equipment"]) || "").toString().trim();
    if(k) data1ByKey.set(k, r);
  });

  const data2ByMat = new Map();
  data2Data.forEach(r => {
    const matKey = (getVal(r, ["MAT","Mat","Material","Key"]) || "").toString().trim();
    if(matKey) data2ByMat.set(matKey, r);
  });

  const planningByOrder = new Map();
  planningData.forEach(r => {
    const ord = (getVal(r, ["Order","Order No","Order_No","Key"]) || "").toString().trim();
    if(ord) planningByOrder.set(ord, r);
  });

  // iterate IW39 rows
  iw39Data.forEach(row => {
    const order = (getVal(row, ["Order","Order No","Order_No","ORD","Key"]) || "").toString().trim();
    const room = getVal(row, ["Room","ROOM","Location","Lokasi"]) || "";
    const orderType = getVal(row, ["Order Type","OrderType","Type","Order_Type"]) || "";
    const description = getVal(row, ["Description","Desc","Keterangan"]) || "";
    const createdOn = getVal(row, ["Created On","CreatedOn","Created_Date","Tanggal"]) || "";
    const userStatus = getVal(row, ["User Status","UserStatus","Status User"]) || "";
    const mat = (getVal(row, ["MAT","Mat","Material"]) || "").toString().trim();

    // Cost formula uses TotalPlan & TotalActual in IW39
    const totalPlan = safeNum(getVal(row, ["TotalPlan","Total Plan","Plan","Total_Plan","TotalPlan."]));
    const totalActual = safeNum(getVal(row, ["TotalActual","Total Actual","Actual","Total_Actual"]));
    let cost = NaN;
    if(!isNaN(totalPlan) && !isNaN(totalActual)){
      cost = (totalPlan - totalActual) / 16500;
      if(cost < 0) cost = "-";
      else cost = Math.round((cost + Number.EPSILON) * 100) / 100; // 2 decimals
    } else {
      cost = "-";
    }

    // CPH: if MAT starts with JR => "JR" else lookup in data2 by MAT, field named CPH or similar
    let cph = "";
    if(mat && mat.toUpperCase().startsWith("JR")) {
      cph = "JR";
    } else {
      const d2 = data2ByMat.get(mat);
      if(d2){
        cph = getVal(d2, ["CPH","Cph","cph","CPH Code","Code"]) || "";
      } else {
        cph = "";
      }
    }

    // Section: lookup in Data1 by key (we try by Order OR MAT)
    let section = "";
    let secSource = null;
    if(order && data1ByKey.has(order)) {
      secSource = data1ByKey.get(order);
    } else if(mat && data1ByKey.has(mat)){
      secSource = data1ByKey.get(mat);
    } else {
      // try fuzzy by some column in Data1 that matches Room or other
      for(const v of data1Data){
        const k = (getVal(v, ["Key","Order","MAT","Equipment"]) || "").toString();
        if(k && k === order){ secSource = v; break; }
      }
    }
    if(secSource){
      section = getVal(secSource, ["Section","Section Name","SectionName","SECTION"]) || "";
    }

    // Status Part & Aging: lookup in SUM57 by Order or MAT
    let statusPart = "";
    let aging = "";
    const sumRow = sum57ByKey.get(order) || sum57ByKey.get(mat);
    if(sumRow){
      statusPart = getVal(sumRow, ["Status Part","StatusPart","Part Status","Status"]) || "";
      aging = getVal(sumRow, ["Aging","Age"]) || "";
    }

    // Planning & Status AMT from Planning lookup by Order
    let planning = "";
    let statusAMT = "";
    const pRow = planningByOrder.get(order);
    if(pRow){
      planning = getVal(pRow, ["Event Start","Planning","Start","EventStart"]) || "";
      statusAMT = getVal(pRow, ["Status AMT","StatusAMT","AMT Status","Status"]) || "";
    }

    // Default month & reman if previously saved? We'll default to "".
    let month = getVal(row, ["Month"]) || "";
    let reman = getVal(row, ["Reman"]) || "";

    // Include calculation
    let include = "-";
    if(cost === "-") {
      include = "-";
    } else {
      const costNum = Number(cost);
      if(String(reman).toLowerCase() === "reman".toLowerCase()){
        include = Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100;
      } else {
        include = costNum;
      }
    }

    // Exclude calculation
    let exclude = "";
    if(String(orderType).trim().toUpperCase() === "PM38") {
      exclude = "-";
    } else {
      exclude = include;
    }

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
      "Status AMT": statusAMT
    };

    mergedData.push(mergedRow);
  });

  // if we previously saved user edits in localStorage, re-apply Month/Reman edits
  const saved = loadFromLocal();
  if(saved && Array.isArray(saved.mergedData)){
    // match by Order and reapply month/reman
    const savedMap = new Map(saved.mergedData.map(r => [r.Order, r]));
    mergedData = mergedData.map(r => {
      const s = savedMap.get(r.Order);
      if(s){
        if(s.Month !== undefined) r.Month = s.Month;
        if(s.Reman !== undefined) r.Reman = s.Reman;
        // recalc include/exclude based on possibly updated Reman
        if(r.Cost !== "-" && !isNaN(Number(r.Cost))){
          const costNum = Number(r.Cost);
          r.Include = (String(r.Reman).toLowerCase() === "reman") ? Math.round((costNum*0.25+Number.EPSILON)*100)/100 : costNum;
        } else {
          r.Include = "-";
        }
        r.Exclude = (String(r["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : r.Include;
      }
      return r;
    });
  }

  // finished
}

// ---------- Render table ----------
function renderTable(data){
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";

  if(!data || data.length === 0){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 19;
    td.style.textAlign = "center";
    td.textContent = "Tidak ada data";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  data.forEach((row, idx) => {
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
    // Month (editable)
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

    // Actions
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
      if(confirm("Hapus baris order " + row.Order + " ?")){
        mergedData.splice(idx,1);
        saveToLocal(); // persist
        renderTable(mergedData);
      }
    });
    tdAction.appendChild(delBtn);

    tr.appendChild(tdAction);

    tbody.appendChild(tr);
  });
}

// ---------- Edit row inline ----------
function startEditRow(index, trElement){
  // replace Month cell with select and Reman cell with input, change Edit -> Save/Cancel
  const row = mergedData[index];
  if(!row) return;

  // get Month td (12th cell: index 11)
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

  // action cell buttons
  const actionTd = trElement.querySelector('td:last-child');
  actionTd.innerHTML = "";

  const saveBtn = document.createElement("button");
  saveBtn.textContent = "Save";
  saveBtn.className = "action-btn save-btn";
  saveBtn.addEventListener("click", () => {
    // apply edits
    row.Month = sel.value;
    row.Reman = remInput.value;

    // recalc include/exclude
    if(row.Cost !== "-" && !isNaN(Number(row.Cost))){
      const costNum = Number(row.Cost);
      row.Include = (String(row.Reman).toLowerCase() === "reman") ? Math.round((costNum*0.25+Number.EPSILON)*100)/100 : costNum;
    } else {
      row.Include = "-";
    }
    row.Exclude = (String(row["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : row.Include;

    saveToLocal();
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

// ---------- Filter & Reset ----------
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
    document.getElementById(id).value = "";
  });
  renderTable(mergedData);
}

// ---------- Add Orders (manual) ----------
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
    // check if already exists
    if(mergedData.find(r => r.Order === o)) return;
    // try to find IW39 row for that order
    const iw = iw39Data.find(r => {
      const val = (getVal(r, ["Order","Order No","Order_No","Key"]) || "").toString().trim();
      return val === o;
    });
    if(iw){
      // rebuild single merged row from iw
      iw39Data = iw39Data; // no-op
      // run mergeData for all then push? For simplicity, push one built row by using merge approach
      const tempIw = [iw];
      const prevIw = iw39Data;
      // quick local merge for this one:
      // reuse mergeData logic by temporarily assigning iw39Data to [iw], but better to reconstruct this row:
      // We'll just push object with core fields; user can refresh to get full lookup
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
      // create minimal row if IW39 doesn't have it
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
  saveToLocal();
  renderTable(mergedData);
}

// ---------- Save / Load local ----------
function saveToLocal(){
  const payload = {
    timestamp: Date.now(),
    iw39Data, sum57Data, planningData, data1Data, data2Data, budgetData,
    mergedData
  };
  localStorage.setItem(LS_KEY, JSON.stringify(payload));
  console.log("Disimpan ke localStorage");
}
function loadFromLocal(){
  const raw = localStorage.getItem(LS_KEY);
  if(!raw) return null;
  try{
    return JSON.parse(raw);
  }catch(e){
    return null;
  }
}

// ---------- UI handlers ----------
function wireUp(){
  // menu
  document.querySelectorAll(".menu-item").forEach(it => {
    it.addEventListener("click", () => {
      document.querySelectorAll(".menu-item").forEach(i=>i.classList.remove("active"));
      document.querySelectorAll(".content-section").forEach(s=>s.classList.remove("active"));
      it.classList.add("active");
      const m = it.dataset.menu;
      const sec = document.getElementById(m);
      if(sec) sec.classList.add("active");
    });
  });

  // upload
  document.getElementById("upload-btn").addEventListener("click", async () => {
    const sel = document.getElementById("file-select").value;
    const f = document.getElementById("file-input").files[0];
    if(!f){ alert("Pilih file terlebih dahulu"); return; }
    document.getElementById("upload-status").textContent = `Parsing ${f.name} ...`;
    try{
      const json = await parseFile(f, sel);
      if(sel === "IW39") iw39Data = json;
      else if(sel === "SUM57") sum57Data = json;
      else if(sel === "Planning") planningData = json;
      else if(sel === "Data1") data1Data = json;
      else if(sel === "Data2") data2Data = json;
      else if(sel === "Budget") budgetData = json;
      document.getElementById("upload-status").textContent = `${sel} loaded (${json.length} rows)`;
      // keep file input empty for next
      document.getElementById("file-input").value = "";
      saveToLocal();
    }catch(err){
      console.error(err);
      alert("Gagal parsing file: " + err.message);
    }
  });

  document.getElementById("clear-files-btn").addEventListener("click", () => {
    if(confirm("Clear semua data yang di-upload (data di memory) ?")){
      iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[];
      mergedData=[];
      localStorage.removeItem(LS_KEY);
      document.getElementById("upload-status").textContent = "Data cleared";
      renderTable(mergedData);
    }
  });

  // Lembar kerja controls
  document.getElementById("filter-btn").addEventListener("click", filterData);
  document.getElementById("reset-btn").addEventListener("click", resetFilter);
  document.getElementById("refresh-btn").addEventListener("click", () => {
    // require IW39
    if(!iw39Data || iw39Data.length === 0){
      alert("Upload IW39 dulu sebelum Refresh.");
      return;
    }
    mergeData();
    saveToLocal();
    renderTable(mergedData);
    document.getElementById("add-order-status").textContent = "";
  });

  document.getElementById("add-order-btn").addEventListener("click", addOrders);

  document.getElementById("save-btn").addEventListener("click", () => {
    saveToLocal();
    alert("Tersimpan ke localStorage.");
  });
  document.getElementById("load-btn").addEventListener("click", () => {
    const saved = loadFromLocal();
    if(saved){
      iw39Data = saved.iw39Data || [];
      sum57Data = saved.sum57Data || [];
      planningData = saved.planningData || [];
      data1Data = saved.data1Data || [];
      data2Data = saved.data2Data || [];
      budgetData = saved.budgetData || [];
      mergedData = saved.mergedData || [];
      renderTable(mergedData);
      alert("Data dimuat dari localStorage.");
    } else {
      alert("Tidak ada data di localStorage.");
    }
  });

  // try to auto-load previously saved
  const prev = loadFromLocal();
  if(prev){
    iw39Data = prev.iw39Data || [];
    sum57Data = prev.sum57Data || [];
    planningData = prev.planningData || [];
    data1Data = prev.data1Data || [];
    data2Data = prev.data2Data || [];
    budgetData = prev.budgetData || [];
    mergedData = prev.mergedData || [];
    renderTable(mergedData);
  } else {
    renderTable([]);
  }
}

// ---------- init ----------
window.addEventListener("DOMContentLoaded", () => {
  wireUp();
});
