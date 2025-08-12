/* script.js - FINAL Revised
   - Upload & parse XLSX: IW39, SUM57, Planning, Data1, Data2, Budget
   - mergeData() with lookups:
       Section <- Data1 by Room
       Status Part, Aging <- SUM57 by Order
       CPH <- Data2 (JR rule + lookup by MAT)
       Cost = (TotalPlan - TotalActual)/16500 -> "-" if <0 or missing
   - Format dates dd-MMM-yyyy for Created On & Planning
   - Render table (Menu 2) with Edit (Month, Reman) & Delete
   - Add Orders, Filter, Reset, Refresh, Export XLSX, JSON backup
   - Persist only small UI edits (Month/Reman) in localStorage (UI_LS_KEY)
*/

// ---------------- Global stores ----------------
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];

let mergedData = []; // merged result used for rendering

const UI_LS_KEY = "ndarboe_ui_state_v1"; // small persisted edits

// ---------------- Utility helpers ----------------
function getVal(row, candidates) {
  if (!row) return undefined;
  for (const c of candidates) {
    if (c in row && row[c] !== undefined && row[c] !== null && row[c] !== "") return row[c];
    // case-insensitive fallback
    for (const k of Object.keys(row)) {
      if (k.toLowerCase() === c.toLowerCase() && row[k] !== "" && row[k] !== null && row[k] !== undefined) {
        return row[k];
      }
    }
  }
  return undefined;
}
function safeNum(v) {
  if (v === undefined || v === null || v === "") return NaN;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^0-9\.\-]/g, "");
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}
// Excel date serial to JS Date (approx)
function excelDateToJS(serial) {
  if (typeof serial === "number") {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional = serial - Math.floor(serial);
    if (fractional > 0) {
      const seconds = Math.round(fractional * 86400);
      date_info.setSeconds(date_info.getSeconds() + seconds);
    }
    return date_info;
  }
  const d = new Date(serial);
  if (!isNaN(d)) return d;
  return null;
}
function formatDateDDMMMYYYY(input) {
  if (input === undefined || input === null || input === "") return "";
  let d = null;
  if (typeof input === "number") d = excelDateToJS(input);
  else {
    // try parse ISO or other
    d = new Date(input);
    if (isNaN(d)) {
      const alt = new Date(String(input).replace(/\//g, "-"));
      d = isNaN(alt) ? null : alt;
    }
  }
  if (!d || isNaN(d)) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const day = String(d.getDate()).padStart(2,"0");
  const mon = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${mon}-${year}`;
}
function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ---------------- Parse XLSX ----------------
function parseFileToJson(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const wb = XLSX.read(data, { type: "binary" });
        const first = wb.SheetNames[0];
        const ws = wb.Sheets[first];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
}

// ---------------- Merge logic ----------------
function mergeData() {
  mergedData = [];

  if (!iw39Data || iw39Data.length === 0) {
    alert("IW39 belum di-upload. Upload IW39 lalu klik Refresh.");
    return;
  }

  // Build quick lookup maps
  const sum57ByOrder = new Map();
  sum57Data.forEach(r => {
    const k = String(getVal(r, ["Order","Order No","Order_No","Key","No"]) || "").trim();
    if (k) sum57ByOrder.set(k, r);
  });

  const data1ByRoom = new Map();
  data1Data.forEach(r => {
    const k = String(getVal(r, ["Room","ROOM","Lokasi","Location"]) || "").trim();
    if (k) data1ByRoom.set(k, r);
  });

  const data2ByMat = new Map();
  data2Data.forEach(r => {
    const mk = String(getVal(r, ["MAT","Mat","Material","Key"]) || "").trim();
    if (mk) data2ByMat.set(mk, r);
  });

  const planningByOrder = new Map();
  planningData.forEach(r => {
    const ord = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
    if (ord) planningByOrder.set(ord, r);
  });

  // Iterate IW39 rows and create merged rows
  iw39Data.forEach(row => {
    const order = String(getVal(row, ["Order","Order No","Order_No","ORD","Key"]) || "").trim();
    const room = getVal(row, ["Room","ROOM","Location","Lokasi"]) || "";
    const orderType = getVal(row, ["Order Type","OrderType","Type","Order_Type"]) || "";
    const description = getVal(row, ["Description","Desc","Keterangan"]) || "";
    const createdRaw = getVal(row, ["Created On","CreatedOn","Created_Date","Tanggal","Create On","Created"]) || "";
    const userStatus = getVal(row, ["User Status","UserStatus","Status User","Status"]) || "";
    const mat = String(getVal(row, ["MAT","Mat","Material"]) || "").trim();

    const totalPlan = safeNum(getVal(row, ["TotalPlan","Total Plan","Plan","Total_Plan","TotalPlan."]));
    const totalActual = safeNum(getVal(row, ["TotalActual","Total Actual","Actual","Total_Actual"]));

    // Cost calculation
    let cost;
    if (!isNaN(totalPlan) && !isNaN(totalActual)) {
      const calc = (totalPlan - totalActual) / 16500;
      cost = (isNaN(calc) || calc === null) ? "-" : (calc < 0 ? "-" : (Math.round((calc + Number.EPSILON) * 100) / 100).toFixed(2));
    } else {
      cost = "-";
    }

    // CPH
    let cph = "";
    if (mat && mat.toUpperCase().startsWith("JR")) cph = "JR";
    else if (mat && data2ByMat.has(mat)) cph = getVal(data2ByMat.get(mat), ["CPH","Cph","cph","CPH Code","Code"]) || "";
    else cph = "";

    // Section by Room from Data1
    let section = "-";
    if (room && data1ByRoom.has(String(room).trim())) {
      section = getVal(data1ByRoom.get(String(room).trim()), ["Section","Section Name","SectionName","SECTION"]) || "-";
    } else {
      // fallback search
      const f = data1Data.find(r => {
        const k = String(getVal(r, ["Room","ROOM","Location","Lokasi"]) || "").trim();
        return k && String(k) === String(room).trim();
      });
      if (f) section = getVal(f, ["Section","Section Name","SectionName","SECTION"]) || "-";
    }

    // Status Part & Aging from SUM57 by Order
    let statusPart = "-";
    let aging = "-";
    if (order && sum57ByOrder.has(order)) {
      const s = sum57ByOrder.get(order);
      statusPart = getVal(s, ["Status Part","StatusPart","Part Status","Status"]) || "-";
      aging = getVal(s, ["Aging","Age","Aging Days"]) || "-";
    }

    // Planning & Status AMT
    let planning = "";
    let statusAMT = "";
    if (order && planningByOrder.has(order)) {
      const p = planningByOrder.get(order);
      planning = getVal(p, ["Event Start","Planning","Start","EventStart","Start Date","Start_Date"]) || "";
      statusAMT = getVal(p, ["Status AMT","StatusAMT","AMT Status","Status"]) || "";
    }

    // Month & Reman defaults (from IW39 if present)
    let month = getVal(row, ["Month"]) || "";
    let reman = getVal(row, ["Reman"]) || "";

    // Include
    let include = "-";
    if (cost === "-" || cost === undefined) include = "-";
    else {
      const cnum = Number(cost);
      include = (String(reman).toLowerCase() === "reman") ? (Math.round((cnum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : cnum.toFixed(2);
    }
    // Exclude
    let exclude = (String(orderType).trim().toUpperCase() === "PM38") ? "-" : include;

    mergedData.push({
      Room: room || "",
      "Order Type": orderType || "",
      Order: order || "",
      Description: description || "",
      "Created On": createdRaw || "",
      "User Status": userStatus || "",
      MAT: mat || "",
      CPH: cph || "",
      Section: section || "-",
      "Status Part": statusPart || "-",
      Aging: aging || "-",
      Month: month || "",
      Cost: cost,
      Reman: reman || "",
      Include: include,
      Exclude: exclude,
      Planning: planning || "",
      "Status AMT": statusAMT || "",
      _IW39_totalPlan: isNaN(totalPlan) ? "" : totalPlan,
      _IW39_totalActual: isNaN(totalActual) ? "" : totalActual
    });
  });

  // Reapply small UI edits (Month/Reman) from localStorage
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const ui = JSON.parse(raw);
      if (ui && Array.isArray(ui.userEdits)) {
        const map = new Map(ui.userEdits.map(e => [e.Order, e]));
        mergedData = mergedData.map(r => {
          const s = map.get(r.Order);
          if (s) {
            if (s.Month !== undefined) r.Month = s.Month;
            if (s.Reman !== undefined) r.Reman = s.Reman;
            // recalc include/exclude
            if (r.Cost !== "-" && !isNaN(Number(r.Cost))) {
              const costNum = Number(r.Cost);
              r.Include = (String(r.Reman).toLowerCase() === "reman") ? (Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : costNum.toFixed(2);
            } else {
              r.Include = "-";
            }
            r.Exclude = (String(r["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : r.Include;
          }
          return r;
        });
      }
    }
  } catch (e) {
    console.warn("UI reapply failed:", e);
  }
}

// ---------------- Render table ----------------
function renderTable(dataToRender) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) {
    console.error("#output-table tbody not found");
    return;
  }
  tbody.innerHTML = "";

  const rows = Array.isArray(dataToRender) ? dataToRender : mergedData;
  if (!rows || rows.length === 0) {
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

    function addCell(val) {
      const td = document.createElement("td");
      td.textContent = (val === undefined || val === null) ? "" : val;
      return td;
    }

    const createdDisplay = row["Created On"] ? formatDateDDMMMYYYY(row["Created On"]) : "";
    const planningDisplay = row.Planning ? formatDateDDMMMYYYY(row.Planning) : "";

    tr.appendChild(addCell(row.Room));
    tr.appendChild(addCell(row["Order Type"]));
    tr.appendChild(addCell(row.Order));
    tr.appendChild(addCell(row.Description));
    tr.appendChild(addCell(createdDisplay));
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
    tr.appendChild(addCell(planningDisplay));
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
      if (confirm("Hapus baris order " + (row.Order || "") + " ?")) {
        const gi = mergedData.findIndex(r => r.Order === row.Order);
        if (gi !== -1) mergedData.splice(gi, 1);
        removeUserEdit(row.Order);
        renderTable(mergedData);
      }
    });
    tdAction.appendChild(delBtn);

    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

// ---------------- Edit inline ----------------
function startEditRow(index, trElement) {
  const row = mergedData[index];
  if (!row) return;

  const monthTd = trElement.querySelector('td[data-col="Month"]');
  const remanTd = trElement.querySelector('td[data-col="Reman"]');
  if (!monthTd || !remanTd) return;

  const months = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const sel = document.createElement("select");
  months.forEach(m => {
    const o = document.createElement("option");
    o.value = m;
    o.text = m || "--";
    if (m === row.Month) o.selected = true;
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

  // action cell save/cancel
  const actionTd = trElement.querySelector("td:last-child");
  actionTd.innerHTML = "";

  const saveBtn = document.createElement("button");
  saveBtn.textContent = "Save";
  saveBtn.className = "action-btn save-btn";
  saveBtn.addEventListener("click", () => {
    row.Month = sel.value;
    row.Reman = remInput.value;

    // recalc include/exclude
    if (row.Cost !== "-" && !isNaN(Number(row.Cost))) {
      const costNum = Number(row.Cost);
      row.Include = (String(row.Reman).toLowerCase() === "reman") ? (Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : costNum.toFixed(2);
    } else {
      row.Include = "-";
    }
    row.Exclude = (String(row["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : row.Include;

    // persist small edit
    saveUserEdit(row.Order, { Order: row.Order, Month: row.Month, Reman: row.Reman });

    renderTable(mergedData);
  });
  actionTd.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "action-btn cancel-btn";
  cancelBtn.addEventListener("click", () => renderTable(mergedData));
  actionTd.appendChild(cancelBtn);
}

// ---------------- small UI edits persistence ----------------
function saveUserEdit(orderKey, editObj) {
  let ui = { userEdits: [] };
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) ui = JSON.parse(raw);
  } catch (e) { ui = { userEdits: [] }; }
  ui.userEdits = ui.userEdits.filter(r => r.Order !== orderKey);
  ui.userEdits.push(editObj);
  try {
    localStorage.setItem(UI_LS_KEY, JSON.stringify(ui));
  } catch (e) {
    console.warn("Could not save UI edits:", e);
  }
}
function removeUserEdit(orderKey) {
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (!raw) return;
    const ui = JSON.parse(raw);
    ui.userEdits = ui.userEdits.filter(r => r.Order !== orderKey);
    localStorage.setItem(UI_LS_KEY, JSON.stringify(ui));
  } catch (e) {}
}

// ---------------- Filter / Reset ----------------
function filterData() {
  let filtered = mergedData.slice();
  const room = document.getElementById("filter-room").value.trim().toLowerCase();
  const order = document.getElementById("filter-order").value.trim().toLowerCase();
  const cph = document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat = document.getElementById("filter-mat").value.trim().toLowerCase();
  const section = document.getElementById("filter-section").value.trim().toLowerCase();

  if (room) filtered = filtered.filter(d => (d.Room || "").toString().toLowerCase().includes(room));
  if (order) filtered = filtered.filter(d => (d.Order || "").toString().toLowerCase().includes(order));
  if (cph) filtered = filtered.filter(d => (d.CPH || "").toString().toLowerCase().includes(cph));
  if (mat) filtered = filtered.filter(d => (d.MAT || "").toString().toLowerCase().includes(mat));
  if (section) filtered = filtered.filter(d => (d.Section || "").toString().toLowerCase().includes(section));

  renderTable(filtered);
}
function resetFilter() {
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  renderTable(mergedData);
}

// ---------------- Add Orders manual ----------------
function addOrders() {
  const input = document.getElementById("add-order-input").value.trim();
  const statusEl = document.getElementById("add-order-status");
  if (!input) {
    if (statusEl) { statusEl.textContent = "Masukkan order dulu ya bro!"; statusEl.style.color = "red"; }
    return;
  }
  const orders = input.split(/[\s,]+/).filter(o => o.length > 0);
  let added = 0;
  orders.forEach(o => {
    if (mergedData.find(r => r.Order === o)) return;
    // try find IW39 row
    const iw = iw39Data.find(r => {
      const v = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
      return v === o;
    });
    if (iw) {
      // push minimal merged record; user can Refresh later to rebuild full merged
      mergedData.push({
        Room: getVal(iw, ["Room","ROOM","Location"]) || "",
        "Order Type": getVal(iw, ["Order Type","OrderType"]) || "",
        Order: (getVal(iw, ["Order","Order No","Key"]) || "").toString().trim(),
        Description: getVal(iw, ["Description","Desc"]) || "",
        "Created On": getVal(iw, ["Created On","CreatedOn","Tanggal"]) || "",
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
    }
    added++;
  });
  if (statusEl) { statusEl.textContent = `${added} order berhasil ditambahkan.`; statusEl.style.color = "green"; }
  document.getElementById("add-order-input").value = "";
  renderTable(mergedData);
}

// ---------------- Export to Excel ----------------
function exportMergedToExcel() {
  if (!mergedData || mergedData.length === 0) { alert("Tidak ada data untuk diexport."); return; }
  const rows = mergedData.map(r => {
    const c = Object.assign({}, r);
    delete c._IW39_totalPlan;
    delete c._IW39_totalActual;
    return c;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Merged");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i=0;i<s.length;i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  downloadFile("Lembar_Kerja_merged.xlsx", s2ab(wbout), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// ---------------- JSON backup ----------------
function downloadJSONBackup() {
  const payload = { iw39Data, sum57Data, planningData, data1Data, data2Data, budgetData, mergedData, timestamp: new Date().toISOString() };
  downloadFile("ndarboe_backup.json", JSON.stringify(payload, null, 2), "application/json");
}
function loadJSONBackupFile(file) {
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
    } catch (err) {
      alert("Gagal memuat JSON: " + err.message);
    }
  };
  reader.readAsText(file);
}

// ---------------- UI wiring ----------------
function wireUp() {
  // menu
  document.querySelectorAll(".menu-item").forEach(it => {
    it.addEventListener("click", () => {
      document.querySelectorAll(".menu-item").forEach(i => i.classList.remove("active"));
      document.querySelectorAll(".content-section").forEach(s => s.classList.remove("active"));
      it.classList.add("active");
      const m = it.dataset.menu;
      const sec = document.getElementById(m);
      if (sec) sec.classList.add("active");
    });
  });

  // upload button
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) {
    uploadBtn.addEventListener("click", async () => {
      const sel = document.getElementById("file-select").value;
      const f = document.getElementById("file-input").files[0];
      if (!f) { alert("Pilih file terlebih dahulu"); return; }
      document.getElementById("upload-status").textContent = `Parsing ${f.name} ...`;
      try {
        const json = await parseFileToJson(f);
        switch (sel) {
          case "IW39": iw39Data = json; break;
          case "SUM57": sum57Data = json; break;
          case "Planning": planningData = json; break;
          case "Data1": data1Data = json; break;
          case "Data2": data2Data = json; break;
          case "Budget": budgetData = json; break;
        }
        document.getElementById("upload-status").textContent = `${sel} loaded (${json.length} rows)`;
        document.getElementById("file-input").value = "";
      } catch (err) {
        console.error(err);
        alert("Gagal parsing file: " + err.message);
      }
    });
  }

  // clear files memory
  const clearBtn = document.getElementById("clear-files-btn");
  if (clearBtn) clearBtn.addEventListener("click", () => {
    if (!confirm("Clear semua data yang sudah di-upload di memory?")) return;
    iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[];
    mergedData=[];
    document.getElementById("upload-status").textContent = "Data cleared";
    renderTable([]);
  });

  // refresh (merge)
  const refreshBtn = document.getElementById("refresh-btn");
  if (refreshBtn) refreshBtn.addEventListener("click", () => {
    if (!iw39Data || iw39Data.length === 0) { alert("Upload IW39 dulu sebelum Refresh."); return; }
    mergeData();
    renderTable(mergedData);
    const s = document.getElementById("add-order-status");
    if (s) s.textContent = "";
  });

  // add orders
  const addBtn = document.getElementById("add-order-btn");
  if (addBtn) addBtn.addEventListener("click", addOrders);

  // filter / reset
  const filterBtn = document.getElementById("filter-btn");
  if (filterBtn) filterBtn.addEventListener("click", filterData);
  const resetBtn = document.getElementById("reset-btn");
  if (resetBtn) resetBtn.addEventListener("click", resetFilter);

  // save (export) button
  const saveBtn = document.getElementById("save-btn");
  if (saveBtn) saveBtn.addEventListener("click", () => exportMergedToExcel());

  // load (backup JSON) button
  const loadBtn = document.getElementById("load-btn");
  if (loadBtn) loadBtn.addEventListener("click", () => {
    const inpf = document.createElement("input");
    inpf.type = "file";
    inpf.accept = ".json";
    inpf.addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (f) loadJSONBackupFile(f);
    });
    inpf.click();
  });

  // initial empty render
  renderTable([]);
}

// ---------------- Init ----------------
window.addEventListener("DOMContentLoaded", () => {
  wireUp();
});
