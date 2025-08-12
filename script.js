// ----------------- Global variables -----------------
const UI_LS_KEY = "ndarboe-ui-edits";

let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];

// ----------------- Helper function to format date (DD MMM YYYY) -----------------
function formatDateDDMMMYYYY(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

// ----------------- Render Table (Menu 2) -----------------
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

    const createdDisp = row["Created On"] ? formatDateDDMMMYYYY(row["Created On"]) : "";
    const planningDisp = row.Planning ? formatDateDDMMMYYYY(row.Planning) : "";

    tr.appendChild(addCell(row.Room));
    tr.appendChild(addCell(row["Order Type"]));
    tr.appendChild(addCell(row.Order));
    tr.appendChild(addCell(row.Description));
    tr.appendChild(addCell(createdDisp));
    tr.appendChild(addCell(row["User Status"]));
    tr.appendChild(addCell(row.MAT));
    tr.appendChild(addCell(row.CPH));
    tr.appendChild(addCell(row.Section));
    tr.appendChild(addCell(row["Status Part"]));
    tr.appendChild(addCell(row.Aging));

    // month editable cell
    const tdMonth = document.createElement("td");
    tdMonth.textContent = row.Month || "";
    tdMonth.dataset.col = "Month";
    tr.appendChild(tdMonth);

    tr.appendChild(addCell(row.Cost));
    // reman editable
    const tdReman = document.createElement("td");
    tdReman.textContent = row.Reman || "";
    tdReman.dataset.col = "Reman";
    tr.appendChild(tdReman);

    tr.appendChild(addCell(row.Include));
    tr.appendChild(addCell(row.Exclude));
    tr.appendChild(addCell(planningDisp));
    tr.appendChild(addCell(row["Status AMT"]));

    // actions
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
        populateMonthFilter(); // refresh month dropdown juga
        renderTable(mergedData);
      }
    });
    tdAction.appendChild(delBtn);

    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

// ----------------- Edit row inline (Month, Reman) -----------------
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
    } else row.Include = "-";
    row.Exclude = (String(row["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : row.Include;

    saveUserEdit(row.Order, { Order: row.Order, Month: row.Month, Reman: row.Reman });
    populateMonthFilter(); // refresh month dropdown
    renderTable(mergedData);
  });
  actionTd.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "action-btn cancel-btn";
  cancelBtn.addEventListener("click", () => renderTable(mergedData));
  actionTd.appendChild(cancelBtn);
}

// ---------------- small UI edits persistence -----------------
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

// ---------------- Filter / Reset -----------------
function filterData() {
  let filtered = mergedData.slice();
  const room = (document.getElementById("filter-room").value || "").trim().toLowerCase();
  const order = (document.getElementById("filter-order").value || "").trim().toLowerCase();
  const cph = (document.getElementById("filter-cph").value || "").trim().toLowerCase();
  const mat = (document.getElementById("filter-mat").value || "").trim().toLowerCase();
  const section = (document.getElementById("filter-section").value || "").trim().toLowerCase();
  const month = document.getElementById("filter-month").value || "";

  if (room) filtered = filtered.filter(d => (d.Room || "").toString().toLowerCase().includes(room));
  if (order) filtered = filtered.filter(d => (d.Order || "").toString().toLowerCase().includes(order));
  if (cph) filtered = filtered.filter(d => (d.CPH || "").toString().toLowerCase().includes(cph));
  if (mat) filtered = filtered.filter(d => (d.MAT || "").toString().toLowerCase().includes(mat));
  if (section) filtered = filtered.filter(d => (d.Section || "").toString().toLowerCase().includes(section));
  if (month) filtered = filtered.filter(d => (d.Month || "") === month);

  renderTable(filtered);
}
function resetFilter() {
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section","filter-month"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  renderTable(mergedData);
}

// ---------------- Populate Month filter dropdown -----------------
function populateMonthFilter() {
  const select = document.getElementById("filter-month");
  if (!select) return;
  const monthsSet = new Set(mergedData.map(r => r.Month).filter(m => m && m.trim() !== ""));
  select.innerHTML = "";
  const optEmpty = document.createElement("option");
  optEmpty.value = "";
  optEmpty.textContent = "-- All Months --";
  select.appendChild(optEmpty);
  const monthOrder = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  monthOrder.forEach(m => {
    if (monthsSet.has(m)) {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      select.appendChild(opt);
    }
  });
}

// ---------------- Add Orders manual -----------------
function addOrders() {
  const input = (document.getElementById("add-order-input").value || "").trim();
  const statusEl = document.getElementById("add-order-status");
  if (!input) {
    if (statusEl) { statusEl.textContent = "Masukkan order dulu ya bro!"; statusEl.style.color = "red"; }
    return;
  }
  const orders = input.split(/[\s,]+/).filter(o => o.length > 0);
  let added = 0;
  orders.forEach(o => {
    if (mergedData.find(r => r.Order === o)) return;
    const iw = iw39Data.find(r => {
      const v = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
      return v === o;
    });
    if (iw) {
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
  populateMonthFilter();
  renderTable(mergedData);
}

// ---------------- Export merged to XLSX -----------------
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

// ---------------- Export merged to JSON -----------------
function exportMergedToJSON() {
  if (!mergedData || mergedData.length === 0) {
    alert("Tidak ada data untuk diexport.");
    return;
  }
  const payload = {
    mergedData: mergedData.map(r => {
      const c = { ...r };
      delete c._IW39_totalPlan;
      delete c._IW39_totalActual;
      return c;
    }),
    timestamp: new Date().toISOString()
  };
  const jsonStr = JSON.stringify(payload, null, 2);
  downloadFile("merged_data_backup.json", jsonStr, "application/json");
}

// ---------------- Download helper -----------------
function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

// ---------------- Load JSON backup -----------------
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
      populateMonthFilter();
      renderTable(mergedData);
      alert("Backup JSON dimuat.");
    } catch (err) {
      alert("Gagal memuat JSON: " + err.message);
    }
  };
  reader.readAsText(file);
}

// ---------------- UI wiring -----------------
function wireUp() {
  // menu switching
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

  // Upload local file (existing upload control)
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) {
    uploadBtn.addEventListener("click", async () => {
      const sel = document.getElementById("file-select").value;
      const f = document.getElementById("file-input").files[0];
      if (!f) { alert("Pilih file terlebih dahulu"); return; }
      document.getElementById("upload-status").textContent = `Parsing ${f.name} ...`;
      try {
        const json = await parseFile(f);
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

  // GitHub load buttons (you can make a small UI to call these)
  const ghLoadBtn = document.getElementById("gh-load-btn");
  if (ghLoadBtn) {
    ghLoadBtn.addEventListener("click", async () => {
      const owner = prompt("GitHub owner (username/org):");
      const repo = prompt("Repo name:");
      const branch = prompt("Branch (default: main):", "main");
      if (!owner || !repo) return alert("Owner & repo required");
      try {
        document.getElementById("upload-status").textContent = "Loading IW39 from GitHub...";
        const iw39json = await loadJSONFromGitHub(owner, repo, branch, "excel/IW39.json");
        iw39Data = iw39json;
        document.getElementById("upload-status").textContent = `IW39 loaded (${iw39json.length} rows)`;

        document.getElementById("upload-status").textContent = "Loading SUM57 from GitHub...";
        const sum57json = await loadJSONFromGitHub(owner, repo, branch, "excel/SUM57.json");
        sum57Data = sum57json;
        document.getElementById("upload-status").textContent = `SUM57 loaded (${sum57json.length} rows)`;

        document.getElementById("upload-status").textContent = "Loading Planning from GitHub...";
        const planningjson = await loadJSONFromGitHub(owner, repo, branch, "excel/Planning.json");
        planningData = planningjson;
        document.getElementById("upload-status").textContent = `Planning loaded (${planningjson.length} rows)`;

        // Do similar for data1Data, data2Data, budgetData if needed

        mergeData();
        renderTable(mergedData);
        populateMonthFilter();
      } catch (e) {
        alert("Gagal load data dari GitHub: " + e.message);
      }
    });
  }

  // Save button
  const saveBtn = document.getElementById("save-btn");
  if (saveBtn) {
    // kalau mau save ke JSON
    saveBtn.removeEventListener("click", exportMergedToExcel);
    saveBtn.addEventListener("click", exportMergedToJSON);
  }

  // Filter inputs event binding
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section","filter-month"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", filterData);
  });

  // Reset filter button
  const resetBtn = document.getElementById("reset-filter-btn");
  if (resetBtn) resetBtn.addEventListener("click", resetFilter);

  // Add orders button
  const addOrderBtn = document.getElementById("add-order-btn");
  if (addOrderBtn) addOrderBtn.addEventListener("click", addOrders);

  // Load local backup JSON
  const backupInput = document.getElementById("backup-json-input");
  if (backupInput) {
    backupInput.addEventListener("change", e => {
      if (e.target.files.length > 0) loadJSONBackupFile(e.target.files[0]);
    });
  }
}

// ---------------- Merge Data -----------------
function mergeData() {
  mergedData = [];

  iw39Data.forEach(iw => {
    // Find matching data from other datasets, simplified example:
    const orderKey = (getVal(iw, ["Order","Order No","Key"]) || "").toString().trim();
    const plan = planningData.find(p => (getVal(p, ["Order","Order No","Key"]) || "").toString().trim() === orderKey) || {};
    const sum57 = sum57Data.find(s => (getVal(s, ["Order","Order No","Key"]) || "").toString().trim() === orderKey) || {};

    mergedData.push({
      Room: getVal(iw, ["Room","ROOM","Location"]) || "",
      "Order Type": getVal(iw, ["Order Type","OrderType"]) || "",
      Order: orderKey,
      Description: getVal(iw, ["Description","Desc"]) || "",
      "Created On": getVal(iw, ["Created On","CreatedOn","Tanggal"]) || "",
      "User Status": getVal(iw, ["User Status","UserStatus"]) || "",
      MAT: (getVal(iw, ["MAT","Mat","Material"]) || "").toString().trim(),
      CPH: getVal(sum57, ["CPH"]) || "",
      Section: getVal(sum57, ["Section"]) || "",
      "Status Part": getVal(sum57, ["Status Part"]) || "",
      Aging: getVal(sum57, ["Aging"]) || "",
      Month: getVal(plan, ["Month"]) || "",
      Cost: getVal(sum57, ["Cost"]) || "-",
      Reman: "",
      Include: "-",
      Exclude: "-",
      Planning: getVal(plan, ["Planning","Plan Date"]) || "",
      "Status AMT": getVal(plan, ["Status AMT"]) || ""
    });
  });

  // Apply user edits from localStorage
  let ui = { userEdits: [] };
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) ui = JSON.parse(raw);
  } catch (e) {
    ui = { userEdits: [] };
  }
  ui.userEdits.forEach(edit => {
    const idx = mergedData.findIndex(r => r.Order === edit.Order);
    if (idx !== -1) {
      mergedData[idx] = { ...mergedData[idx], ...edit };
    }
  });
}

// ---------------- Get Value helper -----------------
function getVal(obj, keys) {
  if (!obj) return undefined;
  for (let k of keys) {
    if (obj[k] !== undefined) return obj[k];
  }
  return undefined;
}

// ---------------- Load JSON from GitHub -----------------
async function loadJSONFromGitHub(owner, repo, branch, path) {
  const url = `https://raw.githubusercontent.com/${owner}/${repo}/${branch}/${path}`;
  const resp = await fetch(url);
  if (!resp.ok) throw new Error("HTTP error " + resp.status);
  const json = await resp.json();
  return json;
}

// ---------------- Main startup -----------------
window.addEventListener("DOMContentLoaded", () => {
  wireUp();
  // You can load default or test data here, e.g., from local or GitHub, or leave empty
  // mergeData();
  // renderTable(mergedData);
  // populateMonthFilter();
});
