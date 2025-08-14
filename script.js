/****************************************************
 * Ndarboe.net - FULL script.js (menu 1–4, >500 lines)
 * --------------------------------------------------
 * - Upload 6 sumber (IW39, SUM57, Planning, Budget, Data1, Data2)
 * - Merge & render Lembar Kerja
 * - Filter, Add Order, Edit/Save/Delete, Save/Load JSON
 * - Pewarnaan kolom status
 * - Format tanggal dd-MMM-yyyy (Created On, Planning)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];

const UI_LS_KEY = "ndarboe_ui_edits_v2";

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupMenu();
  setupButtons();
  renderTable([]);        // kosong dulu
  updateMonthFilterOptions();
});

/* ===================== MENU HANDLER ===================== */
function setupMenu() {
  const menuItems = document.querySelectorAll(".sidebar .menu-item");
  const contentSections = document.querySelectorAll(".content-section");
  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      item.classList.add("active");
      const menuId = item.dataset.menu;
      contentSections.forEach(sec => {
        if (sec.id === menuId) sec.classList.add("active");
        else sec.classList.remove("active");
      });
    });
  });
}

/* ===================== HELPERS: DATE PARSING/FORMATTING ===================== */
/**
 * Parse apa pun (Excel serial number / string MM/DD/YYYY / string MM/DD/YYYY HH:mm:ss AM/PM /
 * ISO yyyy-mm-dd) ke Date object (valid) atau null
 */
function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;

  // Numeric Excel serial
  if (typeof anyDate === "number") {
    const dec = XLSX && XLSX.SSF && XLSX.SSF.parse_date_code
      ? XLSX.SSF.parse_date_code(anyDate)
      : null;
    if (dec && typeof dec === "object") {
      return new Date(dec.y, (dec.m || 1) - 1, dec.d || 1, dec.H || 0, dec.M || 0, dec.S || 0);
    }
  }

  if (anyDate instanceof Date && !isNaN(anyDate)) {
    return anyDate;
  }

  if (typeof anyDate === "string") {
    const s = anyDate.trim();
    if (!s) return null;

    // ISO yyyy-mm-dd / yyyy-mm-ddTHH:mm:ss
    const iso = new Date(s);
    if (!isNaN(iso)) return iso;

    // Coba pattern US "M/D/YYYY" atau "M/D/YYYY HH:mm:ss AM"
    // Ambil digit
    const parts = s.match(/(\d{1,4})/g);
    if (parts && parts.length >= 3) {
      // deteksi AM/PM
      const ampm = /am|pm/i.test(s) ? s.match(/am|pm/i)[0] : "";
      // heuristic: kalau 1st <=12 dan 2nd <=31 → asumsikan M/D/YYYY
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      let year, month, day, hour = 0, min = 0, sec = 0;
      if (p1 <= 12 && p2 <= 31) {
        month = p1; day = p2; year = p3;
      } else {
        // fallback: D/M/YYYY
        day = p1; month = p2; year = p3;
      }
      if (parts.length >= 5) {
        hour = parseInt(parts[3], 10);
        min  = parseInt(parts[4], 10);
        if (parts.length >= 6) sec = parseInt(parts[5], 10) || 0;
        if (ampm) {
          if (/pm/i.test(ampm) && hour < 12) hour += 12;
          if (/am/i.test(ampm) && hour === 12) hour = 0;
        }
      }
      const d = new Date(year, (month || 1) - 1, day || 1, hour, min, sec);
      if (!isNaN(d)) return d;
    }
  }

  return null;
}

/** format dd-MMM-yyyy */


function formatDateDDMMMYYYY(input) {
  if (input === undefined || input === null || input === "") return "";
  let d = null;
  if (typeof input === "number") d = excelDateToJS(input);
  else {
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

/** format value="yyyy-mm-dd" untuk input[type=date] */
function formatDateISO(anyDate) {
  const d = toDateObj(anyDate);
  if (!d || isNaN(d)) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}

/* ===================== UPLOAD & PARSE EXCEL ===================== */
async function parseFile(file, jenis) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        let sheetName = "";
        if (workbook.SheetNames.includes(jenis)) {
          sheetName = jenis;
        } else {
          sheetName = workbook.SheetNames[0]; // fallback
          console.warn(`Sheet "${jenis}" tidak ditemukan, pakai sheet pertama: ${sheetName}`);
        }
        const ws = workbook.Sheets[sheetName];
        if (!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan di file.`);

        // important: raw:false agar tanggal string tetap bisa di-parse manual
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = err => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

async function handleUpload() {
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const status = document.getElementById("upload-status");

  if (!fileInput.files.length) {
    alert("Pilih file terlebih dahulu.");
    return;
  }
  const file = fileInput.files[0];
  const jenis = fileSelect.value;

  status.textContent = `Memproses file ${file.name} sebagai ${jenis}...`;
  try {
    const json = await parseFile(file, jenis);
    switch (jenis) {
      case "IW39":     iw39Data = json; break;
      case "SUM57":    sum57Data = json; break;
      case "Planning": planningData = json; break;
      case "Data1":    data1Data = json; break;
      case "Data2":    data2Data = json; break;
      case "Budget":   budgetData = json; break;
      default: break;
    }
    status.textContent = `File ${file.name} berhasil diupload sebagai ${jenis} (rows: ${json.length}).`;

    // clear file input agar bisa upload file yang sama lagi
    fileInput.value = "";

  } catch (e) {
    status.textContent = `Error saat membaca file: ${e.message}`;
  }
}

/* ===================== MERGE ===================== */
function mergeData() {
  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }

  // base rows dari IW39
  mergedData = iw39Data.map(row => ({
    Room: (row.Room || "").toString(),
    "Order Type": (row["Order Type"] || "").toString(),
    Order: (row.Order || "").toString(),
    Description: (row.Description || "").toString(),
    "Created On": row["Created On"] || "",          // raw, nanti format saat render
    "User Status": (row["User Status"] || "").toString(),
    MAT: (row.MAT || "").toString(),
    CPH: "",
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: (row.Month || "").toString(),
    Cost: "-",
    Reman: (row.Reman || "").toString(),
    Include: "-",
    Exclude: "-",
    Planning: "",                // Event Start (Planning)
    "Status AMT": ""             // Status (Planning)
  }));

  // CPH: Description startsWith JR → "External Job", else lookup Data2 by MAT
  mergedData.forEach(md => {
    if ((md.Description || "").trim().toUpperCase().startsWith("JR")) {
      md.CPH = "External Job";
    } else {
      const d2 = data2Data.find(d => (d.MAT || "").toString().trim() === md.MAT.trim());
      md.CPH = d2 ? (d2.CPH || "").toString() : "";
    }
  });

  // Section: Data1 via Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => (d.Room || "").toString().trim() === md.Room.trim());
    md.Section = d1 ? (d1.Section || "").toString() : "";
  });

  // SUM57: Aging & Status Part via Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => (s.Order || "").toString() === md.Order);
    if (s57) {
      md.Aging = (s57.Aging || "").toString();
      md["Status Part"] = (s57["Part Complete"] || "").toString();
    }
  });

  // Planning: Planning(Event Start) & Status AMT(Status) by Order
  mergedData.forEach(md => {
    const pl = planningData.find(p => (p.Order || "").toString() === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] || "";           // raw, format saat render
      md["Status AMT"] = (pl.Status || "").toString();
    }
  });

  // Hitung Cost/Include/Exclude dari IW39 plan vs actual
  mergedData.forEach(md => {
    const src = iw39Data.find(i => (i.Order || "").toString() === md.Order);
    if (!src) return;

    const plan = parseFloat((src["Total sum (plan)"] || "").toString().replace(/,/g,"")) || 0;
    const actual = parseFloat((src["Total sum (actual)"] || "").toString().replace(/,/g,"")) || 0;
    let cost = (plan - actual) / 16500;
    if (!isFinite(cost) || cost < 0) {
      md.Cost = "-";
      md.Include = "-";
      md.Exclude = md["Order Type"] === "PM38" ? "-" : "-";
    } else {
      // 2 decimal
      const costStr = cost.toFixed(2);
      md.Cost = costStr;

      const isReman = (md.Reman || "").toLowerCase().includes("reman");
      const includeNum = isReman ? (cost * 0.25) : cost;
      md.Include = includeNum.toFixed(2);

      md.Exclude = (md["Order Type"] === "PM38") ? "-" : md.Include;
    }
  });

  // Restore user edits
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      if (saved && Array.isArray(saved.userEdits)) {
        saved.userEdits.forEach(edit => {
          const idx = mergedData.findIndex(r => r.Order === edit.Order);
          if (idx !== -1) {
            mergedData[idx] = { ...mergedData[idx], ...edit };
          }
        });
      }
    }
  } catch {}

  updateMonthFilterOptions();
}

/* ===================== RENDER TABLE ===================== */

function renderTable(data = [], tableId = "output-table") {
  const tbody = document.querySelector(`#${tableId} tbody`);
  if (!tbody) return;

  tbody.innerHTML = "";
  data.forEach((row, index) => {
    const tr = document.createElement("tr");
    Object.values(row).forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.appendChild(td);
    });

    // Kolom action
    const actionTd = document.createElement("td");
    actionTd.innerHTML = `
      <button class="action-btn edit-btn" data-index="${index}">Edit</button>
      <button class="action-btn delete-btn" data-index="${index}">Delete</button>
    `;
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });

  attachTableEvents();
}

/* ===================== ATTACH TABLE EVENTS ===================== */
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.addEventListener("click", function () {
      const tr = this.closest("tr");
      const tds = tr.querySelectorAll("td");

      const currentMonth = tds[11].textContent.trim();
      const currentCost  = tds[12].textContent.trim();
      const currentReman = tds[13].textContent.trim();

      const monthOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        .map(m => `<option value="${m}" ${m===currentMonth?"selected":""}>${m}</option>`).join("");
      tds[11].innerHTML = `<select class="edit-month">${monthOptions}</select>`;

      tds[12].innerHTML = `<input type="number" class="edit-cost" value="${currentCost}" style="width:80px;text-align:right;">`;

      tds[13].innerHTML = `
        <select class="edit-reman">
          <option value="Reman" ${currentReman==="Reman"?"selected":""}>Reman</option>
          <option value="-" ${currentReman==="-"?"selected":""}>-</option>
        </select>`;

      this.outerHTML = `<button class="action-btn save-btn" data-index="${btn.dataset.index}">Save</button>
                        <button class="action-btn cancel-btn">Cancel</button>`;

      tr.querySelector(".save-btn").addEventListener("click", function () {
        const index = this.dataset.index;
        data[index][11] = tr.querySelector(".edit-month").value;
        data[index][12] = tr.querySelector(".edit-cost").value;
        data[index][13] = tr.querySelector(".edit-reman").value;
        renderTable();
      });

      tr.querySelector(".cancel-btn").addEventListener("click", function () {
        renderTable();
      });
    });
  });

  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.addEventListener("click", function () {
      const index = this.dataset.index;
      if (confirm("Yakin mau hapus data ini?")) {
        data.splice(index, 1);
        renderTable();
      }
    });
  });
}

// Pastikan fungsi dipanggil setelah DOM siap
document.addEventListener("DOMContentLoaded", () => {
  renderTable();
});
/* ===================== CELL COLORING ===================== */
function asColoredStatusUser(val) {
  const v = (val || "").toString().toUpperCase();
  let bg = "", fg = "";
  if (v === "OUTS") { bg = "#ffeb3b"; fg = "#000"; }
  else if (v === "RELE") { bg = "#2e7d32"; fg = "#fff"; }
  else if (v === "PROG") { bg = "#1976d2"; fg = "#fff"; }
  else if (v === "COMP") { bg = "#000"; fg = "#fff"; }
  else return safe(val);
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${safe(val)}</span>`;
}

function asColoredStatusPart(val) {
  const s = (val || "").toString().toLowerCase();
  if (s === "complete") {
    return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  }
  if (s === "not complete") {
    return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  }
  return safe(val);
}

function asColoredStatusAMT(val) {
  const v = (val || "").toString().toUpperCase();
  if (v === "O") {
    return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#ffeb3b;color:#000;">${safe(val)}</span>`;
  }
  if (v === "IP") {
    return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  }
  if (v === "YTS") {
    return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  }
  return safe(val);
}

/* ===================== EDIT / DELETE ===================== */
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.onclick = () => startEdit(btn.dataset.order);
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.onclick = () => deleteOrder(btn.dataset.order);
  });
}

function startEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;
  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];
  const row = mergedData[rowIndex];

  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort();
  const monthOptions = [`<option value="">--Select Month--</option>`, ...months.map(m => `<option value="${m}">${m}</option>`)].join("");

  tr.innerHTML = `
    <td><input type="text" value="${safe(row.Room)}" data-field="Room"/></td>
    <td><input type="text" value="${safe(row["Order Type"])}" data-field="Order Type"/></td>
    <td>${safe(row.Order)}</td>
    <td><input type="text" value="${safe(row.Description)}" data-field="Description"/></td>
    <td><input type="date" value="${formatDateISO(row["Created On"])}" data-field="Created On"/></td>
    <td><input type="text" value="${safe(row["User Status"])}" data-field="User Status"/></td>
    <td><input type="text" value="${safe(row.MAT)}" data-field="MAT"/></td>
    <td><input type="text" value="${safe(row.CPH)}" data-field="CPH"/></td>
    <td><input type="text" value="${safe(row.Section)}" data-field="Section"/></td>
    <td><input type="text" value="${safe(row["Status Part"])}" data-field="Status Part"/></td>
    <td><input type="text" value="${safe(row.Aging)}" data-field="Aging"/></td>
    <td>
      <select data-field="Month">${monthOptions}</select>
    </td>
    <td><input type="text" value="${safe(row.Cost)}" data-field="Cost" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${safe(row.Reman)}" data-field="Reman"/></td>
    <td><input type="text" value="${safe(row.Include)}" data-field="Include" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${safe(row.Exclude)}" data-field="Exclude" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="date" value="${formatDateISO(row.Planning)}" data-field="Planning"/></td>
    <td><input type="text" value="${safe(row["Status AMT"])}" data-field="Status AMT"/></td>
    <td>
      <button class="action-btn save-btn" data-order="${safe(order)}">Save</button>
      <button class="action-btn cancel-btn" data-order="${safe(order)}">Cancel</button>
    </td>
  `;

  // Set selected month
  const monthSel = tr.querySelector("select[data-field='Month']");
  monthSel.value = row.Month || "";

  tr.querySelector(".save-btn").onclick = () => saveEdit(order);
  tr.querySelector(".cancel-btn").onclick = () => cancelEdit();
}

function cancelEdit() {
  renderTable(mergedData);
}

function saveEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;
  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];

  const inputs = tr.querySelectorAll("input[data-field], select[data-field]");
  inputs.forEach(input => {
    const field = input.dataset.field;
    let val = input.value;
    mergedData[rowIndex][field] = val;
  });

  // persist
  saveUserEdits();

  // re-merge utk kalkulasi & render ulang
  mergeData();
  renderTable(mergedData);
}

function deleteOrder(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx === -1) return;
  if (!confirm(`Hapus data order ${order} ?`)) return;
  mergedData.splice(idx, 1);
  saveUserEdits();
  renderTable(mergedData);
}

/* ===================== FILTERS ===================== */
function filterData() {
  const roomFilter    = (document.getElementById("filter-room").value || "").trim().toLowerCase();
  const orderFilter   = (document.getElementById("filter-order").value || "").trim().toLowerCase();
  const cphFilter     = (document.getElementById("filter-cph").value || "").trim().toLowerCase();
  const matFilter     = (document.getElementById("filter-mat").value || "").trim().toLowerCase();
  const sectionFilter = (document.getElementById("filter-section").value || "").trim().toLowerCase();
  const monthFilter   = (document.getElementById("filter-month").value || "").trim().toLowerCase();

  const filtered = mergedData.filter(row => {
    if (roomFilter && !String(row.Room).toLowerCase().includes(roomFilter)) return false;
    if (orderFilter && !String(row.Order).toLowerCase().includes(orderFilter)) return false;
    if (cphFilter && !String(row.CPH).toLowerCase().includes(cphFilter)) return false;
    if (matFilter && !String(row.MAT).toLowerCase().includes(matFilter)) return false;
    if (sectionFilter && !String(row.Section).toLowerCase().includes(sectionFilter)) return false;
    if (monthFilter && String(row.Month).toLowerCase() !== monthFilter) return false;
    return true;
  });

  renderTable(filtered);
}

function resetFilters() {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-month").value = "";
  renderTable(mergedData);
}

function updateMonthFilterOptions() {
  const monthSelect = document.getElementById("filter-month");
  if (!monthSelect) return;
  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort();
  monthSelect.innerHTML = `<option value="">-- All --</option>` + months.map(m => `<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDERS ===================== */
function addOrders() {
  const input = document.getElementById("add-order-input");
  const status = document.getElementById("add-order-status");
  const text = (input.value || "").trim();
  if (!text) {
    alert("Masukkan Order terlebih dahulu.");
    return;
  }
  const orders = text.split(/[\s,]+/).filter(Boolean);
  let added = 0;
  orders.forEach(o => {
    if (!mergedData.some(r => r.Order === o)) {
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
  if (added) {
    saveUserEdits();
    renderTable(mergedData);
    status.textContent = `${added} Order berhasil ditambahkan.`;
  } else {
    status.textContent = "Order sudah ada di data.";
  }
  input.value = "";
}

/* ===================== SAVE / LOAD JSON (export/import) ===================== */
function saveToJSON() {
  if (!mergedData.length) {
    alert("Tidak ada data untuk disimpan.");
    return;
  }
  const dataStr = JSON.stringify({ mergedData, timestamp: new Date().toISOString() }, null, 2);
  const blob = new Blob([dataStr], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function loadFromJSON(file) {
  try {
    const text = await file.text();
    const obj = JSON.parse(text);
    if (obj.mergedData && Array.isArray(obj.mergedData)) {
      mergedData = obj.mergedData;
      renderTable(mergedData);
      updateMonthFilterOptions();
      alert("Data berhasil dimuat dari JSON.");
    } else {
      alert("File JSON tidak valid.");
    }
  } catch (e) {
    alert("Gagal membaca file JSON: " + e.message);
  }
}

/* ===================== USER EDITS PERSISTENCE ===================== */
function saveUserEdits() {
  try {
    const userEdits = mergedData.map(item => ({
      Order: item.Order,
      Room: item.Room,
      "Order Type": item["Order Type"],
      Description: item.Description,
      "Created On": item["Created On"],
      "User Status": item["User Status"],
      MAT: item.MAT,
      CPH: item.CPH,
      Section: item.Section,
      "Status Part": item["Status Part"],
      Aging: item.Aging,
      Month: item.Month,
      Cost: item.Cost,
      Reman: item.Reman,
      Include: item.Include,
      Exclude: item.Exclude,
      Planning: item.Planning,
      "Status AMT": item["Status AMT"]
    }));
    localStorage.setItem(UI_LS_KEY, JSON.stringify({ userEdits }));
  } catch {}
}

/* ===================== CLEAR ALL ===================== */
function clearAllData() {
  if (!confirm("Yakin ingin menghapus semua data yang telah diupload?")) return;
  iw39Data = [];
  sum57Data = [];
  planningData = [];
  data1Data = [];
  data2Data = [];
  budgetData = [];
  mergedData = [];
  renderTable([]);
  document.getElementById("upload-status").textContent = "Data dihapus.";
  updateMonthFilterOptions();
}

/* ===================== BUTTON WIRING ===================== */
function setupButtons() {
  // Upload
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) uploadBtn.onclick = handleUpload;

  const clearBtn = document.getElementById("clear-files-btn");
  if (clearBtn) clearBtn.onclick = clearAllData;

  // Lembar Kerja
  const refreshBtn = document.getElementById("refresh-btn");
  if (refreshBtn) refreshBtn.onclick = () => { mergeData(); renderTable(mergedData); };

  const filterBtn = document.getElementById("filter-btn");
  if (filterBtn) filterBtn.onclick = filterData;

  const resetBtn = document.getElementById("reset-btn");
  if (resetBtn) resetBtn.onclick = resetFilters;

  const saveBtn = document.getElementById("save-btn");
  if (saveBtn) saveBtn.onclick = saveToJSON;

  const loadBtn = document.getElementById("load-btn");
  if (loadBtn) {
    loadBtn.onclick = () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "application/json";
      input.onchange = () => {
        if (input.files.length) loadFromJSON(input.files[0]);
      };
      input.click();
    };
  }

  const addOrderBtn = document.getElementById("add-order-btn");
  if (addOrderBtn) addOrderBtn.onclick = addOrders;
}








