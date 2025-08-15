/****************************************************
 * Ndarboe.net - FULL script.js (Menu 1–5, revisi LOM)
 * --------------------------------------------------
 * - Upload 7 sumber (IW39, SUM57, Planning, Budget, Data1, Data2, LOM)
 * - Merge & render Lembar Kerja (lookup Month/Cost/Reman dari LOM)
 * - Cost: jika LOM.Cost == 0 → pakai perhitungan (plan-actual)/16500
 * - Format angka (Cost/Include/Exclude) Indonesia, 1 desimal, rata kanan
 * - Filter, Add Order (pindah ke Menu 3), Save/Load JSON (tetap)
 * - Pewarnaan kolom status (tetap)
 * - Format tanggal dd-MMM-yyyy (Created On, Planning) (tetap)
 * - Menu baru: LOM (filter Order + tabel Order, Month, Cost, Reman, Planning, Status)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let lomData = [];
let mergedData = [];

const UI_LS_KEY = "ndarboe_ui_edits_v3";

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupMenu();
  setupButtons();
  renderTable([]);
  updateMonthFilterOptions();
  renderLOMTable(lomData);
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

/* ===================== HELPERS ===================== */
function safe(v) {
  return String(v ?? "").replace(/[<>&"]/g, s => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[s]));
}

function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;
  if (typeof anyDate === "number") return excelDateToJS(anyDate);
  if (anyDate instanceof Date && !isNaN(anyDate)) return anyDate;
  if (typeof anyDate === "string") {
    const s = anyDate.trim();
    if (!s) return null;
    const iso = new Date(s);
    if (!isNaN(iso)) return iso;
    const parts = s.match(/(\d{1,4})/g);
    if (parts && parts.length >= 3) {
      const ampm = /am|pm/i.test(s) ? s.match(/am|pm/i)[0] : "";
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      let year, month, day, hour = 0, min = 0, sec = 0;
      if (p1 <= 12 && p2 <= 31) { month = p1; day = p2; year = p3; }
      else { day = p1; month = p2; year = p3; }
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

function excelDateToJS(serial) {
  if (!serial || isNaN(serial)) return null;
  const utc_days  = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400 * 1000;
  const date_info = new Date(utc_value);
  const fractional_day = serial - Math.floor(serial);
  const totalSeconds = Math.round(86400 * fractional_day);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  date_info.setHours(hours, minutes, seconds, 0);
  return date_info;
}

function formatDateDDMMMYYYY(input) {
  if (input === undefined || input === null || input === "") return "";
  let d = null;
  if (typeof input === "number" || !isNaN(Number(input))) {
    d = excelDateToJS(Number(input));
  } else {
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

function formatDateISO(anyDate) {
  const d = toDateObj(anyDate);
  if (!d || isNaN(d)) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}

function formatNumberID(num) {
  if (num === "-" || num === "" || num === null || num === undefined) return num ?? "-";
  const n = typeof num === "string" ? Number(num.toString().replace(/[^\d.-]/g,"")) : Number(num);
  if (!isFinite(n)) return "-";
  return n.toLocaleString("id-ID", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
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
          sheetName = workbook.SheetNames[0];
          console.warn(`Sheet "${jenis}" tidak ditemukan, pakai sheet pertama: ${sheetName}`);
        }
        const ws = workbook.Sheets[sheetName];
        if (!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan di file.`);
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        resolve(json);
      } catch (err) { reject(err); }
    };
    reader.onerror = err => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

async function handleUpload() {
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const status = document.getElementById("upload-status");
  if (!fileInput || !fileSelect) return;
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
      case "LOM":      lomData = normalizeLOM(json); break;
      default: break;
    }
    status.textContent = `File ${file.name} berhasil diupload sebagai ${jenis} (rows: ${json.length}).`;
    fileInput.value = "";
    if (jenis === "LOM") renderLOMTable(lomData);
  } catch (e) {
    status.textContent = `Error saat membaca file: ${e.message}`;
  }
}

/* ===================== NORMALIZE LOM ===================== */
function normalizeLOM(rows) {
  const mapMonth = m => {
    const t = (m || "").toString().trim().slice(0,3).toLowerCase();
    const dict = { jan:"Jan",feb:"Feb",mar:"Mar",apr:"Apr",may:"May",jun:"Jun",jul:"Jul",aug:"Aug",sep:"Sep",oct:"Oct",nov:"Nov",dec:"Dec" };
    return dict[t] || (m || "");
  };
  return rows.map(r => {
    const order = (r.Order || r.ORDER || r.order || "").toString().trim();
    const month = mapMonth(r.Month || r.MONTH || r.month || "");
    const cost = Number((r.Cost || r.COST || r.cost || "0").toString().replace(/[^\d.-]/g,"")) || 0;
    const reman = (r.Reman || r.REMAN || r.reman || r["Reman Status"] || "").toString();
    let planning = r.Planning || r["Event Start"] || "";
    let status   = r.Status   || r["Status"]      || "";
    if (!planning || !status) {
      const pl = planningData.find(p => (p.Order || "").toString() === order);
      if (pl) {
        if (!planning) planning = pl["Event Start"] || "";
        if (!status)   status   = (pl.Status || "").toString();
      }
    }
    return { Order: order, Month: month, Cost: cost, Reman: reman, Planning: planning, Status: status };
  }).filter(x => x.Order);
}

/* ===================== RENDER LOM TABLE ===================== */
function renderLOMTable(data) {
  const tbody = document.querySelector("#lom-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  data.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${safe(row.Order)}</td>
      <td>${safe(row.Month)}</td>
      <td style="text-align:right;">${formatNumberID(row.Cost)}</td>
      <td>${safe(row.Reman)}</td>
      <td>${safe(formatDateDDMMMYYYY(row.Planning))}</td>
      <td>${safe(row.Status)}</td>
    `;
    tbody.appendChild(tr);
  });
}

/* ===================== MERGE DATA ===================== */
function mergeAllData() {
  if (!iw39Data.length) return [];
  mergedData = iw39Data.map(row => {
    const orderNum = (row.Order || row.ORDER || "").toString().trim();
    const sum57Row = sum57Data.find(s => (s.Order || s.ORDER || "").toString().trim() === orderNum) || {};
    const planRow  = planningData.find(p => (p.Order || "").toString().trim() === orderNum) || {};
    const lomRow   = lomData.find(l => (l.Order || "").toString().trim() === orderNum) || {};

    let month = lomRow.Month || planRow.Month || row.Month || "";
    let cost  = 0;
    if (lomRow && lomRow.Cost && lomRow.Cost !== 0) {
      cost = lomRow.Cost;
    } else {
      // perhitungan lama
      const actual = Number(sum57Row["Actual Cost"] || 0);
      const plan   = Number(sum57Row["Planned Cost"] || 0);
      cost = (plan - actual) / 16500;
    }
    const reman   = lomRow.Reman || planRow.Reman || "";
    const include = Number(sum57Row.Include || 0);
    const exclude = Number(sum57Row.Exclude || 0);
    const planning = lomRow.Planning || planRow.Planning || "";
    const statusAMT = lomRow.Status || planRow.Status || "";

    return {
      Room: row.Room || "",
      OrderType: row["Order Type"] || "",
      Order: orderNum,
      Description: row.Description || "",
      CreatedOn: row["Created On"] || "",
      UserStatus: row["User Status"] || "",
      MAT: row.MAT || "",
      CPH: row.CPH || "",
      Section: row.Section || "",
      StatusPart: row["Status Part"] || "",
      Aging: row.Aging || "",
      Month: month,
      Cost: cost,
      Reman: reman,
      Include: include,
      Exclude: exclude,
      Planning: planning,
      StatusAMT: statusAMT
    };
  });
  return mergedData;
}

/* ===================== RENDER MAIN TABLE ===================== */
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  data.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${safe(row.Room)}</td>
      <td>${safe(row.OrderType)}</td>
      <td>${safe(row.Order)}</td>
      <td>${safe(row.Description)}</td>
      <td>${safe(formatDateDDMMMYYYY(row.CreatedOn))}</td>
      <td>${safe(row.UserStatus)}</td>
      <td>${safe(row.MAT)}</td>
      <td>${safe(row.CPH)}</td>
      <td>${safe(row.Section)}</td>
      <td>${safe(row.StatusPart)}</td>
      <td>${safe(row.Aging)}</td>
      <td>${safe(row.Month)}</td>
      <td style="text-align:right;">${formatNumberID(row.Cost)}</td>
      <td>${safe(row.Reman)}</td>
      <td style="text-align:right;">${formatNumberID(row.Include)}</td>
      <td style="text-align:right;">${formatNumberID(row.Exclude)}</td>
      <td>${safe(formatDateDDMMMYYYY(row.Planning))}</td>
      <td>${safe(row.StatusAMT)}</td>
      <td>
        <button class="action-btn edit-btn" data-order="${safe(row.Order)}">Edit</button>
        <button class="action-btn delete-btn" data-order="${safe(row.Order)}">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/* ===================== FILTER ===================== */
function applyFilter() {
  const room = document.getElementById("filter-room").value.trim().toLowerCase();
  const order = document.getElementById("filter-order").value.trim().toLowerCase();
  const cph = document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat = document.getElementById("filter-mat").value.trim().toLowerCase();
  const section = document.getElementById("filter-section").value.trim().toLowerCase();
  const month = document.getElementById("filter-month").value;

  const filtered = mergedData.filter(row =>
    (!room || (row.Room || "").toLowerCase().includes(room)) &&
    (!order || (row.Order || "").toLowerCase().includes(order)) &&
    (!cph || (row.CPH || "").toLowerCase().includes(cph)) &&
    (!mat || (row.MAT || "").toLowerCase().includes(mat)) &&
    (!section || (row.Section || "").toLowerCase().includes(section)) &&
    (!month || (row.Month || "") === month)
  );
  renderTable(filtered);
}

function resetFilter() {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-month").value = "";
  renderTable(mergedData);
}

function updateMonthFilterOptions() {
  const sel = document.getElementById("filter-month");
  if (!sel) return;
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  sel.innerHTML = `<option value="">-- All --</option>` +
    months.map(m => `<option value="${m}">${m}</option>`).join("");
}

/* ===================== FILTER LOM ===================== */
function applyLOMFilter() {
  const order = document.getElementById("lom-filter-order").value.trim().toLowerCase();
  const filtered = lomData.filter(row =>
    (!order || (row.Order || "").toLowerCase().includes(order))
  );
  renderLOMTable(filtered);
}

function resetLOMFilter() {
  document.getElementById("lom-filter-order").value = "";
  renderLOMTable(lomData);
}

/* ===================== ADD ORDER (PINDAH KE LOM) ===================== */
function addOrderToLOM() {
  const input = document.getElementById("add-order-input-lom");
  if (!input) return;
  const orders = input.value
    .split(/[\s,]+/)
    .map(o => o.trim())
    .filter(o => o.length > 0);

  if (!orders.length) {
    document.getElementById("add-order-status-lom").textContent = "Tidak ada order yang valid.";
    return;
  }

  orders.forEach(orderNum => {
    lomData.push({
      Order: orderNum,
      Month: "",
      Cost: 0,
      Reman: "",
      Planning: "",
      Status: ""
    });
  });

  input.value = "";
  document.getElementById("add-order-status-lom").textContent = "Order berhasil ditambahkan ke LOM.";
  renderLOMTable(lomData);
}

/* ===================== SAVE / LOAD JSON ===================== */
function saveJSON() {
  const blob = new Blob([JSON.stringify(mergedData, null, 2)], {type: "application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "data.json";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function loadJSON(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      if (Array.isArray(data)) {
        mergedData = data;
        renderTable(mergedData);
      } else {
        alert("Format JSON tidak valid.");
      }
    } catch (err) {
      alert("Gagal memuat JSON.");
    }
  };
  reader.readAsText(file);
}

/* ===================== EDIT / DELETE ===================== */
function handleTableClick(evt) {
  const target = evt.target;
  if (target.classList.contains("edit-btn")) {
    startEditRow(target.dataset.order);
  } else if (target.classList.contains("delete-btn")) {
    deleteRow(target.dataset.order);
  } else if (target.classList.contains("save-btn")) {
    saveEditRow(target.dataset.order);
  } else if (target.classList.contains("cancel-btn")) {
    cancelEditRow(target.dataset.order);
  }
}

function startEditRow(orderNum) {
  const tbody = document.querySelector("#output-table tbody");
  const tr = Array.from(tbody.querySelectorAll("tr")).find(row =>
    row.children[2].textContent.trim() === orderNum
  );
  if (!tr) return;

  const rowData = mergedData.find(r => r.Order === orderNum);
  if (!rowData) return;

  const editableCols = ["Month","Cost","Reman"];
  Array.from(tr.children).forEach((td, idx) => {
    const colName = document.querySelector(`#output-table thead tr`).children[idx].textContent;
    if (editableCols.includes(colName)) {
      td.innerHTML = `<input type="text" value="${rowData[colName]}" data-field="${colName}" />`;
    }
  });

  tr.children[18].innerHTML = `
    <button class="action-btn save-btn" data-order="${orderNum}">Save</button>
    <button class="action-btn cancel-btn" data-order="${orderNum}">Cancel</button>
  `;
}

function saveEditRow(orderNum) {
  const tbody = document.querySelector("#output-table tbody");
  const tr = Array.from(tbody.querySelectorAll("tr")).find(row =>
    row.children[2].textContent.trim() === orderNum
  );
  if (!tr) return;

  const rowData = mergedData.find(r => r.Order === orderNum);
  if (!rowData) return;

  const inputs = tr.querySelectorAll("input[data-field]");
  inputs.forEach(input => {
    const field = input.dataset.field;
    let val = input.value;
    if (field === "Cost" || field === "Include" || field === "Exclude") {
      val = parseFloat(val) || 0;
    }
    rowData[field] = val;
  });

  renderTable(mergedData);
}

function cancelEditRow(orderNum) {
  renderTable(mergedData);
}

function deleteRow(orderNum) {
  mergedData = mergedData.filter(r => r.Order !== orderNum);
  renderTable(mergedData);
}

/* ===================== CLEAR ALL DATA ===================== */
function clearAllData() {
  iw39Data = [];
  sum57Data = [];
  planningData = [];
  budgetData = [];
  data1 = [];
  data2 = [];
  lomData = [];
  mergedData = [];
  renderTable([]);
  renderLOMTable([]);
}

/* ===================== SETUP BUTTONS ===================== */
function setupButtons() {
  document.getElementById("upload-btn").addEventListener("click", handleUpload);
  document.getElementById("clear-files-btn").addEventListener("click", clearAllData);

  document.getElementById("filter-btn").addEventListener("click", applyFilters);
  document.getElementById("reset-btn").addEventListener("click", resetFilters);
  document.getElementById("refresh-btn").addEventListener("click", refreshData);

  document.getElementById("lom-filter-btn").addEventListener("click", applyLOMFilter);
  document.getElementById("lom-reset-btn").addEventListener("click", resetLOMFilter);
  document.getElementById("lom-refresh-btn").addEventListener("click", refreshData);

  const addOrderBtnLom = document.getElementById("add-order-btn-lom");
  if (addOrderBtnLom) {
    addOrderBtnLom.addEventListener("click", addOrderToLOM);
  }

  document.getElementById("save-btn").addEventListener("click", saveJSON);
  document.getElementById("load-btn").addEventListener("change", loadJSON);

  document.querySelector("#output-table tbody").addEventListener("click", handleTableClick);
}

/* ===================== STATUS CHIP COLORS ===================== */
function styleStatusChips() {
  document.querySelectorAll("#output-table tbody tr").forEach(tr => {
    const statusCell = tr.children[5];
    if (statusCell) {
      const status = statusCell.textContent.trim().toUpperCase();
      if (status === "OUTS") {
        statusCell.style.background = "yellow";
        statusCell.style.color = "black";
      } else if (status === "RELE") {
        statusCell.style.background = "green";
        statusCell.style.color = "white";
      } else if (status === "PROG") {
        statusCell.style.background = "orange";
        statusCell.style.color = "black";
      }
    }
  });
}

/* ===================== INIT ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupButtons();
  styleStatusChips();
});
