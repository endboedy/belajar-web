/****************************************************
 * Ndarboe.net - FULL script.js (Menu 1–5, revisi LOM pindah Add Order ke Menu 3)
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
/** Parse apapun ke Date object atau null */
function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;

  if (typeof anyDate === "number") {
    return excelDateToJS(anyDate);
  }
  if (anyDate instanceof Date && !isNaN(anyDate)) {
    return anyDate;
  }
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

/** Excel serial → JS Date */
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

/** Format → dd-MMM-yyyy */
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

/** Format → yyyy-mm-dd untuk input[type=date] */
function formatDateISO(anyDate) {
  const d = toDateObj(anyDate);
  if (!d || isNaN(d)) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}

/** Format angka Indonesia (1 desimal, titik ribuan) */
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
          sheetName = workbook.SheetNames[0]; // fallback
          console.warn(`Sheet "${jenis}" tidak ditemukan, pakai sheet pertama: ${sheetName}`);
        }
        const ws = workbook.Sheets[sheetName];
        if (!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan di file.`);

        // raw:false agar tanggal string tetap bisa di-parse manual
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
      case "LOM":      lomData = normalizeLOM(json); break;   // ⬅️ baru
      default: break;
    }
    status.textContent = `File ${file.name} berhasil diupload sebagai ${jenis} (rows: ${json.length}).`;
    fileInput.value = "";

    // auto re-render LOM table kalau upload LOM
    if (jenis === "LOM") renderLOMTable(lomData);

  } catch (e) {
    status.textContent = `Error saat membaca file: ${e.message}`;
  }
}

/* ===================== NORMALIZE LOM ===================== */
/** 
 * Normalisasi data LOM agar kolom-kolom penting tersedia:
 * - Order
 * - Month (Jan, Feb, ... Dec)
 * - Cost (number)
 * - Reman (string)
 * - Planning (ambil dari Planning file via Order kalau kolom ini kosong)
 * - Status   (ambil dari Planning file via Order kalau kolom ini kosong)
 */
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

/* ===================== MERGE DATA ===================== */
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
    "Created On": row["Created On"] || "",
    "User Status": (row["User Status"] || "").toString(),
    MAT: (row.MAT || "").toString(),
    CPH: "",
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: (row.Month || "").toString(), // default (bisa ditimpa LOM)
    Cost: "-",                            // default (akan dihitung/lookup)
    Reman: (row.Reman || "").toString(),  // default (bisa ditimpa LOM)
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
      md.Planning = pl["Event Start"] || "";
      md["Status AMT"] = (pl.Status || "").toString();
    }
  });

  // Hitung Cost default (plan-actual)/16500
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
      md.Cost = Number(cost.toFixed(1)); // simpan numeric 1 desimal (render akan format ID)
      const isReman = (md.Reman || "").toLowerCase().includes("reman");
      const includeNum = isReman ? (cost * 0.25) : cost;
      md.Include = Number(includeNum.toFixed(1));
      md.Exclude = (md["Order Type"] === "PM38") ? "-" : Number(includeNum.toFixed(1));
    }
  });

  // Lookup dari LOM (Month / Reman / Cost override jika ≠ 0)
  if (lomData.length) {
    const lomMap = {};
    lomData.forEach(l => { if (l && l.Order) lomMap[l.Order] = l; });

    mergedData.forEach(md => {
      const lom = lomMap[md.Order];
      if (!lom) return;

      // Month & Reman langsung ambil dari LOM jika ada
      if (lom.Month) md.Month = lom.Month;
      if (lom.Reman) md.Reman = lom.Reman;

      // Cost: jika LOM.Cost bukan 0 → override; jika 0 → biarkan perhitungan lama
      const lomCostNum = Number(lom.Cost) || 0;
      if (lomCostNum !== 0) {
        md.Cost = Number(lomCostNum.toFixed(1));
        // Recalculate Include/Exclude berdasar Cost baru
        const isReman = (md.Reman || "").toLowerCase().includes("reman");
        const includeNum = isReman ? (lomCostNum * 0.25) : lomCostNum;
        md.Include = Number(includeNum.toFixed(1));
        md.Exclude = (md["Order Type"] === "PM38") ? "-" : Number(includeNum.toFixed(1));
      }
    });
  }

  // Restore user edits lama (jika ada) — tetapi hanya untuk Month/Cost/Reman (kompat)
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      if (saved && Array.isArray(saved.userEdits)) {
        const editMap = {};
        saved.userEdits.forEach(edit => { editMap[edit.Order] = edit; });
        mergedData.forEach(md => {
          const e = editMap[md.Order];
          if (e) {
            if (e.Month)  md.Month  = e.Month;
            if (e.Reman)  md.Reman  = e.Reman;
            if (e.Cost !== undefined && e.Cost !== "-" && e.Cost !== "") {
              const n = Number(e.Cost);
              if (isFinite(n)) {
                md.Cost = Number(n.toFixed(1));
                const isReman = (md.Reman || "").toLowerCase().includes("reman");
                const includeNum = isReman ? (n * 0.25) : n;
                md.Include = Number(includeNum.toFixed(1));
                md.Exclude = (md["Order Type"] === "PM38") ? "-" : Number(includeNum.toFixed(1));
              }
            }
          }
        });
      }
    }
  } catch {}

  updateMonthFilterOptions();
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender = mergedData) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) {
    console.warn("Tabel #output-table tidak ditemukan.");
    return;
  }
  tbody.innerHTML = "";

  dataToRender.forEach((row) => {
    const tr = document.createElement("tr");

    const columns = [
      "Room", "Order Type", "Order", "Description", "Created On",
      "User Status", "MAT", "CPH", "Section", "Status Part", "Aging",
      "Month", "Cost", "Reman", "Include", "Exclude", "Planning", "Status AMT"
    ];

    columns.forEach(col => {
      const td = document.createElement("td");

      if (col === "Created On" || col === "Planning") {
        td.textContent = formatDateDDMMMYYYY(row[col]);
      } else if (["Cost", "Include", "Exclude"].includes(col)) {
        td.textContent = formatNumberID(row[col]);
        td.style.textAlign = "right";
      } else {
        td.textContent = row[col] ?? "";
      }

      // style opsional untuk status chips
      if (col === "User Status") {
        // bisa diterapkan nanti kalau ingin chip warna
      }
      tr.appendChild(td);
    });

    // ⛔ Tidak ada kolom Action lagi (hapus edit/delete)
    tbody.appendChild(tr);
  });
}

/* ===================== MENU LOM (LEMBAR ORDER MONTHLY) ===================== */
function renderLOMTable(rows = lomData) {
  const tbody = document.querySelector("#lom-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  rows.forEach(r => {
    const tr = document.createElement("tr");

    // === Order ===
    const tdOrder = document.createElement("td");
    tdOrder.textContent = r.Order ?? "";
    tr.appendChild(tdOrder);

    // === Month (dropdown) ===
    const tdMonth = document.createElement("td");
    const monthSelect = document.createElement("select");
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    months.forEach(m => {
      const option = document.createElement("option");
      option.value = m;
      option.textContent = m;
      if (r.Month === m) option.selected = true;
      monthSelect.appendChild(option);
    });
    // bind perubahan ke lomData
    monthSelect.addEventListener("change", e => r.Month = e.target.value);
    tdMonth.appendChild(monthSelect);
    tr.appendChild(tdMonth);

    // === Cost (editable input) ===
    const tdCost = document.createElement("td");
    const costInput = document.createElement("input");
    costInput.type = "number";
    costInput.value = r.Cost ?? "";
    costInput.style.width = "100px";
    costInput.style.textAlign = "right";
    costInput.addEventListener("input", e => r.Cost = parseFloat(e.target.value) || 0);
    tdCost.appendChild(costInput);
    tr.appendChild(tdCost);

    // === Reman (dropdown) ===
    const tdReman = document.createElement("td");
    const remanSelect = document.createElement("select");
    ["Reman", "-"].forEach(opt => {
      const option = document.createElement("option");
      option.value = opt;
      option.textContent = opt;
      if (r.Reman === opt || r.Reman === "0,0" || !r.Reman) option.selected = true;
      remanSelect.appendChild(option);
    });
    // bind perubahan ke lomData
    remanSelect.addEventListener("change", e => r.Reman = e.target.value);
    tdReman.appendChild(remanSelect);
    tr.appendChild(tdReman);

    // === Planning (lookup dari Excel) ===
    const tdPlanning = document.createElement("td");
    // pastikan r.Order sudah ada, lakukan lookup
    r.Planning = lookupPlanning(r.Order);
    tdPlanning.textContent = r.Planning ? formatDateDDMMMYYYY(r.Planning) : "";
    tr.appendChild(tdPlanning);

    // === Status (lookup dari Excel) ===
    const tdStatus = document.createElement("td");
    r.Status = lookupStatus(r.Order);
    tdStatus.textContent = r.Status ?? "";
    tr.appendChild(tdStatus);

    tbody.appendChild(tr);
  });
}

// ===================== Helper lookup Excel =====================
function lookupPlanning(orderID) {
  const row = excelData.find(r => r.OrderID === orderID);
  return row ? row.EventStart : "";
}
function lookupStatus(orderID) {
  const row = excelData.find(r => r.OrderID === orderID);
  return row ? row.Status : "";
}

// ===================== Helper format tanggal =====================
function formatDateDDMMMYYYY(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  const day = String(d.getDate()).padStart(2, "0");
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const month = monthNames[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

/* ===================== FILTER LOM ===================== */
function filterLOM() {
  const monthFilter = document.querySelector("#filter-month")?.value;
  const remanFilter = document.querySelector("#filter-reman")?.value;

  let filtered = lomData;

  if (monthFilter && monthFilter !== "All") {
    filtered = filtered.filter(r => r.Month === monthFilter);
  }
  if (remanFilter && remanFilter !== "All") {
    filtered = filtered.filter(r => r.Reman === remanFilter);
  }

  renderLOMTable(filtered);
}

/* ===================== FILTERS ===================== */
function filterData() {
  const roomFilter    = (document.getElementById("filter-room")?.value || "").trim().toLowerCase();
  const orderFilter   = (document.getElementById("filter-order")?.value || "").trim().toLowerCase();
  const cphFilter     = (document.getElementById("filter-cph")?.value || "").trim().toLowerCase();
  const matFilter     = (document.getElementById("filter-mat")?.value || "").trim().toLowerCase();
  const sectionFilter = (document.getElementById("filter-section")?.value || "").trim().toLowerCase();
  const monthFilter   = (document.getElementById("filter-month")?.value || "").trim().toLowerCase();

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
  const ids = ["filter-room","filter-order","filter-cph","filter-mat","filter-section","filter-month"];
  ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
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
    const ordersText = document.getElementById("lom-add-order-input").value.trim();
    if (!ordersText) {
        alert("Masukkan Order terlebih dahulu.");
        return;
    }

    const orders = ordersText.split(/[\s,]+/).filter(o => o);
    orders.forEach(order => {
        lomData.push({
            Order: order,
            Month: "",
            Cost: 0,
            Reman: "",
            Planning: "",
            Status: ""
        });
    });

    renderLOMTable(lomData);
    document.getElementById("lom-add-order-input").value = "";
} // <-- cukup ini, jangan ada '});' ekstra

/* ===================== SAVE / LOAD JSON ===================== */
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
      Month: item.Month,
      Cost: item.Cost,
      Reman: item.Reman
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
  lomData = [];
  mergedData = [];
  renderTable([]);
  renderLOMTable([]);
  const statusEl = document.getElementById("upload-status");
  if (statusEl) statusEl.textContent = "Data dihapus.";
  updateMonthFilterOptions();
}

/* ===================== BUTTON WIRING ===================== */
function setupButtons() {
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) uploadBtn.onclick = handleUpload;

  const clearBtn = document.getElementById("clear-files-btn");
  if (clearBtn) clearBtn.onclick = clearAllData;

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

  // REVISI: pindahkan Add Order ke Menu 3 (LOM)
  const addOrderBtn = document.getElementById("lom-add-order-btn");
  if (addOrderBtn) addOrderBtn.onclick = addOrders;

  const lomFilterBtn = document.getElementById("lom-filter-btn");
  if (lomFilterBtn) lomFilterBtn.onclick = filterLOM;

  const lomResetBtn = document.getElementById("lom-reset-btn");
  if (lomResetBtn) lomResetBtn.onclick = resetFilterLOM;

  const lomRefreshBtn = document.getElementById("lom-refresh-btn");
  if (lomRefreshBtn) lomRefreshBtn.onclick = () => renderLOMTable(lomData);
}

/* ===================== STATUS COLOR ===================== */
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
















