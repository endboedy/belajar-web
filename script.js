/****************************************************
 * Ndarboe.net - FULL script.js (Menu 1–4)
 * Upload 6 sumber (IW39, SUM57, Planning, Budget, Data1, Data2)
 * Merge & render Lembar Kerja
 * Filter, Add Order, Edit/Save/Delete, Save/Load JSON
 * Pewarnaan status + format tanggal
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

/* ===================== HELPERS: DATE ===================== */
function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;

  // Excel serial number
  if (typeof anyDate === "number") {
    const dec = XLSX && XLSX.SSF && XLSX.SSF.parse_date_code
      ? XLSX.SSF.parse_date_code(anyDate)
      : null;
    if (dec && typeof dec === "object") {
      return new Date(dec.y, (dec.m || 1) - 1, dec.d || 1, dec.H || 0, dec.M || 0, dec.S || 0);
    }
  }

  if (anyDate instanceof Date && !isNaN(anyDate)) return anyDate;

  if (typeof anyDate === "string") {
    const s = anyDate.trim();
    if (!s) return null;

    // ISO or browser-parseable
    const iso = new Date(s);
    if (!isNaN(iso)) return iso;

    // Try M/D/YYYY (with or without time AM/PM)
    const parts = s.match(/(\d{1,4})/g);
    if (parts && parts.length >= 3) {
      const ampm = /am|pm/i.test(s) ? s.match(/am|pm/i)[0] : "";
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      let y, m, d;
      if (p1 <= 12 && p2 <= 31) {
        m = p1; d = p2; y = p3;
      } else {
        y = p1; m = p2; d = p3;
      }
      if (!y || !m || !d) return null;
      let H = 0, M = 0, S = 0;
      if (parts.length >= 5) {
        H = parseInt(parts[3], 10) || 0;
        M = parseInt(parts[4], 10) || 0;
        S = parseInt(parts[5], 10) || 0;
      }
      if (ampm) {
        if (/pm/i.test(ampm) && H < 12) H += 12;
        if (/am/i.test(ampm) && H === 12) H = 0;
      }
      return new Date(y, (m - 1), d, H, M, S);
    }
  }
  return null;
}
function safe(v){ return (v == null ? "" : String(v)); }
function formatDateISO(v){
  const d = toDateObj(v);
  if (!d) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}
function fmtDateDisplay(v){ // dd-MMM-yyyy
  const d = toDateObj(v);
  if (!d) return "";
  return d.toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" }).replace(/ /g,"-");
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

        // raw:false agar tanggal string tetap terbaca & di-parse manual
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

  status.textContent = `Memproses file ${file.name} sebagai ${jenis}.`;
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
    fileInput.value = ""; // boleh upload file yang sama lagi
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
    Room: safe(row.Room),
    "Order Type": safe(row["Order Type"]),
    Order: safe(row.Order),
    Description: safe(row.Description),
    "Created On": row["Created On"] || "",   // raw, format saat render
    "User Status": safe(row["User Status"]),
    MAT: safe(row.MAT),
    CPH: "",                    // diisi dari rules
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: safe(row.Month),
    Cost: "-",
    Reman: safe(row.Reman),
    Include: "-",
    Exclude: "-",
    Planning: "",               // Event Start (Planning)
    "Status AMT": ""            // Status (Planning)
  }));

  // CPH: Description startsWith JR → "External Job", else lookup Data2 by MAT
  mergedData.forEach(md => {
    if ((md.Description || "").trim().toUpperCase().startsWith("JR")) {
      md.CPH = "External Job";
    } else {
      const d2 = data2Data.find(d => safe(d.MAT).trim() === md.MAT.trim());
      md.CPH = d2 ? safe(d2.CPH) : "";
    }
  });

  // Section: Data1 via Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => safe(d.Room).trim() === md.Room.trim());
    md.Section = d1 ? safe(d1.Section) : "";
  });

  // SUM57: Aging & Status Part via Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => safe(s.Order) === md.Order);
    if (s57) {
      md.Aging = safe(s57.Aging);
      md["Status Part"] = safe(s57["Part Complete"]);
    }
  });

  // Planning: Event Start & Status
  mergedData.forEach(md => {
    const pl = planningData.find(p => safe(p.Order) === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] || "";
      md["Status AMT"] = safe(pl.Status);
    }
  });

  // Hitung Cost/Include/Exclude dari IW39 plan vs actual
  mergedData.forEach(md => {
    const src = iw39Data.find(i => safe(i.Order) === md.Order);
    if (!src) return;

    const plan = parseFloat(safe(src["Total sum (plan)"]).replace(/,/g,"")) || 0;
    const actual = parseFloat(safe(src["Total sum (actual)"]).replace(/,/g,"")) || 0;
    let cost = (plan - actual) / 16500;

    if (!isFinite(cost) || cost < 0) {
      md.Cost = "-";
      md.Include = "-";
      md.Exclude = (md["Order Type"] === "PM38") ? "-" : "-";
    } else {
      const costStr = cost.toFixed(2);
      md.Cost = costStr;

      const isReman = (md.Reman || "").toLowerCase().includes("reman");
      const includeNum = isReman ? (cost * 0.25) : cost;
      md.Include = includeNum.toFixed(2);

      md.Exclude = (md["Order Type"] === "PM38") ? "-" : md.Include;
    }
  });

  // Restore user edits dari localStorage (FIX: pakai spread object yang benar) :contentReference[oaicite:3]{index=3}
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      if (saved && Array.isArray(saved.userEdits)) {
        saved.userEdits.forEach(edit => {
          const idx = mergedData.findIndex(r => r.Order === edit.Order);
          if (idx !== -1) {
            mergedData[idx] = { ...mergedData[idx], ...edit }; // <- perbaikan
          }
        });
      }
    }
  } catch {}

  updateMonthFilterOptions();
}

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
  if (s === "complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if (s === "not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}
function asColoredStatusAMT(val) {
  const v = (val || "").toString().toUpperCase();
  if (v === "O")   return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#ffeb3b;color:#000;">${safe(val)}</span>`;
  if (v === "IP")  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if (v === "YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataSet = mergedData) {
  const tbody = document.querySelector("#data-table tbody"); // pastikan id sesuai HTML
  if (!tbody) return;
  tbody.innerHTML = "";

  dataSet.forEach((row, index) => {
    const tr = document.createElement("tr");

    // Tulis semua kolom sesuai urutan field di mergedData
    Object.values(row).forEach(val => {
      const td = document.createElement("td");
      td.textContent = val ?? "";
      tr.appendChild(td);
    });

    // Kolom aksi edit/delete
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
  // Tombol Edit
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.addEventListener("click", function () {
      const tr = this.closest("tr");
      const tds = tr.querySelectorAll("td");

      // Ambil nilai lama
      const currentMonth = tds[11].textContent.trim();
      const currentCost  = tds[12].textContent.trim();
      const currentReman = tds[13].textContent.trim();

      // Ganti hanya kolom Month
      const monthOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        .map(m => `<option value="${m}" ${m===currentMonth?"selected":""}>${m}</option>`).join("");
      tds[11].innerHTML = `<select class="edit-month">${monthOptions}</select>`;

      // Ganti hanya kolom Cost
      tds[12].innerHTML = `<input type="number" class="edit-cost" value="${currentCost}" style="width:80px;text-align:right;">`;

      // Ganti hanya kolom Reman
      tds[13].innerHTML = `
        <select class="edit-reman">
          <option value="Reman" ${currentReman==="Reman"?"selected":""}>Reman</option>
          <option value="-" ${currentReman==="-"?"selected":""}>-</option>
        </select>`;

      // Ganti tombol jadi Save & Cancel
      this.outerHTML = `<button class="action-btn save-btn" data-index="${btn.dataset.index}">Save</button>
                        <button class="action-btn cancel-btn">Cancel</button>`;

      // Handler Save
      tr.querySelector(".save-btn").addEventListener("click", function () {
        const index = parseInt(this.dataset.index, 10);
        mergedData[index].Month = tr.querySelector(".edit-month").value;
        mergedData[index].Cost  = tr.querySelector(".edit-cost").value;
        mergedData[index].Reman = tr.querySelector(".edit-reman").value;
        saveMergedData(); // kalau ada fitur simpan
        renderTable();
      });

      // Handler Cancel
      tr.querySelector(".cancel-btn").addEventListener("click", function () {
        renderTable();
      });
    });
  });

  // Tombol Delete
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.addEventListener("click", function () {
      const index = parseInt(this.dataset.index, 10);
      if (confirm("Yakin mau hapus data ini?")) {
        mergedData.splice(index, 1);
        saveMergedData(); // kalau ada fitur simpan
        renderTable();
      }
    });
  });
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

  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort(); // FIX: tidak pakai .months :contentReference[oaicite:4]{index=4}
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
    <td><select data-field="Month">${monthOptions}</select></td>
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

  const monthSel = tr.querySelector("select[data-field='Month']");
  monthSel.value = row.Month || "";

  tr.querySelector(".save-btn").onclick = () => saveEdit(order);
  tr.querySelector(".cancel-btn").onclick = () => cancelEdit();
}

function cancelEdit() { renderTable(mergedData); }

function saveEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];

  const inputs = tr.querySelectorAll("input[data-field], select[data-field]");
  inputs.forEach(input => {
    const field = input.dataset.field;
    mergedData[rowIndex][field] = input.value;
  });

  // persist
  saveUserEdits();

  // re-merge utk kalkulasi kembali (Cost/Include/Exclude) & render
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
  if (!text) { alert("Masukkan Order terlebih dahulu."); return; }

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

/* ===================== SAVE / LOAD JSON ===================== */
function saveToJSON() {
  if (!mergedData.length) { alert("Tidak ada data untuk disimpan."); return; }
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
  const st = document.getElementById("upload-status");
  if (st) st.textContent = "Data dihapus.";
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

