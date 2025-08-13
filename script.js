// Global data arrays
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];

const UI_LS_KEY = "ndarboe_ui_edits";

// -------- Utility function to parse XLSX file to JSON ----------
async function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

// -------- Format date DD MMM YYYY ----------
function formatDateDDMMMYYYY(dt) {
  if (!dt) return "";

  let d;
  if (dt instanceof Date) {
    d = dt;
  } else if (typeof dt === "string" || typeof dt === "number") {
    d = new Date(dt);
  } else if (typeof dt === "object" && dt !== null) {
    if (dt.$date) d = new Date(dt.$date);
    else if (dt.date) d = new Date(dt.date);
    else return "";
  } else {
    return "";
  }

  if (isNaN(d.getTime())) return "";

  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${d.getDate().toString().padStart(2,"0")} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

// -------- Format date ISO yyyy-mm-dd for input[type=date] ----------
function formatDateISO(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return "";
  return d.toISOString().split("T")[0];
}

// -------- Merge function to combine data from multiple sheets --------
function mergeData() {
  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }

  // Start with iw39Data as base
  mergedData = iw39Data.map(item => ({
    Room: item.Room || "",
    "Order Type": item["Order Type"] || "",
    Order: item.Order || "",
    Description: item.Description || "",
    "Created On": item["Created On"] || "",
    "User Status": item["User Status"] || "",
    MAT: item.MAT || "",
    CPH: "", // nanti isi di bawah
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: item.Month || "",
    Cost: "",
    Reman: item.Reman || "",
    Include: "",
    Exclude: "",
    Planning: "",
    "Status AMT": ""
  }));

  // Merge CPH kolom
  mergedData.forEach(md => {
    if (md.Description && md.Description.substring(0,2).toUpperCase() === "JR") {
      md.CPH = "External Job";
    } else {
      // Lookup dari Data2 berdasarkan MAT
      const d2 = data2Data.find(d => d.MAT && d.MAT.trim() === md.MAT.trim());
      md.CPH = d2 ? (d2.CPH || "") : "";
    }
  });

  // Merge Section dari data1Data berdasar Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => d.Room && d.Room.trim() === md.Room.trim());
    md.Section = d1 ? (d1.Section || "") : "";
  });

  // Merge Aging & Status Part dari sum57Data berdasar Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => s.Order === md.Order);
    if (s57) {
      md.Aging = s57.Aging || "";
      md["Status Part"] = s57["Part Complete"] || "";
    }
  });

  // Rumus Cost dari IW39: (Total sum (plan) - Total sum (actual))/16500, jika < 0 maka "-"
  mergedData.forEach(md => {
    const iw = iw39Data.find(i => i.Order === md.Order);
    if (iw && typeof iw["Total sum (plan)"] !== "undefined" && typeof iw["Total sum (actual)"] !== "undefined") {
      const plan = Number(iw["Total sum (plan)"]) || 0;
      const actual = Number(iw["Total sum (actual)"]) || 0;
      let costVal = (plan - actual) / 16500;
      if (costVal < 0) md.Cost = "-";
      else md.Cost = costVal.toFixed(2);
    } else {
      md.Cost = "";
    }
  });

  // Kolom Include = jika Reman = "Reman" maka Cost*0.25, jika kosong sama dengan Cost
  mergedData.forEach(md => {
    if (md.Reman && md.Reman.toLowerCase() === "reman") {
      md.Include = (isNaN(Number(md.Cost)) ? 0 : Number(md.Cost)*0.25).toFixed(2);
    } else {
      md.Include = md.Cost;
    }
  });

  // Kolom Exclude = jika Order Type = "PM38" maka "-", jika bukan maka sama dengan Include
  mergedData.forEach(md => {
    if (md["Order Type"] && md["Order Type"].toUpperCase() === "PM38") {
      md.Exclude = "-";
    } else {
      md.Exclude = md.Include;
    }
  });

  // Merge Planning data berdasarkan Order
  mergedData.forEach(md => {
    const pl = planningData.find(p => p.Order === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] ? formatDateDDMMMYYYY(pl["Event Start"]) : "";
      md["Status AMT"] = pl.Status || "";
    }
  });

  // Restore user edits dari localStorage
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      if (saved.userEdits && Array.isArray(saved.userEdits)) {
        saved.userEdits.forEach(edit => {
          const idx = mergedData.findIndex(r => r.Order === edit.Order);
          if (idx !== -1) {
            mergedData[idx] = { ...mergedData[idx], ...edit };
          }
        });
      }
    }
  } catch (e) {
    console.warn("Gagal load user edits:", e);
  }

  updateMonthFilterOptions();
}

// -------- Render mergedData ke tabel ----------
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#999">Tidak ada data</td></tr>`;
    return;
  }
  data.forEach(row => {
    tbody.insertAdjacentHTML("beforeend", `
      <tr>
        <td>${row.Room}</td>
        <td>${row["Order Type"]}</td>
        <td>${row.Order}</td>
        <td>${row.Description}</td>
        <td>${formatDateDDMMMYYYY(row["Created On"])}</td>
        <td>${row["User Status"]}</td>
        <td>${row.MAT}</td>
        <td>${row.CPH}</td>
        <td>${row.Section}</td>
        <td>${row["Status Part"]}</td>
        <td>${row.Aging}</td>
        <td>${row.Month}</td>
        <td>${row.Cost}</td>
        <td>${row.Reman}</td>
        <td>${row.Include}</td>
        <td>${row.Exclude}</td>
        <td>${row.Planning}</td>
        <td>${row["Status AMT"]}</td>
        <td>
          <button class="action-btn edit-btn" data-order="${row.Order}">Edit</button>
          <button class="action-btn delete-btn" data-order="${row.Order}">Delete</button>
        </td>
      </tr>
    `);
  });
  attachTableEvents();
}

// -------- Pasang event tombol Edit/Delete ----------
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.onclick = () => startEdit(btn.dataset.order);
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.onclick = () => deleteOrder(btn.dataset.order);
  });
}

// -------- Mulai Edit ----------
function startEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];
  const rowData = mergedData[rowIndex];

  tr.innerHTML = `
    <td><input type="text" value="${rowData.Room}" data-field="Room" /></td>
    <td><input type="text" value="${rowData["Order Type"]}" data-field="Order Type" /></td>
    <td>${rowData.Order}</td>
    <td><input type="text" value="${rowData.Description}" data-field="Description" /></td>
    <td><input type="date" value="${formatDateISO(rowData["Created On"])}" data-field="Created On" /></td>
    <td><input type="text" value="${rowData["User Status"]}" data-field="User Status" /></td>
    <td><input type="text" value="${rowData.MAT}" data-field="MAT" /></td>
    <td><input type="text" value="${rowData.CPH}" data-field="CPH" readonly /></td>
    <td><input type="text" value="${rowData.Section}" data-field="Section" /></td>
    <td><input type="text" value="${rowData["Status Part"]}" data-field="Status Part" /></td>
    <td><input type="text" value="${rowData.Aging}" data-field="Aging" /></td>
    <td><input type="text" value="${rowData.Month}" data-field="Month" /></td>
    <td><input type="text" value="${rowData.Cost}" data-field="Cost" readonly /></td>
    <td><input type="text" value="${rowData.Reman}" data-field="Reman" /></td>
    <td><input type="text" value="${rowData.Include}" data-field="Include" readonly /></td>
    <td><input type="text" value="${rowData.Exclude}" data-field="Exclude" readonly /></td>
    <td><input type="text" value="${rowData.Planning}" data-field="Planning" /></td>
    <td><input type="text" value="${rowData["Status AMT"]}" data-field="Status AMT" /></td>
    <td>
      <button class="action-btn save-btn" data-order="${order}">Save</button>
      <button class="action-btn cancel-btn" data-order="${order}">Cancel</button>
    </td>
  `;

  tr.querySelector(".save-btn").onclick = () => saveEdit(order);
  tr.querySelector(".cancel-btn").onclick = () => cancelEdit(order);
}

// -------- Cancel edit ----------
function cancelEdit(order) {
  renderTable(mergedData);
}

// -------- Save edit ----------
function saveEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];
  const inputs = tr.querySelectorAll("input[data-field]");

  inputs.forEach(input => {
    const field = input.dataset.field;
    // Ignore readonly (CPH, Cost, Include, Exclude) karena hasil formula
    if (input.hasAttribute("readonly")) return;
    mergedData[rowIndex][field] = input.value;
  });

  // Setelah save, recalculate Cost, Include, Exclude dan CPH jika perlu
  // CPH recalculated sama seperti sebelumnya
  const md = mergedData[rowIndex];
  if (md.Description && md.Description.substring(0,2).toUpperCase() === "JR") {
    md.CPH = "External Job";
  } else {
    const d2 = data2Data.find(d => d.MAT && d.MAT.trim() === md.MAT.trim());
    md.CPH = d2 ? (d2.CPH || "") : "";
  }

  // Cost recalculated dari iw39Data berdasar order
  const iw = iw39Data.find(i => i.Order === md.Order);
  if (iw && typeof iw["Total sum (plan)"] !== "undefined" && typeof iw["Total sum (actual)"] !== "undefined") {
    const plan = Number(iw["Total sum (plan)"]) || 0;
    const actual = Number(iw["Total sum (actual)"]) || 0;
    let costVal = (plan - actual) / 16500;
    if (costVal < 0) md.Cost = "-";
    else md.Cost = costVal.toFixed(2);
  } else {
    md.Cost = "";
  }

  if (md.Reman && md.Reman.toLowerCase() === "reman") {
    md.Include = (isNaN(Number(md.Cost)) ? 0 : Number(md.Cost)*0.25).toFixed(2);
  } else {
    md.Include = md.Cost;
  }

  if (md["Order Type"] && md["Order Type"].toUpperCase() === "PM38") {
    md.Exclude = "-";
  } else {
    md.Exclude = md.Include;
  }

  saveUserEdits();
  renderTable(mergedData);
}

// -------- Save user edits to localStorage ----------
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
  } catch (e) {
    console.warn("Gagal simpan user edits:", e);
  }
}

// -------- Delete order ----------
function deleteOrder(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx === -1) return;
  if (!confirm(`Hapus data order ${order} ?`)) return;

  mergedData.splice(idx, 1);
  saveUserEdits();
  renderTable(mergedData);
}

// -------- Filter function ----------
function filterData() {
  const roomFilter = document.getElementById("filter-room").value.trim().toLowerCase();
  const orderFilter = document.getElementById("filter-order").value.trim().toLowerCase();
  const cphFilter = document.getElementById("filter-cph").value.trim().toLowerCase();
  const matFilter = document.getElementById("filter-mat").value.trim().toLowerCase();
  const sectionFilter = document.getElementById("filter-section").value.trim().toLowerCase();
  const monthFilter = document.getElementById("filter-month").value.trim().toLowerCase();

  const filtered = mergedData.filter(item => {
    return (
      (!roomFilter || (item.Room && item.Room.toLowerCase().includes(roomFilter))) &&
      (!orderFilter || (item.Order && item.Order.toLowerCase().includes(orderFilter))) &&
      (!cphFilter || (item.CPH && item.CPH.toLowerCase().includes(cphFilter))) &&
      (!matFilter || (item.MAT && item.MAT.toLowerCase().includes(matFilter))) &&
      (!sectionFilter || (item.Section && item.Section.toLowerCase().includes(sectionFilter))) &&
      (!monthFilter || (item.Month && item.Month.toLowerCase() === monthFilter))
    );
  });

  renderTable(filtered);
}

// -------- Reset filters ----------
function resetFilters() {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-month").value = "";
  renderTable(mergedData);
}

// -------- Update Month dropdown options ----------
function updateMonthFilterOptions() {
  const monthSelect = document.getElementById("filter-month");
  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== "")));
  months.sort();

  while (monthSelect.options.length > 1) {
    monthSelect.remove(1);
  }

  months.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m.toLowerCase();
    opt.textContent = m;
    monthSelect.appendChild(opt);
  });
}

// -------- Handle Upload File ----------
async function handleUpload() {
  const fileInput = document.getElementById("file-input");
  const fileSelect = document.getElementById("file-select");
  const uploadStatus = document.getElementById("upload-status");

  if (!fileInput.files.length) {
    alert("Pilih file dulu!");
    return;
  }

  const file = fileInput.files[0];
  const jenis = fileSelect.value;

  uploadStatus.textContent = `Uploading ${jenis}...`;

  try {
    const data = await parseFile(file);

    switch (jenis) {
      case "IW39": iw39Data = data; break;
      case "SUM57": sum57Data = data; break;
      case "Planning": planningData = data; break;
      case "Data1": data1Data = data; break;
      case "Data2": data2Data = data; break;
      case "Budget": budgetData = data; break;
      default: alert("Jenis file tidak dikenal!"); return;
    }

    uploadStatus.textContent = `${jenis} berhasil diupload.`;
    fileInput.value = "";

  } catch (e) {
    uploadStatus.textContent = `Error upload: ${e.message || e}`;
  }
}

// -------- Clear all data ----------
function clearAllData() {
  if (!confirm("Yakin hapus semua data dan reset?")) return;

  iw39Data = [];
  sum57Data = [];
  planningData = [];
  data1Data = [];
  data2Data = [];
  budgetData = [];
  mergedData = [];
  localStorage.removeItem(UI_LS_KEY);

  document.getElementById("upload-status").textContent = "";
  renderTable([]);
  updateMonthFilterOptions();
}

// -------- Refresh data: merge dan render ----------
function refreshData() {
  mergeData();
  renderTable(mergedData);
}

// -------- Add Orders by input textarea ----------
function addOrders() {
  const input = document.getElementById("add-order-input");
  const status = document.getElementById("add-order-status");
  const raw = input.value.trim();

  if (!raw) {
    status.textContent = "Masukkan minimal 1 order.";
    return;
  }

  // Pisah input per spasi atau koma atau newline
  const orders = raw.split(/[\s,]+/).filter(o => o);

  let addedCount = 0;

  orders.forEach(o => {
    if (!mergedData.find(r => r.Order === o)) {
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
        Cost: "",
        Reman: "",
        Include: "",
        Exclude: "",
        Planning: "",
        "Status AMT": ""
      });
      addedCount++;
    }
  });

  if (addedCount) {
    saveUserEdits();
    renderTable(mergedData);
    status.textContent = `${addedCount} order berhasil ditambahkan.`;
    input.value = "";
  } else {
    status.textContent = `Order sudah ada di tabel.`;
  }
}

// -------- Setup menu navigation ----------
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

// -------- Setup all event listeners ----------
function setupEvents() {
  document.getElementById("upload-btn").onclick = handleUpload;
  document.getElementById("clear-files-btn").onclick = clearAllData;
  document.getElementById("refresh-btn").onclick = refreshData;
  document.getElementById("add-order-btn").onclick = addOrders;
  document.getElementById("filter-btn").onclick = filterData;
  document.getElementById("reset-btn").onclick = resetFilters;
  document.getElementById("save-btn").onclick = () => {
    const dataStr = JSON.stringify(mergedData, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "ndarboe_data.json";
    a.click();
  };

  document.getElementById("load-btn").onclick = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";
    input.onchange = e => {
      const file = e.target.files[0];
      if (file) loadFromJSON(file);
    };
    input.click();
  };
}

// -------- Load data from JSON file ----------
function loadFromJSON(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      if (!Array.isArray(data)) throw new Error("JSON tidak valid");
      mergedData = data;
      saveUserEdits();
      renderTable(mergedData);
      updateMonthFilterOptions();
      alert("Data berhasil dimuat dari JSON.");
    } catch (err) {
      alert("Gagal load JSON: " + err.message);
    }
  };
  reader.readAsText(file);
}

// -------- Init ----------
function init() {
  setupMenu();
  setupEvents();
  renderTable([]); // empty awal
}

window.onload = init;
