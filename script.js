// Global data arrays
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];

const UI_LS_KEY = "ndarboe_ui_edits";

// Parse Excel date string robustly (handles "MM/DD/YYYY" and "MM/DD/YYYY hh:mm:ss AM/PM")
function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  let d = new Date(value);
  if (!isNaN(d.getTime())) return d;

  const regex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?)?$/i;
  const m = regex.exec(value.trim());
  if (m) {
    let month = parseInt(m[1], 10) - 1;
    let day = parseInt(m[2], 10);
    let year = parseInt(m[3], 10);
    let hour = m[4] ? parseInt(m[4], 10) : 0;
    let minute = m[5] ? parseInt(m[5], 10) : 0;
    let second = m[6] ? parseInt(m[6], 10) : 0;
    const ampm = m[7];

    if (ampm && ampm.toUpperCase() === "PM" && hour < 12) hour += 12;
    if (ampm && ampm.toUpperCase() === "AM" && hour === 12) hour = 0;

    return new Date(year, month, day, hour, minute, second);
  }

  return null;
}

// Format date dd-MMM-yyyy
function formatDateDDMMMYYYY(dt) {
  if (!dt) return "";
  if (!(dt instanceof Date)) return "";
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${dt.getDate().toString().padStart(2, "0")}-${monthNames[dt.getMonth()]}-${dt.getFullYear()}`;
}

// Format date ISO yyyy-mm-dd for input type=date
function formatDateISO(dateStr) {
  const d = parseExcelDate(dateStr);
  if (!d) return "";
  return d.toISOString().split("T")[0];
}

// Merge data from all sheets into mergedData array
function mergeData() {
  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }

  mergedData = iw39Data.map(item => ({
    Room: item.Room || "",
    "Order Type": item["Order Type"] || "",
    Order: item.Order || "",
    Description: item.Description || "",
    "Created On": item["Created On"] || "",
    "User Status": item["User Status"] || "",
    MAT: item.MAT || "",
    CPH: "",
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: item.Month || "",
    Cost: "-",
    Reman: item.Reman || "",
    Include: "-",
    Exclude: "-",
    Planning: "",
    "Status AMT": ""
  }));

  // CPH logic:
  // If Description starts with "JR" => "External Job"
  // Else lookup from Data2 by MAT
  mergedData.forEach(md => {
    if (md.Description && md.Description.trim().toUpperCase().startsWith("JR")) {
      md.CPH = "External Job";
    } else {
      const d2 = data2Data.find(d => d.MAT && d.MAT.trim() === md.MAT.trim());
      md.CPH = d2 ? d2.CPH || "" : "";
    }
  });

  // Section from data1Data by matching Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => d.Room && d.Room.trim() === md.Room.trim());
    md.Section = d1 ? d1.Section || "" : "";
  });

  // Status Part & Aging from sum57Data by matching Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => s.Order === md.Order);
    if (s57) {
      md["Status Part"] = s57["Part Complete"] || "";
      md.Aging = s57.Aging || "";
    }
  });

  // Planning and Status AMT from planningData by matching Order
  mergedData.forEach(md => {
    const pl = planningData.find(p => p.Order === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] || "";
      md["Status AMT"] = pl.Status || "";
    }
  });

  // Calculate Cost, Include, Exclude
  mergedData.forEach(md => {
    // Cost = (IW39.Total sum (plan) - IW39.Total sum (actual))/16500
    const iw = iw39Data.find(i => i.Order === md.Order);
    if (iw) {
      const planVal = parseFloat(iw["Total sum (plan)"]) || 0;
      const actualVal = parseFloat(iw["Total sum (actual)"]) || 0;
      let costVal = (planVal - actualVal) / 16500;
      if (costVal < 0) costVal = "-";
      else costVal = costVal.toFixed(2);
      md.Cost = costVal;

      if (md.Reman && md.Reman.toLowerCase().includes("reman")) {
        md.Include = (costVal === "-" ? "-" : (costVal * 0.25).toFixed(2));
      } else {
        md.Include = costVal;
      }

      if (md["Order Type"] === "PM38") {
        md.Exclude = "-";
      } else {
        md.Exclude = md.Include;
      }
    }
  });

  // Restore saved user edits from localStorage
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      saved.userEdits.forEach(edit => {
        const idx = mergedData.findIndex(r => r.Order === edit.Order);
        if (idx !== -1) {
          mergedData[idx] = { ...mergedData[idx], ...edit };
        }
      });
    }
  } catch {}

  updateMonthFilterOptions();
}

// Render mergedData to table
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#999">Tidak ada data</td></tr>`;
    return;
  }
  data.forEach(row => {
    const createdOnDate = parseExcelDate(row["Created On"]);
    const planningDate = parseExcelDate(row.Planning);

    tbody.insertAdjacentHTML("beforeend", `
      <tr>
        <td>${row.Room}</td>
        <td>${row["Order Type"]}</td>
        <td>${row.Order}</td>
        <td>${row.Description}</td>
        <td>${createdOnDate ? formatDateDDMMMYYYY(createdOnDate) : ""}</td>
        <td>${row["User Status"]}</td>
        <td>${row.MAT}</td>
        <td>${row.CPH}</td>
        <td>${row.Section}</td>
        <td>${row["Status Part"]}</td>
        <td>${row.Aging}</td>
        <td>${row.Month}</td>
        <td style="text-align:right;">${row.Cost}</td>
        <td>${row.Reman}</td>
        <td style="text-align:right;">${row.Include}</td>
        <td style="text-align:right;">${row.Exclude}</td>
        <td>${planningDate ? formatDateDDMMMYYYY(planningDate) : ""}</td>
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

// Attach event listeners to Edit/Delete buttons
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.onclick = () => {
      const order = btn.dataset.order;
      startEdit(order);
    };
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.onclick = () => {
      const order = btn.dataset.order;
      deleteOrder(order);
    };
  });
}

// Start inline editing a row
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
    <td><input type="text" value="${rowData.CPH}" data-field="CPH" /></td>
    <td><input type="text" value="${rowData.Section}" data-field="Section" /></td>
    <td><input type="text" value="${rowData["Status Part"]}" data-field="Status Part" /></td>
    <td><input type="text" value="${rowData.Aging}" data-field="Aging" /></td>
    <td><input type="text" value="${rowData.Month}" data-field="Month" /></td>
    <td><input type="text" style="text-align:right" value="${rowData.Cost}" data-field="Cost" /></td>
    <td><input type="text" value="${rowData.Reman}" data-field="Reman" /></td>
    <td><input type="text" style="text-align:right" value="${rowData.Include}" data-field="Include" /></td>
    <td><input type="text" style="text-align:right" value="${rowData.Exclude}" data-field="Exclude" /></td>
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

// Fungsi setupMenu untuk navigasi sidebar
function setupMenu() {
  const menuItems = document.querySelectorAll(".sidebar .menu-item");
  const contentSections = document.querySelectorAll(".content-section");

  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      // Hapus active dari semua menu item
      menuItems.forEach(i => i.classList.remove("active"));
      // Set active menu yang diklik
      item.classList.add("active");

      // Tampilkan content sesuai menu yang dipilih
      const menuId = item.dataset.menu;
      contentSections.forEach(sec => {
        if (sec.id === menuId) sec.classList.add("active");
        else sec.classList.remove("active");
      });
    });
  });
}

// Cancel editing: just re-render table
function cancelEdit(order) {
  renderTable(mergedData);
}

// Save edited data
function saveEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];
  const inputs = tr.querySelectorAll("input[data-field]");

  inputs.forEach(input => {
    const field = input.dataset.field;
    let val = input.value;
    if (field === "Created On" || field === "Planning") val = val || "";
    mergedData[rowIndex][field] = val;
  });

  saveUserEdits();
  renderTable(mergedData);
}

// Save user edits to localStorage
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

// Delete order row
function deleteOrder(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx !== -1) {
    if (!confirm(`Hapus data order ${order} ?`)) return;
    mergedData.splice(idx, 1);
    saveUserEdits();
    renderTable(mergedData);
  }
}

// Filter data based on inputs
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

// Reset filters to empty and show all data
function resetFilters() {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-month").value = "";
  renderTable(mergedData);
}

// Update Month filter dropdown options based on mergedData
function updateMonthFilterOptions() {
  const monthSelect = document.getElementById("filter-month");
  const months = new Set();
  mergedData.forEach(row => {
    if (row.Month) months.add(row.Month.trim());
  });
  const prevVal = monthSelect.value;
  monthSelect.innerHTML = `<option value="">-- All --</option>`;
  [...months].sort().forEach(m => {
    monthSelect.insertAdjacentHTML("beforeend", `<option value="${m.toLowerCase()}">${m}</option>`);
  });
  if ([...months].map(m => m.toLowerCase()).includes(prevVal)) {
    monthSelect.value = prevVal;
  }
}

// File upload handlers
document.getElementById("upload-btn").addEventListener("click", () => {
  const fileInput = document.getElementById("file-input");
  const fileType = document.getElementById("file-select").value;
  const status = document.getElementById("upload-status");

  if (!fileInput.files.length) {
    alert("Pilih file terlebih dahulu.");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Ambil sheet pertama saja
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    switch (fileType) {
      case "IW39":
        iw39Data = jsonData;
        status.textContent = `File IW39 berhasil diupload, ${jsonData.length} baris.`;
        break;
      case "SUM57":
        sum57Data = jsonData;
        status.textContent = `File SUM57 berhasil diupload, ${jsonData.length} baris.`;
        break;
      case "Planning":
        planningData = jsonData;
        status.textContent = `File Planning berhasil diupload, ${jsonData.length} baris.`;
        break;
      case "Data1":
        data1Data = jsonData;
        status.textContent = `File Data1 berhasil diupload, ${jsonData.length} baris.`;
        break;
      case "Data2":
        data2Data = jsonData;
        status.textContent = `File Data2 berhasil diupload, ${jsonData.length} baris.`;
        break;
      case "Budget":
        budgetData = jsonData;
        status.textContent = `File Budget berhasil diupload, ${jsonData.length} baris.`;
        break;
      default:
        status.textContent = "Jenis file tidak dikenal.";
        break;
    }
  };

  reader.onerror = () => {
    status.textContent = "Gagal membaca file.";
  };

  reader.readAsArrayBuffer(file);
});

// Buttons
document.getElementById("filter-btn").addEventListener("click", filterData);
document.getElementById("reset-btn").addEventListener("click", resetFilters);
document.getElementById("refresh-btn").addEventListener("click", () => {
  mergeData();
  renderTable(mergedData);
});
document.getElementById("save-btn").addEventListener("click", saveUserEdits);

// Initialize
window.onload = () => {
  renderTable([]);
};


