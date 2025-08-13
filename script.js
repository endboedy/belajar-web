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
  if (typeof dt === "string") d = new Date(dt);
  else d = dt;
  if (isNaN(d)) return dt;
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${d.getDate().toString().padStart(2,"0")} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

// -------- Format date to yyyy-mm-dd for input type=date ----------
function formatDateISO(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return d.toISOString().split("T")[0];
}

// -------- Merge function to combine data from multiple sheets --------
function mergeData() {
  // map iw39Data order as base
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
    Reman: "",
    Include: "-",
    Exclude: "-",
    Planning: "",
    "Status AMT": ""
  }));

  // Merge CPH from data2Data by matching MAT
  mergedData.forEach(md => {
    const d2 = data2Data.find(d => d.MAT && d.MAT.trim() === md.MAT.trim());
    md.CPH = d2 ? d2.CPH || "" : "";
  });

  // Merge Section from data1Data by matching Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => d.Room && d.Room.trim() === md.Room.trim());
    md.Section = d1 ? d1.Section || "" : "";
  });

  // Merge Aging & Status Part from sum57Data by matching Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => s.Order === md.Order);
    if (s57) {
      md.Aging = s57.Aging || "";
      md["Status Part"] = s57["Part Complete"] || "";
    }
  });

  // Merge Planning data by matching Order and picking Event Start or Status
  mergedData.forEach(md => {
    const pl = planningData.find(p => p.Order === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] || pl.Status || "";
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

  // Update filter month dropdown options based on mergedData
  updateMonthFilterOptions();
}

// -------- Render mergedData to table ----------
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#999">Tidak ada data</td></tr>`;
    return;
  }
  data.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
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
    `;
    tbody.appendChild(tr);
  });
  attachTableEvents();
}

// -------- Attach event listeners for Edit/Delete buttons ----------
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.onclick = () => startEdit(btn.dataset.order);
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.onclick = () => deleteOrder(btn.dataset.order);
  });
}

// -------- Start editing a row ----------
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
    <td><input type="text" value="${rowData.Cost}" data-field="Cost" /></td>
    <td><input type="text" value="${rowData.Reman}" data-field="Reman" /></td>
    <td><input type="text" value="${rowData.Include}" data-field="Include" /></td>
    <td><input type="text" value="${rowData.Exclude}" data-field="Exclude" /></td>
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

// -------- Cancel editing, restore original row ----------
function cancelEdit(order) {
  renderTable(mergedData);
}

// -------- Save edited data ----------
function saveEdit(order) {
  const rowIndex = mergedData.findIndex(r => r.Order === order);
  if (rowIndex === -1) return;
  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[rowIndex];

  const inputs = tr.querySelectorAll("input[data-field]");
  inputs.forEach(input => {
    const field = input.dataset.field;
    mergedData[rowIndex][field] = input.value;
  });

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
  } catch {}

}

// -------- Delete order ----------
function deleteOrder(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx !== -1) {
    if (!confirm(`Hapus data order ${order} ?`)) return;
    mergedData.splice(idx, 1);
    saveUserEdits();
    renderTable(mergedData);
  }
}

// -------- Filter function ----------
function filterData() {
  const roomFilter = document.getElementById("filter-room").value.trim().toLowerCase();
  const orderFilter = document.getElementById("filter-order").value.trim().toLowerCase();
  const cphFilter = document.getElementById("filter-cph").value.trim().toLowerCase();
  const matFilter = document.getElementById("filter-mat").value.trim().toLowerCase();
  const sectionFilter = document.getElementById("filter-section").value.trim().toLowerCase();
  const monthFilter = document.getElementById("filter-month").value.trim().toLowerCase();

  let filtered = mergedData.filter(item => {
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

// -------- Save mergedData to JSON file ----------
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

// -------- Load JSON file and update mergedData ----------
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

// -------- Upload button handler ----------
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
    const json = await parseFile(file);

    switch(jenis) {
      case "IW39": iw39Data = json; break;
      case "SUM57": sum57Data = json; break;
      case "Planning": planningData = json; break;
      case "Data1": data1Data = json; break;
      case "Data2": data2Data = json; break;
      case "Budget": budgetData = json; break;
    }

    status.textContent = `File ${file.name} berhasil diupload sebagai ${jenis}.`;

    // Clear file input to allow same file upload again
    fileInput.value = "";

  } catch (e) {
    status.textContent = `Error saat membaca file: ${e.message}`;
  }
}

// -------- Clear all uploaded data ----------
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

// -------- Refresh button handler: merge + render ----------
function refreshData() {
  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }
  mergeData();
  renderTable(mergedData);
}

// -------- Add Order button handler (append new order) ----------
function addOrders() {
  const input = document.getElementById("add-order-input");
  const text = input.value.trim();
  if (!text) {
    alert("Masukkan Order terlebih dahulu.");
    return;
  }

  // Split by comma, space or newline
  const orders = text.split(/[\s,]+/).filter(o => o);
  let added = 0;
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
      added++;
    }
  });

  if (added) {
    saveUserEdits();
    renderTable(mergedData);
    alert(`${added} Order berhasil ditambahkan.`);
  } else {
    alert("Order sudah ada di data.");
  }
  input.value = "";
}

// -------- Setup menu click switching ----------
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

window.onload = () => {
  setupMenu();
};

// -------- Init main --------
function init() {
  setupMenu();

  document.getElementById("upload-btn").onclick = handleUpload;
  document.getElementById("clear-files-btn").onclick = clearAllData;
  document.getElementById("refresh-btn").onclick = refreshData;
  document.getElementById("add-order-btn").onclick = addOrders;

  // Filters
  document.getElementById("filter-room").oninput = filterData;
  document.getElementById("filter-order").oninput = filterData;
  document.getElementById("filter-cph").oninput = filterData;
  document.getElementById("filter-mat").oninput = filterData;
  document.getElementById("filter-section").oninput = filterData;
  document.getElementById("filter-month").onchange = filterData;

  document.getElementById("filter-btn").onclick = filterData;
  document.getElementById("reset-btn").onclick = resetFilters;
  document.getElementById("save-btn").onclick = saveToJSON;
  document.getElementById("load-btn").onclick = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "application/json";
    input.onchange = () => {
      if (input.files.length) {
        loadFromJSON(input.files[0]);
      }
    };
    input.click();
  };

  renderTable([]);
}

window.onload = init;

