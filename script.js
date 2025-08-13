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
  if (typeof dt === "string") {
    // Try parse date from string in various formats
    // Try Date parse directly first
    d = new Date(dt);
    if (isNaN(d)) {
      // try manual parsing dd/mm/yyyy or mm/dd/yyyy
      const parts = dt.split(/[\/\-\s:]+/);
      if (parts.length >= 3) {
        // detect format based on first part
        // let's assume mm/dd/yyyy or dd/mm/yyyy - tricky, so try both
        let mm = parseInt(parts[0], 10);
        let dd = parseInt(parts[1], 10);
        let yyyy = parseInt(parts[2], 10);
        if (yyyy < 100) yyyy += 2000;
        if (mm > 12) {
          // swap if mm invalid
          [mm, dd] = [dd, mm];
        }
        d = new Date(yyyy, mm - 1, dd);
      } else {
        return dt; // return original if cannot parse
      }
    }
  } else if (dt instanceof Date) {
    d = dt;
  } else {
    return dt;
  }
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
  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }
  // map iw39Data order as base
  mergedData = iw39Data.map(item => {
    // Parse Create On date properly
    let createdOn = item["Created On"];
    if (createdOn) {
      createdOn = formatDateDDMMMYYYY(createdOn);
    }
    // Cost calculation (Total sum (plan) - Total sum (actual)) / 16500
    // fields: "Total sum (plan)", "Total sum (actual)"
    let costVal = "-";
    if (item["Total sum (plan)"] !== undefined && item["Total sum (actual)"] !== undefined) {
      const planNum = Number(item["Total sum (plan)"]) || 0;
      const actualNum = Number(item["Total sum (actual)"]) || 0;
      const costCalc = (planNum - actualNum) / 16500;
      costVal = costCalc < 0 ? "-" : costCalc.toFixed(2);
    }
    // Include calculation based on Reman
    let includeVal = costVal;
    if (item.Reman && item.Reman.toLowerCase().includes("reman")) {
      if (costVal !== "-") includeVal = (parseFloat(costVal) * 0.25).toFixed(2);
    }
    // Exclude calculation based on Order Type
    let excludeVal = includeVal;
    if (item["Order Type"] && item["Order Type"].toUpperCase() === "PM38") {
      excludeVal = "-";
    }

    // Planning lookup
    let planningVal = "";
    let statusAMTVal = "";
    const pl = planningData.find(p => p.Order === item.Order);
    if (pl) {
      planningVal = pl["Event Start"] ? formatDateDDMMMYYYY(pl["Event Start"]) : "";
      statusAMTVal = pl["Status"] || "";
    }

    // CPH column: If Description starts with "JR" => "External Job" else lookup Data2 by MAT
    let cphVal = "";
    if (item.Description && item.Description.startsWith("JR")) {
      cphVal = "External Job";
    } else if (item.MAT && data2Data.length) {
      const d2 = data2Data.find(d => d.MAT === item.MAT);
      cphVal = d2 ? (d2.CPH || "") : "";
    }

    return {
      Room: item.Room || "",
      "Order Type": item["Order Type"] || "",
      Order: item.Order || "",
      Description: item.Description || "",
      "Created On": createdOn || "",
      "User Status": item["User Status"] || "",
      MAT: item.MAT || "",
      CPH: cphVal || "",
      Section: "", // Will fill later
      "Status Part": "", // Will fill later
      Aging: "", // Will fill later
      Month: item.Month || "",
      Cost: costVal,
      Reman: item.Reman || "",
      Include: includeVal,
      Exclude: excludeVal,
      Planning: planningVal,
      "Status AMT": statusAMTVal
    };
  });

  // Fill Section from data1Data by Room match
  mergedData.forEach(md => {
    if (md.Room && data1Data.length) {
      const d1 = data1Data.find(d => d.Room === md.Room);
      if (d1) md.Section = d1.Section || "";
    }
  });

  // Fill Aging and Status Part from sum57Data by Order match
  mergedData.forEach(md => {
    if (md.Order && sum57Data.length) {
      const s57 = sum57Data.find(s => s.Order === md.Order);
      if (s57) {
        md.Aging = s57.Aging || "";
        md["Status Part"] = s57["Part Complete"] || "";
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
      <td>${row["Created On"]}</td>
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
    <td><input type="text" value="${rowData.Cost}" data-field="Cost" style="text-align:right;" /></td>
    <td><input type="text" value="${rowData.Reman}" data-field="Reman" /></td>
    <td><input type="text" value="${rowData.Include}" data-field="Include" style="text-align:right;" /></td>
    <td><input type="text" value="${rowData.Exclude}" data-field="Exclude" style="text-align:right;" /></td>
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
    let val = input.value;
    if (field === "Created On") val = formatDateDDMMMYYYY(val);
    mergedData[rowIndex][field] = val;
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
  if (!file) return;
  const text = await file.text();
  try {
    const json = JSON.parse(text);
    if (!json.mergedData) throw new Error("Format JSON tidak sesuai");
    mergedData = json.mergedData;
    renderTable(mergedData);
    updateMonthFilterOptions();
  } catch (e) {
    alert("File JSON tidak valid atau format salah");
  }
}

// -------- Upload button handler ----------
async function handleUpload(file, type) {
  if (!file) {
    alert("Pilih file terlebih dahulu");
    return;
  }
  try {
    const jsonData = await parseFile(file);
    switch(type) {
      case "IW39": iw39Data = jsonData; break;
      case "SUM57": sum57Data = jsonData; break;
      case "Planning": planningData = jsonData; break;
      case "Data1": data1Data = jsonData; break;
      case "Data2": data2Data = jsonData; break;
      case "Budget": budgetData = jsonData; break;
      default: alert("Tipe file tidak dikenali"); return;
    }
    alert(`${type} berhasil diupload, total baris: ${jsonData.length}`);
  } catch (e) {
    alert(`Gagal upload file ${type}: ${e.message}`);
  }
}

// -------- Setup sidebar menu navigation ----------
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

// -------- Initialize UI, event listeners --------
function init() {
  // Upload handler
  document.getElementById("upload-btn").onclick = async () => {
    const fileInput = document.getElementById("file-input");
    const fileType = document.getElementById("file-select").value;
    const file = fileInput.files[0];
    await handleUpload(file, fileType);
  };

  // Filter button
  document.getElementById("filter-btn").onclick = filterData;
  // Reset filter button
  document.getElementById("reset-btn").onclick = resetFilters;

  // Refresh data (merge + render)
  document.getElementById("refresh-btn").onclick = () => {
    mergeData();
    renderTable(mergedData);
  };

  // Save JSON button
  document.getElementById("save-btn").onclick = saveToJSON;

  // Load JSON button
  document.getElementById("load-btn").onclick = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";
    input.onchange = async (e) => {
      if (e.target.files.length) await loadFromJSON(e.target.files[0]);
    };
    input.click();
  };

  // Setup sidebar menu
  setupMenu();

  // Initial empty table render
  renderTable([]);
}

window.onload = init;
