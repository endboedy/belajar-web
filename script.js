// ======= GLOBAL DATA =======
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];

const UI_LS_KEY = "ndarboe_ui_edits";

// ======= UTILITIES =======
function formatDateDDMMMYYYY(dateInput) {
  if (!dateInput) return "";
  let d;
  if (dateInput instanceof Date) {
    d = dateInput;
  } else if (typeof dateInput === "string") {
    d = new Date(dateInput);
    if (isNaN(d)) {
      // try dd/mm/yyyy or mm/dd/yyyy
      let parts = dateInput.split(/[\s\/:-]+/);
      if (parts.length >= 3) {
        // swap month/day if month > 12
        let mm = parseInt(parts[0], 10);
        let dd = parseInt(parts[1], 10);
        let yyyy = parseInt(parts[2], 10);
        if (yyyy < 100) yyyy += 2000;
        if (mm > 12) [mm, dd] = [dd, mm];
        d = new Date(yyyy, mm - 1, dd);
      } else {
        return dateInput;
      }
    }
  } else {
    return "";
  }
  if (isNaN(d)) return "";
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${d.getDate().toString().padStart(2, "0")} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

function formatDateISO(dateInput) {
  if (!dateInput) return "";
  const d = new Date(dateInput);
  if (isNaN(d)) return "";
  return d.toISOString().split("T")[0];
}

async function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
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
    reader.onerror = err => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

// ======= MERGE DATA =======
function mergeData() {
  if (!iw39Data.length) {
    alert("Upload file IW39 dulu ya");
    return;
  }

  mergedData = iw39Data.map(item => {
    // Format Created On tanggal
    const createdOnRaw = item["Created On"];
    const createdOn = formatDateDDMMMYYYY(createdOnRaw);

    // Cost calculation
    let cost = "-";
    const planNum = Number(item["Total sum (plan)"]) || 0;
    const actualNum = Number(item["Total sum (actual)"]) || 0;
    const diff = (planNum - actualNum) / 16500;
    if (diff >= 0) cost = diff.toFixed(2);

    // Include dan Exclude
    let includeVal = cost;
    if (item.Reman && item.Reman.toLowerCase().includes("reman")) {
      includeVal = cost === "-" ? "-" : (parseFloat(cost) * 0.25).toFixed(2);
    }
    let excludeVal = includeVal;
    if (item["Order Type"] && item["Order Type"].toUpperCase() === "PM38") excludeVal = "-";

    // Planning dan Status AMT dari planningData (lookup Order)
    const pl = planningData.find(p => p.Order === item.Order);
    const planningVal = pl ? formatDateDDMMMYYYY(pl["Event Start"]) : "";
    const statusAMTVal = pl ? (pl.Status || "") : "";

    // CPH logic
    let cphVal = "";
    if (item.Description && item.Description.startsWith("JR")) {
      cphVal = "External Job";
    } else if (item.MAT && data2Data.length) {
      const d2 = data2Data.find(d => d.MAT === item.MAT);
      cphVal = d2 ? (d2.CPH || "") : "";
    }

    // Section dari data1Data (lookup Room)
    let sectionVal = "";
    if (item.Room && data1Data.length) {
      const d1 = data1Data.find(d => d.Room === item.Room);
      sectionVal = d1 ? (d1.Section || "") : "";
    }

    // Status Part dan Aging dari sum57Data (lookup Order)
    let statusPart = "";
    let agingVal = "";
    if (item.Order && sum57Data.length) {
      const s57 = sum57Data.find(s => s.Order === item.Order);
      statusPart = s57 ? (s57["Part Complete"] || "") : "";
      agingVal = s57 ? (s57.Aging || "") : "";
    }

    return {
      Room: item.Room || "",
      "Order Type": item["Order Type"] || "",
      Order: item.Order || "",
      Description: item.Description || "",
      "Created On": createdOn,
      "User Status": item["User Status"] || "",
      MAT: item.MAT || "",
      CPH: cphVal,
      Section: sectionVal,
      "Status Part": statusPart,
      Aging: agingVal,
      Month: item.Month || "",
      Cost: cost,
      Reman: item.Reman || "",
      Include: includeVal,
      Exclude: excludeVal,
      Planning: planningVal,
      "Status AMT": statusAMTVal
    };
  });

  // Restore user edits dari localStorage
  try {
    const lsRaw = localStorage.getItem(UI_LS_KEY);
    if (lsRaw) {
      const ls = JSON.parse(lsRaw);
      ls.userEdits.forEach(edit => {
        const idx = mergedData.findIndex(r => r.Order === edit.Order);
        if (idx >= 0) {
          mergedData[idx] = { ...mergedData[idx], ...edit };
        }
      });
    }
  } catch {}

  updateMonthFilterOptions();
}

// ======= RENDER TABLE =======
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

// ======= ATTACH EDIT & DELETE EVENTS =======
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.onclick = () => startEdit(btn.dataset.order);
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.onclick = () => deleteOrder(btn.dataset.order);
  });
}

// ======= START EDIT =======
function startEdit(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[idx];
  const d = mergedData[idx];

  tr.innerHTML = `
    <td><input type="text" value="${d.Room}" data-field="Room"></td>
    <td><input type="text" value="${d["Order Type"]}" data-field="Order Type"></td>
    <td>${d.Order}</td>
    <td><input type="text" value="${d.Description}" data-field="Description"></td>
    <td><input type="date" value="${formatDateISO(d["Created On"])}" data-field="Created On"></td>
    <td><input type="text" value="${d["User Status"]}" data-field="User Status"></td>
    <td><input type="text" value="${d.MAT}" data-field="MAT"></td>
    <td><input type="text" value="${d.CPH}" data-field="CPH"></td>
    <td><input type="text" value="${d.Section}" data-field="Section"></td>
    <td><input type="text" value="${d["Status Part"]}" data-field="Status Part"></td>
    <td><input type="text" value="${d.Aging}" data-field="Aging"></td>
    <td><input type="text" value="${d.Month}" data-field="Month"></td>
    <td><input style="text-align:right;" type="text" value="${d.Cost}" data-field="Cost"></td>
    <td><input type="text" value="${d.Reman}" data-field="Reman"></td>
    <td><input style="text-align:right;" type="text" value="${d.Include}" data-field="Include"></td>
    <td><input style="text-align:right;" type="text" value="${d.Exclude}" data-field="Exclude"></td>
    <td><input type="text" value="${d.Planning}" data-field="Planning"></td>
    <td><input type="text" value="${d["Status AMT"]}" data-field="Status AMT"></td>
    <td>
      <button class="action-btn save-btn" data-order="${order}">Save</button>
      <button class="action-btn cancel-btn" data-order="${order}">Cancel</button>
    </td>
  `;

  tr.querySelector(".save-btn").onclick = () => saveEdit(order);
  tr.querySelector(".cancel-btn").onclick = () => cancelEdit(order);
}

// ======= CANCEL EDIT =======
function cancelEdit(order) {
  renderTable(mergedData);
}

// ======= SAVE EDIT =======
function saveEdit(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx === -1) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[idx];
  const inputs = tr.querySelectorAll("input[data-field]");

  inputs.forEach(input => {
    const field = input.dataset.field;
    let val = input.value;
    if (field === "Created On") val = formatDateDDMMMYYYY(val);
    mergedData[idx][field] = val;
  });

  saveUserEdits();
  renderTable(mergedData);
}

// ======= SAVE USER EDITS KE localStorage =======
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
    console.error("Error saving edits:", e);
  }
}

// ======= DELETE ORDER =======
function deleteOrder(order) {
  const idx = mergedData.findIndex(r => r.Order === order);
  if (idx === -1) return;
  if (!confirm(`Hapus data order ${order}?`)) return;
  mergedData.splice(idx, 1);
  saveUserEdits();
  renderTable(mergedData);
}

// ======= FILTER DATA =======
function filterData() {
  const room = document.getElementById("filter-room").value.trim().toLowerCase();
  const order = document.getElementById("filter-order").value.trim().toLowerCase();
  const cph = document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat = document.getElementById("filter-mat").value.trim().toLowerCase();
  const section = document.getElementById("filter-section").value.trim().toLowerCase();
  const month = document.getElementById("filter-month").value.trim().toLowerCase();

  const filtered = mergedData.filter(row => {
    return (
      (!room || (row.Room && row.Room.toLowerCase().includes(room))) &&
      (!order || (String(row.Order || "").toLowerCase().includes(order))) &&
      (!cph || (row.CPH && row.CPH.toLowerCase().includes(cph))) &&
      (!mat || (row.MAT && row.MAT.toLowerCase().includes(mat))) &&
      (!section || (row.Section && row.Section.toLowerCase().includes(section))) &&
      (!month || (row.Month && row.Month.toLowerCase() === month))
    );
  });

  renderTable(filtered);
}
// ======= RESET FILTERS =======
function resetFilters() {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-month").value = "";
  renderTable(mergedData);
}

// ======= UPDATE MONTH SELECT OPTIONS =======
function updateMonthFilterOptions() {
  const monthSelect = document.getElementById("filter-month");
  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m)));
  months.sort();
  // Clear existing except first option
  for (let i = monthSelect.options.length - 1; i > 0; i--) {
    monthSelect.remove(i);
  }
  months.forEach(m => {
    const option = document.createElement("option");
    option.value = m.toLowerCase();
    option.textContent = m;
    monthSelect.appendChild(option);
  });
}

// ======= SAVE TO JSON =======
function saveToJSON() {
  if (!mergedData.length) {
    alert("Tidak ada data untuk disimpan.");
    return;
  }
  const dataStr = JSON.stringify({ mergedData, savedAt: new Date().toISOString() }, null, 2);
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

// ======= LOAD FROM JSON FILE =======
async function loadFromJSON(file) {
  if (!file) return;
  const text = await file.text();
  try {
    const json = JSON.parse(text);
    if (!json.mergedData) throw new Error("File JSON tidak valid");
    mergedData = json.mergedData;
    renderTable(mergedData);
    updateMonthFilterOptions();
  } catch (err) {
    alert("File JSON tidak valid atau format salah.");
  }
}

// ======= HANDLE FILE UPLOAD =======
async function handleUpload(file, type) {
  if (!file) {
    alert("Pilih file dulu ya.");
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
      default: alert("Tipe file tidak dikenali."); return;
    }
    alert(`${type} berhasil diupload, baris: ${jsonData.length}`);
  } catch (err) {
    alert(`Gagal upload ${type}: ${err.message}`);
  }
}

// ======= SETUP MENU =======
function setupMenu() {
  const menuItems = document.querySelectorAll(".sidebar .menu-item");
  const contentSections = document.querySelectorAll(".content-section");

  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      item.classList.add("active");

      const menuId = item.dataset.menu;
      contentSections.forEach(sec => {
        sec.classList.toggle("active", sec.id === menuId);
      });
    });
  });
}

// ======= INIT =======
function init() {
  document.getElementById("upload-btn").onclick = async () => {
    const fileInput = document.getElementById("file-input");
    const fileType = document.getElementById("file-select").value;
    const file = fileInput.files[0];
    await handleUpload(file, fileType);
  };

  document.getElementById("filter-btn").onclick = filterData;
  document.getElementById("reset-btn").onclick = resetFilters;
  document.getElementById("refresh-btn").onclick = () => {
    mergeData();
    renderTable(mergedData);
  };
  document.getElementById("save-btn").onclick = saveToJSON;
  document.getElementById("load-btn").onclick = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";
    input.onchange = async (e) => {
      if (e.target.files.length) await loadFromJSON(e.target.files[0]);
    };
    input.click();
  };

  setupMenu();

  renderTable([]);
}

window.onload = init;

