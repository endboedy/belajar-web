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
async function parseFile(file, jenis) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        // Cari sheet sesuai jenis file
        let sheetName = "";
        if (workbook.SheetNames.includes(jenis)) {
          sheetName = jenis;
        } else {
          sheetName = workbook.SheetNames[0];
          console.warn(`Sheet "${jenis}" tidak ditemukan, pakai sheet pertama: ${sheetName}`);
        }

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          reject(new Error(`Sheet "${sheetName}" tidak ditemukan di file.`));
          return;
        }

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
    // Coba parse dengan Date
    d = new Date(dt);
    if (isNaN(d)) {
      // Coba parsing manual untuk format mm/dd/yyyy atau dd/mm/yyyy
      const parts = dt.match(/(\d+)/g);
      if (parts && parts.length >= 3) {
        // Coba interpretasi: jika bulan lebih dari 12 kemungkinan format dd/mm/yyyy
        const monthNum = parseInt(parts[0], 10);
        const dayNum = parseInt(parts[1], 10);
        if (monthNum > 12) {
          d = new Date(parts[2], monthNum - 1, dayNum);
        } else {
          d = new Date(parts[2], dayNum - 1, monthNum);
        }
      }
    }
  } else if (dt instanceof Date) {
    d = dt;
  } else if (typeof dt === "number") {
    // Excel serial date
    d = XLSX.SSF.parse_date_code(dt);
    if (d) {
      d = new Date(d.y, d.m - 1, d.d);
    } else {
      return "";
    }
  } else {
    return "";
  }
  if (!d || isNaN(d)) return "";

  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${d.getDate().toString().padStart(2,"0")} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

// -------- Format date yyyy-mm-dd for input type=date ----------
function formatDateISO(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return d.toISOString().split("T")[0];
}

// -------- Merge function to combine data from multiple sheets --------
function mergeData() {
  // Base from iw39Data order
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

  // Merge CPH from data2Data by logic:
  mergedData.forEach(md => {
    if(md.Description && md.Description.startsWith("JR")) {
      md.CPH = "External Job";
    } else {
      const d2 = data2Data.find(d => d.MAT && d.MAT.trim() === md.MAT.trim());
      md.CPH = d2 ? d2.CPH || "" : "";
    }
  });

  // Merge Section from data1Data by Room
  mergedData.forEach(md => {
    const d1 = data1Data.find(d => d.Room && d.Room.trim() === md.Room.trim());
    md.Section = d1 ? d1.Section || "" : "";
  });

  // Merge Aging & Status Part from sum57Data by Order
  mergedData.forEach(md => {
    const s57 = sum57Data.find(s => s.Order === md.Order);
    if (s57) {
      md.Aging = s57.Aging || "";
      md["Status Part"] = s57["Part Complete"] || "";
    }
  });

  // Merge Planning data from Planning by Order
  mergedData.forEach(md => {
    const pl = planningData.find(p => p.Order === md.Order);
    if (pl) {
      md.Planning = pl["Event Start"] || "";
      md["Status AMT"] = pl.Status || "";
    }
  });

  // Hitung Cost, Include, Exclude
  mergedData.forEach(md => {
    const iw39Item = iw39Data.find(i => i.Order === md.Order);
    if (iw39Item) {
      const plan = parseFloat(iw39Item["Total sum (plan)"] || 0);
      const actual = parseFloat(iw39Item["Total sum (actual)"] || 0);
      let costCalc = (plan - actual)/16500;
      if (isNaN(costCalc) || costCalc < 0) costCalc = "-";
      else costCalc = costCalc.toFixed(2);
      md.Cost = costCalc;

      if(md.Reman && md.Reman.toLowerCase().includes("reman")) {
        md.Include = (costCalc === "-") ? "-" : (costCalc * 0.25).toFixed(2);
      } else {
        md.Include = costCalc;
      }

      if(md["Order Type"] === "PM38") {
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

// -------- Render mergedData to table ----------
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#999">Tidak ada data</td></tr>`;
    return;
  }
  data.forEach((row, idx) => {
    const tr = document.createElement("tr");

    // Format Create On & Planning dates
    const createdOnFormatted = formatDateDDMMMYYYY(row["Created On"]);
    const planningFormatted = formatDateDDMMMYYYY(row.Planning);

    tr.innerHTML = `
      <td>${row.Room}</td>
      <td>${row["Order Type"]}</td>
      <td>${row.Order}</td>
      <td>${row.Description}</td>
      <td>${createdOnFormatted}</td>
      <td>${row["User Status"]}</td>
      <td>${row.MAT}</td>
      <td>${row.CPH}</td>
      <td>${row.Section}</td>
      <td>${row["Status Part"]}</td>
      <td>${row.Aging}</td>
      <td>${row.Month}</td>
      <td style="text-align:right;">${row.Cost}</td>
      <td>${row.Reman || ""}</td>
      <td style="text-align:right;">${row.Include}</td>
      <td style="text-align:right;">${row.Exclude}</td>
      <td>${planningFormatted}</td>
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

  // Build month options for dropdown (ambil unique dari mergedData)
  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort();

  const monthOptions = months.map(m => {
    const selected = (m === rowData.Month) ? "selected" : "";
    return `<option value="${m}" ${selected}>${m}</option>`;
  }).join("");

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
    <td>
      <select data-field="Month">
        <option value="">--Select Month--</option>
        ${monthOptions}
      </select>
    </td>
    <td style="width:80px;"><input type="text" value="${rowData.Cost}" data-field="Cost" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${rowData.Reman || ""}" data-field="Reman" /></td>
    <td style="width:80px;"><input type="text" value="${rowData.Include}" data-field="Include" readonly style="text-align:right;background:#eee;"/></td>
    <td style="width:80px;"><input type="text" value="${rowData.Exclude}" data-field="Exclude" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="date" value="${formatDateISO(rowData.Planning)}" data-field="Planning" /></td>
    <td><input type="text" value="${rowData["Status AMT"]}" data-field="Status AMT" /></td>
    <td>
      <button class="action-btn save-btn" data-order="${order}">Save</button>
      <button class="action-btn cancel-btn" data-order="${order}">Cancel</button>
    </td>
  `;

  // Set selected month value explicitly in case blank
  const monthSelect = tr.querySelector("select[data-field='Month']");
  monthSelect.value = rowData.Month;

  // Attach save/cancel handlers
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

  // Read inputs
  const inputs = tr.querySelectorAll("input[data-field], select[data-field]");
  inputs.forEach(input => {
    const field = input.dataset.field;
    let val = input.value;
    if (field === "Created On" || field === "Planning") {
      // Format date to ISO for storage
      val = val || "";
    }
    mergedData[rowIndex][field] = val;
  });

  // Save edits to localStorage
  saveUserEdits();

  // Re-merge for calculated fields and refresh table
  mergeData();
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

  let filtered = mergedData.filter(row => {
    if (roomFilter && !row.Room.toLowerCase().includes(roomFilter)) return false;
    if (orderFilter && !row.Order.toLowerCase().includes(orderFilter)) return false;
    if (cphFilter && !row.CPH.toLowerCase().includes(cphFilter)) return false;
    if (matFilter && !row.MAT.toLowerCase().includes(matFilter)) return false;
    if (sectionFilter && !row.Section.toLowerCase().includes(sectionFilter)) return false;
    if (monthFilter && row.Month.toLowerCase() !== monthFilter) return false;
    return true;
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

// -------- Update Month filter options dropdown ----------
function updateMonthFilterOptions() {
  const monthSelect = document.getElementById("filter-month");
  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort();
  monthSelect.innerHTML = `<option value="">-- All --</option>` + months.map(m => `<option value="${m}">${m}</option>`).join("");
}

// -------- Upload handler ----------
async function handleUpload() {
  const fileInput = document.getElementById("file-input");
  const jenis = document.getElementById("file-select").value;

  if (!fileInput.files.length) {
    alert("Pilih file terlebih dahulu!");
    return;
  }
  const file = fileInput.files[0];

  try {
    const json = await parseFile(file, jenis);

    switch(jenis) {
      case "IW39": iw39Data = json; break;
      case "SUM57": sum57Data = json; break;
      case "Planning": planningData = json; break;
      case "Data1": data1Data = json; break;
      case "Data2": data2Data = json; break;
      case "Budget": budgetData = json; break;
      default: alert("Jenis file tidak dikenali!"); return;
    }

    document.getElementById("upload-status").textContent = `File ${jenis} berhasil diupload dengan ${json.length} baris.`;

    // Jika sudah upload IW39 dan Planning dan SUM57 dan Data1 dan Data2, merge data
    if (iw39Data.length && sum57Data.length && planningData.length && data1Data.length && data2Data.length) {
      mergeData();
      renderTable(mergedData);
    }
  } catch(err) {
    alert("Gagal membaca file: " + err.message);
  }
}

// -------- Setup menu click ----------
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

// -------- Add orders (dummy example, you can customize) --------
function addOrders() {
  const input = document.getElementById("add-order-input").value.trim();
  const status = document.getElementById("add-order-status");

  if (!input) {
    status.textContent = "Masukkan order terlebih dahulu.";
    return;
  }

  const orders = input.split(/[\s,]+/);
  orders.forEach(o => {
    if (!mergedData.some(d => d.Order === o)) {
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
  });
  status.textContent = `Berhasil menambah ${orders.length} order.`;
  renderTable(mergedData);
  document.getElementById("add-order-input").value = "";
}

// -------- Event listeners setup --------
function setupEventListeners() {
  document.getElementById("upload-btn").onclick = handleUpload;
  document.getElementById("filter-btn").onclick = filterData;
  document.getElementById("reset-btn").onclick = resetFilters;
  document.getElementById("refresh-btn").onclick = () => renderTable(mergedData);
  document.getElementById("save-btn").onclick = saveUserEdits;
  document.getElementById("load-btn").onclick = () => {
    try {
      const raw = localStorage.getItem(UI_LS_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        if (saved.userEdits) {
          saved.userEdits.forEach(edit => {
            const idx = mergedData.findIndex(r => r.Order === edit.Order);
            if (idx !== -1) mergedData[idx] = { ...mergedData[idx], ...edit };
          });
          renderTable(mergedData);
          alert("Data user edit berhasil dimuat.");
        }
      }
    } catch {
      alert("Gagal memuat data user edit.");
    }
  };
  document.getElementById("add-order-btn").onclick = addOrders;
}

// -------- Main init ----------
function init() {
  setupMenu();
  setupEventListeners();
  renderTable(mergedData);
}

init();
