/****************************************************
 * Ndarboe.net - script1.js (Full Revisi)
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
const MERGED_LS_KEY = "ndarboe_merged_v1";
const DATASETS_LS_KEY = "ndarboe_datasets_v1";

/* ===================== SIMPAN & LOAD MERGEDDATA ===================== */
function saveMergedData() {
  try {
    localStorage.setItem(MERGED_LS_KEY, JSON.stringify(mergedData));
  } catch (e) {
    console.error("Gagal simpan mergedData:", e);
  }
}

function loadMergedData() {
  try {
    const raw = localStorage.getItem(MERGED_LS_KEY);
    if (raw) {
      mergedData = JSON.parse(raw);
      console.log("mergedData dimuat dari localStorage");
    }
  } catch (e) {
    console.error("Gagal load mergedData:", e);
  }
}

/* ===================== SIMPAN & LOAD DATASET MENTAH ===================== */
function saveAllDatasets() {
  try {
    localStorage.setItem(DATASETS_LS_KEY, JSON.stringify({
      iw39Data, sum57Data, planningData, data1Data, data2Data, budgetData
    }));
  } catch (e) {
    console.error("Gagal simpan datasets:", e);
  }
}

function loadAllDatasets() {
  try {
    const raw = localStorage.getItem(DATASETS_LS_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      iw39Data = saved.iw39Data || [];
      sum57Data = saved.sum57Data || [];
      planningData = saved.planningData || [];
      data1Data = saved.data1Data || [];
      data2Data = saved.data2Data || [];
      budgetData = saved.budgetData || [];
      console.log("datasets dimuat dari localStorage");
    }
  } catch (e) {
    console.error("Gagal load datasets:", e);
  }
}

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  loadAllDatasets();
  loadMergedData();
  setupMenu();
  setupButtons();
  renderTable(mergedData);
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
        sec.classList.toggle("active", sec.id === menuId);
      });
    });
  });
}

/* ===================== RENDER TABLE ===================== */
function renderTable(rows = mergedData) {
  const tbody = document.querySelector("#data-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  rows.forEach((row, index) => {
    const tr = document.createElement("tr");

    Object.values(row).forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.appendChild(td);
    });

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
        const index = parseInt(this.dataset.index, 10);
        mergedData[index].Month  = tr.querySelector(".edit-month").value;
        mergedData[index].Cost   = tr.querySelector(".edit-cost").value;
        mergedData[index].Reman  = tr.querySelector(".edit-reman").value;
        saveUserEdits();
        saveMergedData();
        renderTable();
      });

      tr.querySelector(".cancel-btn").addEventListener("click", function () {
        renderTable();
      });
    });
  });

  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.addEventListener("click", function () {
      const index = parseInt(this.dataset.index, 10);
      if (confirm("Yakin mau hapus data ini?")) {
        mergedData.splice(index, 1);
        saveUserEdits();
        saveMergedData();
        renderTable();
      }
    });
  });
}

/* ===================== MERGE ===================== */
function mergeData() {
  if (!iw39Data.length) {
    loadAllDatasets(); // coba load dari localStorage
  }

  if (!iw39Data.length) {
    alert("Upload data IW39 dulu sebelum refresh.");
    return;
  }

  // ... isi mergeData() lama kamu ...

  updateMonthFilterOptions();
  saveMergedData();
  saveAllDatasets();
}

/* ===================== SAVE USER EDITS ===================== */
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

/* ===================== BUTTON WIRING ===================== */
function setupButtons() {
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) uploadBtn.onclick = () => { handleUpload(); saveAllDatasets(); };

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

  const addOrderBtn = document.getElementById("add-order-btn");
  if (addOrderBtn) addOrderBtn.onclick = addOrders;
}
