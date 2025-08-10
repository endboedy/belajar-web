// ----- Pastikan kamu sudah load XLSX.js di HTML -----
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

//// GLOBAL DATA ////
let IW39 = [];
let Data1 = {};
let Data2 = {};
let SUM57 = {};
let Planning = {};

let dataLembarKerja = [];

// ----- Format angka 1 decimal ----- 
function formatNumber(num) {
  return Number(num).toFixed(1);
}

// ----- Parse multi file Excel sekaligus ----- 
async function parseMultipleExcelFiles(files) {
  // Reset semua data dulu
  IW39 = [];
  Data1 = {};
  Data2 = {};
  SUM57 = {};
  Planning = {};

  for (const file of files) {
    await parseExcelFile(file);
  }
}

// ----- Parse 1 file Excel ----- 
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(event) {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (workbook.SheetNames.includes('IW39')) {
          IW39 = XLSX.utils.sheet_to_json(workbook.Sheets['IW39']);
        }

        if (workbook.SheetNames.includes('Data1')) {
          const data1Arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data1']);
          data1Arr.forEach(row => {
            if(row.Order && row.Section) Data1[row.Order] = row.Section;
          });
        }

        if (workbook.SheetNames.includes('Data2')) {
          const data2Arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data2']);
          data2Arr.forEach(row => {
            if(row.MAT && row.CPH) Data2[row.MAT] = row.CPH;
          });
        }

        if (workbook.SheetNames.includes('SUM57')) {
          const sumArr = XLSX.utils.sheet_to_json(workbook.Sheets['SUM57']);
          sumArr.forEach(row => {
            if(row.Order) {
              SUM57[row.Order] = { StatusPart: row.StatusPart || '', Aging: row.Aging || '' };
            }
          });
        }

        if (workbook.SheetNames.includes('Planning')) {
          const planArr = XLSX.utils.sheet_to_json(workbook.Sheets['Planning']);
          planArr.forEach(row => {
            if(row.Order) {
              Planning[row.Order] = { Planning: row.Planning || '', StatusAMT: row.StatusAMT || '' };
            }
          });
        }

        resolve();
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ----- Build dataLembarKerja dari orders input dan lookup ----- 
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    const orderKey = (row.Order || "").toLowerCase();

    // Lookup IW39 by Order (case-insensitive)
    const iw = IW39.find(i => i.Order && i.Order.toLowerCase() === orderKey) || {};

    // Isi data dari IW39 dan lain-lain
    row.Room = iw.Room || row.Room || "";
    row.OrderType = iw.OrderType || row.OrderType || "";
    row.Description = iw.Description || row.Description || "";
    row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
    row.UserStatus = iw.UserStatus || row.UserStatus || "";
    row.MAT = iw.MAT || row.MAT || "";

    // CPH logic
    if ((row.Description || "").substring(0,2).toUpperCase() === "JR") {
      row.CPH = "JR";
    } else {
      row.CPH = Data2[row.MAT] || "";
    }

    // Section from Data1 by Order
    row.Section = Data1[row.Order] || "";

    // StatusPart & Aging from SUM57 by Order
    if (SUM57[row.Order]) {
      row.StatusPart = SUM57[row.Order].StatusPart || "";
      row.Aging = SUM57[row.Order].Aging || "";
    } else {
      row.StatusPart = "";
      row.Aging = "";
    }

    // Cost calculation
    if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const costCalc = (iw.TotalPlan - iw.TotalActual) / 16500;
      row.Cost = costCalc < 0 ? "-" : costCalc;
    } else {
      row.Cost = "-";
    }

    // Include calculation
    if ((row.Reman || "").toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    // Exclude calculation
    if ((row.OrderType || "").toUpperCase() === "PM38") {
      row.Exclude = "-";
    } else {
      row.Exclude = row.Include;
    }

    // Planning & StatusAMT lookup
    if (Planning[row.Order]) {
      row.Planning = Planning[row.Order].Planning || "";
      row.StatusAMT = Planning[row.Order].StatusAMT || "";
    } else {
      row.Planning = "";
      row.StatusAMT = "";
    }

    return row;
  });
}

// ----- Render tabel ----- 
const outputTableBody = document.querySelector("#output-table tbody");

function renderTable(data) {
  outputTableBody.innerHTML = "";

  if (data.length === 0) {
    outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center; font-style:italic; color:#888;">Tidak ada data sesuai filter.</td></tr>`;
    return;
  }

  // Cari duplikat Order (case-insensitive)
  const ordersLower = data.map(d => d.Order.toLowerCase());
  const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

  data.forEach(row => {
    const tr = document.createElement("tr");
    if (duplicates.includes(row.Order.toLowerCase())) tr.classList.add("duplicate");

    function createTD(text, className) {
      const td = document.createElement("td");
      if(className) td.classList.add(className);
      td.textContent = text;
      return td;
    }

    tr.appendChild(createTD(row.Room));
    tr.appendChild(createTD(row.OrderType));
    tr.appendChild(createTD(row.Order));
    tr.appendChild(createTD(row.Description));
    tr.appendChild(createTD(row.CreatedOn));
    tr.appendChild(createTD(row.UserStatus));
    tr.appendChild(createTD(row.MAT));
    tr.appendChild(createTD(row.CPH));
    tr.appendChild(createTD(row.Section));
    tr.appendChild(createTD(row.StatusPart));
    tr.appendChild(createTD(row.Aging));

    // Month editable
    const tdMonth = document.createElement("td");
    tdMonth.classList.add("editable");
    tdMonth.textContent = row.Month || "";
    tdMonth.title = "Klik untuk edit bulan";
    tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
    tr.appendChild(tdMonth);

    // Cost number right aligned
    tr.appendChild(createTD(typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost, "cost"));

    // Reman editable
    const tdReman = document.createElement("td");
    tdReman.classList.add("editable");
    tdReman.textContent = row.Reman || "";
    tdReman.title = "Klik untuk edit Reman";
    tdReman.addEventListener("click", () => editReman(tdReman, row));
    tr.appendChild(tdReman);

    // Include number right aligned
    tr.appendChild(createTD(typeof row.Include === "number" ? formatNumber(row.Include) : row.Include, "include"));

    // Exclude number right aligned
    tr.appendChild(createTD(typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude, "exclude"));

    tr.appendChild(createTD(row.Planning));
    tr.appendChild(createTD(row.StatusAMT));

    // Action buttons: Edit (focus Month & Reman) & Delete
    const tdAction = document.createElement("td");

    const btnEdit = document.createElement("button");
    btnEdit.textContent = "Edit";
    btnEdit.classList.add("btn-action", "btn-edit");
    btnEdit.addEventListener("click", () => {
      editMonthAction(row);
      editRemanAction(row);
    });
    tdAction.appendChild(btnEdit);

    const btnDelete = document.createElement("button");
    btnDelete.textContent = "Delete";
    btnDelete.classList.add("btn-action", "btn-delete");
    btnDelete.addEventListener("click", () => {
      if(confirm(`Hapus order ${row.Order}?`)) {
        dataLembarKerja = dataLembarKerja.filter(d => d.Order.toLowerCase() !== row.Order.toLowerCase());
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToStorage();
      }
    });
    tdAction.appendChild(btnDelete);

    tr.appendChild(tdAction);

    outputTableBody.appendChild(tr);
  });
}

// ----- Edit Month & Reman inline ----- 
function editMonth(td, row) {
  const select = document.createElement("select");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  months.forEach(m => {
    const option = document.createElement("option");
    option.value = m;
    option.textContent = m;
    if (m === row.Month) option.selected = true;
    select.appendChild(option);
  });

  select.addEventListener("change", () => {
    row.Month = select.value;
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    saveDataToStorage();
  });

  select.addEventListener("blur", () => renderTable(dataLembarKerja));

  td.textContent = "";
  td.appendChild(select);
  select.focus();
}

function editReman(td, row) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = row.Reman || "";

  input.addEventListener("keydown", e => {
    if (e.key === "Enter") {
      row.Reman = input.value.trim();
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToStorage();
    } else if (e.key === "Escape") {
      renderTable(dataLembarKerja);
    }
  });

  input.addEventListener("blur", () => {
    row.Reman = input.value.trim();
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    saveDataToStorage();
  });

  td.textContent = "";
  td.appendChild(input);
  input.focus();
}

function editMonthAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      editMonth(tr.children[11], row);
    }
  });
}
function editRemanAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      editReman(tr.children[13], row);
    }
  });
}

// ----- Validasi order input ----- 
function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ----- Add order multi input ----- 
const addOrderBtn = document.getElementById("add-order-btn");
const addOrderInput = document.getElementById("add-order-input");
const addOrderStatus = document.getElementById("add-order-status");

addOrderBtn.addEventListener("click", () => {
  let rawInput = addOrderInput.value.trim();
  if (!rawInput) {
    alert("Masukkan minimal satu Order bro!");
    return;
  }

  let orders = rawInput.split(/[\s,\n]+/).map(s => s.trim()).filter(s => s.length > 0);

  let addedCount = 0;
  let skippedOrders = [];
  let invalidOrders = [];

  orders.forEach(order => {
    if (!isValidOrder(order)) {
      invalidOrders.push(order);
      return;
    }
    const exists = dataLembarKerja.some(d => d.Order.toLowerCase() === order.toLowerCase());
    if (!exists) {
      dataLembarKerja.push({
        Room: "",
        OrderType: "",
        Order: order,
        Description: "",
        CreatedOn: "",
        UserStatus: "",
        MAT: "",
        CPH: "",
        Section: "",
        StatusPart: "",
        Aging: "",
        Month: "",
        Cost: "-",
        Reman: "",
        Include: "-",
        Exclude: "-",
        Planning: "",
        StatusAMT: ""
      });
      addedCount++;
    } else {
      skippedOrders.push(order);
    }
  });

  buildDataLembarKerja();
  renderTable(dataLembarKerja);
  addOrderInput.value = "";

  let msg = `${addedCount} Order berhasil ditambahkan.`;
  if (invalidOrders.length) {
    msg += ` Order tidak valid (ada titik atau koma): ${invalidOrders.join(", ")}.`;
  }
  if (skippedOrders.length) {
    msg += ` Order sudah ada dan tidak dimasukkan ulang: ${skippedOrders.join(", ")}.`;
  }
  addOrderStatus.textContent = msg;
  saveDataToStorage();
});

// ----- Filter input fields ----- 
const filterRoom = document.getElementById("filter-room");
const filterOrder = document.getElementById("filter-order");
const filterCPH = document.getElementById("filter-cph");
const filterMAT = document.getElementById("filter-mat");
const filterSection = document.getElementById("filter-section");

const filterInputs = [filterRoom, filterOrder, filterCPH, filterMAT, filterSection];

filterInputs.forEach(input => {
  input.addEventListener("input", () => {
    const filtered = dataLembarKerja.filter(row => {
      const roomMatch = row.Room.toLowerCase().includes(filterRoom.value.toLowerCase());
      const orderMatch = row.Order.toLowerCase().includes(filterOrder.value.toLowerCase());
      const cphMatch = row.CPH.toLowerCase().includes(filterCPH.value.toLowerCase());
      const matMatch = row.MAT.toLowerCase().includes(filterMAT.value.toLowerCase());
      const sectionMatch = row.Section.toLowerCase().includes(filterSection.value.toLowerCase());
      return roomMatch && orderMatch && cphMatch && matMatch && sectionMatch;
    });
    renderTable(filtered);
  });
});

// ----- Save & Load dari localStorage ----- 
const saveBtn = document.getElementById("save-btn");
const loadBtn = document.getElementById("load-btn");

function saveDataToStorage() {
  localStorage.setItem("dataLembarKerja", JSON.stringify(dataLembarKerja));
}
function loadDataFromStorage() {
  const stored = localStorage.getItem("dataLembarKerja");
  if(stored) {
    dataLembarKerja = JSON.parse(stored);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  }
}

saveBtn.addEventListener("click", () => {
  saveDataToStorage();
  alert("Data berhasil disimpan ke browser localStorage!");
});

loadBtn.addEventListener("click", () => {
  loadDataFromStorage();
});

// ----- Upload multi file & progress ----- 
const fileInput = document.getElementById("file-input");
const uploadBtn = document.getElementById("upload-btn");
const progressContainer = document.getElementById("progress-container");
const progressBar = document.getElementById("upload-progress");
const progressText = document.getElementById("progress-text");
const uploadStatus = document.getElementById("upload-status");

fileInput.setAttribute("multiple", true);

uploadBtn.addEventListener("click", async () => {
  const files = fileInput.files;
  if (!files || files.length === 0) {
    alert("Pilih minimal 1 file Excel untuk diupload bro!");
    return;
  }

  progressContainer.classList.remove("hidden");
  progressBar.value = 0;
  progressText.textContent = "0%";
  uploadStatus.textContent = "";

  try {
    // Karena FileReader async dan per file, kita buat progress manual
    for(let i = 0; i < files.length; i++) {
      await parseExcelFile(files[i]);
      let percent = Math.round(((i + 1) / files.length) * 100);
      progressBar.value = percent;
      progressText.textContent = `${percent}%`;
    }
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    uploadStatus.textContent = `Upload selesai, total ${files.length} file berhasil di-load.`;
  } catch (err) {
    alert("Error saat membaca file Excel: " + err.message);
    uploadStatus.textContent = `Upload gagal: ${err.message}`;
  }
  progressContainer.classList.add("hidden");
});

// ----- Menu sidebar ----- 
const menuItems = document.querySelectorAll(".menu-item");
const contentSections = document.querySelectorAll(".content-section");

menuItems.forEach(item => {
  item.addEventListener("click", () => {
    menuItems.forEach(i => i.classList.remove("active"));
    contentSections.forEach(section => section.classList.remove("active"));

    item.classList.add("active");
    const menuName = item.getAttribute("data-menu");
    const activeSection = document.getElementById(menuName);
    if(activeSection) activeSection.classList.add("active");
  });
});

// ----- Init load data kalau ada -----
loadDataFromStorage();
