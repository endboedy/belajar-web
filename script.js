// Pastikan library XLSX sudah load di HTML seperti ini:
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

// ===== GLOBAL DATA =====
let IW39 = [], Data1 = {}, Data2 = {}, SUM57 = {}, Planning = {};
let dataLembarKerja = [];

// ===== Format angka 1 decimal =====
function formatNumber(num) {
  return Number(num).toFixed(1);
}

// ===== Parsing file Excel & update lookup global =====
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(event) {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (workbook.SheetNames.includes('IW39')) {
          IW39 = XLSX.utils.sheet_to_json(workbook.Sheets['IW39']);
        } else IW39 = [];

        if (workbook.SheetNames.includes('Data1')) {
          const arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data1']);
          Data1 = {};
          arr.forEach(r => { if(r.Order && r.Section) Data1[r.Order.toString().toLowerCase()] = r.Section; });
        } else Data1 = {};

        if (workbook.SheetNames.includes('Data2')) {
          const arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data2']);
          Data2 = {};
          arr.forEach(r => { if(r.MAT && r.CPH) Data2[r.MAT.toString().toLowerCase()] = r.CPH; });
        } else Data2 = {};

        if (workbook.SheetNames.includes('SUM57')) {
          const arr = XLSX.utils.sheet_to_json(workbook.Sheets['SUM57']);
          SUM57 = {};
          arr.forEach(r => {
            if(r.Order) {
              SUM57[r.Order.toString().toLowerCase()] = {
                StatusPart: r.StatusPart || "",
                Aging: r.Aging || ""
              };
            }
          });
        } else SUM57 = {};

        if (workbook.SheetNames.includes('Planning')) {
          const arr = XLSX.utils.sheet_to_json(workbook.Sheets['Planning']);
          Planning = {};
          arr.forEach(r => {
            if(r.Order) {
              Planning[r.Order.toString().toLowerCase()] = {
                Planning: r.Planning || "",
                StatusAMT: r.StatusAMT || ""
              };
            }
          });
        } else Planning = {};

        resolve();
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ===== Build dataLembarKerja with lookup and formulas =====
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    const orderKey = row.Order ? row.Order.toString().toLowerCase() : "";

    // Cari data lengkap di IW39 berdasarkan Order key
    const iw = IW39.find(i => i.Order && i.Order.toString().toLowerCase() === orderKey) || {};

    // Ambil dari IW39 kalau ada
    row.Room = iw.Room || row.Room || "";
    row.OrderType = iw.OrderType || row.OrderType || "";
    row.Description = iw.Description || row.Description || "";
    row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
    row.UserStatus = iw.UserStatus || row.UserStatus || "";
    row.MAT = iw.MAT || row.MAT || "";

    // CPH: kalau 2 huruf pertama Description JR => JR, else lookup Data2 pakai MAT
    if ((row.Description || "").substring(0,2).toUpperCase() === "JR") {
      row.CPH = "JR";
    } else {
      row.CPH = Data2[(row.MAT || "").toString().toLowerCase()] || "";
    }

    // Section lookup Data1 by Order (case insensitive)
    row.Section = Data1[orderKey] || "";

    // Status Part & Aging lookup SUM57 by Order
    if (SUM57[orderKey]) {
      row.StatusPart = SUM57[orderKey].StatusPart || "";
      row.Aging = SUM57[orderKey].Aging || "";
    } else {
      row.StatusPart = "";
      row.Aging = "";
    }

    // Cost = (IW39.TotalPlan - IW39.TotalActual) / 16500, jika < 0 maka "-"
    if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const costCalc = (iw.TotalPlan - iw.TotalActual) / 16500;
      row.Cost = costCalc < 0 ? "-" : costCalc;
    } else {
      row.Cost = "-";
    }

    // Include: jika Reman = "Reman" maka Cost * 0.25, jika kosong sama dengan Cost
    if ((row.Reman || "").toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    // Exclude: jika OrderType = PM38 maka "-", else sama dengan Include
    if ((row.OrderType || "").toUpperCase() === "PM38") {
      row.Exclude = "-";
    } else {
      row.Exclude = row.Include;
    }

    // Planning & StatusAMT lookup dari Planning by Order
    if (Planning[orderKey]) {
      row.Planning = Planning[orderKey].Planning || "";
      row.StatusAMT = Planning[orderKey].StatusAMT || "";
    } else {
      row.Planning = "";
      row.StatusAMT = "";
    }

    return row;
  });

  // Setelah update data, simpan otomatis ke localStorage
  saveDataToLocalStorage();
}

// ===== Validasi order input =====
function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ===== Render Tabel Output Menu 2 =====
const outputTableBody = document.querySelector("#output-table tbody");

function renderTable(data) {
  const ordersLower = data.map(d => d.Order.toLowerCase());
  const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

  outputTableBody.innerHTML = "";
  if (data.length === 0) {
    outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center; font-style:italic; color:#888;">Tidak ada data sesuai filter.</td></tr>`;
    return;
  }

  data.forEach(row => {
    const tr = document.createElement("tr");
    if (duplicates.includes(row.Order.toLowerCase())) {
      tr.classList.add("duplicate");
    }

    // Room
    let td = document.createElement("td");
    td.textContent = row.Room;
    tr.appendChild(td);

    // OrderType
    td = document.createElement("td");
    td.textContent = row.OrderType;
    tr.appendChild(td);

    // Order
    td = document.createElement("td");
    td.textContent = row.Order;
    tr.appendChild(td);

    // Description
    td = document.createElement("td");
    td.textContent = row.Description;
    tr.appendChild(td);

    // CreatedOn
    td = document.createElement("td");
    td.textContent = row.CreatedOn;
    tr.appendChild(td);

    // UserStatus
    td = document.createElement("td");
    td.textContent = row.UserStatus;
    tr.appendChild(td);

    // MAT
    td = document.createElement("td");
    td.textContent = row.MAT;
    tr.appendChild(td);

    // CPH
    td = document.createElement("td");
    td.textContent = row.CPH;
    tr.appendChild(td);

    // Section
    td = document.createElement("td");
    td.textContent = row.Section;
    tr.appendChild(td);

    // StatusPart
    td = document.createElement("td");
    td.textContent = row.StatusPart;
    tr.appendChild(td);

    // Aging
    td = document.createElement("td");
    td.textContent = row.Aging;
    tr.appendChild(td);

    // Month (editable)
    td = document.createElement("td");
    td.classList.add("editable");
    td.textContent = row.Month || "";
    td.title = "Klik untuk edit bulan";
    td.addEventListener("click", () => editMonth(td, row));
    tr.appendChild(td);

    // Cost (right align, 1 decimal)
    td = document.createElement("td");
    td.classList.add("cost");
    td.textContent = typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost;
    tr.appendChild(td);

    // Reman (editable)
    td = document.createElement("td");
    td.classList.add("editable");
    td.textContent = row.Reman || "";
    td.title = "Klik untuk edit Reman";
    td.addEventListener("click", () => editReman(td, row));
    tr.appendChild(td);

    // Include (right align, 1 decimal)
    td = document.createElement("td");
    td.classList.add("include");
    td.textContent = typeof row.Include === "number" ? formatNumber(row.Include) : row.Include;
    tr.appendChild(td);

    // Exclude (right align, 1 decimal)
    td = document.createElement("td");
    td.classList.add("exclude");
    td.textContent = typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude;
    tr.appendChild(td);

    // Planning
    td = document.createElement("td");
    td.textContent = row.Planning;
    tr.appendChild(td);

    // StatusAMT
    td = document.createElement("td");
    td.textContent = row.StatusAMT;
    tr.appendChild(td);

    // Action (Edit & Delete)
    td = document.createElement("td");
    // Edit button
    const btnEdit = document.createElement("button");
    btnEdit.textContent = "Edit";
    btnEdit.classList.add("btn-action", "btn-edit");
    btnEdit.addEventListener("click", () => {
      editMonthAction(row);
      editRemanAction(row);
    });
    td.appendChild(btnEdit);

    // Delete button
    const btnDelete = document.createElement("button");
    btnDelete.textContent = "Delete";
    btnDelete.classList.add("btn-action", "btn-delete");
    btnDelete.addEventListener("click", () => {
      if (confirm(`Hapus order ${row.Order}?`)) {
        dataLembarKerja = dataLembarKerja.filter(d => d.Order.toLowerCase() !== row.Order.toLowerCase());
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
      }
    });
    td.appendChild(btnDelete);

    tr.appendChild(td);

    outputTableBody.appendChild(tr);
  });
}

// ===== Inline Edit Month =====
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
  });

  select.addEventListener("blur", () => renderTable(dataLembarKerja));

  td.textContent = "";
  td.appendChild(select);
  select.focus();
}

// ===== Inline Edit Reman =====
function editReman(td, row) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = row.Reman || "";

  input.addEventListener("keydown", e => {
    if (e.key === "Enter") {
      row.Reman = input.value.trim();
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
    } else if (e.key === "Escape") {
      renderTable(dataLembarKerja);
    }
  });

  input.addEventListener("blur", () => {
    row.Reman = input.value.trim();
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  });

  td.textContent = "";
  td.appendChild(input);
  input.focus();
}

// ===== Edit action via Edit button (focus Month & Reman) =====
function editMonthAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if (tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      const tdMonth = tr.children[11];
      editMonth(tdMonth, row);
    }
  });
}
function editRemanAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if (tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      const tdReman = tr.children[13];
      editReman(tdReman, row);
    }
  });
}

// ===== Add order multi input =====
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
});

// ===== Filter elements & filter action =====
const filterRoom = document.getElementById("filter-room");
const filterOrder = document.getElementById("filter-order");
const filterCPH = document.getElementById("filter-cph");
const filterMAT = document.getElementById("filter-mat");
const filterSection = document.getElementById("filter-section");
const filterBtn = document.getElementById("filter-btn");
const resetBtn = document.getElementById("reset-btn");

filterBtn.addEventListener("click", () => {
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

resetBtn.addEventListener("click", () => {
  filterRoom.value = "";
  filterOrder.value = "";
  filterCPH.value = "";
  filterMAT.value = "";
  filterSection.value = "";
  renderTable(dataLembarKerja);
});

// ===== Save & Load =====
const saveBtn = document.getElementById("save-btn");
const loadBtn = document.getElementById("load-btn");

// Save: export Excel file
saveBtn.addEventListener("click", () => {
  if (!dataLembarKerja.length) {
    alert("Data kosong, tidak ada yang disimpan.");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(dataLembarKerja);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "LembarKerja");
  XLSX.writeFile(wb, "LembarKerja.xlsx");
});

// Load: load data dari localStorage ke tabel
loadBtn.addEventListener("click", () => {
  loadDataFromLocalStorage();
  alert("Data berhasil dimuat dari penyimpanan lokal.");
});

// Simpan otomatis ke localStorage
function saveDataToLocalStorage() {
  localStorage.setItem("dataLembarKerja", JSON.stringify(dataLembarKerja));
}

// Load dari localStorage
function loadDataFromLocalStorage() {
  const saved = localStorage.getItem("dataLembarKerja");
  if (saved) {
    dataLembarKerja = JSON.parse(saved);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  } else {
    alert("Tidak ada data tersimpan di localStorage.");
  }
}

// ===== Upload file dan parsing =====
const fileInput = document.getElementById("file-input");
const uploadBtn = document.getElementById("upload-btn");
const progressContainer = document.getElementById("progress-container");
const uploadProgress = document.getElementById("upload-progress");
const progressText = document.getElementById("progress-text");
const uploadStatus = document.getElementById("upload-status");

uploadBtn.addEventListener("click", () => {
  const file = fileInput.files[0];
  if (!file) {
    alert("Pilih file Excel terlebih dahulu bro.");
    return;
  }
  // Mulai proses upload (read)
  progressContainer.classList.remove("hidden");
  uploadProgress.value = 0;
  progressText.textContent = "0%";
  uploadStatus.textContent = "";

  parseExcelFile(file).then(() => {
    uploadProgress.value = 100;
    progressText.textContent = "100%";
    uploadStatus.textContent = "File berhasil diupload dan data di-load.";
    // Setelah upload berhasil, rebuild data lookup agar lookup baru bisa jalan
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  }).catch(err => {
    uploadStatus.textContent = "Error saat memproses file: " + err.message;
  });
});

// ===== Sidebar Menu Switching =====
const menuItems = document.querySelectorAll(".menu-item");
const contentSections = document.querySelectorAll(".content-section");

menuItems.forEach(item => {
  item.addEventListener("click", () => {
    // Remove active class from all menu items
    menuItems.forEach(i => i.classList.remove("active"));
    // Add active class to clicked menu
    item.classList.add("active");
    // Hide all content sections
    contentSections.forEach(section => section.classList.remove("active"));
    // Show corresponding section
    const menuName = item.dataset.menu;
    document.getElementById(menuName).classList.add("active");
  });
});

// ===== Onload Load Data if ada di localStorage =====
window.onload = () => {
  loadDataFromLocalStorage();
  renderTable(dataLembarKerja);
};

