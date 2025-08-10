// ----- Dummy Data Global -----
let IW39 = [
  {
    Room: "R001",
    OrderType: "Type A",
    Order: "ORD001",
    Description: "JR Sample description",
    CreatedOn: "2025-08-01",
    UserStatus: "Open",
    MAT: "MAT001",
    TotalPlan: 50000,
    TotalActual: 30000
  },
  {
    Room: "R002",
    OrderType: "Type B",
    Order: "ORD002",
    Description: "No JR here",
    CreatedOn: "2025-08-02",
    UserStatus: "Closed",
    MAT: "MAT002",
    TotalPlan: 40000,
    TotalActual: 40000
  }
];

// Dummy Data1 lookup (Section by Order)
const Data1 = {
  "ORD001": "Section A",
  "ORD002": "Section B"
};

// Dummy Data2 lookup (CPH by MAT)
const Data2 = {
  "MAT001": "CPH1",
  "MAT002": "CPH2",
  "MAT999": "CPH999"
};

// Dummy SUM57 lookup (Status Part & Aging by Order)
const SUM57 = {
  "ORD001": { StatusPart: "OK", Aging: "5" },
  "ORD002": { StatusPart: "NG", Aging: "10" },
  "ORD010": { StatusPart: "OK", Aging: "2" }
};

// Dummy Planning lookup (Planning & Status AMT by Order)
const Planning = {
  "ORD001": { Planning: "2025-08-10", StatusAMT: "On Track" },
  "ORD002": { Planning: "2025-08-12", StatusAMT: "Delayed" },
  "ORD010": { Planning: "2025-08-15", StatusAMT: "On Hold" }
};

// ----- Data Lembar Kerja -----
let dataLembarKerja = [];

// ----- Format angka 1 decimal -----
function formatNumber(num) {
  return Number(num).toFixed(1);
}

// ----- Switch Menu -----
function switchMenu(menuName) {
  document.querySelectorAll('.menu-item').forEach(m => m.classList.remove('active'));
  document.querySelectorAll('.content-section').forEach(c => c.classList.remove('active'));

  const menu = document.querySelector(`.menu-item[data-menu="${menuName}"]`);
  const content = document.getElementById(menuName);

  if (menu && content) {
    menu.classList.add('active');
    content.classList.add('active');
  }
}

// Event klik menu
document.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', () => {
    switchMenu(item.dataset.menu);
  });
});

// ----- Build Data Lembar Kerja: kalkulasi lookup dan rumus -----
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    // Cari data lengkap IW39 berdasar Order (key)
    const iw = IW39.find(i => i.Order.toLowerCase() === row.Order.toLowerCase()) || {};

    row.Room = iw.Room || "";
    row.OrderType = iw.OrderType || "";
    row.Description = iw.Description || "";
    row.CreatedOn = iw.CreatedOn || "";
    row.UserStatus = iw.UserStatus || "";
    row.MAT = iw.MAT || "";

    // CPH: jika 2 huruf pertama Description = "JR" maka JR, else lookup Data2 by MAT
    if ((iw.Description || "").substring(0, 2).toUpperCase() === "JR") {
      row.CPH = "JR";
    } else {
      row.CPH = Data2[row.MAT] || "";
    }

    // Section lookup dari Data1 berdasarkan Order
    row.Section = Data1[row.Order] || "";

    // SUM57 lookup by Order
    const sum = SUM57[row.Order] || {};
    row.StatusPart = sum.StatusPart || "";
    row.Aging = sum.Aging || "";

    // Cost rumus (IW39.TotalPlan - IW39.TotalActual)/16500, <0 jadi "-"
    if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const costCalc = (iw.TotalPlan - iw.TotalActual) / 16500;
      row.Cost = costCalc < 0 ? "-" : costCalc;
    } else {
      row.Cost = "-";
    }

    // Include rumus
    if ((row.Reman || "").toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    // Exclude rumus
    if ((row.OrderType || "").toUpperCase() === "PM38") {
      row.Exclude = "-";
    } else {
      row.Exclude = row.Include;
    }

    // Planning lookup
    const plan = Planning[row.Order] || {};
    row.Planning = plan.Planning || "";
    row.StatusAMT = plan.StatusAMT || "";

    return row;
  });
}

// ----- Validasi order input (tidak boleh titik atau koma) -----
function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ----- Render tabel -----
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

    // Kolom statis
    const cols = [
      "Room", "OrderType", "Order", "Description", "CreatedOn", "UserStatus",
      "MAT", "CPH", "Section", "StatusPart", "Aging"
    ];
    cols.forEach(col => {
      const td = document.createElement("td");
      td.textContent = row[col] || "";
      tr.appendChild(td);
    });

    // Editable Month
    const tdMonth = document.createElement("td");
    tdMonth.textContent = row.Month || "";
    tdMonth.classList.add("month-cell");
    tr.appendChild(tdMonth);

    // Cost
    const tdCost = document.createElement("td");
    tdCost.textContent = (typeof row.Cost === "number") ? formatNumber(row.Cost) : row.Cost;
    tdCost.classList.add("num-cell");
    tr.appendChild(tdCost);

    // Editable Reman
    const tdReman = document.createElement("td");
    tdReman.textContent = row.Reman || "";
    tdReman.classList.add("reman-cell");
    tr.appendChild(tdReman);

    // Include
    const tdInclude = document.createElement("td");
    tdInclude.textContent = (typeof row.Include === "number") ? formatNumber(row.Include) : row.Include;
    tdInclude.classList.add("num-cell");
    tr.appendChild(tdInclude);

    // Exclude
    const tdExclude = document.createElement("td");
    tdExclude.textContent = (typeof row.Exclude === "number") ? formatNumber(row.Exclude) : row.Exclude;
    tdExclude.classList.add("num-cell");
    tr.appendChild(tdExclude);

    // Planning & Status AMT
    const tdPlanning = document.createElement("td");
    tdPlanning.textContent = row.Planning || "";
    tr.appendChild(tdPlanning);

    const tdStatusAMT = document.createElement("td");
    tdStatusAMT.textContent = row.StatusAMT || "";
    tr.appendChild(tdStatusAMT);

    // Action (Edit & Delete)
    const tdAction = document.createElement("td");

    const btnEdit = document.createElement("button");
    btnEdit.textContent = "Edit";
    btnEdit.classList.add("btn-action", "btn-edit");
    btnEdit.addEventListener("click", () => editRow(tr, row));
    tdAction.appendChild(btnEdit);

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
    tdAction.appendChild(btnDelete);

    tr.appendChild(tdAction);

    outputTableBody.appendChild(tr);
  });
}

// ----- Edit baris untuk Month dan Reman -----
function editRow(tr, row) {
  // Cek apakah sudah dalam mode edit
  if (tr.classList.contains("editing")) {
    alert("Sudah dalam mode edit bro!");
    return;
  }
  tr.classList.add("editing");

  // Cari td Month dan Reman
  const tdMonth = tr.querySelector(".month-cell");
  const tdReman = tr.querySelector(".reman-cell");

  // Buat select untuk Month
  const selectMonth = document.createElement("select");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  months.forEach(m => {
    const option = document.createElement("option");
    option.value = m;
    option.textContent = m;
    if (row.Month === m) option.selected = true;
    selectMonth.appendChild(option);
  });

  // Buat input text untuk Reman
  const inputReman = document.createElement("input");
  inputReman.type = "text";
  inputReman.value = row.Reman || "";

  // Ganti cell dengan form input
  tdMonth.textContent = "";
  tdMonth.appendChild(selectMonth);

  tdReman.textContent = "";
  tdReman.appendChild(inputReman);

  // Ganti tombol Edit jadi Save dan Cancel
  const tdAction = tr.querySelector("td:last-child");
  tdAction.innerHTML = "";

  const btnSave = document.createElement("button");
  btnSave.textContent = "Save";
  btnSave.classList.add("btn-action", "btn-save");
  btnSave.addEventListener("click", () => {
    // Simpan nilai baru ke row
    row.Month = selectMonth.value;
    row.Reman = inputReman.value.trim();

    // Update lookup rumus setelah simpan
    buildDataLembarKerja();

    // Render ulang tabel
    renderTable(dataLembarKerja);
  });
  tdAction.appendChild(btnSave);

  const btnCancel = document.createElement("button");
  btnCancel.textContent = "Cancel";
  btnCancel.classList.add("btn-action", "btn-cancel");
  btnCancel.addEventListener("click", () => {
    renderTable(dataLembarKerja);
  });
  tdAction.appendChild(btnCancel);
}

// ----- Validasi order input (tidak boleh titik atau koma) -----
function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ----- Add Order multi input -----
const addOrderBtn = document.getElementById("add-order-btn");
const addOrderInput = document.getElementById("add-order-input");
const addOrderStatus = document.getElementById("add-order-status");

addOrderBtn.addEventListener("click", () => {
  let rawInput = addOrderInput.value.trim();
  if (!rawInput) {
    alert("Masukkan minimal satu Order bro!");
    return;
  }

  // Split by whitespace, comma, or new line
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
        Order: order,
        Month: "",
        Reman: ""
      });
      addedCount++;
    } else {
      skippedOrders.push(order);
    }
  });

  // Update lookup data lengkap
  buildDataLembarKerja();

  // Render tabel
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

// ----- Filter data -----
const filterBtn = document.getElementById("filter-btn");
const resetBtn = document.getElementById("reset-btn");

filterBtn.addEventListener("click", () => {
  const fRoom = document.getElementById("filter-room").value.trim().toLowerCase();
  const fOrder = document.getElementById("filter-order").value.trim().toLowerCase();
  const fCPH = document.getElementById("filter-cph").value.trim().toLowerCase();
  const fMAT = document.getElementById("filter-mat").value.trim().toLowerCase();
  const fSection = document.getElementById("filter-section").value.trim().toLowerCase();

  const filtered = dataLembarKerja.filter(d => {
    return (
      d.Room.toLowerCase().includes(fRoom) &&
      d.Order.toLowerCase().includes(fOrder) &&
      d.CPH.toLowerCase().includes(fCPH) &&
      d.MAT.toLowerCase().includes(fMAT) &&
      d.Section.toLowerCase().includes(fSection)
    );
  });

  renderTable(filtered);
});

resetBtn.addEventListener("click", () => {
  document.getElementById("filter-room").value = "";
  document.getElementById("filter-order").value = "";
  document.getElementById("filter-cph").value = "";
  document.getElementById("filter-mat").value = "";
  document.getElementById("filter-section").value = "";
  renderTable(dataLembarKerja);
});

// ----- Save & Load data ke localStorage -----
const saveBtn = document.getElementById("save-btn");
const loadBtn = document.getElementById("load-btn");

saveBtn.addEventListener("click", () => {
  localStorage.setItem("dataLembarKerja", JSON.stringify(dataLembarKerja));
  alert("Data berhasil disimpan ke localStorage 👍");
});

loadBtn.addEventListener("click", () => {
  const saved = localStorage.getItem("dataLembarKerja");
  if (saved) {
    dataLembarKerja = JSON.parse(saved);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    alert("Data berhasil dimuat dari localStorage 👍");
  } else {
    alert("Tidak ada data tersimpan di localStorage.");
  }
});

// ----- Update data dari file upload (menu 1) -----
const fileUpload = document.getElementById("file-upload");
fileUpload.addEventListener("change", () => {
  const file = fileUpload.files[0];
  if (!file) return;
  // Simulasi update IW39 dari file upload
  // TODO: Implement parsing file sebenarnya jika perlu
  // Contoh dummy update data:
  IW39 = [
    {
      Room: "R010",
      OrderType: "Type Z",
      Order: "ORD010",
      Description: "JR New uploaded",
      CreatedOn: "2025-08-10",
      UserStatus: "Open",
      MAT: "MAT999",
      TotalPlan: 90000,
      TotalActual: 20000
    }
  ];
  // Reset data lembar kerja dengan Order dari IW39 (input manual kosong)
  dataLembarKerja = IW39.map(iw => ({
    Order: iw.Order,
    Month: "",
    Reman: ""
  }));

  // Build lookup rumus lengkap
  buildDataLembarKerja();

  // Render tabel
  renderTable(dataLembarKerja);

  // Pindah menu ke lembar (menu2)
  switchMenu("lembar");
  alert("File IW39 berhasil diupload dan data lembar kerja diperbarui.");
});

// ----- Inisialisasi awal -----
window.onload = () => {
  // Inisialisasi data lembar kerja dari IW39 awal
  dataLembarKerja = IW39.map(iw => ({
    Order: iw.Order,
    Month: "",
    Reman: ""
  }));
  buildDataLembarKerja();
  renderTable(dataLembarKerja);
  switchMenu("upload");
};
