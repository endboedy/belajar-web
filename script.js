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
  "MAT002": "CPH2"
};

// Dummy SUM57 lookup (Status Part & Aging by Order)
const SUM57 = {
  "ORD001": { StatusPart: "OK", Aging: "5" },
  "ORD002": { StatusPart: "NG", Aging: "10" }
};

// Dummy Planning lookup (Planning & Status AMT by Order)
const Planning = {
  "ORD001": { Planning: "2025-08-10", StatusAMT: "On Track" },
  "ORD002": { Planning: "2025-08-12", StatusAMT: "Delayed" }
};

// ----- Data Lembar Kerja -----
let dataLembarKerja = [];

// ----- Format angka 1 decimal -----
function formatNumber(num) {
  return Number(num).toFixed(1);
}

// ----- Build Data Lembar Kerja: kalkulasi lookup dan rumus -----
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    // Cari data IW39 berdasar order (case insensitive)
    const iw = IW39.find(i => i.Order.toLowerCase() === row.Order.toLowerCase()) || {};

    // Update kolom-kolom yang berasal dari IW39 langsung
    row.Room = iw.Room || "";
    row.OrderType = iw.OrderType || "";
    row.Description = iw.Description || "";
    row.CreatedOn = iw.CreatedOn || "";
    row.UserStatus = iw.UserStatus || "";
    row.MAT = iw.MAT || "";

    // CPH: jika 2 huruf pertama Description = "JR" maka JR, else lookup Data2 by MAT
    if ((row.Description || "").substring(0, 2).toUpperCase() === "JR") {
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

  data.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    if (duplicates.includes(row.Order.toLowerCase())) {
      tr.classList.add("duplicate");
    }

    // Buat sel-sel td
    function createTd(text, className = "") {
      const td = document.createElement("td");
      td.textContent = text;
      if (className) td.classList.add(className);
      return td;
    }

    tr.appendChild(createTd(row.Room));
    tr.appendChild(createTd(row.OrderType));
    tr.appendChild(createTd(row.Order));
    tr.appendChild(createTd(row.Description));
    tr.appendChild(createTd(row.CreatedOn));
    tr.appendChild(createTd(row.UserStatus));
    tr.appendChild(createTd(row.MAT));
    tr.appendChild(createTd(row.CPH));
    tr.appendChild(createTd(row.Section));
    tr.appendChild(createTd(row.StatusPart));
    tr.appendChild(createTd(row.Aging));

    // Month (editable only saat mode edit)
    const tdMonth = document.createElement("td");
    if (row.isEditing) {
      const select = document.createElement("select");
      ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].forEach(m => {
        const option = document.createElement("option");
        option.value = m;
        option.textContent = m;
        if (m === row.Month) option.selected = true;
        select.appendChild(option);
      });
      tdMonth.appendChild(select);
      select.addEventListener("change", () => {
        row.Month = select.value;
      });
    } else {
      tdMonth.textContent = row.Month || "";
    }
    tr.appendChild(tdMonth);

    // Cost (format angka 1 decimal, rata kanan)
    const tdCost = createTd(typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost, "right");
    tr.appendChild(tdCost);

    // Reman (editable only saat mode edit)
    const tdReman = document.createElement("td");
    if (row.isEditing) {
      const input = document.createElement("input");
      input.type = "text";
      input.value = row.Reman || "";
      tdReman.appendChild(input);
      input.addEventListener("input", () => {
        row.Reman = input.value;
      });
    } else {
      tdReman.textContent = row.Reman || "";
    }
    tr.appendChild(tdReman);

    // Include (format angka 1 decimal, rata kanan)
    const tdInclude = createTd(typeof row.Include === "number" ? formatNumber(row.Include) : row.Include, "right");
    tr.appendChild(tdInclude);

    // Exclude (format angka 1 decimal, rata kanan)
    const tdExclude = createTd(typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude, "right");
    tr.appendChild(tdExclude);

    tr.appendChild(createTd(row.Planning));
    tr.appendChild(createTd(row.StatusAMT));

    // Action: tombol Edit / Save / Cancel dan Delete
    const tdAction = document.createElement("td");

    if (row.isEditing) {
      // Tombol Save
      const btnSave = document.createElement("button");
      btnSave.textContent = "Save";
      btnSave.classList.add("btn-action", "btn-save");
      btnSave.addEventListener("click", () => {
        row.isEditing = false;
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
      });
      tdAction.appendChild(btnSave);

      // Tombol Cancel
      const btnCancel = document.createElement("button");
      btnCancel.textContent = "Cancel";
      btnCancel.classList.add("btn-action", "btn-cancel");
      btnCancel.addEventListener("click", () => {
        row.isEditing = false;
        // Reload data dari saved (bisa implementasikan reload khusus jika perlu)
        renderTable(dataLembarKerja);
      });
      tdAction.appendChild(btnCancel);
    } else {
      // Tombol Edit
      const btnEdit = document.createElement("button");
      btnEdit.textContent = "Edit";
      btnEdit.classList.add("btn-action", "btn-edit");
      btnEdit.addEventListener("click", () => {
        // Set hanya row ini yang mode edit, reset row lain
        dataLembarKerja.forEach(r => r.isEditing = false);
        row.isEditing = true;
        renderTable(dataLembarKerja);
      });
      tdAction.appendChild(btnEdit);

      // Tombol Delete
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
    }

    tr.appendChild(tdAction);

    // Tandai warna merah untuk order duplicate
    if (duplicates.includes(row.Order.toLowerCase())) {
      tr.style.backgroundColor = "#d9534f"; // merah
      tr.style.color = "white";
    }

    outputTableBody.appendChild(tr);
  });
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

  // Pisah berdasarkan spasi, koma, atau enter
  let orders = rawInput.split(/[\s,]+/).map(s => s.trim()).filter(s => s.length > 0);

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
function updateDataFromUpload(fileName) {
  if (fileName.toLowerCase().includes('iw39')) {
    // Contoh data dummy baru, kamu bisa sesuaikan dengan file sebenarnya
    IW39 = [
      {
        Room: "R010",
        OrderType: "Type Z",
        Order: "ORD010",
        Description: "JR New description",
        CreatedOn: "2025-08-10",
        UserStatus: "Open",
        MAT: "MAT999",
        TotalPlan: 90000,
        TotalActual: 20000
      }
    ];
  }
  // Update dataLembarKerja dari IW39
  dataLembarKerja = IW39.map(iw => ({
    Room: iw.Room,
    OrderType: iw.OrderType,
    Order: iw.Order,
    Description: iw.Description,
    CreatedOn: iw.CreatedOn,
    UserStatus: iw.UserStatus,
    MAT: iw.MAT,
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
  }));
  buildDataLembarKerja();
  renderTable(dataLembarKerja);

  // Tampilkan notif upload sukses saja, jangan switch menu otomatis
  alert(`File "${fileName}" berhasil diupload. Silakan klik menu "Lembar Kerja" untuk melihat data.`);
}

// ----- Event Upload File (menu 1) -----
const uploadBtn = document.getElementById("upload-btn");
const fileInput = document.getElementById("file-input");
const uploadStatus = document.getElementById("upload-status");
const progressContainer = document.getElementById("progress-container");
const uploadProgress = document.getElementById("upload-progress");
const fileTypeSelect = document.getElementById("file-select");

uploadBtn.addEventListener("click", () => {
  const files = fileInput.files;
  if (!files.length) {
    alert("Pilih file dulu bro!");
    return;
  }
  const file = files[0];
  const selectedFileType = fileTypeSelect.value;

  uploadBtn.disabled = true;
  uploadStatus.textContent = "";
  progressContainer.classList.remove("hidden");
  uploadProgress.value = 0;

  let progress = 0;
  const interval = setInterval(() => {
    progress += 10;
    uploadProgress.value = progress;
    if (progress >= 100) {
      clearInterval(interval);
      uploadStatus.textContent = `File "${file.name}" untuk kategori ${selectedFileType} berhasil diupload! 🎉`;
      uploadBtn.disabled = false;
      fileInput.value = '';
      progressContainer.classList.add('hidden');

      updateDataFromUpload(file.name);
    }
  }, 150);
});

// ----- Switch Menu Manual -----
const menuItems = document.querySelectorAll('.menu-item');
menuItems.forEach(item => {
  item.addEventListener('click', () => {
    const menu = item.getAttribute('data-menu');
    switchMenu(menu);
  });
});

function switchMenu(menuName) {
  document.querySelectorAll('.menu-item').forEach(mi => mi.classList.remove('active'));
  document.querySelectorAll('.content-section').forEach(cs => cs.classList.remove('active'));

  document.querySelector(`.menu-item[data-menu="${menuName}"]`).classList.add('active');
  document.getElementById(menuName).classList.add('active');
}

// ----- Inisialisasi awal: set menu default ke Upload -----
window.onload = () => {
  switchMenu("upload");
  // kosongkan data awal
  dataLembarKerja = [];
  renderTable(dataLembarKerja);
};
