// ----- Sidebar menu switching -----
const menuItems = document.querySelectorAll('.menu-item');
const sections = document.querySelectorAll('.content-section');

menuItems.forEach(item => {
  item.addEventListener('click', () => {
    menuItems.forEach(i => i.classList.remove('active'));
    sections.forEach(s => s.classList.remove('active'));

    item.classList.add('active');
    const target = item.dataset.menu;
    document.getElementById(target).classList.add('active');
  });
});

// ----- Upload simulation -----
const uploadBtn = document.getElementById('upload-btn');
const fileInput = document.getElementById('file-input');
const fileSelect = document.getElementById('file-select');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('upload-progress');
const progressText = document.getElementById('progress-text');
const uploadStatus = document.getElementById('upload-status');

uploadBtn.addEventListener('click', () => {
  const file = fileInput.files[0];
  const selectedFileType = fileSelect.value;

  if (!file) {
    alert('Pilih file terlebih dahulu ya bro!');
    return;
  }

  uploadBtn.disabled = true;
  uploadStatus.textContent = '';
  progressBar.value = 0;
  progressText.textContent = '0%';
  progressContainer.classList.remove('hidden');

  let progress = 0;
  const interval = setInterval(() => {
    progress += Math.floor(Math.random() * 15) + 5;
    if (progress >= 100) {
      progress = 100;
      clearInterval(interval);
      uploadStatus.textContent = `File "${file.name}" untuk kategori ${selectedFileType} berhasil diupload! 🎉`;
      uploadBtn.disabled = false;
      fileInput.value = '';
      progressContainer.classList.add('hidden');
    }
    progressBar.value = progress;
    progressText.textContent = progress + '%';
  }, 300);
});

// ----- Data Dummy (simulasi import dari file Excel) -----

// Simulasi Data IW39 (key utama Order)
const IW39 = [
  {
    Room: "R001",
    OrderType: "Type A",
    Order: "ORD001",
    Description: "JR Check engine",
    CreatedOn: "2025-08-01",
    UserStatus: "Open",
    MAT: "MAT123",
    TotalPlan: 100000,
    TotalActual: 50000
  },
  {
    Room: "R002",
    OrderType: "Type B",
    Order: "ORD002",
    Description: "Check valve",
    CreatedOn: "2025-08-03",
    UserStatus: "Closed",
    MAT: "MAT456",
    TotalPlan: 80000,
    TotalActual: 40000
  },
  {
    Room: "R003",
    OrderType: "Type C",
    Order: "ORD003",
    Description: "JR Inspect pump",
    CreatedOn: "2025-08-05",
    UserStatus: "Open",
    MAT: "MAT789",
    TotalPlan: 120000,
    TotalActual: 60000
  }
];

// Simulasi Data Data2 (lookup CPH by MAT)
const Data2 = {
  MAT123: "CPH45",
  MAT456: "CPH67",
  MAT789: "CPH89"
};

// Simulasi Data Data1 (lookup Section by Order)
const Data1 = {
  ORD001: "Sect1",
  ORD002: "Sect2",
  ORD003: "Sect3"
};

// Simulasi Data SUM57 (lookup Status Part & Aging by Order)
const SUM57 = {
  ORD001: { StatusPart: "Ready", Aging: "2 days" },
  ORD002: { StatusPart: "On order", Aging: "5 days" },
  ORD003: { StatusPart: "Ready", Aging: "1 day" }
};

// Simulasi Data Planning (lookup Planning & Status AMT by Order)
const Planning = {
  ORD001: { Planning: "Plan A", StatusAMT: "Active" },
  ORD002: { Planning: "Plan B", StatusAMT: "Inactive" },
  ORD003: { Planning: "Plan C", StatusAMT: "Active" }
};

// ----- Data Lembar Kerja utama -----
// Ini gabungan data hasil lookup dan input manual
let dataLembarKerja = [];

// ----- Utility: Format angka 1 digit desimal (koma) -----
function formatNumber(num) {
  if (typeof num !== "number" || isNaN(num)) return "-";
  return num.toFixed(1);
}

// ----- Fungsi lookup dan kalkulasi cost, include, exclude -----
function buildDataLembarKerja() {
  // Untuk setiap Order di dataLembarKerja, kita update kolom dari sumber data dan hitung rumus
  dataLembarKerja = dataLembarKerja.map(row => {
    // Cari data IW39 by order
    const iw = IW39.find(i => i.Order.toLowerCase() === row.Order.toLowerCase()) || {};
    // Update kolom dari IW39
    row.Room = iw.Room || "";
    row.OrderType = iw.OrderType || "";
    row.Description = iw.Description || "";
    row.CreatedOn = iw.CreatedOn || "";
    row.UserStatus = iw.UserStatus || "";
    row.MAT = iw.MAT || "";

    // CPH logic
    if (row.Description.substring(0,2).toUpperCase() === "JR") {
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

    // Month (input manual) tetap apa adanya

    // Cost rumus (IW39.TotalPlan - IW39.TotalActual)/16500, <0 jadi "-"
    if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const costCalc = (iw.TotalPlan - iw.TotalActual)/16500;
      row.Cost = costCalc < 0 ? "-" : costCalc;
    } else {
      row.Cost = "-";
    }

    // Reman input manual tetap apa adanya

    // Include rumus
    if (row.Reman.toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    // Exclude rumus
    if (row.OrderType.toUpperCase() === "PM38") {
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
  // cek duplikat order
  const ordersLower = data.map(d => d.Order.toLowerCase());
  const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

  outputTableBody.innerHTML = "";
  if (data.length === 0) {
    outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center; font-style:italic; color:#888;">Tidak ada data sesuai filter.</td></tr>`;
    return;
  }

  data.forEach(row => {
    const tr = document.createElement("tr");
    // Tandai jika duplikat order
    if (duplicates.includes(row.Order.toLowerCase())) {
      tr.classList.add("duplicate");
    }

    // Buat kolom dengan kelas khusus untuk angka (cost/include/exclude)
    const tdRoom = document.createElement("td"); tdRoom.textContent = row.Room; tr.appendChild(tdRoom);
    const tdOrderType = document.createElement("td"); tdOrderType.textContent = row.OrderType; tr.appendChild(tdOrderType);
    const tdOrder = document.createElement("td"); tdOrder.textContent = row.Order; tr.appendChild(tdOrder);
    const tdDescription = document.createElement("td"); tdDescription.textContent = row.Description; tr.appendChild(tdDescription);
    const tdCreatedOn = document.createElement("td"); tdCreatedOn.textContent = row.CreatedOn; tr.appendChild(tdCreatedOn);
    const tdUserStatus = document.createElement("td"); tdUserStatus.textContent = row.UserStatus; tr.appendChild(tdUserStatus);
    const tdMAT = document.createElement("td"); tdMAT.textContent = row.MAT; tr.appendChild(tdMAT);
    const tdCPH = document.createElement("td"); tdCPH.textContent = row.CPH; tr.appendChild(tdCPH);
    const tdSection = document.createElement("td"); tdSection.textContent = row.Section; tr.appendChild(tdSection);
    const tdStatusPart = document.createElement("td"); tdStatusPart.textContent = row.StatusPart; tr.appendChild(tdStatusPart);
    const tdAging = document.createElement("td"); tdAging.textContent = row.Aging; tr.appendChild(tdAging);

    // Editable Month
    const tdMonth = document.createElement("td");
    tdMonth.classList.add("editable");
    tdMonth.textContent = row.Month || "";
    tdMonth.title = "Klik untuk edit bulan";
    tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
    tr.appendChild(tdMonth);

    // Cost (format angka 1 decimal, rata kanan)
    const tdCost = document.createElement("td");
    tdCost.classList.add("cost");
    tdCost.textContent = typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost;
    tr.appendChild(tdCost);

    // Editable Reman
    const tdReman = document.createElement("td");
    tdReman.classList.add("editable");
    tdReman.textContent = row.Reman || "";
    tdReman.title = "Klik untuk edit Reman";
    tdReman.addEventListener("click", () => editReman(tdReman, row));
    tr.appendChild(tdReman);

    // Include (format angka 1 decimal, rata kanan)
    const tdInclude = document.createElement("td");
    tdInclude.classList.add("include");
    tdInclude.textContent = typeof row.Include === "number" ? formatNumber(row.Include) : row.Include;
    tr.appendChild(tdInclude);

    // Exclude (format angka 1 decimal, rata kanan)
    const tdExclude = document.createElement("td");
    tdExclude.classList.add("exclude");
    tdExclude.textContent = typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude;
    tr.appendChild(tdExclude);

    // Planning
    const tdPlanning = document.createElement("td"); tdPlanning.textContent = row.Planning; tr.appendChild(tdPlanning);
    // Status AMT
    const tdStatusAMT = document.createElement("td"); tdStatusAMT.textContent = row.StatusAMT; tr.appendChild(tdStatusAMT);

    // Action (edit & delete buttons)
    const tdAction = document.createElement("td");
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

// ----- Edit inline Month -----
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

// ----- Edit inline Reman -----
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

  // Split by space, newline, or comma (tapi nanti validasi)
  let orders = rawInput.split(/[\s,\n]+/).map(s => s.trim()).filter(s => s.length > 0);

  let addedCount = 0;
  let skippedOrders = [];
  let invalidOrders = [];

  orders.forEach(order => {
    if (!isValidOrder(order)) {
      invalidOrders.push(order);
      return;
    }
    // cek duplicate di dataLembarKerja
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

// ----- Inisialisasi -----
// Mulai dengan data dari IW39 (Order dari IW39 otomatis masuk)
dataLembarKerja = IW39.map(iw => ({
  Room: iw.Room,
  OrderType: iw.OrderType,
  Order: iw.Order,
  Description: iw.Description,
  CreatedOn: iw.CreatedOn,
  UserStatus: iw.UserStatus,
  MAT: iw.MAT,
  CPH: "",       // nanti di kalkulasi
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
