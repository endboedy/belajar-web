// Simulasi data asli untuk contoh, kamu bisa ganti dengan data asli kamu
const originalData = [
  {
    Room: "A101", "Order Type": "Type1", Order: "ORD001", Description: "Desc1",
    "Created On": "2025-08-01", "User Status": "Open", MAT: "MAT1", CPH: "CPH1",
    Section: "Sec1", "Status Part": "OK", Aging: 5, Month: "Aug", Cost: 100,
    Reman: "No", Include: "Yes", Exclude: "No", Planning: "Plan1", "Status AMT": "Active",
    Action: "Action1"
  },
  {
    Room: "B202", "Order Type": "Type2", Order: "ORD002", Description: "Desc2",
    "Created On": "2025-08-02", "User Status": "Close", MAT: "MAT2", CPH: "CPH2",
    Section: "Sec2", "Status Part": "Pending", Aging: 10, Month: "Aug", Cost: 200,
    Reman: "Yes", Include: "No", Exclude: "Yes", Planning: "Plan2", "Status AMT": "Inactive",
    Action: "Action2"
  },
  // ... tambah data contoh lainnya jika perlu
];

// Fungsi untuk render data ke tabel (#output-table tbody)
function renderTable(data) {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = ""; // bersihkan dulu

  if(data.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 19;
    td.style.textAlign = "center";
    td.textContent = "Tidak ada data";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  data.forEach(row => {
    const tr = document.createElement("tr");
    Object.keys(row).forEach(key => {
      const td = document.createElement("td");
      td.textContent = row[key];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

// Fungsi load data asli ke tabel (dipanggil saat refresh)
function loadDataOriginal() {
  renderTable(originalData);
  console.log("Data asli berhasil diload");
}

// Filter data berdasarkan filter input
function filterData() {
  let filtered = originalData.slice();

  const room = document.getElementById("filter-room").value.trim().toLowerCase();
  const order = document.getElementById("filter-order").value.trim().toLowerCase();
  const cph = document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat = document.getElementById("filter-mat").value.trim().toLowerCase();
  const section = document.getElementById("filter-section").value.trim().toLowerCase();

  if(room) filtered = filtered.filter(d => (d.Room || "").toLowerCase().includes(room));
  if(order) filtered = filtered.filter(d => (d.Order || "").toLowerCase().includes(order));
  if(cph) filtered = filtered.filter(d => (d.CPH || "").toLowerCase().includes(cph));
  if(mat) filtered = filtered.filter(d => (d.MAT || "").toLowerCase().includes(mat));
  if(section) filtered = filtered.filter(d => (d.Section || "").toLowerCase().includes(section));

  renderTable(filtered);
}

// Reset filter inputs dan tabel kosong
function resetFilter() {
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section"].forEach(id => {
    document.getElementById(id).value = "";
  });
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";
}

// Add orders dari textarea (contoh simulasi)
function addOrders() {
  const input = document.getElementById("add-order-input").value.trim();
  const status = document.getElementById("add-order-status");

  if(!input) {
    status.textContent = "Masukkan order dulu ya bro!";
    status.style.color = "red";
    return;
  }

  const orders = input.split(/[\s,]+/).filter(o => o.length > 0);

  // Simulasi menambahkan data ke originalData
  orders.forEach(order => {
    originalData.push({
      Room: "NewRoom",
      "Order Type": "NewType",
      Order: order,
      Description: "Auto added",
      "Created On": new Date().toISOString().slice(0,10),
      "User Status": "Open",
      MAT: "MAT-New",
      CPH: "CPH-New",
      Section: "Sec-New",
      "Status Part": "New",
      Aging: 0,
      Month: "Aug",
      Cost: 0,
      Reman: "No",
      Include: "No",
      Exclude: "No",
      Planning: "Plan-New",
      "Status AMT": "Active",
      Action: "None"
    });
  });

  status.textContent = orders.length + " order berhasil ditambahkan.";
  status.style.color = "green";

  // Render ulang data
  loadDataOriginal();

  // Clear textarea
  document.getElementById("add-order-input").value = "";
}

// Handle menu navigasi
function setupMenu() {
  const menuItems = document.querySelectorAll(".menu-item");
  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      // Remove active class dari semua menu dan content
      menuItems.forEach(i => i.classList.remove("active"));
      document.querySelectorAll(".content-section").forEach(sec => sec.classList.remove("active"));

      // Set active pada menu dan section sesuai data-menu
      item.classList.add("active");
      const menu = item.getAttribute("data-menu");
      const section = document.getElementById(menu);
      if(section) section.classList.add("active");
    });
  });
}

window.addEventListener("DOMContentLoaded", () => {
  setupMenu();

  loadDataOriginal();

  document.getElementById("filter-btn").addEventListener("click", filterData);
  document.getElementById("reset-btn").addEventListener("click", () => {
    resetFilter();
    renderTable(originalData);
  });
  document.getElementById("refresh-btn").addEventListener("click", () => {
    resetFilter();
    loadDataOriginal();
    document.getElementById("add-order-status").textContent = "";
  });

  document.getElementById("add-order-btn").addEventListener("click", addOrders);

  // Save & Load tombol (dummy example)
  document.getElementById("save-btn").addEventListener("click", () => {
    alert("Fitur Save belum tersedia");
  });
  document.getElementById("load-btn").addEventListener("click", () => {
    alert("Fitur Load belum tersedia");
  });
});
