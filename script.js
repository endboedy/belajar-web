// Helper parse tanggal Excel (number or string)
function parseExcelDate(excelDate) {
  if (!excelDate) return null;
  if (excelDate instanceof Date) return excelDate;
  if (typeof excelDate === "number") {
    return new Date((excelDate - (25569 + 2)) * 86400 * 1000);
  }
  const d = new Date(excelDate);
  return isNaN(d) ? null : d;
}

function formatDateDDMMMYYYY(date) {
  if (!(date instanceof Date)) return "";
  const day = date.getDate().toString().padStart(2, "0");
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const month = monthNames[date.getMonth()];
  const year = date.getFullYear();
  return `${day} ${month} ${year}`;
}

// Global data dan state
let allData = [];
let isEditingIndex = null;

const monthOptions = [
  "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

// Inisialisasi menu sidebar
function setupMenu() {
  const menuItems = document.querySelectorAll(".sidebar .menu-item");
  const contentSections = document.querySelectorAll(".content-section");
  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      item.classList.add("active");
      const menuId = item.dataset.menu;
      contentSections.forEach(sec => {
        sec.id === menuId ? sec.classList.add("active") : sec.classList.remove("active");
      });
    });
  });
}

// Load file Excel
function handleFileUpload() {
  const input = document.getElementById("file-input");
  const file = input.files[0];
  if (!file) {
    alert("Pilih file Excel dulu");
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Contoh: baca sheet IW39 (sesuaikan nama sheet dan parsing sesuai kebutuhan)
    const sheetName = "IW39";
    if (!workbook.SheetNames.includes(sheetName)) {
      alert(`Sheet ${sheetName} tidak ditemukan di file.`);
      return;
    }
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    // Simpan data, contoh ambil kolom2 yang diperlukan
    allData = jsonData.map(row => {
      // Contoh map data kolom sesuai nama header kamu
      return {
        Room: row.Room || "",
        "Order Type": row["Order Type"] || "",
        Order: row.Order || "",
        Description: row.Description || "",
        "Created On": row["Created On"], // parsing nanti
        "User Status": row["User Status"] || "",
        MAT: row.MAT || "",
        CPH: row.CPH || "",
        Section: row.Section || "",
        "Status Part": row["Status Part"] || "",
        Aging: row.Aging || "",
        Month: row.Month || "",
        Cost: row.Cost || "",
        Reman: row.Reman || "",
        Include: row.Include || "",
        Exclude: row.Exclude || "",
        Planning: row.Planning || "", // dari Planning/Event Start nanti
        "Status AMT": row["Status AMT"] || ""
      };
    });

    // Update status upload
    document.getElementById("upload-status").textContent = `File ${file.name} berhasil diupload dan diproses.`;

    renderTable();
  };
  reader.readAsArrayBuffer(file);
}

// Render tabel data
function renderTable() {
  const tbody = document.querySelector("#output-table tbody");
  tbody.innerHTML = "";

  allData.forEach((row, index) => {
    const tr = document.createElement("tr");
    const editing = (index === isEditingIndex);

    // Helper buat td teks biasa
    const tdText = (text, align = "left") => {
      const td = document.createElement("td");
      td.textContent = text || "";
      if (align === "right") td.style.textAlign = "right";
      return td;
    };

    tr.appendChild(tdText(row.Room));
    tr.appendChild(tdText(row["Order Type"]));
    tr.appendChild(tdText(row.Order));
    tr.appendChild(tdText(row.Description));

    // Created On format tanggal
    const createdOnDate = parseExcelDate(row["Created On"]);
    tr.appendChild(tdText(formatDateDDMMMYYYY(createdOnDate)));

    tr.appendChild(tdText(row["User Status"]));
    tr.appendChild(tdText(row.MAT));
    tr.appendChild(tdText(row.CPH));
    tr.appendChild(tdText(row.Section));
    tr.appendChild(tdText(row["Status Part"]));
    tr.appendChild(tdText(row.Aging));

    // Month kolom edit dropdown
    if (editing) {
      const tdMonth = document.createElement("td");
      const select = document.createElement("select");
      monthOptions.forEach(m => {
        const opt = document.createElement("option");
        opt.value = m;
        opt.textContent = m || "--";
        if (m === row.Month) opt.selected = true;
        select.appendChild(opt);
      });
      tdMonth.appendChild(select);
      tr.appendChild(tdMonth);
    } else {
      tr.appendChild(tdText(row.Month));
    }

    tr.appendChild(tdText(row.Cost, "right"));

    // Reman kolom edit input text
    if (editing) {
      const tdReman = document.createElement("td");
      const input = document.createElement("input");
      input.type = "text";
      input.value = row.Reman || "";
      tdReman.appendChild(input);
      tr.appendChild(tdReman);
    } else {
      tr.appendChild(tdText(row.Reman));
    }

    tr.appendChild(tdText(row.Include, "right"));
    tr.appendChild(tdText(row.Exclude, "right"));

    // Planning format tanggal
    const planningDate = parseExcelDate(row.Planning);
    tr.appendChild(tdText(formatDateDDMMMYYYY(planningDate)));

    tr.appendChild(tdText(row["Status AMT"]));

    // Action tombol edit/delete/save/cancel
    const tdAction = document.createElement("td");

    if (editing) {
      const btnSave = document.createElement("button");
      btnSave.textContent = "Save";
      btnSave.className = "action-btn save-btn";
      btnSave.addEventListener("click", () => saveEdit(index));
      tdAction.appendChild(btnSave);

      const btnCancel = document.createElement("button");
      btnCancel.textContent = "Cancel";
      btnCancel.className = "action-btn cancel-btn";
      btnCancel.addEventListener("click", cancelEdit);
      tdAction.appendChild(btnCancel);
    } else {
      const btnEdit = document.createElement("button");
      btnEdit.textContent = "Edit";
      btnEdit.className = "action-btn edit-btn";
      btnEdit.addEventListener("click", () => startEdit(index));
      tdAction.appendChild(btnEdit);

      const btnDelete = document.createElement("button");
      btnDelete.textContent = "Delete";
      btnDelete.className = "action-btn delete-btn";
      btnDelete.addEventListener("click", () => deleteRow(index));
      tdAction.appendChild(btnDelete);
    }

    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

// Start edit baris
function startEdit(index) {
  if (isEditingIndex !== null) {
    alert("Selesaikan dulu edit yang sedang aktif.");
    return;
  }
  isEditingIndex = index;
  renderTable();
}

// Cancel edit
function cancelEdit() {
  isEditingIndex = null;
  renderTable();
}

// Save edit
function saveEdit(index) {
  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.children[index];
  const selectMonth = tr.querySelector("select");
  const inputReman = tr.querySelector("input[type=text]");

  if (!selectMonth || !inputReman) {
    alert("Gagal ambil data edit.");
    return;
  }

  allData[index].Month = selectMonth.value;
  allData[index].Reman = inputReman.value.trim();

  isEditingIndex = null;
  renderTable();
}

// Delete baris
function deleteRow(index) {
  if (confirm("Yakin ingin menghapus baris ini?")) {
    allData.splice(index, 1);
    if (isEditingIndex === index) isEditingIndex = null;
    renderTable();
  }
}

// Filter data (contoh filter order)
function filterData() {
  const filterOrder = document.getElementById("filter-order").value.toLowerCase();
  const filterRoom = document.getElementById("filter-room").value.toLowerCase();
  const filterCPH = document.getElementById("filter-cph").value.toLowerCase();
  const filterMAT = document.getElementById("filter-mat").value.toLowerCase();
  const filterSection = document.getElementById("filter-section").value.toLowerCase();
  const filterMonth = document.getElementById("filter-month").value;

  let filtered = allData.filter(row => {
    const orderLower = (row.Order || "").toString().toLowerCase();
    const roomLower = (row.Room || "").toString().toLowerCase();
    const cphLower = (row.CPH || "").toString().toLowerCase();
    const matLower = (row.MAT || "").toString().toLowerCase();
    const sectionLower = (row.Section || "").toString().toLowerCase();

    return (!filterOrder || orderLower.includes(filterOrder)) &&
           (!filterRoom || roomLower.includes(filterRoom)) &&
           (!filterCPH || cphLower.includes(filterCPH)) &&
           (!filterMAT || matLower.includes(filterMAT)) &&
           (!filterSection || sectionLower.includes(filterSection)) &&
           (!filterMonth || row.Month === filterMonth);
  });

  allData = filtered;
  isEditingIndex = null;
  renderTable();
}

// Reset filter dan reload data (butuh reload asli dari file atau backup)
function resetFilter() {
  // Reload ulang file atau simpan backup data asal supaya reset bisa pakai itu
  alert("Reset filter: silakan upload ulang file agar data kembali.");
}

// Setup event listeners tombol dll
function setupEvents() {
  document.getElementById("upload-btn").addEventListener("click", handleFileUpload);
  document.getElementById("filter-btn").addEventListener("click", filterData);
  document.getElementById("reset-btn").addEventListener("click", resetFilter);
  document.getElementById("refresh-btn").addEventListener("click", renderTable);

  // Tombol add order (buat nambah data dummy contoh, bisa dihapus atau dikembangkan)
  document.getElementById("add-order-btn").addEventListener("click", () => {
    const val = document.getElementById("add-order-input").value.trim();
    if (!val) {
      alert("Isi order dulu");
      return;
    }
    const orders = val.split(/[\s,]+/);
    orders.forEach(ord => {
      allData.push({
        Room: "",
        "Order Type": "",
        Order: ord,
        Description: "",
        "Created On": new Date(),
        "User Status": "",
        MAT: "",
        CPH: "",
        Section: "",
        "Status Part": "",
        Aging: "",
        Month: "",
        Cost: "",
        Reman: "",
        Include: "",
        Exclude: "",
        Planning: new Date(),
        "Status AMT": ""
      });
    });
    renderTable();
    document.getElementById("add-order-input").value = "";
    document.getElementById("add-order-status").textContent = `${orders.length} order berhasil ditambahkan.`;
    setTimeout(() => {
      document.getElementById("add-order-status").textContent = "";
    }, 3000);
  });
}

// Inisialisasi semua
window.onload = () => {
  setupMenu();
  setupEvents();
  renderTable(); // kosong dulu
};
