// Pastikan kamu sudah load xlsx.js di HTML
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

// ---------- GLOBAL VAR ----------
let IW39 = [];
let Data1 = {};
let Data2 = {};
let SUM57 = {};
let Planning = {};

let dataLembarKerja = [];

// ---------- UTIL ----------
function formatNumber(num) {
  if (typeof num !== "number") return num;
  return num.toFixed(1);
}

function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ---------- PARSE FILE EXCEL ----------
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });

        // Parse sheet IW39
        if (wb.SheetNames.includes("IW39")) {
          IW39 = XLSX.utils.sheet_to_json(wb.Sheets["IW39"]);
        }

        // Parse Data1 sheet
        if (wb.SheetNames.includes("Data1")) {
          const arr = XLSX.utils.sheet_to_json(wb.Sheets["Data1"]);
          Data1 = {};
          arr.forEach(r => { if(r.Order && r.Section) Data1[r.Order] = r.Section; });
        }

        // Parse Data2 sheet
        if (wb.SheetNames.includes("Data2")) {
          const arr = XLSX.utils.sheet_to_json(wb.Sheets["Data2"]);
          Data2 = {};
          arr.forEach(r => { if(r.MAT && r.CPH) Data2[r.MAT] = r.CPH; });
        }

        // Parse SUM57 sheet
        if (wb.SheetNames.includes("SUM57")) {
          const arr = XLSX.utils.sheet_to_json(wb.Sheets["SUM57"]);
          SUM57 = {};
          arr.forEach(r => {
            if (r.Order) SUM57[r.Order] = { StatusPart: r.StatusPart || "", Aging: r.Aging || "" };
          });
        }

        // Parse Planning sheet
        if (wb.SheetNames.includes("Planning")) {
          const arr = XLSX.utils.sheet_to_json(wb.Sheets["Planning"]);
          Planning = {};
          arr.forEach(r => {
            if(r.Order) Planning[r.Order] = { Planning: r.Planning || "", StatusAMT: r.StatusAMT || "" };
          });
        }

        resolve();
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}

// ---------- BUILD DATA LEMBAR KERJA (lookup) ----------
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    const orderKey = (row.Order || "").toLowerCase();

    // Lookup IW39 by order
    const iw = IW39.find(i => i.Order && i.Order.toLowerCase() === orderKey) || {};

    row.Room = iw.Room || row.Room || "";
    row.OrderType = iw.OrderType || row.OrderType || "";
    row.Description = iw.Description || row.Description || "";
    row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
    row.UserStatus = iw.UserStatus || row.UserStatus || "";
    row.MAT = iw.MAT || row.MAT || "";

    // CPH logic
    if ((row.Description || "").toUpperCase().startsWith("JR")) {
      row.CPH = "JR";
    } else {
      row.CPH = Data2[row.MAT] || "";
    }

    row.Section = Data1[row.Order] || "";

    if (SUM57[row.Order]) {
      row.StatusPart = SUM57[row.Order].StatusPart || "";
      row.Aging = SUM57[row.Order].Aging || "";
    } else {
      row.StatusPart = "";
      row.Aging = "";
    }

    // Cost
    if(iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const val = (iw.TotalPlan - iw.TotalActual) / 16500;
      row.Cost = val < 0 ? "-" : val;
    } else {
      row.Cost = "-";
    }

    // Include
    if ((row.Reman || "").toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    // Exclude
    row.Exclude = (row.OrderType || "").toUpperCase() === "PM38" ? "-" : row.Include;

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

// ---------- RENDER TABEL ----------
const tbody = document.querySelector("#output-table tbody");

function renderTable(data) {
  tbody.innerHTML = "";
  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="19" style="text-align:center; font-style:italic; color:#888;">Tidak ada data sesuai filter.</td></tr>`;
    return;
  }

  // Cari duplikat order
  const lowerOrders = data.map(d => d.Order.toLowerCase());
  const duplicates = lowerOrders.filter((v, i) => lowerOrders.indexOf(v) !== i);

  data.forEach(row => {
    const tr = document.createElement("tr");
    if (duplicates.includes(row.Order.toLowerCase())) tr.classList.add("duplicate");

    function createTD(text, className) {
      const td = document.createElement("td");
      if (className) td.classList.add(className);
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
    tdMonth.title = "Klik untuk edit Month";
    tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
    tr.appendChild(tdMonth);

    tr.appendChild(createTD(typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost, "cost"));

    // Reman editable
    const tdReman = document.createElement("td");
    tdReman.classList.add("editable");
    tdReman.textContent = row.Reman || "";
    tdReman.title = "Klik untuk edit Reman";
    tdReman.addEventListener("click", () => editReman(tdReman, row));
    tr.appendChild(tdReman);

    tr.appendChild(createTD(typeof row.Include === "number" ? formatNumber(row.Include) : row.Include, "include"));
    tr.appendChild(createTD(typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude, "exclude"));
    tr.appendChild(createTD(row.Planning));
    tr.appendChild(createTD(row.StatusAMT));

    // Action buttons
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
      if (confirm(`Hapus order ${row.Order}?`)) {
        dataLembarKerja = dataLembarKerja.filter(d => d.Order.toLowerCase() !== row.Order.toLowerCase());
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToStorage();
      }
    });
    tdAction.appendChild(btnDelete);

    tr.appendChild(tdAction);

    tbody.appendChild(tr);
  });
}

// ---------- EDIT Month & Reman Inline ----------
function editMonth(td, row) {
  const select = document.createElement("select");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  months.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    if (m === row.Month) opt.selected = true;
    select.appendChild(opt);
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
  const trs = tbody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      editMonth(tr.children[11], row);
    }
  });
}

function editRemanAction(row) {
  const trs = tbody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
      editReman(tr.children[13], row);
    }
  });
}

// ---------- ADD ORDER ----------
const addOrderBtn = document.getElementById("add-order-btn");
const addOrderInput = document.getElementById("add-order-input");
const addOrderStatus = document.getElementById("add-order-status");

addOrderBtn.addEventListener("click", () => {
  const raw = addOrderInput.value.trim();
  if (!raw) {
    alert("Masukkan minimal satu Order bro!");
    return;
  }
  let orders = raw.split(/[\s,]+/).map(o => o.trim()).filter(o => o.length);
  let added = 0;
  let invalid = [];
  let skipped = [];

  orders.forEach(o => {
    if(!isValidOrder(o)) {
      invalid.push(o);
      return;
    }
    if (dataLembarKerja.find(d => d.Order.toLowerCase() === o.toLowerCase())) {
      skipped.push(o);
      return;
    }
    dataLembarKerja.push({
      Room: "", OrderType: "", Order: o, Description: "", CreatedOn: "", UserStatus: "",
      MAT: "", CPH: "", Section: "", StatusPart: "", Aging: "", Month: "", Cost: "-",
      Reman: "", Include: "-", Exclude: "-", Planning: "", StatusAMT: ""
    });
    added++;
  });

  buildDataLembarKerja();
  renderTable(dataLembarKerja);
  addOrderInput.value = "";

  let msg = `${added} order berhasil ditambahkan.`;
  if (invalid.length) msg += ` Order tidak valid: ${invalid.join(", ")}.`;
  if (skipped.length) msg += ` Order sudah ada: ${skipped.join(", ")}.`;
  addOrderStatus.textContent = msg;
  saveDataToStorage();
});

// ---------- FILTER ----------
const filterRoom = document.getElementById("filter-room");
const filterOrder = document.getElementById("filter-order");
const filterCPH = document.getElementById("filter-cph");
const filterMAT = document.getElementById("filter-mat");
const filterSection = document.getElementById("filter-section");

const filters = [filterRoom, filterOrder, filterCPH, filterMAT, filterSection];

filters.forEach(f => f.addEventListener("input", () => {
  const filtered = dataLembarKerja.filter(r => {
    return r.Room.toLowerCase().includes(filterRoom.value.toLowerCase())
      && r.Order.toLowerCase().includes(filterOrder.value.toLowerCase())
      && r.CPH.toLowerCase().includes(filterCPH.value.toLowerCase())
      && r.MAT.toLowerCase().includes(filterMAT.value.toLowerCase())
      && r.Section.toLowerCase().includes(filterSection.value.toLowerCase());
  });
  renderTable(filtered);
}));

// ---------- SAVE & LOAD LOCALSTORAGE ----------
const saveBtn = document.getElementById("save-btn");
const loadBtn = document.getElementById("load-btn");

function saveDataToStorage() {
  localStorage.setItem("dataLembarKerja", JSON.stringify(dataLembarKerja));
}

function loadDataFromStorage() {
  const data = localStorage.getItem("dataLembarKerja");
  if(data) {
    dataLembarKerja = JSON.parse(data);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  }
}

saveBtn.addEventListener("click", () => {
  saveDataToStorage();
  alert("Data berhasil disimpan di browser.");
});

loadBtn.addEventListener("click", () => {
  loadDataFromStorage();
});

// ---------- UPLOAD FILE ----------
const fileInput = document.getElementById("file-input");
const uploadBtn = document.getElementById("upload-btn");
const progressContainer = document.getElementById("progress-container");
const progressBar = document.getElementById("upload-progress");
const progressText = document.getElementById("progress-text");
const uploadStatus = document.getElementById("upload-status");

fileInput.setAttribute("multiple", true);

uploadBtn.addEventListener("click", async () => {
  const files = fileInput.files;
  if (!files.length) {
    alert("Pilih minimal 1 file Excel untuk upload bro!");
    return;
  }

  progressContainer.classList.remove("hidden");
  progressBar.value = 0;
  progressText.textContent = "0%";
  uploadStatus.textContent = "";

  try {
    // Reset all global data before parse
    IW39 = [];
    Data1 = {};
    Data2 = {};
    SUM57 = {};
    Planning = {};

    for(let i=0; i < files.length; i++) {
      await parseExcelFile(files[i]);
      let perc = Math.round(((i+1)/files.length)*100);
      progressBar.value = perc;
      progressText.textContent = `${perc}%`;
    }

    buildDataLembarKerja();
    renderTable(dataLembarKerja);

    uploadStatus.textContent = `Upload selesai: ${files.length} file berhasil diproses.`;
  } catch (err) {
    uploadStatus.textContent = "Upload gagal: " + err.message;
  }
  progressContainer.classList.add("hidden");
});

// ---------- SIDEBAR MENU ----------
const menuItems = document.querySelectorAll(".menu-item");
const sections = document.querySelectorAll(".content-section");

menuItems.forEach(item => {
  item.addEventListener("click", () => {
    menuItems.forEach(i => i.classList.remove("active"));
    item.classList.add("active");
    const target = item.dataset.menu;
    sections.forEach(s => {
      s.classList.toggle("active", s.id === target);
    });
  });
});

// ---------- INIT ----------
window.addEventListener("DOMContentLoaded", () => {
  loadDataFromStorage();
  renderTable(dataLembarKerja);
});
