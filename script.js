// ----- Library XLSX sudah harus kamu load di HTML -----
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

// ----- Global Data ----- 
let IW39 = [], Data1 = {}, Data2 = {}, SUM57 = {}, Planning = {};
let dataLembarKerja = [];

// ----- Format angka 1 decimal ----- 
function formatNumber(num) {
  return Number(num).toFixed(1);
}

// ----- Fungsi parsing file Excel dan update lookup data global -----
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onloadstart = () => {
      updateProgress(0);
      setProgressVisible(true);
      setUploadStatus("Membaca file...");
    };

    reader.onprogress = (event) => {
      if (event.lengthComputable) {
        let percent = Math.round((event.loaded / event.total) * 100);
        updateProgress(percent);
        setUploadStatus(`Loading file... ${percent}%`);
      }
    };

    reader.onload = function(event) {
      try {
        updateProgress(50);
        setUploadStatus("Parsing file...");
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (workbook.SheetNames.includes('IW39')) {
          IW39 = XLSX.utils.sheet_to_json(workbook.Sheets['IW39']);
        } else {
          IW39 = [];
        }

        if (workbook.SheetNames.includes('Data1')) {
          const data1Arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data1']);
          Data1 = {};
          data1Arr.forEach(row => {
            if(row.Order && row.Section) Data1[row.Order] = row.Section;
          });
        } else {
          Data1 = {};
        }

        if (workbook.SheetNames.includes('Data2')) {
          const data2Arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data2']);
          Data2 = {};
          data2Arr.forEach(row => {
            if(row.MAT && row.CPH) Data2[row.MAT] = row.CPH;
          });
        } else {
          Data2 = {};
        }

        if (workbook.SheetNames.includes('SUM57')) {
          const sumArr = XLSX.utils.sheet_to_json(workbook.Sheets['SUM57']);
          SUM57 = {};
          sumArr.forEach(row => {
            if(row.Order) {
              SUM57[row.Order] = { StatusPart: row.StatusPart || '', Aging: row.Aging || '' };
            }
          });
        } else {
          SUM57 = {};
        }

        if (workbook.SheetNames.includes('Planning')) {
          const planArr = XLSX.utils.sheet_to_json(workbook.Sheets['Planning']);
          Planning = {};
          planArr.forEach(row => {
            if(row.Order) {
              Planning[row.Order] = { Planning: row.Planning || '', StatusAMT: row.StatusAMT || '' };
            }
          });
        } else {
          Planning = {};
        }

        updateProgress(100);
        setUploadStatus(`File ${file.name} berhasil diupload dan data siap.`);
        setTimeout(() => setProgressVisible(false), 1000);
        resolve();

      } catch (error) {
        setProgressVisible(false);
        setUploadStatus("Error: " + error.message);
        reject(error);
      }
    };
    reader.onerror = function(e) {
      setProgressVisible(false);
      setUploadStatus("Error saat membaca file");
      reject(e);
    };
    reader.readAsArrayBuffer(file);
  });
}

// ----- UI Helpers for progress -----
function updateProgress(value) {
  const prog = document.getElementById("upload-progress");
  const txt = document.getElementById("progress-text");
  if(prog && txt){
    prog.value = value;
    txt.textContent = value + "%";
  }
}
function setProgressVisible(visible) {
  const cont = document.getElementById("progress-container");
  if(cont){
    if(visible) cont.classList.remove("hidden");
    else cont.classList.add("hidden");
  }
}
function setUploadStatus(text) {
  const stat = document.getElementById("upload-status");
  if(stat){
    stat.textContent = text;
  }
}

// ----- Build dataLembarKerja with lookup and formulas -----
function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    const iw = IW39.find(i => i.Order && row.Order && i.Order.toLowerCase() === row.Order.toLowerCase()) || {};

    row.Room = iw.Room || row.Room || "";
    row.OrderType = iw.OrderType || row.OrderType || "";
    row.Description = iw.Description || row.Description || "";
    row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
    row.UserStatus = iw.UserStatus || row.UserStatus || "";
    row.MAT = iw.MAT || row.MAT || "";

    if ((row.Description || "").substring(0,2).toUpperCase() === "JR") {
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

    if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
      const costCalc = (iw.TotalPlan - iw.TotalActual) / 16500;
      row.Cost = costCalc < 0 ? "-" : costCalc;
    } else {
      row.Cost = "-";
    }

    if ((row.Reman || "").toLowerCase() === "reman") {
      row.Include = typeof row.Cost === "number" ? row.Cost * 0.25 : "-";
    } else {
      row.Include = row.Cost;
    }

    if ((row.OrderType || "").toUpperCase() === "PM38") {
      row.Exclude = "-";
    } else {
      row.Exclude = row.Include;
    }

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

// ----- Validasi order input -----
function isValidOrder(order) {
  return !/[.,]/.test(order);
}

// ----- Render Tabel Output Menu 2 -----
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

    // Kolom-kolom ...
    ["Room","OrderType","Order","Description","CreatedOn","UserStatus","MAT","CPH","Section","StatusPart","Aging"].forEach((key, idx) => {
      const td = document.createElement("td");
      td.textContent = row[key] || "";
      tr.appendChild(td);
    });

    // Month (editable)
    const tdMonth = document.createElement("td");
    tdMonth.classList.add("editable");
    tdMonth.textContent = row.Month || "";
    tdMonth.title = "Klik untuk edit bulan";
    tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
    tr.appendChild(tdMonth);

    // Cost
    const tdCost = document.createElement("td");
    tdCost.classList.add("cost");
    tdCost.textContent = typeof row.Cost === "number" ? formatNumber(row.Cost) : row.Cost;
    tr.appendChild(tdCost);

    // Reman (editable)
    const tdReman = document.createElement("td");
    tdReman.classList.add("editable");
    tdReman.textContent = row.Reman || "";
    tdReman.title = "Klik untuk edit Reman";
    tdReman.addEventListener("click", () => editReman(tdReman, row));
    tr.appendChild(tdReman);

    // Include
    const tdInclude = document.createElement("td");
    tdInclude.classList.add("include");
    tdInclude.textContent = typeof row.Include === "number" ? formatNumber(row.Include) : row.Include;
    tr.appendChild(tdInclude);

    // Exclude
    const tdExclude = document.createElement("td");
    tdExclude.classList.add("exclude");
    tdExclude.textContent = typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude;
    tr.appendChild(tdExclude);

    // Planning & StatusAMT
    const tdPlanning = document.createElement("td");
    tdPlanning.textContent = row.Planning;
    tr.appendChild(tdPlanning);

    const tdStatusAMT = document.createElement("td");
    tdStatusAMT.textContent = row.StatusAMT;
    tr.appendChild(tdStatusAMT);

    // Action buttons
    const tdAction = document.createElement("td");

    const btnEdit = document.createElement("button");
    btnEdit.textContent = "Edit";
    btnEdit.classList.add("btn-action","btn-edit");
    btnEdit.addEventListener("click", () => {
      editMonthAction(row);
      editRemanAction(row);
    });
    tdAction.appendChild(btnEdit);

    const btnDelete = document.createElement("button");
    btnDelete.textContent = "Delete";
    btnDelete.classList.add("btn-action","btn-delete");
    btnDelete.addEventListener("click", () => {
      if(confirm(`Hapus order ${row.Order}?`)){
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
    if(m === row.Month) option.selected = true;
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
    if(e.key === "Enter") {
      row.Reman = input.value.trim();
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
    } else if(e.key === "Escape") {
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

// ----- Edit action tombol Edit -----
function editMonthAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()){
      const tdMonth = tr.children[11];
      editMonth(tdMonth, row);
    }
  });
}
function editRemanAction(row) {
  const trs = outputTableBody.querySelectorAll("tr");
  trs.forEach(tr => {
    if(tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()){
      const tdReman = tr.children[13];
      editReman(tdReman, row);
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
  if(!rawInput){
    alert("Masukkan minimal satu Order bro!");
    return;
  }

  let orders = rawInput.split(/[\s,\n]+/).map(s => s.trim()).filter(s => s.length > 0);

  let addedCount = 0;
  let skippedOrders = [];
  let invalidOrders = [];

  orders.forEach(order => {
    if(!isValidOrder(order)){
      invalidOrders.push(order);
      return;
    }
    const exists = dataLembarKerja.some(d => d.Order.toLowerCase() === order.toLowerCase());
    if(!exists){
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
  if(invalidOrders.length){
    msg += ` Order tidak valid (ada titik atau koma): ${invalidOrders.join(", ")}.`;
  }
  if(skippedOrders.length){
    msg += ` Order sudah ada dan tidak dimasukkan ulang: ${skippedOrders.join(", ")}.`;
  }
  addOrderStatus.textContent = msg;
});

// ----- Filter buttons -----
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
      const orderTypeMatch = row.OrderType.toLowerCase().includes(filterOrder.value.toLowerCase());
      const cphMatch = row.CPH.toLowerCase().includes(filterCPH.value.toLowerCase());
      const matMatch = row.MAT.toLowerCase().includes(filterMAT.value.toLowerCase());
      const sectionMatch = row.Section.toLowerCase().includes(filterSection.value.toLowerCase());
      return roomMatch && orderTypeMatch && cphMatch && matMatch && sectionMatch;
    });
    renderTable(filtered);
  });
});

// ----- File upload handler -----
const fileInput = document.getElementById("file-input");
fileInput.value = ""; // reset input awal

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if(!file) return;

  try {
    await parseExcelFile(file);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  } catch(err) {
    alert("Error saat membaca file Excel: " + err.message);
  }
});

// ----- Sidebar menu handling -----
const menuItems = document.querySelectorAll(".menu-item");
const contentSections = document.querySelectorAll(".content-section");

menuItems.forEach(menu => {
  menu.addEventListener("click", () => {
    // Remove active class from all menu and sections
    menuItems.forEach(m => m.classList.remove("active"));
    contentSections.forEach(c => c.classList.remove("active"));

    // Add active class to clicked menu and corresponding section
    menu.classList.add("active");
    const selectedMenu = menu.getAttribute("data-menu");
    const section = document.getElementById(selectedMenu);
    if(section){
      section.classList.add("active");
    }
  });
});
