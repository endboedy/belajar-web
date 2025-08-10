// ----- Library XLSX sudah harus kamu load di HTML -----
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

document.addEventListener("DOMContentLoaded", () => {
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
      reader.onload = function(event) {
        try {
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
              if(row.Order && row.Section) Data1[row.Order.toString().trim()] = row.Section;
            });
          } else {
            Data1 = {};
          }

          if (workbook.SheetNames.includes('Data2')) {
            const data2Arr = XLSX.utils.sheet_to_json(workbook.Sheets['Data2']);
            Data2 = {};
            data2Arr.forEach(row => {
              if(row.MAT && row.CPH) Data2[row.MAT.toString().trim()] = row.CPH;
            });
          } else {
            Data2 = {};
          }

          if (workbook.SheetNames.includes('SUM57')) {
            const sumArr = XLSX.utils.sheet_to_json(workbook.Sheets['SUM57']);
            SUM57 = {};
            sumArr.forEach(row => {
              if(row.Order) {
                SUM57[row.Order.toString().trim()] = { 
                  StatusPart: row.StatusPart || '', 
                  Aging: row.Aging || '' 
                };
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
                Planning[row.Order.toString().trim()] = { 
                  Planning: row.Planning || '', 
                  StatusAMT: row.StatusAMT || '' 
                };
              }
            });
          } else {
            Planning = {};
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

  // ----- Build dataLembarKerja with lookup and formulas -----
  function buildDataLembarKerja() {
    dataLembarKerja = dataLembarKerja.map(row => {
      // Cari data lengkap di IW39 berdasarkan Order key
      const iw = IW39.find(i => i.Order && row.Order && i.Order.toString().trim().toLowerCase() === row.Order.toString().trim().toLowerCase()) || {};

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
        row.CPH = Data2[row.MAT.toString().trim()] || "";
      }

      // Section lookup Data1 by Order
      row.Section = Data1[row.Order.toString().trim()] || "";

      // Status Part & Aging lookup SUM57 by Order
      if (SUM57[row.Order.toString().trim()]) {
        row.StatusPart = SUM57[row.Order.toString().trim()].StatusPart || "";
        row.Aging = SUM57[row.Order.toString().trim()].Aging || "";
      } else {
        row.StatusPart = "";
        row.Aging = "";
      }

      // Cost = (IW39.TotalPlan - IW39.TotalActual) / 16500, jika < 0 maka "-"
      if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined) {
        const costCalc = (Number(iw.TotalPlan) - Number(iw.TotalActual)) / 16500;
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
      if (Planning[row.Order.toString().trim()]) {
        row.Planning = Planning[row.Order.toString().trim()].Planning || "";
        row.StatusAMT = Planning[row.Order.toString().trim()].StatusAMT || "";
      } else {
        row.Planning = "";
        row.StatusAMT = "";
      }

      return row;
    });
  }

  // ----- Validasi order input (tidak boleh titik atau koma) -----
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

      // Room
      const tdRoom = document.createElement("td");
      tdRoom.textContent = row.Room;
      tr.appendChild(tdRoom);

      // OrderType
      const tdOrderType = document.createElement("td");
      tdOrderType.textContent = row.OrderType;
      tr.appendChild(tdOrderType);

      // Order
      const tdOrder = document.createElement("td");
      tdOrder.textContent = row.Order;
      tr.appendChild(tdOrder);

      // Description
      const tdDescription = document.createElement("td");
      tdDescription.textContent = row.Description;
      tr.appendChild(tdDescription);

      // CreatedOn
      const tdCreatedOn = document.createElement("td");
      tdCreatedOn.textContent = row.CreatedOn;
      tr.appendChild(tdCreatedOn);

      // UserStatus
      const tdUserStatus = document.createElement("td");
      tdUserStatus.textContent = row.UserStatus;
      tr.appendChild(tdUserStatus);

      // MAT
      const tdMAT = document.createElement("td");
      tdMAT.textContent = row.MAT;
      tr.appendChild(tdMAT);

      // CPH
      const tdCPH = document.createElement("td");
      tdCPH.textContent = row.CPH;
      tr.appendChild(tdCPH);

      // Section
      const tdSection = document.createElement("td");
      tdSection.textContent = row.Section;
      tr.appendChild(tdSection);

      // StatusPart
      const tdStatusPart = document.createElement("td");
      tdStatusPart.textContent = row.StatusPart;
      tr.appendChild(tdStatusPart);

      // Aging
      const tdAging = document.createElement("td");
      tdAging.textContent = row.Aging;
      tr.appendChild(tdAging);

      // Month (editable)
      const tdMonth = document.createElement("td");
      tdMonth.classList.add("editable");
      tdMonth.textContent = row.Month || "";
      tdMonth.title = "Klik untuk edit bulan";
      tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost (rata kanan, 1 decimal)
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

      // Include (rata kanan, 1 decimal)
      const tdInclude = document.createElement("td");
      tdInclude.classList.add("include");
      tdInclude.textContent = typeof row.Include === "number" ? formatNumber(row.Include) : row.Include;
      tr.appendChild(tdInclude);

      // Exclude (rata kanan, 1 decimal)
      const tdExclude = document.createElement("td");
      tdExclude.classList.add("exclude");
      tdExclude.textContent = typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude;
      tr.appendChild(tdExclude);

      // Planning
      const tdPlanning = document.createElement("td");
      tdPlanning.textContent = row.Planning;
      tr.appendChild(tdPlanning);

      // StatusAMT
      const tdStatusAMT = document.createElement("td");
      tdStatusAMT.textContent = row.StatusAMT;
      tr.appendChild(tdStatusAMT);

      // Action: Edit & Delete Buttons
      const tdAction = document.createElement("td");

      // Edit button (fokus ke Month & Reman)
      const btnEdit = document.createElement("button");
      btnEdit.textContent = "Edit";
      btnEdit.classList.add("btn-action", "btn-edit");
      btnEdit.addEventListener("click", () => {
        editMonthAction(row);
        editRemanAction(row);
      });
      tdAction.appendChild(btnEdit);

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
      tdAction.appendChild(btnDelete);

      tr.appendChild(tdAction);

      outputTableBody.appendChild(tr);
    });
  }

  // ----- Edit fungsi inline untuk Month -----
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

  // ----- Edit fungsi inline untuk Reman -----
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

  // ----- Fungsi Edit action via tombol Edit (fokus Month & Reman) -----
  function editMonthAction(row) {
    // Cari elemen baris sesuai order
    const trs = outputTableBody.querySelectorAll("tr");
    trs.forEach(tr => {
      if (tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
        // Month td ada di index 11 (count dari 0)
        const tdMonth = tr.children[11];
        editMonth(tdMonth, row);
      }
    });
  }

  function editRemanAction(row) {
    const trs = outputTableBody.querySelectorAll("tr");
    trs.forEach(tr => {
      if (tr.children[2].textContent.toLowerCase() === row.Order.toLowerCase()) {
        // Reman td ada di index 13
        const tdReman = tr.children[13];
        editReman(tdReman, row);
      }
    });
  }

  // ----- Validasi order input (tidak boleh titik atau koma) -----
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
  });

  // ----- Filter input elements -----
  const filterRoom = document.getElementById("filter-room");
  const filterOrderType = document.getElementById("filter-order");
  const filterSection = document.getElementById("filter-section");
  // Kalau filter month dan reman belum ada di HTML, skip dulu
  const filterMonth = document.getElementById("filter-month"); // bisa null
  const filterReman = document.getElementById("filter-reman"); // bisa null

  // Buat array filter input yg valid (exist)
  const filterInputs = [filterRoom, filterOrderType, filterSection];
  if(filterMonth) filterInputs.push(filterMonth);
  if(filterReman) filterInputs.push(filterReman);

  filterInputs.forEach(input => {
    if(!input) return;
    input.addEventListener("input", () => {
      const filtered = dataLembarKerja.filter(row => {
        const roomMatch = row.Room.toLowerCase().includes(filterRoom.value.toLowerCase());
        const orderTypeMatch = row.OrderType.toLowerCase().includes(filterOrderType.value.toLowerCase());
        const sectionMatch = row.Section.toLowerCase().includes(filterSection.value.toLowerCase());
        const monthMatch = filterMonth ? row.Month.toLowerCase().includes(filterMonth.value.toLowerCase()) : true;
        const remanMatch = filterReman ? row.Reman.toLowerCase().includes(filterReman.value.toLowerCase()) : true;
        return roomMatch && orderTypeMatch && sectionMatch && monthMatch && remanMatch;
      });
      renderTable(filtered);
    });
  });

  // ----- File upload handler -----
  const fileInput = document.getElementById("file-input");
  const fileStatus = document.getElementById("file-status");

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      await parseExcelFile(file);
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      fileStatus.textContent = `File ${file.name} berhasil diupload dan data siap.`;
    } catch (err) {
      alert("Error saat membaca file Excel: " + err.message);
    }
  });

  // --- Inisialisasi awal ---
  renderTable(dataLembarKerja);
});
