// Pastikan kode ini bungkus DOMContentLoaded supaya element pasti ada saat akses
window.addEventListener("DOMContentLoaded", () => {

  // ----- Library XLSX sudah di load di HTML -----

  // ----- Global Data -----
  let IW39 = [], Data1 = {}, Data2 = {}, SUM57 = {}, Planning = {};
  let dataLembarKerja = [];

  // ----- Format angka 1 decimal -----
  function formatNumber(num) {
    return Number(num).toFixed(1);
  }

  // ----- Parsing Excel dan update data global -----
  function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function(event) {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          if (workbook.SheetNames.includes("IW39")) {
            IW39 = XLSX.utils.sheet_to_json(workbook.Sheets["IW39"]);
          } else {
            IW39 = [];
          }

          if (workbook.SheetNames.includes("Data1")) {
            const data1Arr = XLSX.utils.sheet_to_json(workbook.Sheets["Data1"]);
            Data1 = {};
            data1Arr.forEach(row => {
              if (row.Order && row.Section) Data1[row.Order] = row.Section;
            });
          } else {
            Data1 = {};
          }

          if (workbook.SheetNames.includes("Data2")) {
            const data2Arr = XLSX.utils.sheet_to_json(workbook.Sheets["Data2"]);
            Data2 = {};
            data2Arr.forEach(row => {
              if (row.MAT && row.CPH) Data2[row.MAT] = row.CPH;
            });
          } else {
            Data2 = {};
          }

          if (workbook.SheetNames.includes("SUM57")) {
            const sumArr = XLSX.utils.sheet_to_json(workbook.Sheets["SUM57"]);
            SUM57 = {};
            sumArr.forEach(row => {
              if (row.Order) {
                SUM57[row.Order] = { StatusPart: row.StatusPart || "", Aging: row.Aging || "" };
              }
            });
          } else {
            SUM57 = {};
          }

          if (workbook.SheetNames.includes("Planning")) {
            const planArr = XLSX.utils.sheet_to_json(workbook.Sheets["Planning"]);
            Planning = {};
            planArr.forEach(row => {
              if (row.Order) {
                Planning[row.Order] = { Planning: row.Planning || "", StatusAMT: row.StatusAMT || "" };
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

  // ----- Build data dengan lookup dan rumus -----
  function buildDataLembarKerja() {
    dataLembarKerja = dataLembarKerja.map(row => {
      const iw = IW39.find(i => i.Order && row.Order && i.Order.toLowerCase() === row.Order.toLowerCase()) || {};

      row.Room = iw.Room || row.Room || "";
      row.OrderType = iw.OrderType || row.OrderType || "";
      row.Description = iw.Description || row.Description || "";
      row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
      row.UserStatus = iw.UserStatus || row.UserStatus || "";
      row.MAT = iw.MAT || row.MAT || "";

      // CPH: jika 2 huruf pertama Description = "JR" => "JR", else lookup Data2 by MAT
      if ((row.Description || "").substring(0, 2).toUpperCase() === "JR") {
        row.CPH = "JR";
      } else {
        row.CPH = Data2[row.MAT] || "";
      }

      // Section lookup Data1 by Order
      row.Section = Data1[row.Order] || "";

      // Status Part & Aging lookup SUM57 by Order
      if (SUM57[row.Order]) {
        row.StatusPart = SUM57[row.Order].StatusPart || "";
        row.Aging = SUM57[row.Order].Aging || "";
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

      // Planning & StatusAMT lookup Planning by Order
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

  // ----- Validasi order input (tidak boleh titik atau koma) -----
  function isValidOrder(order) {
    return !/[.,]/.test(order);
  }

  // ----- Render tabel output menu 2 -----
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

      // Buat sel kolom
      ["Room","OrderType","Order","Description","CreatedOn","UserStatus","MAT","CPH","Section","StatusPart","Aging"].forEach(field => {
        const td = document.createElement("td");
        td.textContent = row[field] || "";
        tr.appendChild(td);
      });

      // Month (editable)
      const tdMonth = document.createElement("td");
      tdMonth.classList.add("editable");
      tdMonth.textContent = row.Month || "";
      tdMonth.title = "Klik untuk edit bulan";
      tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost (rata kanan)
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

      // Include (rata kanan)
      const tdInclude = document.createElement("td");
      tdInclude.classList.add("include");
      tdInclude.textContent = typeof row.Include === "number" ? formatNumber(row.Include) : row.Include;
      tr.appendChild(tdInclude);

      // Exclude (rata kanan)
      const tdExclude = document.createElement("td");
      tdExclude.classList.add("exclude");
      tdExclude.textContent = typeof row.Exclude === "number" ? formatNumber(row.Exclude) : row.Exclude;
      tr.appendChild(tdExclude);

      // Planning
      const tdPlanning = document.createElement("td");
      tdPlanning.textContent = row.Planning || "";
      tr.appendChild(tdPlanning);

      // StatusAMT
      const tdStatusAMT = document.createElement("td");
      tdStatusAMT.textContent = row.StatusAMT || "";
      tr.appendChild(tdStatusAMT);

      // Action: Edit & Delete
      const tdAction = document.createElement("td");

      // Edit button (fokus Month & Reman)
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

  // ----- Tombol menu sidebar -----
  const menuItems = document.querySelectorAll(".menu-item");
  const contentSections = document.querySelectorAll(".content-section");

  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      contentSections.forEach(section => section.classList.remove("active"));
      item.classList.add("active");

      const target = item.getAttribute("data-menu");
      const targetSection = document.getElementById(target);
      if (targetSection) targetSection.classList.add("active");
    });
  });

  // ----- Upload File Menu -----
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const uploadBtn = document.getElementById("upload-btn");
  const progressContainer = document.getElementById("progress-container");
  const uploadProgress = document.getElementById("upload-progress");
  const progressText = document.getElementById("progress-text");
  const uploadStatus = document.getElementById("upload-status");

  uploadBtn.addEventListener("click", async () => {
    const selectedFileName = fileSelect.value;
    const file = fileInput.files[0];

    if (!file) {
      alert("Pilih file terlebih dahulu!");
      return;
    }

    progressContainer.classList.remove("hidden");
    uploadProgress.value = 0;
    progressText.textContent = "0%";
    uploadStatus.textContent = "";

    try {
      // Parse file dan update data sesuai sheet yang ada
      await parseExcelFile(file);

      uploadProgress.value = 100;
      progressText.textContent = "100%";
      uploadStatus.style.color = "green";
      uploadStatus.textContent = `File "${file.name}" berhasil di-upload dan diproses.`;

      // Jika sheet IW39 ada, update dataLembarKerja langsung pakai IW39
      if (selectedFileName === "IW39") {
        dataLembarKerja = IW39.map(row => ({ ...row }));
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
      }
      // Kalau upload file lain bisa kamu tambah sesuai logic
    } catch (err) {
      uploadStatus.style.color = "red";
      uploadStatus.textContent = "Error saat memproses file: " + err.message;
    }
  });

  // ----- Lembar Kerja: Add Orders -----
  const addOrderInput = document.getElementById("add-order-input");
  const addOrderBtn = document.getElementById("add-order-btn");
  const addOrderStatus = document.getElementById("add-order-status");

  addOrderBtn.addEventListener("click", () => {
    const inputText = addOrderInput.value.trim();
    if (!inputText) {
      addOrderStatus.style.color = "red";
      addOrderStatus.textContent = "Input order tidak boleh kosong!";
      return;
    }

    const ordersRaw = inputText.split(/[\s,]+/);
    const orders = ordersRaw.filter(o => o && isValidOrder(o));

    if (orders.length === 0) {
      addOrderStatus.style.color = "red";
      addOrderStatus.textContent = "Order tidak valid atau mengandung karakter terlarang!";
      return;
    }

    // Tambah order ke dataLembarKerja tanpa duplikat
    orders.forEach(order => {
      if (!dataLembarKerja.some(d => d.Order.toLowerCase() === order.toLowerCase())) {
        dataLembarKerja.push({ Order: order });
      }
    });

    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    addOrderStatus.style.color = "green";
    addOrderStatus.textContent = "Order berhasil ditambahkan.";
    addOrderInput.value = "";
  });

  // ----- Filter data di Lembar Kerja -----
  const filterRoom = document.getElementById("filter-room");
  const filterOrder = document.getElementById("filter-order");
  const filterCPH = document.getElementById("filter-cph");
  const filterMAT = document.getElementById("filter-mat");
  const filterSection = document.getElementById("filter-section");
  const filterBtn = document.getElementById("filter-btn");
  const resetBtn = document.getElementById("reset-btn");

  filterBtn.addEventListener("click", () => {
    let filtered = dataLembarKerja;

    if (filterRoom.value.trim() !== "") {
      filtered = filtered.filter(d => (d.Room || "").toLowerCase().includes(filterRoom.value.toLowerCase()));
    }
    if (filterOrder.value.trim() !== "") {
      filtered = filtered.filter(d => (d.Order || "").toLowerCase().includes(filterOrder.value.toLowerCase()));
    }
    if (filterCPH.value.trim() !== "") {
      filtered = filtered.filter(d => (d.CPH || "").toLowerCase().includes(filterCPH.value.toLowerCase()));
    }
    if (filterMAT.value.trim() !== "") {
      filtered = filtered.filter(d => (d.MAT || "").toLowerCase().includes(filterMAT.value.toLowerCase()));
    }
    if (filterSection.value.trim() !== "") {
      filtered = filtered.filter(d => (d.Section || "").toLowerCase().includes(filterSection.value.toLowerCase()));
    }

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

  // ----- Save & Load dataLembarKerja ke localStorage -----
  const saveBtn = document.getElementById("save-btn");
  const loadBtn = document.getElementById("load-btn");

  saveBtn.addEventListener("click", () => {
    try {
      localStorage.setItem("lembarKerjaData", JSON.stringify(dataLembarKerja));
      alert("Data berhasil disimpan di browser.");
    } catch (e) {
      alert("Gagal menyimpan data: " + e.message);
    }
  });

  loadBtn.addEventListener("click", () => {
    const savedData = localStorage.getItem("lembarKerjaData");
    if (savedData) {
      try {
        dataLembarKerja = JSON.parse(savedData);
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        alert("Data berhasil dimuat dari browser.");
      } catch (e) {
        alert("Data tidak valid: " + e.message);
      }
    } else {
      alert("Tidak ada data yang disimpan.");
    }
  });

  // ----- Inisialisasi awal -----
  renderTable(dataLembarKerja);

});
