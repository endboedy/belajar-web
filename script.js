// script.js
// Pastikan XLSX sudah dimuat di HTML <head> sebelum file ini

window.addEventListener("DOMContentLoaded", () => {

  // -----------------------
  // Utility helpers
  // -----------------------
  const safeStr = v => (v === null || v === undefined) ? "" : String(v).trim();
  const safeLower = v => safeStr(v).toLowerCase();
  const isNumberLike = v => !isNaN(Number(v)) && v !== "";

  function mapRowKeys(row) {
    // Normalize various header names to our expected keys
    const mapped = {};
    for (const k in row) {
      const lk = k.trim().toLowerCase();
      const v = row[k];
      if (lk === "order" || lk === "order no" || lk === "ordernumber" || lk.includes("order")) mapped.Order = safeStr(v);
      else if (lk === "room" || lk.includes("room")) mapped.Room = safeStr(v);
      else if (lk === "ordertype" || lk.includes("order type")) mapped.OrderType = safeStr(v);
      else if (lk === "description" || lk.includes("desc")) mapped.Description = safeStr(v);
      else if (lk === "created on" || lk.includes("created")) mapped.CreatedOn = safeStr(v);
      else if (lk === "user status" || lk.includes("user status") || lk.includes("userstatus")) mapped.UserStatus = safeStr(v);
      else if (lk === "mat" || lk.includes("mat")) mapped.MAT = safeStr(v);
      else if (lk === "totalplan" || lk.includes("total plan")) mapped.TotalPlan = Number(v || 0);
      else if (lk === "totalactual" || lk.includes("total actual")) mapped.TotalActual = Number(v || 0);
      else if (lk === "section") mapped.Section = safeStr(v);
      else if (lk === "cph") mapped.CPH = safeStr(v);
      else if (lk === "statuspart" || lk.includes("status part")) mapped.StatusPart = safeStr(v);
      else if (lk === "aging") mapped.Aging = safeStr(v);
      else if (lk === "planning" || lk.includes("event start")) mapped.Planning = safeStr(v);
      else if (lk === "statusamt" || lk.includes("status amt")) mapped.StatusAMT = safeStr(v);
      else {
        // keep unknown keys as-is (in case)
        mapped[k] = v;
      }
    }
    return mapped;
  }

  // -----------------------
  // Global datasets (filled from uploads)
  // -----------------------
  let IW39 = [];     // array of rows (normalized)
  let Data1 = {};    // mapping Order -> Section
  let Data2 = {};    // mapping MAT -> CPH
  let SUM57 = {};    // mapping Order -> {StatusPart, Aging}
  let Planning = {}; // mapping Order -> {Planning, StatusAMT}

  // Master table (Lembar Kerja) - contains objects with Order at least
  let dataLembarKerja = [];

  // -----------------------
  // DOM references
  // -----------------------
  const fileSelect = document.getElementById("file-select"); // choose which dataset you're uploading
  const fileInput = document.getElementById("file-input");
  const uploadBtn = document.getElementById("upload-btn");
  const progressContainer = document.getElementById("progress-container");
  const uploadProgress = document.getElementById("upload-progress");
  const progressText = document.getElementById("progress-text");
  const uploadStatus = document.getElementById("upload-status");

  const addOrderInput = document.getElementById("add-order-input");
  const addOrderBtn = document.getElementById("add-order-btn");
  const addOrderStatus = document.getElementById("add-order-status");

  const filterRoom = document.getElementById("filter-room");
  const filterOrder = document.getElementById("filter-order");
  const filterCPH = document.getElementById("filter-cph");
  const filterMAT = document.getElementById("filter-mat");
  const filterSection = document.getElementById("filter-section");
  const filterBtn = document.getElementById("filter-btn");
  const resetBtn = document.getElementById("reset-btn");

  const saveBtn = document.getElementById("save-btn");
  const loadBtn = document.getElementById("load-btn");

  const outputTableBody = document.querySelector("#output-table tbody");

  // -----------------------
  // Parse Excel (single file)
  // -----------------------
  function parseExcelFileToJSON(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          // choose first sheet
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          resolve(json);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = err => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  // -----------------------
  // Upload handler
  // -----------------------
  uploadBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    const selectedType = fileSelect.value; // IW39, SUM57, Planning, Budget, Data1, Data2
    if (!file) {
      alert("Pilih file dulu bro!");
      return;
    }

    progressContainer.classList.remove("hidden");
    uploadProgress.value = 0;
    progressText.textContent = "0%";
    uploadStatus.textContent = "";

    try {
      // read file
      const json = await parseExcelFileToJSON(file);

      // normalize each row keys
      const normalized = json.map(r => mapRowKeys(r));

      // depending on selected type, populate datasets
      if (selectedType === "IW39") {
        // map rows: ensure Order is string
        IW39 = normalized.map(r => ({
          Order: safeStr(r.Order),
          Room: safeStr(r.Room),
          OrderType: safeStr(r.OrderType),
          Description: safeStr(r.Description),
          CreatedOn: safeStr(r.CreatedOn),
          UserStatus: safeStr(r.UserStatus),
          MAT: safeStr(r.MAT),
          TotalPlan: Number(r.TotalPlan || 0),
          TotalActual: Number(r.TotalActual || 0)
        }));
        // If master is empty, initialize from IW39; else update existing matching orders
        if (dataLembarKerja.length === 0) {
          dataLembarKerja = IW39.map(i => ({
            Order: safeStr(i.Order),
            Room: i.Room || "",
            OrderType: i.OrderType || "",
            Description: i.Description || "",
            CreatedOn: i.CreatedOn || "",
            UserStatus: i.UserStatus || "",
            MAT: i.MAT || "",
            Month: "",
            Cost: "-",
            Reman: "",
            Include: "-",
            Exclude: "-",
            Section: "",
            CPH: "",
            StatusPart: "",
            Aging: "",
            Planning: "",
            StatusAMT: ""
          }));
        } else {
          // update fields for existing orders
          dataLembarKerja = dataLembarKerja.map(row => {
            const match = IW39.find(i => safeLower(i.Order) === safeLower(row.Order));
            if (match) {
              return {
                ...row,
                Room: match.Room || row.Room,
                OrderType: match.OrderType || row.OrderType,
                Description: match.Description || row.Description,
                CreatedOn: match.CreatedOn || row.CreatedOn,
                UserStatus: match.UserStatus || row.UserStatus,
                MAT: match.MAT || row.MAT
              };
            }
            return row;
          });
        }
      } else if (selectedType === "Data1") {
        // Data1: map Order -> Section
        Data1 = {};
        normalized.forEach(r => {
          if (r.Order) Data1[safeStr(r.Order)] = safeStr(r.Section || r.Section || "");
        });
      } else if (selectedType === "Data2") {
        // Data2: map MAT -> CPH
        Data2 = {};
        normalized.forEach(r => {
          if (r.MAT) Data2[safeStr(r.MAT)] = safeStr(r.CPH || "");
        });
      } else if (selectedType === "SUM57") {
        SUM57 = {};
        normalized.forEach(r => {
          if (r.Order) SUM57[safeStr(r.Order)] = { StatusPart: safeStr(r.StatusPart), Aging: safeStr(r.Aging) };
        });
      } else if (selectedType === "Planning") {
        Planning = {};
        normalized.forEach(r => {
          if (r.Order) Planning[safeStr(r.Order)] = { Planning: safeStr(r.Planning), StatusAMT: safeStr(r.StatusAMT) };
        });
      } else {
        // other sheets - ignore or extend as needed
        console.warn("Uploaded file type not specifically handled:", selectedType);
      }

      // simulate progress to 100%
      uploadProgress.value = 100;
      progressText.textContent = "100%";
      uploadStatus.style.color = "green";
      uploadStatus.textContent = `File "${file.name}" berhasil diupload untuk ${selectedType}.`;

      // rebuild derived fields and render
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
    } catch (err) {
      uploadStatus.style.color = "red";
      uploadStatus.textContent = `Error saat memproses file: ${err && err.message ? err.message : err}`;
      console.error(err);
    } finally {
      setTimeout(() => progressContainer.classList.add("hidden"), 600);
      // clear file input so user can reupload same file if needed
      fileInput.value = "";
    }
  });

  // -----------------------
  // Build/derive dataLembarKerja fields (lookup & formulas)
  // -----------------------
  function buildDataLembarKerja() {
    // ensure Orders are strings
    dataLembarKerja = dataLembarKerja.map(r => ({ ...r, Order: safeStr(r.Order) }));

    dataLembarKerja = dataLembarKerja.map(row => {
      // find IW39 record by Order
      const iw = IW39.find(i => safeLower(i.Order) === safeLower(row.Order)) || {};

      row.Room = iw.Room || row.Room || "";
      row.OrderType = iw.OrderType || row.OrderType || "";
      row.Description = iw.Description || row.Description || "";
      row.CreatedOn = iw.CreatedOn || row.CreatedOn || "";
      row.UserStatus = iw.UserStatus || row.UserStatus || "";
      row.MAT = iw.MAT || row.MAT || "";

      // CPH: if first 2 chars of Description = JR -> JR, else lookup Data2 by MAT
      if (safeLower((row.Description || "").substring(0,2)) === "jr") {
        row.CPH = "JR";
      } else {
        row.CPH = Data2[row.MAT] || row.CPH || "";
      }

      // Section lookup from Data1 by Order
      row.Section = Data1[row.Order] || row.Section || "";

      // StatusPart & Aging from SUM57
      if (SUM57[row.Order]) {
        row.StatusPart = SUM57[row.Order].StatusPart || "";
        row.Aging = SUM57[row.Order].Aging || "";
      } else {
        row.StatusPart = row.StatusPart || "";
        row.Aging = row.Aging || "";
      }

      // Cost = (TotalPlan - TotalActual) / 16500, if < 0 => "-"
      if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined && !isNaN(iw.TotalPlan) && !isNaN(iw.TotalActual)) {
        const costCalc = (Number(iw.TotalPlan) - Number(iw.TotalActual)) / 16500;
        row.Cost = costCalc < 0 ? "-" : Number(costCalc);
      } else {
        row.Cost = row.Cost || "-";
      }

      // Include: if Reman == "Reman" => Cost*0.25 else same as Cost
      if (safeLower(row.Reman) === "reman") {
        row.Include = typeof row.Cost === "number" ? Number((row.Cost * 0.25)) : "-";
      } else {
        row.Include = row.Cost;
      }

      // Exclude: if OrderType == PM38 => "-" else same as Include
      if (safeLower(row.OrderType) === "pm38") {
        row.Exclude = "-";
      } else {
        row.Exclude = row.Include;
      }

      // Planning & StatusAMT from Planning mapping
      if (Planning[row.Order]) {
        row.Planning = Planning[row.Order].Planning || "";
        row.StatusAMT = Planning[row.Order].StatusAMT || "";
      } else {
        row.Planning = row.Planning || "";
        row.StatusAMT = row.StatusAMT || "";
      }

      return row;
    });
  }

  // -----------------------
  // Render table
  // -----------------------
  function renderTable(data) {
    // detect duplicates by order (case-insensitive)
    const ordersLower = data.map(d => safeLower(d.Order));
    const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

    outputTableBody.innerHTML = "";

    if (!data.length) {
      outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#666;">Tidak ada data.</td></tr>`;
      return;
    }

    data.forEach((row, idx) => {
      const tr = document.createElement("tr");
      if (duplicates.includes(safeLower(row.Order))) {
        tr.classList.add("duplicate"); // style in CSS can set bg red + white text
      }

      // helper to create cell
      const mk = text => {
        const td = document.createElement("td");
        td.textContent = text ?? "";
        return td;
      };

      tr.appendChild(mk(row.Room));
      tr.appendChild(mk(row.OrderType));
      tr.appendChild(mk(row.Order));
      tr.appendChild(mk(row.Description));
      tr.appendChild(mk(row.CreatedOn));
      tr.appendChild(mk(row.UserStatus));
      tr.appendChild(mk(row.MAT));
      tr.appendChild(mk(row.CPH));
      tr.appendChild(mk(row.Section));
      tr.appendChild(mk(row.StatusPart));
      tr.appendChild(mk(row.Aging));

      // Month (editable cell)
      const tdMonth = document.createElement("td");
      tdMonth.classList.add("editable");
      tdMonth.textContent = row.Month || "";
      tdMonth.title = "Klik untuk edit Month";
      tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost (right aligned)
      const tdCost = document.createElement("td");
      tdCost.classList.add("cost");
      tdCost.textContent = (typeof row.Cost === "number") ? Number(row.Cost).toFixed(1) : row.Cost;
      tr.appendChild(tdCost);

      // Reman editable
      const tdReman = document.createElement("td");
      tdReman.classList.add("editable");
      tdReman.textContent = row.Reman || "";
      tdReman.title = "Klik untuk edit Reman";
      tdReman.addEventListener("click", () => editReman(tdReman, row));
      tr.appendChild(tdReman);

      // Include
      const tdInclude = document.createElement("td");
      tdInclude.classList.add("include");
      tdInclude.textContent = (typeof row.Include === "number") ? Number(row.Include).toFixed(1) : row.Include;
      tr.appendChild(tdInclude);

      // Exclude
      const tdExclude = document.createElement("td");
      tdExclude.classList.add("exclude");
      tdExclude.textContent = (typeof row.Exclude === "number") ? Number(row.Exclude).toFixed(1) : row.Exclude;
      tr.appendChild(tdExclude);

      // Planning & StatusAMT
      tr.appendChild(mk(row.Planning));
      tr.appendChild(mk(row.StatusAMT));

      // Action: Edit (focus month & reman) and Delete
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
          dataLembarKerja = dataLembarKerja.filter(d => safeLower(d.Order) !== safeLower(row.Order));
          buildDataLembarKerja();
          renderTable(dataLembarKerja);
          saveDataToLocalStorage();
        }
      });
      tdAction.appendChild(btnDelete);

      tr.appendChild(tdAction);

      outputTableBody.appendChild(tr);
    });
  }

  // -----------------------
  // Inline editing helpers
  // -----------------------
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
      saveDataToLocalStorage();
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
        saveDataToLocalStorage();
      } else if (e.key === "Escape") {
        renderTable(dataLembarKerja);
      }
    });
    input.addEventListener("blur", () => {
      row.Reman = input.value.trim();
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });

    td.textContent = "";
    td.appendChild(input);
    input.focus();
  }

  function editMonthAction(row) {
    // find row in table and open month select
    const trs = Array.from(outputTableBody.querySelectorAll("tr"));
    trs.forEach(tr => {
      const orderCell = tr.children[2];
      if (orderCell && safeLower(orderCell.textContent) === safeLower(row.Order)) {
        editMonth(tr.children[11], row); // month index according to table structure
      }
    });
  }

  function editRemanAction(row) {
    const trs = Array.from(outputTableBody.querySelectorAll("tr"));
    trs.forEach(tr => {
      const orderCell = tr.children[2];
      if (orderCell && safeLower(orderCell.textContent) === safeLower(row.Order)) {
        editReman(tr.children[13], row); // reman index
      }
    });
  }

  // -----------------------
  // Add Order
  // -----------------------
  addOrderBtn.addEventListener("click", () => {
    const raw = addOrderInput.value.trim();
    if (!raw) {
      addOrderStatus.style.color = "red";
      addOrderStatus.textContent = "Masukkan minimal 1 order.";
      return;
    }
    // split by whitespace or comma or newline
    const orders = raw.split(/[\s,]+/).map(s => s.trim()).filter(Boolean);

    let added = 0;
    let skipped = [];
    let invalid = [];

    orders.forEach(o => {
      if (!isValidOrder(o)) { invalid.push(o); return; }
      if (dataLembarKerja.some(d => safeLower(d.Order) === safeLower(o))) { skipped.push(o); return; }
      dataLembarKerja.push({
        Order: safeStr(o),
        Room: "", OrderType: "", Description: "", CreatedOn: "", UserStatus: "",
        MAT: "", CPH: "", Section: "", StatusPart: "", Aging: "", Month: "",
        Cost: "-", Reman: "", Include: "-", Exclude: "-", Planning: "", StatusAMT: ""
      });
      added++;
    });

    buildDataLembarKerja();
    renderTable(dataLembarKerja);
    saveDataToLocalStorage();

    let msg = `${added} order ditambahkan.`;
    if (skipped.length) msg += ` Sudah ada: ${skipped.join(", ")}.`;
    if (invalid.length) msg += ` Invalid: ${invalid.join(", ")}.`;
    addOrderStatus.style.color = added ? "green" : "red";
    addOrderStatus.textContent = msg;
    addOrderInput.value = "";
  });

  // -----------------------
  // Filters
  // -----------------------
  filterBtn.addEventListener("click", () => {
    let filtered = dataLembarKerja;
    if (filterRoom.value.trim()) filtered = filtered.filter(d => (d.Room || "").toLowerCase().includes(filterRoom.value.trim().toLowerCase()));
    if (filterOrder.value.trim()) filtered = filtered.filter(d => (d.Order || "").toLowerCase().includes(filterOrder.value.trim().toLowerCase()));
    if (filterCPH.value.trim()) filtered = filtered.filter(d => (d.CPH || "").toLowerCase().includes(filterCPH.value.trim().toLowerCase()));
    if (filterMAT.value.trim()) filtered = filtered.filter(d => (d.MAT || "").toLowerCase().includes(filterMAT.value.trim().toLowerCase()));
    if (filterSection.value.trim()) filtered = filtered.filter(d => (d.Section || "").toLowerCase().includes(filterSection.value.trim().toLowerCase()));
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

  // -----------------------
  // Save / Load localStorage
  // -----------------------
  function saveDataToLocalStorage() {
    try {
      localStorage.setItem("lembarKerjaData", JSON.stringify(dataLembarKerja));
    } catch (e) {
      console.warn("Gagal simpan ke localStorage:", e);
    }
  }

  saveBtn.addEventListener("click", () => {
    saveDataToLocalStorage();
    alert("Data tersimpan di browser.");
  });

  loadBtn.addEventListener("click", () => {
    const saved = localStorage.getItem("lembarKerjaData");
    if (!saved) { alert("Tidak ada data tersimpan."); return; }
    try {
      dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) }));
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      alert("Data dimuat dari penyimpanan lokal.");
    } catch (e) {
      alert("Gagal memuat data: " + e.message);
    }
  });

  // -----------------------
  // Validation for Order (no dot/comma)
  // -----------------------
  function isValidOrder(order) {
    return !/[.,]/.test(order);
  }

  // -----------------------
  // Sidebar menu switching (no duplicate declaration)
  // -----------------------
  const menuItems = document.querySelectorAll(".menu-item");
  const contentSections = document.querySelectorAll(".content-section");
  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      contentSections.forEach(s => s.classList.remove("active"));
      item.classList.add("active");
      const target = item.getAttribute("data-menu");
      const sec = document.getElementById(target);
      if (sec) sec.classList.add("active");
    });
  });

  // -----------------------
  // Init: try load saved data automatically
  // -----------------------
  (function init() {
    const saved = localStorage.getItem("lembarKerjaData");
    if (saved) {
      try {
        dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) }));
      } catch (e) {
        dataLembarKerja = [];
      }
    }
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  })();

});
