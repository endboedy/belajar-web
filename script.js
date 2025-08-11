// script.js
// Pastikan XLSX sudah dimuat di HTML <head> sebelum file ini.
// Full robust implementation: upload (multi type), lookup, add order, edit inline, filter, save/load.

window.addEventListener("DOMContentLoaded", () => {

  // ---------- Helpers ----------
  const safeStr = v => (v === null || v === undefined) ? "" : String(v).trim();
  const safeLower = v => safeStr(v).toLowerCase();
  const isNumeric = v => !isNaN(Number(v)) && v !== "";
  const monthAbbr = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // Excel serial to JS Date (handles Excel 1900 leap bug)
  function excelDateToJS(n) {
    // If input already a date string, return Date try
    if (!isFinite(n)) return null;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
    const days = Math.floor(Number(n));
    const ms = days * 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + ms);
  }

  function formatToDDMMMYYYY(val) {
    if (val === null || val === undefined || val === "") return "";
    // If numeric -> treat as Excel serial possibly
    if (isNumeric(val)) {
      const d = excelDateToJS(Number(val));
}
      if (d && !isNaN(d.getTime())) {
        return `${String(d.getDate()).padStart(2,'0')}-${monthAbbr[d.getMonth()]}-${d.getFullYear()}`;
}
      }
      // fallback to raw
      return String(val);
    }
    // try parse date string
    const tryDate = new Date(val);
    if (!isNaN(tryDate.getTime())) {
      return `${String(tryDate.getDate()).padStart(2,'0')}-${monthAbbr[tryDate.getMonth()]}-${tryDate.getFullYear()}`;
}
    }
    // fallback: return original
    return String(val);
  }

  // map possible header names to canonical keys
  function mapRowKeys(row) {
    const mapped = {};
    Object.keys(row).forEach(k => {
      const lk = k.trim().toLowerCase();
      const v = row[k];
      if (lk === "order" || lk === "order no" || lk.includes("order") && !lk.includes("order type")) mapped.Order = safeStr(v);
      else if (lk === "room" || lk.includes("room")) mapped.Room = safeStr(v);
      else if (lk === "order type" || lk === "ordertype" || lk.includes("order type")) mapped.OrderType = safeStr(v);
      else if (lk === "description" || lk.includes("desc")) mapped.Description = safeStr(v);
      else if (lk === "created on" || lk.includes("created")) mapped.CreatedOn = safeStr(v);
      else if (lk === "user status" || lk.includes("user status")) mapped.UserStatus = safeStr(v);
      else if (lk === "mat" || lk.includes("mat")) mapped.MAT = safeStr(v);
      else if (lk === "totalplan" || lk.includes("total plan")) mapped.TotalPlan = Number(v || 0);
      else if (lk === "totalactual" || lk.includes("total actual")) mapped.TotalActual = Number(v || 0);
      else if (lk === "section" || lk.includes("section")) mapped.Section = safeStr(v);
      else if (lk === "cph" || lk.includes("cph")) mapped.CPH = safeStr(v);
      else if (lk === "status part" || lk.includes("statuspart")) mapped.StatusPart = safeStr(v);
      else if (lk === "aging" || lk.includes("aging")) mapped.Aging = safeStr(v);
      else if (lk === "planning" || lk.includes("event start") || lk.includes("event_start")) mapped.Planning = safeStr(v);
      else if (lk === "statusamt" || lk.includes("status amt")) mapped.StatusAMT = safeStr(v);
      else mapped[k] = v;
    });
    return mapped;
  }

  // ---------- Global datasets ----------
  let IW39 = [];     // array of normalized objects
  let Data1 = {};    // Order -> Section
  let Data2 = {};    // MAT -> CPH
  let SUM57 = {};    // Order -> {StatusPart, Aging}
  let Planning = {}; // Order -> {Planning, StatusAMT}

  // Master Lembar Kerja rows
  let dataLembarKerja = [];

  // ---------- DOM refs ----------
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const uploadBtn = document.getElementById("upload-btn");
if (uploadBtn) {
  const progressContainer = document.getElementById("progress-container");
}
  const uploadProgress = document.getElementById("upload-progress");
  const progressText = document.getElementById("progress-text");
  const uploadStatus = document.getElementById("upload-status");

  const addOrderInput = document.getElementById("add-order-input");
  const addOrderBtn = document.getElementById("add-order-btn");
if (addOrderBtn) {
  const addOrderStatus = document.getElementById("add-order-status");
}

  const filterRoom = document.getElementById("filter-room");
  const filterOrder = document.getElementById("filter-order");
  const filterCPH = document.getElementById("filter-cph");
  const filterMAT = document.getElementById("filter-mat");
  const filterSection = document.getElementById("filter-section");
  const filterBtn = document.getElementById("filter-btn");
if (filterBtn) {
  const resetBtn = document.getElementById("reset-btn");
}
if (resetBtn) {

}
  const saveBtn = document.getElementById("save-btn");
if (saveBtn) {
  const loadBtn = document.getElementById("load-btn");
}
if (loadBtn) {

}
  const outputTableBody = document.querySelector("#output-table tbody");

  // ---------- Excel parsing ----------
  function parseExcelFileToJSON(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          resolve({sheetName, json});
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = err => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  // ---------- Upload handler ----------
  uploadBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    const selectedType = fileSelect.value; // IW39, SUM57, Planning, Budget, Data1, Data2

    if (!file) {
      alert("Pilih file dulu bro!");
}
      return;
    }

    progressContainer.classList.remove("hidden");
    uploadProgress.value = 0;
    progressText.textContent = "0%";
    uploadStatus.textContent = "";

    try {
      // parse
      const {sheetName, json} = await parseExcelFileToJSON(file);

      // If Budget sheet selected or sheetName includes "budget", skip gracefully
      if (selectedType.toLowerCase() === "budget" || sheetName.toLowerCase().includes("budget")) {
        uploadProgress.value = 100;
}
        progressText.textContent = "100%";
        uploadStatus.style.color = "green";
        uploadStatus.textContent = `File "${file.name}" (Budget) di-skip (tidak digunakan).`;
        // hide progress shortly
        setTimeout(()=> progressContainer.classList.add("hidden"), 600);
        fileInput.value = "";
        return;
      }

      // Normalize rows
      const normalized = json.map(r => mapRowKeys(r));

      // Process into appropriate dataset based on selectedType
      if (selectedType === "IW39") {
        // Build IW39 array with canonical keys
        IW39 = normalized.map(r => ({
          Order: safeStr(r.Order),
}
          Room: safeStr(r.Room),
          OrderType: safeStr(r.OrderType),
          Description: safeStr(r.Description),
          CreatedOn: r.CreatedOn !== undefined ? r.CreatedOn : safeStr(r.CreatedOn),
          UserStatus: safeStr(r.UserStatus),
          MAT: safeStr(r.MAT),
          TotalPlan: Number(r.TotalPlan || 0),
          TotalActual: Number(r.TotalActual || 0)
        }));
        // If master empty -> initialize
        if (dataLembarKerja.length === 0) {
          dataLembarKerja = IW39.map(i => ({
            Order: safeStr(i.Order),
}
            Room: i.Room || "",
            OrderType: i.OrderType || "",
            Description: i.Description || "",
            CreatedOn: i.CreatedOn || "",
            UserStatus: i.UserStatus || "",
            MAT: i.MAT || "",
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
        } else {
          // update existing master rows where order matches
          dataLembarKerja = dataLembarKerja.map(row => {
            const match = IW39.find(i => safeLower(i.Order) === safeLower(row.Order));
            if (match) {
              return {
                ...row,
}
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
        Data1 = {};
}
        normalized.forEach(r => {
          if (r.Order) Data1[safeStr(r.Order)] = safeStr(r.Section || r.Section || "");
        });
      } else if (selectedType === "Data2") {
        Data2 = {};
}
        normalized.forEach(r => {
          if (r.MAT) Data2[safeStr(r.MAT)] = safeStr(r.CPH || "");
        });
      } else if (selectedType === "SUM57") {
        SUM57 = {};
}
        normalized.forEach(r => {
          if (r.Order) SUM57[safeStr(r.Order)] = { StatusPart: safeStr(r.StatusPart), Aging: safeStr(r.Aging) };
        });
      } else if (selectedType === "Planning") {
        Planning = {};
}
        normalized.forEach(r => {
          if (r.Order) Planning[safeStr(r.Order)] = { Planning: r.Planning || "", StatusAMT: r.StatusAMT || "" };
        });
      } else {
        // Unknown type: skip but warn (no console error)
        console.log(`Uploaded file type not specifically handled: ${selectedType}`);
      }

      // success UI
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
      fileInput.value = "";
    }
  });

  // ---------- Build / lookup / compute ----------
  function buildDataLembarKerja() {
    // ensure Order strings
    dataLembarKerja = dataLembarKerja.map(r => ({ ...r, Order: safeStr(r.Order) }));

    dataLembarKerja = dataLembarKerja.map(row => {
      // find IW39 row
      const iw = IW39.find(i => safeLower(i.Order) === safeLower(row.Order)) || {};

      row.Room = iw.Room || row.Room || "";
      row.OrderType = iw.OrderType || row.OrderType || "";
      row.Description = iw.Description || row.Description || "";
      row.CreatedOn = iw.CreatedOn !== undefined ? iw.CreatedOn : row.CreatedOn || "";
      row.UserStatus = iw.UserStatus || row.UserStatus || "";
      row.MAT = iw.MAT || row.MAT || "";

      // CPH logic
      if (safeLower((row.Description || "").substring(0,2)) === "jr") {
        row.CPH = "JR";
}
      } else {
        row.CPH = Data2[row.MAT] || row.CPH || "";
      }

      // Section from Data1
      row.Section = Data1[row.Order] || row.Section || "";

      // SUM57
      if (SUM57[row.Order]) {
        row.StatusPart = SUM57[row.Order].StatusPart || "";
}
        row.Aging = SUM57[row.Order].Aging || "";
      } else {
        row.StatusPart = row.StatusPart || "";
        row.Aging = row.Aging || "";
      }

      // Cost calc
      if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined && isFinite(iw.TotalPlan) && isFinite(iw.TotalActual)) {
        const costCalc = (Number(iw.TotalPlan) - Number(iw.TotalActual)) / 16500;
}
        row.Cost = costCalc < 0 ? "-" : Number(costCalc);
      } else {
        row.Cost = row.Cost || "-";
      }

      // Include
      if (safeLower(row.Reman) === "reman") {
        row.Include = (typeof row.Cost === "number") ? Number(row.Cost * 0.25) : "-";
}
      } else {
        row.Include = row.Cost;
      }

      // Exclude
      if (safeLower(row.OrderType) === "pm38") {
        row.Exclude = "-";
}
      } else {
        row.Exclude = row.Include;
      }

      // Planning
      if (Planning[row.Order]) {
        row.Planning = Planning[row.Order].Planning || "";
}
        row.StatusAMT = Planning[row.Order].StatusAMT || "";
      } else {
        row.Planning = row.Planning || "";
        row.StatusAMT = row.StatusAMT || "";
      }

      return row;
    });
  }

  // ---------- Render table ----------
  function renderTable(data) {
    const dt = Array.isArray(data) ? data : dataLembarKerja;
    const ordersLower = dt.map(d => safeLower(d.Order));
    const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

    outputTableBody.innerHTML = "";

    if (!dt.length) {
      outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#666;">Tidak ada data.</td></tr>`;
}
      return;
    }

    dt.forEach(row => {
      const tr = document.createElement("tr");
      if (duplicates.includes(safeLower(row.Order))) {
        tr.classList.add("duplicate");
}
      }

      function mkCell(val) { const td = document.createElement("td"); td.textContent = val ?? ""; return td; }

      tr.appendChild(mkCell(row.Room));
      tr.appendChild(mkCell(row.OrderType));
      tr.appendChild(mkCell(row.Order));
      tr.appendChild(mkCell(row.Description));
      // CreatedOn formatted
      tr.appendChild(mkCell(formatToDDMMMYYYY(row.CreatedOn)));
      tr.appendChild(mkCell(row.UserStatus));
      tr.appendChild(mkCell(row.MAT));
      tr.appendChild(mkCell(row.CPH));
      tr.appendChild(mkCell(row.Section));
      tr.appendChild(mkCell(row.StatusPart));
      tr.appendChild(mkCell(row.Aging));

      // Month editable
      const tdMonth = document.createElement("td");
      tdMonth.classList.add("editable");
      tdMonth.textContent = row.Month || "";
      tdMonth.title = "Klik untuk edit Month";
      tdMonth.addEventListener("click", () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost (right align)
      const tdCost = document.createElement("td");
      tdCost.classList.add("cost");
      tdCost.textContent = (typeof row.Cost === "number") ? Number(row.Cost).toFixed(1) : row.Cost;
      tdCost.style.textAlign = "right";
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
      tdInclude.style.textAlign = "right";
      tr.appendChild(tdInclude);

      // Exclude
      const tdExclude = document.createElement("td");
      tdExclude.classList.add("exclude");
      tdExclude.textContent = (typeof row.Exclude === "number") ? Number(row.Exclude).toFixed(1) : row.Exclude;
      tdExclude.style.textAlign = "right";
      tr.appendChild(tdExclude);

      // Planning formatted
      tr.appendChild(mkCell(formatToDDMMMYYYY(row.Planning)));
      // StatusAMT
      tr.appendChild(mkCell(row.StatusAMT));

      // Action cell: Edit & Delete (Edit will toggle inline cost editing as well)
      const tdAction = document.createElement("td");

      const btnEdit = document.createElement("button");
      btnEdit.textContent = "Edit";
      btnEdit.classList.add("btn-action","btn-edit");
      btnEdit.addEventListener("click", () => {
        // open inline editors for month, reman and cost
        openInlineEditorsForRow(row);
      });
      tdAction.appendChild(btnEdit);

      const btnDelete = document.createElement("button");
      btnDelete.textContent = "Delete";
      btnDelete.classList.add("btn-action","btn-delete");
      btnDelete.addEventListener("click", () => {
        if (confirm(`Hapus order ${row.Order}?`)) {
          dataLembarKerja = dataLembarKerja.filter(d => safeLower(d.Order) !== safeLower(row.Order));
}
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

  // ---------- Inline editing functionality ----------
  function editMonth(td, row) {
    const sel = document.createElement("select");
    monthAbbr.forEach(m => {
      const opt = document.createElement("option");
      opt.value = m; opt.textContent = m;
      if (row.Month === m) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener("change", () => {
      row.Month = sel.value;
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    sel.addEventListener("blur", () => renderTable(dataLembarKerja));
    td.textContent = ""; td.appendChild(sel); sel.focus();
  }

  function editReman(td, row) {
    const inp = document.createElement("input");
    inp.type = "text"; inp.value = row.Reman || "";
    inp.addEventListener("keydown", e => {
      if (e.key === "Enter") {
        row.Reman = inp.value.trim();
}
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === "Escape") {
        renderTable(dataLembarKerja);
}
      }
    });
    inp.addEventListener("blur", () => {
      row.Reman = inp.value.trim();
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    td.textContent = ""; td.appendChild(inp); inp.focus();
  }

  // For editing Cost inline as well (triggered by Edit button)
  function editCost(td, row) {
    const inp = document.createElement("input");
    inp.type = "number";
    inp.step = "0.1";
    inp.value = (typeof row.Cost === "number") ? Number(row.Cost).toFixed(1) : (row.Cost === "-" ? "" : row.Cost);
    inp.addEventListener("keydown", e => {
      if (e.key === "Enter") {
        const v = inp.value.trim();
}
        row.Cost = v === "" ? "-" : Number(v);
        // recalc include/exclude
        if (safeLower(row.Reman) === "reman" && typeof row.Cost === "number") row.Include = Number(row.Cost * 0.25);
        else row.Include = row.Cost;
        row.Exclude = safeLower(row.OrderType) === "pm38" ? "-" : row.Include;
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === "Escape") {
        renderTable(dataLembarKerja);
}
      }
    });
    inp.addEventListener("blur", () => {
      const v = inp.value.trim();
      row.Cost = v === "" ? "-" : Number(v);
      if (safeLower(row.Reman) === "reman" && typeof row.Cost === "number") row.Include = Number(row.Cost * 0.25);
      else row.Include = row.Cost;
      row.Exclude = safeLower(row.OrderType) === "pm38" ? "-" : row.Include;
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    td.textContent = ""; td.appendChild(inp); inp.focus();
  }

  // Called by Edit button => open Month, Reman, Cost inline for that row
  function openInlineEditorsForRow(row) {
    // find table row element by order
    const trs = Array.from(outputTableBody.querySelectorAll("tr"));
    for (const tr of trs) {
      const orderCell = tr.children[2];
      if (orderCell && safeLower(orderCell.textContent) === safeLower(row.Order)) {
        // month index: 11 (0-based as built)
        const monthTd = tr.children[11];
}
        const costTd = tr.children[12];
        const remanTd = tr.children[13];
        editMonth(monthTd, row);
        editCost(costTd, row);
        editReman(remanTd, row);
        break;
      }
    }
  }

  // ---------- Add Order ----------
  addOrderBtn.addEventListener("click", () => {
    const raw = addOrderInput.value.trim();
    if (!raw) {
      addOrderStatus.style.color = "red";
}
      addOrderStatus.textContent = "Masukkan minimal 1 order.";
      return;
    }
    const orders = raw.split(/[\s,]+/).map(s => s.trim()).filter(Boolean);

    let added = 0, skipped = [], invalid = [];
    for (const o of orders) {
      if (!isValidOrder(o)) { invalid.push(o); continue; }
      if (dataLembarKerja.some(d => safeLower(d.Order) === safeLower(o))) { skipped.push(o); continue; }
      // push base row then fill via buildDataLembarKerja
      dataLembarKerja.push({
        Order: safeStr(o),
        Room: "", OrderType: "", Description: "", CreatedOn: "", UserStatus: "",
        MAT: "", CPH: "", Section: "", StatusPart: "", Aging: "", Month: "",
        Cost: "-", Reman: "", Include: "-", Exclude: "-", Planning: "", StatusAMT: ""
      });
      added++;
    }

    buildDataLembarKerja(); // perform lookups & calculations
    renderTable(dataLembarKerja);
    saveDataToLocalStorage();

    let msg = `${added} order ditambahkan.`;
    if (skipped.length) msg += ` Sudah ada: ${skipped.join(", ")}.`;
    if (invalid.length) msg += ` Invalid: ${invalid.join(", ")}.`;
    addOrderStatus.style.color = added ? "green" : "red";
    addOrderStatus.textContent = msg;
    addOrderInput.value = "";
  });

  // ---------- Filters ----------
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
    filterRoom.value = ""; filterOrder.value = ""; filterCPH.value = ""; filterMAT.value = ""; filterSection.value = "";
    renderTable(dataLembarKerja);
  });

  // ---------- Save / Load ----------
  function saveDataToLocalStorage() {
    try {
      localStorage.setItem("lembarKerjaData", JSON.stringify(dataLembarKerja));
    } catch (e) {
      console.warn("Gagal menyimpan:", e);
    }
  }

  saveBtn.addEventListener("click", () => {
    saveDataToLocalStorage();
    alert("Data disimpan di browser.");
  });

  loadBtn.addEventListener("click", () => {
    const saved = localStorage.getItem("lembarKerjaData");
    if (!saved) { alert("Tidak ada data tersimpan."); return; }
    try {
      dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) }));
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      alert("Data dimuat.");
    } catch (e) {
      alert("Gagal muat data: " + e.message);
    }
  });

  // ---------- Validation ----------
  function isValidOrder(order) {
    return !/[.,]/.test(order);
  }

  // ---------- Sidebar menu ----------
  (function initMenu() {
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
  })();

  // ---------- Init load saved ----------
  (function init() {
    const saved = localStorage.getItem("lembarKerjaData");
    if (saved) {
      try {
        dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) }));
}
      } catch (e) { dataLembarKerja = []; }
    }
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  })();

});
