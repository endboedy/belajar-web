/* script.js - FINAL (GitHub load + merge + admin save)
   - Load Excel from GitHub raw URLs (or upload local via input)
   - Merge logic (Section by Room from Data1, Status Part/Aging from SUM57 by Order,
     CPH: if Description starts "JR" => "External Job", else lookup Data2 by Order/MAT)
   - Cost = (TotalPlan - TotalActual) / 16500  (columns auto-detected)
   - Format Created On & Planning -> dd-MMM-yyyy
   - Edit row inline (Month dropdown, Reman free text), Delete row
   - Add Orders, Filter, Reset, Refresh
   - Export merged to Excel; backup to/from JSON
   - Admin Save -> commit data.json to GitHub (PUT) (requires PAT)
*/

// ----------------- Config / Globals -----------------
const GITHUB_RAW_BASE = "https://raw.githubusercontent.com"; // used when fetching public raw files
// session admin info stored in sessionStorage keys:
const SESSION_ADMIN_KEY = "ndarboe_admin_info_v1";
const UI_LS_KEY = "ndarboe_ui_state_v1"; // small user edits persisted

let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];

let mergedData = [];

// ----------------- Utility helpers -----------------
function safeNum(v) {
  if (v === undefined || v === null || v === "") return NaN;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^0-9\.\-]/g, "");
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}
function getVal(row, candidates) {
  if (!row) return undefined;
  for (const c of candidates) {
    if (c in row && row[c] !== undefined && row[c] !== null && row[c] !== "") return row[c];
    // case-insensitive fallback
    for (const k of Object.keys(row)) {
      if (k.toLowerCase() === c.toLowerCase() && row[k] !== "" && row[k] !== null && row[k] !== undefined) {
        return row[k];
      }
    }
  }
  return undefined;
}
function excelDateToJS(serial) {
  if (typeof serial === "number") {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional = serial - Math.floor(serial);
    if (fractional > 0) {
      const seconds = Math.round(fractional * 86400);
      date_info.setSeconds(date_info.getSeconds() + seconds);
    }
    return date_info;
  }
  const d = new Date(serial);
  if (!isNaN(d)) return d;
  return null;
}
function formatDateDDMMMYYYY(input) {
  if (input === undefined || input === null || input === "") return "";
  let d = null;
  if (typeof input === "number") d = excelDateToJS(input);
  else {
    d = new Date(input);
    if (isNaN(d)) {
      const alt = new Date(String(input).replace(/\//g, "-"));
      d = isNaN(alt) ? null : alt;
    }
  }
  if (!d || isNaN(d)) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const day = String(d.getDate()).padStart(2,"0");
  const mon = months[d.getMonth()];
  const year = d.getFullYear();
  return ${day}-${mon}-${year};
}
function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
function textToBase64Uint8(str) {
  // convert binary string to Uint8Array
  const buf = new ArrayBuffer(str.length);
  const view = new Uint8Array(buf);
  for (let i=0;i<str.length;i++) view[i] = str.charCodeAt(i) & 0xFF;
  return view;
}
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

// ----------------- Parse XLSX (from File input) -----------------
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const wb = XLSX.read(data, { type: "binary" });
        const first = wb.SheetNames[0];
        const ws = wb.Sheets[first];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

// ----------------- Fetch Excel from GitHub raw (public) -----------------
async function fetchExcelFromGitHubRaw(owner, repo, branch, pathInRepo) {
  // raw url: https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{path}
  const rawUrl = ${GITHUB_RAW_BASE}/${owner}/${repo}/${branch}/${pathInRepo};
  // fetch as arrayBuffer then parse with XLSX
  const resp = await fetch(rawUrl);
  if (!resp.ok) throw new Error(Failed to fetch ${rawUrl}: ${resp.statusText});
  const buffer = await resp.arrayBuffer();
  // parse buffer via XLSX
  const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
  const first = wb.SheetNames[0];
  const ws = wb.Sheets[first];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return json;
}

// ----------------- Merge Logic (updated rules) -----------------
function mergeData() {
  mergedData = [];
  if (!iw39Data || iw39Data.length === 0) {
    alert("IW39 belum di-upload / dimuat. Upload atau load IW39 lalu klik Refresh.");
    return;
  }

  // Build lookup maps
  const sum57ByOrder = new Map();
  sum57Data.forEach(r => {
    const k = String(getVal(r, ["Order","Order No","Order_No","Key","No"]) || "").trim();
    if (k) sum57ByOrder.set(k, r);
  });

  const data1ByRoom = new Map();
  data1Data.forEach(r => {
    const k = String(getVal(r, ["Room","ROOM","Lokasi","Location"]) || "").trim();
    if (k) data1ByRoom.set(k, r);
  });

  const data2ByOrder = new Map();
  const data2ByMat = new Map();
  data2Data.forEach(r => {
    const ord = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
    const mat = String(getVal(r, ["MAT","Mat","Material","Key"]) || "").trim();
    if (ord) data2ByOrder.set(ord, r);
    if (mat) data2ByMat.set(mat, r);
  });

  const planningByOrder = new Map();
  planningData.forEach(r => {
    const ord = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
    if (ord) planningByOrder.set(ord, r);
  });

  // For cost calculation: we look for columns that match "Total sum (plan)" and "Total sum (actual)" (or variants)
  const planCandidates = ["Total sum (plan)","TotalPlan","Total Plan","Total_Sum_Plan","Total sum plan","Total Plan (sum)"];
  const actualCandidates = ["Total sum (actual)","TotalActual","Total Actual","Total_Sum_Actual","Total actual"];

  iw39Data.forEach(row => {
    const order = String(getVal(row, ["Order","Order No","Order_No","ORD","Key"]) || "").trim();
    const room = getVal(row, ["Room","ROOM","Location","Lokasi"]) || "";
    const orderType = getVal(row, ["Order Type","OrderType","Type","Order_Type"]) || "";
    const description = getVal(row, ["Description","Desc","Keterangan"]) || "";
    const createdRaw = getVal(row, ["Created On","CreatedOn","Created_Date","Tanggal","Create On","Created"]) || "";
    const userStatus = getVal(row, ["User Status","UserStatus","Status User","Status"]) || "";
    const mat = String(getVal(row, ["MAT","Mat","Material"]) || "").trim();

    // find plan & actual (try candidate column names in iw39 row)
    let totalPlan = NaN, totalActual = NaN;
    for (const c of planCandidates) {
      const v = getVal(row, [c]);
      if (v !== undefined && v !== "") { totalPlan = safeNum(v); break; }
    }
    for (const c of actualCandidates) {
      const v = getVal(row, [c]);
      if (v !== undefined && v !== "") { totalActual = safeNum(v); break; }
    }

    // cost formula
    let cost = "-";
    if (!isNaN(totalPlan) && !isNaN(totalActual)) {
      const calc = (totalPlan - totalActual) / 16500;
      if (!isNaN(calc) && calc >= 0) cost = (Math.round((calc + Number.EPSILON) * 100) / 100).toFixed(2);
      else cost = "-";
    }

    // CPH rule:
    // if first 2 chars of Description are "JR" (case-insensitive) => "External Job"
    // else try lookup Data2: first by Order, then by MAT
    let cph = "";
    const descPrefix = (String(description || "").trim().substr(0,2) || "").toUpperCase();
    if (descPrefix === "JR") {
      cph = "External Job";
    } else {
      // try by Order in Data2
      if (order && data2ByOrder.has(order)) {
        cph = getVal(data2ByOrder.get(order), ["CPH","Cph","cph","CPH Code","Code","Result"]) || "";
      } else if (mat && data2ByMat.has(mat)) {
        cph = getVal(data2ByMat.get(mat), ["CPH","Cph","cph","CPH Code","Code","Result"]) || "";
      } else {
        cph = "";
      }
    }

    // Section from Data1 by Room
    let section = "-";
    if (room && data1ByRoom.has(String(room).trim())) {
      section = getVal(data1ByRoom.get(String(room).trim()), ["Section","Section Name","SectionName","SECTION"]) || "-";
    } else {
      const f = data1Data.find(r => {
        const k = String(getVal(r, ["Room","ROOM","Location","Lokasi"]) || "").trim();
        return k && String(k) === String(room).trim();
      });
      if (f) section = getVal(f, ["Section","Section Name","SectionName","SECTION"]) || "-";
    }

    // Status Part & Aging from SUM57 by Order
    let statusPart = "-";
    let aging = "-";
    if (order && sum57ByOrder.has(order)) {
      const s = sum57ByOrder.get(order);
      // user said status part is "Part Complete" column in SUM57
      statusPart = getVal(s, ["Part Complete","PartComplete","Part_Complete","Status Part","Status"]) || "-";
      aging = getVal(s, ["Aging","Age","Aging Days","Age Days"]) || "-";
    }

    // Planning & Status AMT from planningByOrder
    let planning = "";
    let statusAMT = "";
    if (order && planningByOrder.has(order)) {
      const p = planningByOrder.get(order);
      planning = getVal(p, ["Event Start","Planning","Start","EventStart","Start Date","Start_Date"]) || "";
      statusAMT = getVal(p, ["Status AMT","StatusAMT","AMT Status","Status"]) || "";
    }

    // Month & Reman defaults
    const month = getVal(row, ["Month"]) || "";
    const reman = getVal(row, ["Reman"]) || "";

    // Include / Exclude
    let include = "-";
    if (cost === "-" || cost === undefined) include = "-";
    else {
      const costNum = Number(cost);
      include = (String(reman).toLowerCase() === "reman") ? (Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : costNum.toFixed(2);
    }
    let exclude = (String(orderType).trim().toUpperCase() === "PM38") ? "-" : include;

    const mergedRow = {
      Room: room || "",
      "Order Type": orderType || "",
      Order: order || "",
      Description: description || "",
      "Created On": createdRaw || "",
      "User Status": userStatus || "",
      MAT: mat || "",
      CPH: cph || "",
      Section: section || "-",
      "Status Part": statusPart || "-",
      Aging: aging || "-",
      Month: month || "",
      Cost: cost,
      Reman: reman || "",
      Include: include,
      Exclude: exclude,
      Planning: planning || "",
      "Status AMT": statusAMT || "",
      _IW39_totalPlan: isNaN(totalPlan) ? "" : totalPlan,
      _IW39_totalActual: isNaN(totalActual) ? "" : totalActual
    };

    mergedData.push(mergedRow);
  });

  // reapply UI small edits (Month/Reman)
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) {
      const ui = JSON.parse(raw);
      if (ui && Array.isArray(ui.userEdits)) {
        const map = new Map(ui.userEdits.map(e => [e.Order, e]));
        mergedData = mergedData.map(r => {
          const s = map.get(r.Order);
          if (s) {
            if (s.Month !== undefined) r.Month = s.Month;
            if (s.Reman !== undefined) r.Reman = s.Reman;
            if (r.Cost !== "-" && !isNaN(Number(r.Cost))) {
              const costNum = Number(r.Cost);
              r.Include = (String(r.Reman).toLowerCase() === "reman") ? (Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : costNum.toFixed(2);
            } else r.Include = "-";
            r.Exclude = (String(r["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : r.Include;
          }
          return r;
        });
      }
    }
  } catch (e) {
    console.warn("Failed reapply UI edits:", e);
  }
}

// ----------------- Render Table (Menu 2) -----------------
function renderTable(dataToRender) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) {
    console.error("#output-table tbody not found");
    return;
  }
  tbody.innerHTML = "";

  const rows = Array.isArray(dataToRender) ? dataToRender : mergedData;
  if (!rows || rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 19;
    td.style.textAlign = "center";
    td.textContent = "Tidak ada data";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  rows.forEach((row, idx) => {
    const tr = document.createElement("tr");

    function addCell(val) {
      const td = document.createElement("td");
      td.textContent = (val === undefined || val === null) ? "" : val;
      return td;
    }

    const createdDisp = row["Created On"] ? formatDateDDMMMYYYY(row["Created On"]) : "";
    const planningDisp = row.Planning ? formatDateDDMMMYYYY(row.Planning) : "";

    tr.appendChild(addCell(row.Room));
    tr.appendChild(addCell(row["Order Type"]));
    tr.appendChild(addCell(row.Order));
    tr.appendChild(addCell(row.Description));
    tr.appendChild(addCell(createdDisp));
    tr.appendChild(addCell(row["User Status"]));
    tr.appendChild(addCell(row.MAT));
    tr.appendChild(addCell(row.CPH));
    tr.appendChild(addCell(row.Section));
    tr.appendChild(addCell(row["Status Part"]));
    tr.appendChild(addCell(row.Aging));

    // month editable cell
    const tdMonth = document.createElement("td");
    tdMonth.textContent = row.Month || "";
    tdMonth.dataset.col = "Month";
    tr.appendChild(tdMonth);

    tr.appendChild(addCell(row.Cost));
    // reman editable
    const tdReman = document.createElement("td");
    tdReman.textContent = row.Reman || "";
    tdReman.dataset.col = "Reman";
    tr.appendChild(tdReman);

    tr.appendChild(addCell(row.Include));
    tr.appendChild(addCell(row.Exclude));
    tr.appendChild(addCell(planningDisp));
    tr.appendChild(addCell(row["Status AMT"]));

    // actions
    const tdAction = document.createElement("td");
    const editBtn = document.createElement("button");
    editBtn.textContent = "Edit";
    editBtn.className = "action-btn edit-btn";
    editBtn.addEventListener("click", () => startEditRow(idx, tr));
    tdAction.appendChild(editBtn);

    const delBtn = document.createElement("button");
    delBtn.textContent = "Delete";
    delBtn.className = "action-btn delete-btn";
    delBtn.addEventListener("click", () => {
      if (confirm("Hapus baris order " + (row.Order || "") + " ?")) {
        const gi = mergedData.findIndex(r => r.Order === row.Order);
        if (gi !== -1) mergedData.splice(gi, 1);
        removeUserEdit(row.Order);
        renderTable(mergedData);
      }
    });
    tdAction.appendChild(delBtn);

    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

// ----------------- Edit row inline (Month, Reman) -----------------
function startEditRow(index, trElement) {
  const row = mergedData[index];
  if (!row) return;
  const monthTd = trElement.querySelector('td[data-col="Month"]');
  const remanTd = trElement.querySelector('td[data-col="Reman"]');
  if (!monthTd || !remanTd) return;

  const months = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const sel = document.createElement("select");
  months.forEach(m => {
    const o = document.createElement("option");
    o.value = m;
    o.text = m || "--";
    if (m === row.Month) o.selected = true;
    sel.appendChild(o);
  });
  monthTd.innerHTML = "";
  monthTd.appendChild(sel);

  const remInput = document.createElement("input");
  remInput.type = "text";
  remInput.value = row.Reman || "";
  remInput.style.width = "100%";
  remanTd.innerHTML = "";
  remanTd.appendChild(remInput);

  const actionTd = trElement.querySelector("td:last-child");
  actionTd.innerHTML = "";

  const saveBtn = document.createElement("button");
  saveBtn.textContent = "Save";
  saveBtn.className = "action-btn save-btn";
  saveBtn.addEventListener("click", () => {
    row.Month = sel.value;
    row.Reman = remInput.value;
    // recalc include/exclude
    if (row.Cost !== "-" && !isNaN(Number(row.Cost))) {
      const costNum = Number(row.Cost);
      row.Include = (String(row.Reman).toLowerCase() === "reman") ? (Math.round((costNum * 0.25 + Number.EPSILON) * 100) / 100).toFixed(2) : costNum.toFixed(2);
    } else row.Include = "-";
    row.Exclude = (String(row["Order Type"]).trim().toUpperCase() === "PM38") ? "-" : row.Include;

    saveUserEdit(row.Order, { Order: row.Order, Month: row.Month, Reman: row.Reman });
    renderTable(mergedData);
  });
  actionTd.appendChild(saveBtn);

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "action-btn cancel-btn";
  cancelBtn.addEventListener("click", () => renderTable(mergedData));
  actionTd.appendChild(cancelBtn);
}

// ---------------- small UI edits persistence -----------------
function saveUserEdit(orderKey, editObj) {
  let ui = { userEdits: [] };
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (raw) ui = JSON.parse(raw);
  } catch (e) { ui = { userEdits: [] }; }
  ui.userEdits = ui.userEdits.filter(r => r.Order !== orderKey);
  ui.userEdits.push(editObj);
  try {
    localStorage.setItem(UI_LS_KEY, JSON.stringify(ui));
  } catch (e) {
    console.warn("Could not save UI edits:", e);
  }
}
function removeUserEdit(orderKey) {
  try {
    const raw = localStorage.getItem(UI_LS_KEY);
    if (!raw) return;
    const ui = JSON.parse(raw);
    ui.userEdits = ui.userEdits.filter(r => r.Order !== orderKey);
    localStorage.setItem(UI_LS_KEY, JSON.stringify(ui));
  } catch (e) {}
}

// ---------------- Filter / Reset -----------------
function filterData() {
  let filtered = mergedData.slice();
  const room = (document.getElementById("filter-room").value || "").trim().toLowerCase();
  const order = (document.getElementById("filter-order").value || "").trim().toLowerCase();
  const cph = (document.getElementById("filter-cph").value || "").trim().toLowerCase();
  const mat = (document.getElementById("filter-mat").value || "").trim().toLowerCase();
  const section = (document.getElementById("filter-section").value || "").trim().toLowerCase();

  if (room) filtered = filtered.filter(d => (d.Room || "").toString().toLowerCase().includes(room));
  if (order) filtered = filtered.filter(d => (d.Order || "").toString().toLowerCase().includes(order));
  if (cph) filtered = filtered.filter(d => (d.CPH || "").toString().toLowerCase().includes(cph));
  if (mat) filtered = filtered.filter(d => (d.MAT || "").toString().toLowerCase().includes(mat));
  if (section) filtered = filtered.filter(d => (d.Section || "").toString().toLowerCase().includes(section));

  renderTable(filtered);
}
function resetFilter() {
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  renderTable(mergedData);
}

// ---------------- Add Orders manual -----------------
function addOrders() {
  const input = (document.getElementById("add-order-input").value || "").trim();
  const statusEl = document.getElementById("add-order-status");
  if (!input) {
    if (statusEl) { statusEl.textContent = "Masukkan order dulu ya bro!"; statusEl.style.color = "red"; }
    return;
  }
  const orders = input.split(/[\s,]+/).filter(o => o.length > 0);
  let added = 0;
  orders.forEach(o => {
    if (mergedData.find(r => r.Order === o)) return;
    const iw = iw39Data.find(r => {
      const v = String(getVal(r, ["Order","Order No","Order_No","Key"]) || "").trim();
      return v === o;
    });
    if (iw) {
      mergedData.push({
        Room: getVal(iw, ["Room","ROOM","Location"]) || "",
        "Order Type": getVal(iw, ["Order Type","OrderType"]) || "",
        Order: (getVal(iw, ["Order","Order No","Key"]) || "").toString().trim(),
        Description: getVal(iw, ["Description","Desc"]) || "",
        "Created On": getVal(iw, ["Created On","CreatedOn","Tanggal"]) || "",
        "User Status": getVal(iw, ["User Status","UserStatus"]) || "",
        MAT: (getVal(iw, ["MAT","Mat","Material"]) || "").toString().trim(),
        CPH: "",
        Section: "",
        "Status Part": "",
        Aging: "",
        Month: "",
        Cost: "-",
        Reman: "",
        Include: "-",
        Exclude: "-",
        Planning: "",
        "Status AMT": ""
      });
    } else {
      mergedData.push({
        Room: "",
        "Order Type": "",
        Order: o,
        Description: "",
        "Created On": "",
        "User Status": "",
        MAT: "",
        CPH: "",
        Section: "",
        "Status Part": "",
        Aging: "",
        Month: "",
        Cost: "-",
        Reman: "",
        Include: "-",
        Exclude: "-",
        Planning: "",
        "Status AMT": ""
      });
    }
    added++;
  });
  if (statusEl) { statusEl.textContent = ${added} order berhasil ditambahkan.; statusEl.style.color = "green"; }
  document.getElementById("add-order-input").value = "";
  renderTable(mergedData);
}

// ---------------- Export merged to XLSX -----------------
function exportMergedToExcel() {
  if (!mergedData || mergedData.length === 0) { alert("Tidak ada data untuk diexport."); return; }
  const rows = mergedData.map(r => {
    const c = Object.assign({}, r);
    delete c._IW39_totalPlan;
    delete c._IW39_totalActual;
    return c;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Merged");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i=0;i<s.length;i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  downloadFile("Lembar_Kerja_merged.xlsx", s2ab(wbout), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// ---------------- JSON backup -----------------
function downloadJSONBackup() {
  const payload = { iw39Data, sum57Data, planningData, data1Data, data2Data, budgetData, mergedData, timestamp: new Date().toISOString() };
  downloadFile("ndarboe_backup.json", JSON.stringify(payload, null, 2), "application/json");
}
function loadJSONBackupFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const obj = JSON.parse(e.target.result);
      iw39Data = obj.iw39Data || [];
      sum57Data = obj.sum57Data || [];
      planningData = obj.planningData || [];
      data1Data = obj.data1Data || [];
      data2Data = obj.data2Data || [];
      budgetData = obj.budgetData || [];
      mergedData = obj.mergedData || [];
      renderTable(mergedData);
      alert("Backup JSON dimuat.");
    } catch (err) {
      alert("Gagal memuat JSON: " + err.message);
    }
  };
  reader.readAsText(file);
}

// ---------------- GitHub save (admin) -----------------
// Admin info stored in sessionStorage: { owner, repo, branch, token }
function setAdminSession(info) {
  sessionStorage.setItem(SESSION_ADMIN_KEY, JSON.stringify(info));
}
function getAdminSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_ADMIN_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
}
async function githubGetFileSha(owner, repo, path, branch, token) {
  const url = https://api.github.com/repos/${owner}/${repo}/contents/${encodeURIComponent(path)}?ref=${encodeURIComponent(branch)};
  const res = await fetch(url, {
    headers: { Authorization: token ${token}, Accept: 'application/vnd.github.v3+json' }
  });
  if (!res.ok) throw new Error(GET SHA failed: ${res.status} ${res.statusText});
  const js = await res.json();
  return js.sha;
}
async function githubPutFile(owner, repo, path, branch, token, contentBase64, message, sha) {
  const url = https://api.github.com/repos/${owner}/${repo}/contents/${encodeURIComponent(path)};
  const body = { message, content: contentBase64, branch };
  if (sha) body.sha = sha;
  const res = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: token ${token}, Accept: 'application/vnd.github.v3+json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(PUT failed ${res.status}: ${txt});
  }
  return await res.json();
}
async function saveMergedDataToGitHub(pathInRepo="data/data.json") {
  const admin = getAdminSession();
  if (!admin || !admin.token) { alert("Admin token belum di-set. Klik Admin Login."); return; }
  // prepare payload
  const payload = { mergedData, timestamp: new Date().toISOString() };
  const contentStr = JSON.stringify(payload, null, 2);
  const contentBase64 = btoa(unescape(encodeURIComponent(contentStr)));
  try {
    // try get sha (if exists)
    let sha;
    try { sha = await githubGetFileSha(admin.owner, admin.repo, pathInRepo, admin.branch, admin.token); } catch(e) { sha = undefined; }
    const msg = Update data.json via web by admin at ${new Date().toISOString()};
    const res = await githubPutFile(admin.owner, admin.repo, pathInRepo, admin.branch, admin.token, contentBase64, msg, sha);
    alert("Saved to GitHub: " + res.content.path);
  } catch (err) {
    console.error("GitHub save failed:", err);
    alert("GitHub save failed: " + err.message);
  }
}

// ---------------- Load Excel files either via upload or GitHub raw -----------------
async function loadFromGitHubExcel(owner, repo, branch, pathInRepo, targetArray) {
  try {
    const json = await fetchExcelFromGitHubRaw(owner, repo, branch, pathInRepo);
    targetArray.length = 0;
    targetArray.push(...json);
    return json.length;
  } catch (e) {
    throw e;
  }
}

// ---------------- UI wiring -----------------
function wireUp() {
  // menu switching
  document.querySelectorAll(".menu-item").forEach(it => {
    it.addEventListener("click", () => {
      document.querySelectorAll(".menu-item").forEach(i => i.classList.remove("active"));
      document.querySelectorAll(".content-section").forEach(s => s.classList.remove("active"));
      it.classList.add("active");
      const m = it.dataset.menu;
      const sec = document.getElementById(m);
      if (sec) sec.classList.add("active");
    });
  });

  // Upload local file (existing upload control)
  const uploadBtn = document.getElementById("upload-btn");
  if (uploadBtn) {
    uploadBtn.addEventListener("click", async () => {
      const sel = document.getElementById("file-select").value;
      const f = document.getElementById("file-input").files[0];
      if (!f) { alert("Pilih file terlebih dahulu"); return; }
      document.getElementById("upload-status").textContent = Parsing ${f.name} ...;
      try {
        const json = await parseFile(f);
        switch (sel) {
          case "IW39": iw39Data = json; break;
          case "SUM57": sum57Data = json; break;
          case "Planning": planningData = json; break;
          case "Data1": data1Data = json; break;
          case "Data2": data2Data = json; break;
          case "Budget": budgetData = json; break;
        }
        document.getElementById("upload-status").textContent = ${sel} loaded (${json.length} rows);
        document.getElementById("file-input").value = "";
      } catch (err) {
        console.error(err);
        alert("Gagal parsing file: " + err.message);
      }
    });
  }

  // GitHub load buttons (you can make a small UI to call these)
  // For convenience, add a simple handler for "Load from GitHub" if you want:
  const ghLoadBtn = document.getElementById("gh-load-btn");
  if (ghLoadBtn) {
    ghLoadBtn.addEventListener("click", async () => {
      const owner = prompt("GitHub owner (username/org):");
      const repo = prompt("Repo name:");
      const branch = prompt("Branch (default: main):", "main");
      if (!owner || !repo) return alert("Owner & repo required");
      try {
        document.getElementById("upload-status").textContent = "Loading IW39 from GitHub...";
        await loadFromGitHubExcel(owner, repo, branch, "excel/IW39.xlsx", iw39Data);
        document.getElementById("upload-status").textContent = "Loading Data1...";
        await loadFromGitHubExcel(owner, repo, branch, "excel/Data1.xlsx", data1Data);
        document.getElementById("upload-status").textContent = "Loading SUM57...";
        await loadFromGitHubExcel(owner, repo, branch, "excel/SUM57.xlsx", sum57Data);
        try { await loadFromGitHubExcel(owner, repo, branch, "excel/Planning.xlsx", planningData); } catch(e){}
        try { await loadFromGitHubExcel(owner, repo, branch, "excel/Data2.xlsx", data2Data); } catch(e){}
        document.getElementById("upload-status").textContent = "All loaded from GitHub.";
        mergeData();
        renderTable(mergedData);
      } catch(e) {
        console.error(e);
        alert("GitHub load error: " + e.message);
      }
    });
  }

  // Clear memory
  const clearBtn = document.getElementById("clear-files-btn");
  if (clearBtn) clearBtn.addEventListener("click", () => {
    if (!confirm("Clear uploaded data in memory?")) return;
    iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[];
    mergedData=[];
    document.getElementById("upload-status").textContent = "Data cleared";
    renderTable([]);
  });

  // Refresh -> merge & render
  const refreshBtn = document.getElementById("refresh-btn");
  if (refreshBtn) refreshBtn.addEventListener("click", () => {
    if (!iw39Data || iw39Data.length === 0) { alert("Upload IW39 dulu sebelum Refresh."); return; }
    mergeData();
    renderTable(mergedData);
    const s = document.getElementById("add-order-status");
    if (s) s.textContent = "";
  });

  // Add orders
  const addBtn = document.getElementById("add-order-btn");
  if (addBtn) addBtn.addEventListener("click", addOrders);

  // Filter / Reset
  const filterBtn = document.getElementById("filter-btn");
  if (filterBtn) filterBtn.addEventListener("click", filterData);
  const resetBtn = document.getElementById("reset-btn");
  if (resetBtn) resetBtn.addEventListener("click", resetFilter);

  // Save (export) - downloads XLSX
  const saveBtn = document.getElementById("save-btn");
  if (saveBtn) saveBtn.addEventListener("click", exportMergedToExcel);

  // Load (backup JSON)
  const loadBtn = document.getElementById("load-btn");
  if (loadBtn) loadBtn.addEventListener("click", () => {
    const inpf = document.createElement("input");
    inpf.type = "file";
    inpf.accept = ".json";
    inpf.addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (f) loadJSONBackupFile(f);
    });
    inpf.click();
  });

  // Admin login (for save to GitHub)
  const adminBtn = document.getElementById("admin-login-btn");
  if (adminBtn) adminBtn.addEventListener("click", () => {
    const owner = prompt("GitHub owner (username/org):");
    if (!owner) return;
    const repo = prompt("Repo name:");
    if (!repo) return;
    const branch = prompt("Branch:", "main") || "main";
    const token = prompt("Personal Access Token (will be kept in session only):");
    if (!token) return;
    setAdminSession({ owner, repo, branch, token });
    alert("Admin session set for this tab. Save to GitHub will be enabled.");
  });

  // Admin save to GitHub (save JSON)
  const ghSaveBtn = document.getElementById("gh-save-btn");
  if (ghSaveBtn) ghSaveBtn.addEventListener("click", async () => {
    // default path: data/data.json
    await saveMergedDataToGitHub("data/data.json");
  });

  // initial render empty
  renderTable([]);
}

// ---------------- Init -----------------
window.addEventListener("DOMContentLoaded", () => {
  wireUp();
});
