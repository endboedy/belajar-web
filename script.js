/****************************************************
 * Ndarboe.net - FULL script.js
 * --------------------------------------------------
 * - Menu 1: Upload File (IW39, SUM57, Planning, Budget, Data1, Data2)
 * - Menu 2: Lembar Kerja (filter, lookup Month/Cost/Reman, render tabel)
 * - Menu 3: LOM (Add Order, tabel, filter, save/load)
 * - Menu 4: Summary CRM
 * - Menu 5: Download Excel (Lembar Kerja + LOM + Summary)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
window.iw39Data     = window.iw39Data || [];
window.sum57Data    = window.sum57Data || [];
window.planningData = window.planningData || [];
window.data1Data    = window.data1Data || [];
window.data2Data    = window.data2Data || [];
window.budgetData   = window.budgetData || [];
window.mergedData   = window.mergedData || [];

const UI_LS_KEY  = "ndarboe_ui_edits_v2";
const LOM_LS_KEY = "lomUserEdits";

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupMenu();
  setupLOM();
  setupLembarKerja();
  renderLembarKerjaTable();
  renderLOMTable();
  updateMonthFilterOptions();
  renderSummaryCRM();
});

/* ===================== MENU SETUP ===================== */
function setupMenu(){
  document.querySelectorAll(".sidebar .menu-item").forEach(item=>{
    item.addEventListener("click", ()=>{
      document.querySelectorAll(".sidebar .menu-item").forEach(i=>i.classList.remove("active"));
      item.classList.add("active");
      const menu = item.dataset.menu;
      document.querySelectorAll(".content-section").forEach(s=>s.classList.remove("active"));
      document.getElementById(menu)?.classList.add("active");
    });
  });
}

/* ===================== LOM (LIST ORDER MONTHLY) ===================== */
let lomData = JSON.parse(localStorage.getItem(LOM_LS_KEY)||"[]");

function setupLOM(){
  const addBtn = document.getElementById("lom-add-btn");
  addBtn.addEventListener("click", ()=>{
    const input = document.getElementById("lom-add-order").value;
    const orders = input.split(/[\s,]+/).filter(v=>v);
    orders.forEach(o=>{
      if(!lomData.some(e=>e.Order===o)){
        lomData.push({
          Order:o,
          Month:"",
          Cost:"",
          Reman:"",
          Status:"",
          Planning:"",
          StatusAMT:""
        });
      }
    });
    localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
    renderLOMTable();
    document.getElementById("lom-add-order").value="";
  });

  document.getElementById("lom-filter-btn").addEventListener("click", renderLOMTable);
  document.getElementById("lom-save-btn").addEventListener("click", ()=>{
    localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
    alert("LOM saved!");
  });
  document.getElementById("lom-load-btn").addEventListener("click", ()=>{
    lomData = JSON.parse(localStorage.getItem(LOM_LS_KEY)||"[]");
    renderLOMTable();
  });
}

function renderLOMTable(){
  const tbody = document.querySelector("#lom-table tbody");
  tbody.innerHTML="";
  const filterVal = document.getElementById("lom-filter-order")?.value.toUpperCase() || "";
  lomData.forEach((row,i)=>{
    if(filterVal && !row.Order.toUpperCase().includes(filterVal)) return;
    const tr = document.createElement("tr");

    tr.innerHTML=`
      <td>${row.Order}</td>
      <td>
        <select data-index="${i}" class="lom-month-select">
          <option value="">--</option>
          ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].map(m=>`<option value="${m}" ${row.Month===m?"selected":""}>${m}</option>`).join("")}
        </select>
      </td>
      <td><input type="number" data-index="${i}" class="lom-cost-input" value="${row.Cost}" /></td>
      <td><input type="text" data-index="${i}" class="lom-reman-input" value="${row.Reman}" /></td>
      <td>${row.Status}</td>
      <td>${row.Planning}</td>
      <td>${row.StatusAMT}</td>
      <td>
        <button data-index="${i}" class="delete-btn action-btn">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Event listeners for editable fields
  tbody.querySelectorAll(".lom-month-select").forEach(sel=>{
    sel.addEventListener("change", e=>{
      const idx = e.target.dataset.index;
      lomData[idx].Month = e.target.value;
      localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
    });
  });
  tbody.querySelectorAll(".lom-cost-input").forEach(inp=>{
    inp.addEventListener("input", e=>{
      const idx = e.target.dataset.index;
      lomData[idx].Cost = e.target.value;
      localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
    });
  });
  tbody.querySelectorAll(".lom-reman-input").forEach(inp=>{
    inp.addEventListener("input", e=>{
      const idx = e.target.dataset.index;
      lomData[idx].Reman = e.target.value;
      localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
    });
  });
  tbody.querySelectorAll(".delete-btn").forEach(btn=>{
    btn.addEventListener("click", e=>{
      const idx = e.target.dataset.index;
      lomData.splice(idx,1);
      localStorage.setItem(LOM_LS_KEY, JSON.stringify(lomData));
      renderLOMTable();
    });
  });
}

/* ===================== LEMBAR KERJA ===================== */
let mergedLembarData = [];

function setupLembarKerja(){
  document.getElementById("refresh-btn")?.addEventListener("click", renderLembarKerjaTable);
  document.getElementById("filter-btn")?.addEventListener("click", renderLembarKerjaTable);
  document.getElementById("reset-btn")?.addEventListener("click", ()=>{
    document.querySelectorAll("#filter-room,#filter-order,#filter-cph,#filter-mat,#filter-section,#filter-month").forEach(inp=>inp.value="");
    renderLembarKerjaTable();
  });
}

function updateMonthFilterOptions(){
  const monthFilter = document.getElementById("filter-month");
  if(!monthFilter) return;
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  monthFilter.innerHTML = '<option value="">-- All --</option>' + months.map(m=>`<option value="${m}">${m}</option>`).join("");
}

function renderLembarKerjaTable(){
  const tbody = document.querySelector("#output-table tbody");
  if(!tbody) return;
  tbody.innerHTML="";
  // Merge data
  mergedLembarData = [...window.iw39Data];
  // Lookup LOM for Month/Cost/Reman
  mergedLembarData.forEach(row=>{
    const lomRow = lomData.find(l=>l.Order===row.Order);
    if(lomRow){
      row.Month = lomRow.Month || row.Month;
      row.Cost  = lomRow.Cost  || row.Cost;
      row.Reman = lomRow.Reman || row.Reman;
    }
  });

  const fRoom = document.getElementById("filter-room")?.value.toUpperCase() || "";
  const fOrder = document.getElementById("filter-order")?.value.toUpperCase() || "";
  const fCPH   = document.getElementById("filter-cph")?.value.toUpperCase() || "";
  const fMAT   = document.getElementById("filter-mat")?.value.toUpperCase() || "";
  const fSection = document.getElementById("filter-section")?.value.toUpperCase() || "";
  const fMonth   = document.getElementById("filter-month")?.value.toUpperCase() || "";

  mergedLembarData.forEach((row,i)=>{
    if(fRoom && !row.Room?.toUpperCase().includes(fRoom)) return;
    if(fOrder && !row.Order?.toUpperCase().includes(fOrder)) return;
    if(fCPH && !row.CPH?.toUpperCase().includes(fCPH)) return;
    if(fMAT && !row.MAT?.toUpperCase().includes(fMAT)) return;
    if(fSection && !row.Section?.toUpperCase().includes(fSection)) return;
    if(fMonth && !row.Month?.toUpperCase().includes(fMonth)) return;

    const tr = document.createElement("tr");
    tr.innerHTML=`
      <td>${row.Room||""}</td>
      <td>${row.OrderType||""}</td>
      <td>${row.Order||""}</td>
      <td>${row.Description||""}</td>
      <td>${row.CreatedOn||""}</td>
      <td>${row.UserStatus||""}</td>
      <td>${row.MAT||""}</td>
      <td>${row.CPH||""}</td>
      <td>${row.Section||""}</td>
      <td>${row.StatusPart||""}</td>
      <td>${row.Aging||""}</td>
      <td>${row.Month||""}</td>
      <td>${row.Cost||""}</td>
      <td>${row.Reman||""}</td>
      <td>${row.Include||""}</td>
      <td>${row.Exclude||""}</td>
      <td>${row.Planning||""}</td>
      <td>${row.StatusAMT||""}</td>
      <td></td>
    `;
    tbody.appendChild(tr);
  });
}

/* ===================== UPLOAD FILES ===================== */
document.getElementById("upload-btn")
document.getElementById("upload-btn")?.addEventListener("click", ()=>{
  const fileInput = document.getElementById("file-input");
  const type = document.getElementById("file-select")?.value;
  if(!fileInput || fileInput.files.length===0){ alert("Pilih file terlebih dahulu"); return; }
  const file = fileInput.files[0];
  const reader = new FileReader();
  reader.onload = e=>{
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data,{type:"array"});
    const sheetName = workbook.SheetNames[0];
    const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{defval:""});
    switch(type){
      case "IW39": window.iw39Data = json; break;
      case "SUM57": window.sum57Data = json; break;
      case "Planning": window.planningData = json; break;
      case "Budget": window.budgetData = json; break;
      case "Data1": window.data1Data = json; break;
      case "Data2": window.data2Data = json; break;
    }
    alert(type+" uploaded: "+json.length+" rows");
    renderLembarKerjaTable();
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById("clear-files-btn")?.addEventListener("click", ()=>{
  const fileInput = document.getElementById("file-input");
  if(fileInput) fileInput.value="";
  alert("File input cleared");
});

/* ===================== SUMMARY CRM ===================== */
function renderSummaryCRM(){
  const section = document.getElementById("summary");
  if(!section) return;
  section.innerHTML = "<h2>Summary CRM</h2>";

  const monthCount = {};
  lomData.forEach(row=>{
    const m = row.Month || "Unassigned";
    monthCount[m] = (monthCount[m]||0)+1;
  });

  let html = "<table style='width:100%;border-collapse:collapse;'>";
  html += "<thead><tr><th>Month</th><th>Total Orders</th></tr></thead><tbody>";
  Object.keys(monthCount).forEach(m=>{
    html += `<tr><td>${m}</td><td>${monthCount[m]}</td></tr>`;
  });
  html += "</tbody></table>";
  section.innerHTML += html;
}

/* ===================== DOWNLOAD EXCEL ===================== */
function downloadExcel(){
  const wb = XLSX.utils.book_new();

  // Lembar Kerja Sheet
  const lwData = mergedLembarData.map(r=>({
    Room: r.Room,
    OrderType: r.OrderType,
    Order: r.Order,
    Description: r.Description,
    CreatedOn: r.CreatedOn,
    UserStatus: r.UserStatus,
    MAT: r.MAT,
    CPH: r.CPH,
    Section: r.Section,
    StatusPart: r.StatusPart,
    Aging: r.Aging,
    Month: r.Month,
    Cost: r.Cost,
    Reman: r.Reman,
    Include: r.Include,
    Exclude: r.Exclude,
    Planning: r.Planning,
    StatusAMT: r.StatusAMT
  }));
  const lwSheet = XLSX.utils.json_to_sheet(lwData);
  XLSX.utils.book_append_sheet(wb, lwSheet, "Lembar Kerja");

  // LOM Sheet
  const lomSheet = XLSX.utils.json_to_sheet(lomData);
  XLSX.utils.book_append_sheet(wb, lomSheet, "LOM");

  // Summary CRM Sheet
  const summaryData = Object.keys(
    lomData.reduce((acc,row)=>{
      const m = row.Month || "Unassigned";
      acc[m]=(acc[m]||0)+1;
      return acc;
    },{})
  ).map(m=>({Month:m, TotalOrders:lomData.filter(r=>r.Month===m).length}));

  const summarySheet = XLSX.utils.json_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary CRM");

  XLSX.writeFile(wb, "Ndarboe_Report.xlsx");
}

document.getElementById("download")?.insertAdjacentHTML("beforeend", `
  <h2>Download Excel</h2>
  <button id="download-btn">Download All Data</button>
  <p class="small">Membuat file Excel berisi Lembar Kerja, LOM, dan Summary CRM.</p>
`);

document.getElementById("download-btn")?.addEventListener("click", downloadExcel);
