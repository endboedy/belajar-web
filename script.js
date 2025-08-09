// Data global
let iwData = [];
let data1 = [];
let data2 = [];
let sum57 = [];
let planning = [];
let budget = [];
let merged = [];

// Filter global
const filters = {
  room: '',
  order: '',
  mat: '',
  cph: ''
};

// Helper normalize keys
function normalizeKey(k){
  return String(k||"").toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}
function normalizeRows(rows){
  return (rows||[]).map(r=>{
    const o = {};
    Object.keys(r||{}).forEach(k=>{
      o[normalizeKey(k)] = r[k];
    });
    return o;
  });
}
function asNumber(v){
  if(v===undefined || v===null || v==="") return 0;
  const num = Number(String(v).toString().replace(/[^0-9.\-]/g,''));
  return isNaN(num) ? 0 : num;
}

// Format tanggal "dd-mmm-yyyy"
function formatDateExcel(dateInput){
  if(!dateInput) return "";
  let d;
  if(typeof dateInput === "string"){
    d = new Date(dateInput);
    if(isNaN(d)) {
      const n = Number(dateInput);
      if(!isNaN(n)) d = new Date(Date.UTC(1900,0,n-1));
      else return "";
    }
  } else if(dateInput instanceof Date){
    d = dateInput;
  } else if(typeof dateInput === "number"){
    d = new Date(Date.UTC(1900,0,dateInput-1));
  } else {
    return "";
  }
  const day = String(d.getUTCDate()).padStart(2,'0');
  const monthNames = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const month = monthNames[d.getUTCMonth()] || "";
  const year = d.getUTCFullYear();
  return `${day}-${month}-${year}`;
}

// Read Excel File to JSON rows (sheet 0)
function readExcelFile(file){
  return new Promise((resolve)=>{
    if(!file) return resolve([]);
    const reader = new FileReader();
    reader.onload = e=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {defval:""});
        resolve(json);
      }catch(err){
        console.error("read error", err);
        resolve([]);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// Show page menu helper
function showPage(pageId){
  document.querySelectorAll('.page').forEach(p=> p.classList.remove('active'));
  const page = document.getElementById(pageId);
  if(page) page.classList.add('active');

  document.querySelectorAll('nav button.nav-btn').forEach(b=> b.classList.remove('active'));
  const navBtn = document.querySelector(`nav button#nav${pageId.charAt(0).toUpperCase() + pageId.slice(1)}`);
  if(navBtn) navBtn.classList.add('active');

  if(pageId==="pageLembar"){
    document.getElementById('lmMsg').textContent = "";
  }
  if(pageId==="pageUpload"){
    document.getElementById('loadMsg').textContent = "";
  }
}

// Attach nav button event listeners
document.getElementById('navUpload').addEventListener('click', ()=>showPage('pageUpload'));
document.getElementById('navLembar').addEventListener('click', ()=>showPage('pageLembar'));
document.getElementById('navSummary').addEventListener('click', ()=>showPage('pageSummary'));
document.getElementById('navDownload').addEventListener('click', ()=>showPage('pageDownload'));

// Load files button handler
document.getElementById('btnLoad').addEventListener('click', async ()=>{
  const loadMsg = document.getElementById('loadMsg');
  loadMsg.textContent = "Loading files...";
  const fIW = document.getElementById('fileIW39').files[0];
  if(!fIW) {
    loadMsg.textContent = "File IW39 wajib diupload.";
    return;
  }
  const fSUM = document.getElementById('fileSUM57').files[0];
  const fPlan = document.getElementById('filePlanning').files[0];
  const fBud = document.getElementById('fileBudget').files[0];
  const fD1 = document.getElementById('fileData1').files[0];
  const fD2 = document.getElementById('fileData2').files[0];

  iwData = await readExcelFile(fIW);
  sum57 = fSUM ? await readExcelFile(fSUM) : [];
  planning = fPlan ? await readExcelFile(fPlan) : [];
  budget = fBud ? await readExcelFile(fBud) : [];
  data1 = fD1 ? await readExcelFile(fD1) : [];
  data2 = fD2 ? await readExcelFile(fD2) : [];

  loadMsg.textContent = `Selesai load file IW39 (${iwData.length} baris)`;
});

// Build merged row untuk tabel
function buildMergedRow(origNorm, orderKey){
  const get = (o,names)=>{
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]) !== "") return o[n];
    }
    return "";
  };

  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  const DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  const CreatedOn = formatDateExcel(CreatedOnRaw);

  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);

  // Cari d1 row by MAT dan by Room
  const d1rowByMAT = d1Norm.find(r => (r['mat'] || "") === (MAT || ""));
  const d1rowByRoom = d1Norm.find(r => (r['room'] || "") === (Room || ""));
  const sumrow = sumNorm.find(r => (r['order'] || "") === (Order || ""));
  const planrow = planNorm.find(r => (r['order'] || "") === (Order || ""));

  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(origNorm, ['description','desc','keterangan','shortdesc']) || get(d1rowByMAT, ['description','desc','keterangan','shortdesc']) || "";

  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2Norm.find(r => (r['mat']||"")=== (MAT||"")) || {}, ['cph','costperhour','cphvalue']) || get(d1rowByMAT, ['cph','costperhour','cphvalue']) || "";

  const Section = get(d1rowByRoom, ['section','dept','deptcode','department']) || "";
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";
  const StatusAMT = planrow ? (planrow['statusamt'] || "") : "";

  // Aging dan cost kalkulasi
  let AgingRaw = get(sumrow, ['aging','age']) || "";
  let Aging = "";
  if(typeof AgingRaw === "number") Aging = Math.round(AgingRaw).toString();
  else Aging = AgingRaw.toString().split('.')[0];

  const PlanningRaw = get(planrow, ['planning','eventstart','description']);
  let Planning = formatDateExcel(PlanningRaw);

  const planKey = Object.keys(origNorm||{}).find(k=>k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k=>k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal)/16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(1));

  let Include = "-";
  if(typeof Cost === "number"){
    Include = (String(merged.remanVal||"").toLowerCase().includes("reman") ? Number((Cost*0.25).toFixed(1)) : Number(Cost.toFixed(1)));
  }
  let Exclude = (String(OrderType).toUpperCase() === "PM38") ? "-" : Include;

  const Month = (merged.remanValMonth||"");
  const Reman = (merged.remanVal||"");

  return {
    "Room": Room,
    "Order Type": OrderType,
    "Order": Order,
    "Description": Description,
    "Created On": CreatedOn,
    "User Status": UserStatus,
    "MAT": MAT,
    "CPH": CPH,
    "Section": Section,
    "Status Part": StatusPart,
    "Aging": Aging,
    "Month": Month,
    "Cost": Cost,
    "Reman": Reman,
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning,
    "Status AMT": StatusAMT
  };
}

// Fungsi filter data sesuai filter input global
function filteredData(){
  return merged.filter(row=>{
    return (row["Room"] || "").toString().toLowerCase().includes(filters.room)
        && (row["Order"] || "").toString().toLowerCase().includes(filters.order)
        && (row["MAT"] || "").toString().toLowerCase().includes(filters.mat)
        && (row["CPH"] || "").toString().toLowerCase().includes(filters.cph);
  });
}

// Render tabel Lembar Kerja dengan filter dan action edit/delete
function renderTable(data){
  const container = document.getElementById('tableContainer');
  container.innerHTML = "";
  if(!data || data.length === 0){
    container.textContent = "Data Lembar Kerja kosong.";
    return;
  }

  const columns = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

  // Buat div filter input
  const filterDiv = document.createElement('div');
  filterDiv.style.marginBottom = "10px";
  filterDiv.style.display = "flex";
  filterDiv.style.gap = "10px";

  function createFilterInput(placeholder, key){
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = placeholder;
    inp.value = filters[key] || '';
    inp.style.padding = "5px";
    inp.style.flex = "1";
    inp.addEventListener('input', (e)=>{
      filters[key] = e.target.value.toLowerCase();
      renderTable(filteredData());
    });
    return inp;
  }

  filterDiv.appendChild(createFilterInput('Filter Room', 'room'));
  filterDiv.appendChild(createFilterInput('Filter Order', 'order'));
  filterDiv.appendChild(createFilterInput('Filter MAT', 'mat'));
  filterDiv.appendChild(createFilterInput('Filter CPH', 'cph'));

  container.appendChild(filterDiv);

  const filtered = filteredData();

  const table = document.createElement('table');
  table.style.borderCollapse = "collapse";
  table.style.width = "100%";

  // Header
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  columns.forEach(c=>{
    const th = document.createElement('th');
    th.textContent = c;
    th.style.border = "1px solid #ddd";
    th.style.padding = "8px";
    th.style.backgroundColor = "#f2f2f2";

    if(c === "Section"){
      th.style.width = "15%";
    }

    trh.appendChild(th);
  });

  const thAction = document.createElement('th');
  thAction.textContent = "Action";
  thAction.style.border = "1px solid #ddd";
  thAction.style.padding = "8px";
  thAction.style.backgroundColor = "#f2f2f2";
  trh.appendChild(thAction);
  thead.appendChild(trh);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement('tbody');
  filtered.forEach((row, idx)=>{
    const tr = document.createElement('tr');
    columns.forEach(col=>{
      const td = document.createElement('td');
      td.style.border = "1px solid #ddd";
      td.style.padding = "8px";

      let val = row[col];
      if(col === "Created On" || col === "Planning"){
        val = val || "";
      }
      if(col === "Aging" && val){
        val = String(val).split('.')[0];
      }
      if(["Cost","Include","Exclude"].includes(col)){
        td.style.textAlign = "right";
        if(typeof val === "number"){
          val = val.toFixed(1);
        }
      }
      if(col === "Order" && val){
        val = String(val).split('.')[0];
      }
      td.textContent = val !== undefined && val !== null ? val : "";
      tr.appendChild(td);
    });

    // Action
    const tdAction = document.createElement('td');
    tdAction.style.border = "1px solid #ddd";
    tdAction.style.padding = "8px";

    const btnEdit = document.createElement('button');
    btnEdit.textContent = "Edit";
    btnEdit.style.marginRight = "5px";
    btnEdit.style.cursor = "pointer";
    btnEdit.addEventListener('click', () => openEditModal(idx));

    const btnDelete = document.createElement('button');
    btnDelete.textContent = "Delete";
    btnDelete.style.cursor = "pointer";
    btnDelete.style.color = "white";
    btnDelete.style.backgroundColor = "red";
    btnDelete.addEventListener('click', () => {
      if(confirm(`Hapus order ${row["Order"]}?`)){
        merged.splice(idx, 1);
        renderTable(merged);
      }
    });

    tdAction.appendChild(btnEdit);
    tdAction.appendChild(btnDelete);
    tr.appendChild(tdAction);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  container.appendChild(table);
}

// Add multiple orders dari textarea inputOrders
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim() === ""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  document.getElementById('lmMsg').textContent = "";
  const lines = raw.split(/[\n,]+/).map(l=>l.trim()).filter(l=>l!=="");
  lines.forEach(ord=>{
    const normIW = normalizeRows(iwData);
    const foundIW = normIW.find(r=>String(r['order']).split('.')[0] === ord.split('.')[0]);
    if(foundIW){
      const mergedRow = buildMergedRow(foundIW, ord);
      merged.push(mergedRow);
    } else {
      merged.push({
        "Room":"-",
        "Order Type":"-",
        "Order":ord.split('.')[0],
        "Description":"-",
        "Created On":"",
        "User Status":"",
        "MAT":"",
        "CPH":"",
        "Section":"",
        "Status Part":"",
        "Aging":"",
        "Month":"",
        "Cost":"",
        "Reman":"",
        "Include":"",
        "Exclude":"",
        "Planning":"",
        "Status AMT":""
      });
    }
  });
  renderTable(merged);
  document.getElementById('inputOrders').value = "";
});

// Save / Load Lembar Kerja ke localStorage
document.getElementById('btnSave').addEventListener('click', ()=>{
  localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
  document.getElementById('lmMsg').textContent = "Lembar Kerja tersimpan di localStorage.";
});
function loadSaved(){
  const raw = localStorage.getItem('ndarboe_merged');
  if(raw){
    try{
      merged = JSON.parse(raw);
      renderTable(merged);
      document.getElementById('lmMsg').textContent = "Lembar Kerja dimuat dari localStorage.";
    }catch(e){
      console.error(e);
      document.getElementById('lmMsg').textContent = "Gagal memuat Lembar Kerja dari localStorage.";
    }
  }
}

// Clear Lembar Kerja
document.getElementById('btnClear').addEventListener('click', ()=>{
  if(!confirm("Yakin ingin hapus semua Lembar Kerja?")) return;
  merged = [];
  renderTable(merged);
  localStorage.removeItem('ndarboe_merged');
  document.getElementById('lmMsg').textContent = "Lembar Kerja dihapus.";
});

// Clear pesan saat inputOrders diubah
document.getElementById('inputOrders').addEventListener('input', ()=>{
  document.getElementById('lmMsg').textContent = "";
});

// Edit modal: hanya Month dan Reman
function openEditModal(index){
  if(index < 0 || index >= merged.length) return;
  const row = merged[index];

  const newMonth = prompt("Edit Month:", row.Month || "");
  if(newMonth !== null) row.Month = newMonth;

  const newReman = prompt("Edit Reman:", row.Reman || "");
  if(newReman !== null) row.Reman = newReman;

  renderTable(merged);
}

// Inisialisasi page
showPage('pageUpload');
loadSaved();
