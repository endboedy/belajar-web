// Data global
let iwData = [];
let data1 = [];
let data2 = [];
let sum57 = [];
let planning = [];
let budget = [];
let merged = [];

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
      // Try parsing excel date (number)
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

  // Nav button active
  document.querySelectorAll('nav button.nav-btn').forEach(b=> b.classList.remove('active'));
  const navBtn = document.querySelector(`nav button#nav${pageId.charAt(0).toUpperCase() + pageId.slice(1)}`);
  if(navBtn) navBtn.classList.add('active');

  // Clear messages on page change
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

// Build merged row for tabel
function buildMergedRow(origNorm, orderKey){
  const get = (o,names)=>{
    for(const n of names){
      if(!o) continue;
      if(o[n] !== undefined && o[n] !== null && String(o[n]) !== "") return o[n];
    }
    return "";
  };

  // Ambil data dari normalized rows
  const Room = get(origNorm, ['room','location','area']);
  const OrderType = get(origNorm, ['ordertype','type']);
  const Order = orderKey || get(origNorm, ['order','orderno','no','noorder']);
  const DescriptionRaw = get(origNorm, ['description','desc','keterangan']);
  const CreatedOnRaw = get(origNorm, ['createdon','created','date']);
  const UserStatus = get(origNorm, ['userstatus','status']);
  const MAT = get(origNorm, ['mat','material','materialcode']);

  // Normalize tanggal Created On dan Planning
  const CreatedOn = formatDateExcel(CreatedOnRaw);

  // Lookup data dari Data1, Data2, SUM57, Planning berdasarkan MAT atau Order
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(sum57);
  const planNorm = normalizeRows(planning);

  const d1row = d1Norm.find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const d2row = d2Norm.find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const sumrow = sumNorm.find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};
  const planrow = planNorm.find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};

  // Description rule: jika IW39 description mulai dengan JR, Description = JR
  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(d1row, ['description','desc','keterangan','shortdesc']) || "";

  // CPH rule
  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "";

  const Section = get(d1row, ['section','dept','deptcode','department']);
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";
  let AgingRaw = get(sumrow, ['aging','age']) || "";
  // Buang angka di belakang koma Aging
  let Aging = "";
  if(typeof AgingRaw === "number") Aging = Math.round(AgingRaw).toString();
  else Aging = AgingRaw.toString().split('.')[0];

  // Planning tanggal format "dd-mmm-yyyy"
  const PlanningRaw = get(planrow, ['planning','eventstart','description']);
  let Planning = formatDateExcel(PlanningRaw);

  // Cost calculation dari keys plan dan actual di origNorm
  const planKey = Object.keys(origNorm||{}).find(k=>k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k=>k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal)/16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(2));

  // Include & Exclude
  let Include = "-";
  if(typeof Cost === "number"){
    Include = (String(merged.remanVal||"").toLowerCase().includes("reman") ? Number((Cost*0.25).toFixed(2)) : Number(Cost.toFixed(2)));
  }
  let Exclude = (String(OrderType).toUpperCase() === "PM38") ? "-" : Include;

  // Ambil Month dan Reman dari merged global (disimpan saat Add Orders)
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
    "Status AMT": planrow['statusamt'] || ""
  };
}

// Render tabel Lembar Kerja
function renderTable(data){
  const container = document.getElementById('tableContainer');
  container.innerHTML = "";
  if(!data || data.length === 0){
    container.textContent = "Data Lembar Kerja kosong.";
    return;
  }
  const table = document.createElement('table');
  // Header kolom sesuai urutan
  const columns = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  columns.forEach(c=>{
    const th = document.createElement('th');
    th.textContent = c;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement('tbody');
  data.forEach(row=>{
    const tr = document.createElement('tr');
    columns.forEach(col=>{
      const td = document.createElement('td');
      let val = row[col];
      // Format tanggal untuk Created On dan Planning (string sudah terformat)
      if(col === "Created On" || col === "Planning"){
        val = val || "";
      }
      // Format Aging tanpa koma
      if(col === "Aging" && val){
        val = String(val).split('.')[0];
      }
      // Cost, Include, Exclude - tampilkan 2 desimal kalau number
      if(["Cost","Include","Exclude"].includes(col) && typeof val === "number"){
        val = val.toFixed(2);
      }
      // Hilangkan koma dan angka di belakang koma untuk Order (jika ada)
      if(col === "Order" && val){
        val = String(val).split('.')[0];
      }
      td.textContent = val !== undefined && val !== null ? val : "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  container.appendChild(table);
}

// Tambah order multiple dari textarea inputOrders
document.getElementById('btnAddOrders').addEventListener('click', ()=>{
  const raw = document.getElementById('inputOrders').value || "";
  if(raw.trim() === ""){
    document.getElementById('lmMsg').textContent = "Tidak ada order yang dimasukkan.";
    return;
  }
  document.getElementById('lmMsg').textContent = "";
  const lines = raw.split(/[\n,]+/).map(l=>l.trim()).filter(l=>l!=="");
  lines.forEach(ord=>{
    // Cari data IW39 yang cocok dengan Order no
    const normIW = normalizeRows(iwData);
    const foundIW = normIW.find(r=>String(r['order']).split('.')[0] === ord.split('.')[0]);
    if(foundIW){
      const mergedRow = buildMergedRow(foundIW, ord);
      merged.push(mergedRow);
    } else {
      // Jika tidak ketemu, buat row minimal (Order saja)
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

// Save / Load Lembar Kerja to localStorage
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

// Show hint message clear on inputOrders focus
document.getElementById('inputOrders').addEventListener('focus', ()=>{
  document.getElementById('lmMsg').textContent = "";
});

// Download Excel dari merged data
document.getElementById('btnDownloadExcel').addEventListener('click', ()=>{
  if(!merged || merged.length === 0){
    alert("Data Lembar Kerja kosong, tidak bisa download.");
    return;
  }
  const wb = XLSX.utils.book_new();
  const wsData = [];
  const columns = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
  wsData.push(columns);
  merged.forEach(r=>{
    const row = columns.map(c=>r[c] || "");
    wsData.push(row);
  });
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "LembarKerja");
  XLSX.writeFile(wb, "LembarKerja.xlsx");
});

// On load, load saved data if any
window.addEventListener('load', ()=>{
  loadSaved();
  showPage('pageUpload');
});
