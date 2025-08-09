
document.addEventListener("DOMContentLoaded", () => {
  //====================
  // Navigasi Menu
  //====================
  document.getElementById('btnUpload').addEventListener('click', () => showPage('pageUpload'));
  document.getElementById('btnLembar').addEventListener('click', () => showPage('pageLembar'));
  document.getElementById('btnSummary').addEventListener('click', () => showPage('pageSummary'));
  document.getElementById('btnDownload').addEventListener('click', () => showPage('pageDownload'));
  const btnDownloadExcel = document.getElementById('btnDownloadExcel');
  if (btnDownloadExcel) btnDownloadExcel.addEventListener('click', downloadExcel);
  document.getElementById('btnProcess').addEventListener('click', () => processData());

  showPage('pageUpload');
});

//====================
// Page Switcher
//====================
function showPage(id){
  document.querySelectorAll('.page').forEach(p => p.style.display = 'none');
  document.getElementById(id).style.display = 'block';
  document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
  if (id === 'pageUpload') document.getElementById('btnUpload').classList.add('active');
  if (id === 'pageLembar') document.getElementById('btnLembar').classList.add('active');
  if (id === 'pageSummary') document.getElementById('btnSummary').classList.add('active');
  if (id === 'pageDownload') document.getElementById('btnDownload').classList.add('active');
}

//====================
// Utilities
//====================
function normalizeKey(k){
  if(k===undefined || k===null) return "";
  return String(k).toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}
function normalizeRows(rows){
  return rows.map(r=>{
    const obj = {};
    Object.keys(r||{}).forEach(k=>{
      obj[normalizeKey(k)] = r[k];
    });
    return obj;
  });
}
function asNumber(v){
  if (v === undefined || v === null || v === "") return 0;
  const n = Number(String(v).replace(/[^0-9.\-]/g,''));
  return isNaN(n) ? 0 : n;
}

//====================
// File Reader
//====================
function readExcelFile(file) {
  return new Promise((resolve) => {
    if(!file) return resolve([]);
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {defval: ""});
        resolve(json);
      }catch(err){
        console.error("read error", err);
        resolve([]);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

//====================
// Global Variables
//====================
let dataIW39 = [], data1 = [], data2 = [], dataSUM57 = [], dataPlanning = [];
let mergedData = [];
const DISPLAY_COLUMNS = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

//====================
// Process Data
//====================
async function processData(){
  document.getElementById('uploadMessage').textContent = "Memproses... tunggu sebentar.";
  dataIW39 = await readExcelFile(document.getElementById('fileIW39').files[0]);
  data1 = await readExcelFile(document.getElementById('fileData1').files[0]);
  data2 = await readExcelFile(document.getElementById('fileData2').files[0]);
  dataSUM57 = await readExcelFile(document.getElementById('fileSUM57').files[0]);
  dataPlanning = await readExcelFile(document.getElementById('filePlanning').files[0]);

  const iwNorm = normalizeRows(dataIW39);
  const d1Norm = normalizeRows(data1);
  const d2Norm = normalizeRows(data2);
  const sumNorm = normalizeRows(dataSUM57);
  const planNorm = normalizeRows(dataPlanning);

  mergeData(iwNorm, d1Norm, d2Norm, sumNorm, planNorm);

  document.getElementById('uploadMessage').textContent = `Selesai. Baris diproses: ${mergedData.length}. Lihat di Lembar Kerja.`;
  showPage('pageLembar');
}

//====================
// Merge Data
//====================
function mergeData(iw, d1, d2, sum, plan){
  const month = document.getElementById('monthInput').value || "";
  const remanInput = String(document.getElementById('remanInput').value || "").trim();
  mergedData = [];

  iw.forEach(orig => {
    const Room = orig['room'] || "";
    const Order = orig['order'] || "";
    const OrderType = orig['ordertype'] || "";
    const DescriptionRaw = orig['description'] || "";
    const CreatedOn = orig['createdon'] || "";
    const UserStatus = orig['userstatus'] || "";
    const MAT = orig['mat'] || "";

    const d1row = d1.find(r => r['mat'] === MAT) || {};
    const d2row = d2.find(r => r['mat'] === MAT) || {};
    const sumrow = sum.find(r => r['order'] === Order) || {};
    const planrow = plan.find(r => r['order'] === Order) || {};

    let Description = DescriptionRaw.toUpperCase().startsWith("JR") ? "JR" : (d1row['description'] || "");
    const Section = d1row['section'] || "";
    const CPH = d2row['cph'] || "";
    const StatusPart = sumrow['statuspart'] || "";
    const Aging = sumrow['aging'] || "";
    const Planning = planrow['planning'] || "";
    const StatusAMT = planrow['statusamt'] || "";

    const planVal = asNumber(orig['totalplan']);
    const actualVal = asNumber(orig['totalactual']);
    const rawCost = (planVal - actualVal) / 16500;
    const cost = rawCost >= 0 ? Number(rawCost.toFixed(2)) : "-";

    let Include = cost === "-" ? "-" : (remanInput.toLowerCase().includes("reman") ? Number((cost * 0.25).toFixed(2)) : cost);
    let Exclude = OrderType.toUpperCase() === "PM38" ? "-" : Include;

    const rowOut = {
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
      "Month": month,
      "Cost": cost,
      "Reman": remanInput,
      "Include": Include,
      "Exclude": Exclude,
      "Planning": Planning,
      "Status AMT": StatusAMT
    };

    mergedData.push(rowOut);
  });

  renderTable(mergedData);
  renderSummary();
}

//====================
// Render Table
//====================
function renderTable(data){
  const container = document.getElementById('tableContainer');
  if (!data || data.length === 0) {
    let html = `<div class='table-wrap'><table><thead><tr>`;
    DISPLAY_COLUMNS.forEach(c => html += `<th>${c}</th>`);
    html += `</tr></thead><tbody><tr><td colspan="${DISPLAY_COLUMNS.length}" style="text-align:center;padding:18px;color:#666">Tidak ada data. Silakan upload file dan klik Proses Data.</td></tr></tbody></table></div>`;
    container.innerHTML = html;
    return;
  }

  let html = `<div class='table-wrap'><table><thead><tr>`;
  DISPLAY_COLUMNS.forEach(c => html += `<th>${c}</th>`);
  html += `</tr></thead><tbody>`;
  data.forEach(row => {
    html += `<tr>`;
    DISPLAY_COLUMNS.forEach(col => {
      let v = row[col];
      if (typeof v === "number") v = v.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2});
      html += `<td>${v !== undefined && v !== null ? v : ""}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

//====================
// Filter Table
//====================
function filterTable(){
  const roomVal = document.getElementById('filterRoom').value.toLowerCase();
  const orderVal = document.getElementById('filterOrder').value.toLowerCase();
  const matVal = document.getElementById('filterMAT').value.toLowerCase();
  const sectionVal = document.getElementById('filterSection').value.toLowerCase();
  const cphVal = document.getElementById('filterCPH').value.toLowerCase();

  const filtered = mergedData.filter(r =>
    String(r["Room"] || "").toLowerCase().includes(roomVal) &&
    String(r["Order"] || "").toLowerCase().includes(orderVal) &&
    String(r["MAT"] || "").toLowerCase().includes(matVal) &&
    String(r["Section"] || "").toLowerCase().includes(sectionVal) &&
    String(r["CPH"] || "").toLowerCase().includes(cphVal)
  );

  renderTable(filtered);
}

//====================
// Summary
//====================
function renderSummary(){
  const totalCost = mergedData.reduce((s, r) => s + (typeof r.Cost === 'number' ? r.Cost : 0), 0);
  const totalInclude = mergedData.reduce((s, r) => s + (typeof r.Include === 'number' ? r.Include : 0), 0);
  const totalExclude = mergedData.reduce((s, r) => s + (typeof r.Exclude === 'number' ? r.Exclude : 0), 0);

  const el = document.getElementById('summaryContainer');
  el.innerHTML = `<div>
    <p><strong>Total baris:</strong> ${mergedData.length}</p>
    <p><strong>Total Cost:</strong> ${totalCost.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}</p>
    <p><strong>Total Include:</strong> ${totalInclude.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}</p>
    <p><strong>Total Exclude:</strong> ${totalExclude.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}</p>
  </div>`;
}

//====================
// Download Excel
//====================
function downloadExcel(){
  if (!mergedData || mergedData.length === 0) {
    alert("Tidak ada data untuk didownload. Proses data terlebih dahulu.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(mergedData, {header: DISPLAY_COLUMNS});
  XLSX.utils.sheet_add_aoa(ws, [DISPLAY_COLUMNS], {origin: "A1"});
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hasil_Merge");
  XLSX.writeFile(wb, "Hasil_Merge.xlsx");
}
