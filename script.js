/****************************************************
 * Ndarboe.net - Full script.js (Revisi + Pewarnaan)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1Data = [];
let data2Data = [];
let budgetData = [];
let mergedData = [];
const UI_LS_KEY = "ndarboe_ui_edits_v2";

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupMenu();
  setupButtons();
  renderTable([]);
  updateMonthFilterOptions();
});

/* ===================== MENU HANDLER ===================== */
function setupMenu() {
  const menuItems = document.querySelectorAll(".sidebar .menu-item");
  const contentSections = document.querySelectorAll(".content-section");
  menuItems.forEach(item => {
    item.addEventListener("click", () => {
      menuItems.forEach(i => i.classList.remove("active"));
      item.classList.add("active");
      const menuId = item.dataset.menu;
      contentSections.forEach(sec => {
        sec.classList.toggle("active", sec.id === menuId);
      });
    });
  });
}

/* ===================== HELPERS ===================== */
function toDateObj(anyDate) {
  if (!anyDate) return null;
  if (anyDate instanceof Date && !isNaN(anyDate)) return anyDate;
  if (typeof anyDate === "number") return new Date(Math.round((anyDate - 25569) * 86400 * 1000));
  const d = new Date(anyDate);
  return isNaN(d) ? null : d;
}

function formatDateDDMMMYYYY(input) {
  const d = toDateObj(input);
  if (!d) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

function formatDateISO(input) {
  const d = toDateObj(input);
  if (!d) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${yyyy}-${mm}-${dd}`;
}

function formatNumber1(val){
  const n = parseFloat(val);
  return isNaN(n) ? val : n.toFixed(1);
}

function safe(val){ return val == null ? "" : val; }

function asColoredStatusUser(val) {
  const v = (val || "").toString().toUpperCase();
  let bg="", fg="";
  if(v==="OUTS"){ bg="#ffeb3b"; fg="#000"; }
  else if(v==="RELE"){ bg="#2e7d32"; fg="#fff"; }
  else if(v==="PROG"){ bg="#1976d2"; fg="#fff"; }
  else if(v==="COMP"){ bg="#d32f2f"; fg="#fff"; }
  else return safe(val);
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${safe(val)}</span>`;
}

function asColoredStatusPart(val){
  const s=(val||"").toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

function asColoredStatusAMT(val){
  const v=(val||"").toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#bdbdbd;color:#000;">${safe(val)}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

function asColoredAging(val){
  const n=parseInt(val);
  if(isNaN(n)) return val;
  if(n>=1 && n<30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  if(n>=30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${val}</span>`;
  return val;
}

/* ===================== UPLOAD ===================== */
async function handleUpload() {
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const status = document.getElementById("upload-status");
  if (!fileInput.files.length) { alert("Pilih file terlebih dahulu."); return; }
  const file = fileInput.files[0];
  const jenis = fileSelect.value;
  status.textContent = `Memproses ${file.name} sebagai ${jenis}...`;
  try {
    const data = await parseFile(file, jenis);
    switch (jenis) {
      case "IW39": iw39Data = data; break;
      case "SUM57": sum57Data = data; break;
      case "Planning": planningData = data; break;
      case "Data1": data1Data = data; break;
      case "Data2": data2Data = data; break;
      case "Budget": budgetData = data; break;
      default: break;
    }
    status.textContent = `File ${file.name} berhasil diupload (${data.length} baris)`;
    fileInput.value = "";
  } catch (e) {
    status.textContent = `Error: ${e.message}`;
  }
}

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded", () => {
  setupButtons(); // pastikan handleUpload sudah ada di atas
});

/* ===================== MERGE ===================== */
function mergeData(){
  if(!iw39Data.length){ alert("Upload data IW39 dulu."); return; }
  mergedData = iw39Data.map(row=>({
    Room: row.Room||"", "Order Type": row["Order Type"]||"", Order: row.Order||"", Description: row.Description||"",
    "Created On": row["Created On"]||"", "User Status": row["User Status"]||"", MAT: row.MAT||"", CPH:"", Section:"", "Status Part":"", Aging:"", Month:row.Month||"",
    Cost:"-", Reman:row.Reman||"", Include:"-", Exclude:"-", Planning:"", "Status AMT":""
  }));

  mergedData.forEach(md=>{
    if(md.Description.trim().toUpperCase().startsWith("JR")) md.CPH="External Job";
    else{
      const d2=data2Data.find(d=>d.MAT==md.MAT);
      if(d2) md.CPH=d2.CPH||"";
    }
  });

  mergedData.forEach(md=>{
    const d1=data1Data.find(d=>d.Room==md.Room);
    if(d1) md.Section=d1.Section||"";
  });

  mergedData.forEach(md=>{
    const s57=sum57Data.find(s=>s.Order==md.Order);
    if(s57){ md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||""; }
  });

  mergedData.forEach(md=>{
    const pl=planningData.find(p=>p.Order==md.Order);
    if(pl){ md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||""; }
  });

  mergedData.forEach(md=>{
    const src=iw39Data.find(i=>i.Order==md.Order);
    if(!src) return;
    const plan=parseFloat(src["Total sum (plan)"]?.toString().replace(/,/g,""))||0;
    const actual=parseFloat(src["Total sum (actual)"]?.toString().replace(/,/g,""))||0;
    let cost=(plan-actual)/16500;
    if(!isFinite(cost)||cost<0){ md.Cost="-"; md.Include="-"; md.Exclude="-"; }
    else{
      md.Cost=cost.toFixed(1);
      const isReman=(md.Reman||"").toLowerCase().includes("reman");
      md.Include=(isReman ? cost*0.25 : cost).toFixed(1);
      md.Exclude=md.Include;
    }
  });

  // restore user edits
  try{
    const raw=localStorage.getItem(UI_LS_KEY);
    if(raw){ const saved=JSON.parse(raw); if(saved.userEdits) saved.userEdits.forEach(e=>{
      const idx=mergedData.findIndex(r=>r.Order===e.Order);
      if(idx!==-1) mergedData[idx]={...mergedData[idx],...e};
    }); }
  }catch{}

  updateMonthFilterOptions();
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender=mergedData){
  const tbody=document.querySelector("#output-table tbody");
  if(!tbody){ console.warn("Tabel #output-table tidak ditemukan."); return; }
  tbody.innerHTML="";
  dataToRender.forEach((row,index)=>{
    const tr=document.createElement("tr"); tr.dataset.index=index;
    const cols=["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
    cols.forEach(col=>{
      const td=document.createElement("td");
      const val=row[col]||"";
      if(col==="User Status") td.innerHTML=asColoredStatusUser(val);
      else if(col==="Status Part") td.innerHTML=asColoredStatusPart(val);
      else if(col==="Status AMT") td.innerHTML=asColoredStatusAMT(val);
      else if(col==="Aging") td.innerHTML=asColoredAging(val);
      else if(col==="Created On"||col==="Planning") td.textContent=formatDateDDMMMYYYY(val);
      else if(col==="Cost"||col==="Include"||col==="Exclude"){ td.textContent=formatNumber1(val); td.style.textAlign="right"; }
      else td.textContent=val;
      tr.appendChild(td);
    });
    const tdAct=document.createElement("td");
    tdAct.innerHTML=`<button class="action-btn edit-btn" data-order="${safe(row.Order)}">Edit</button>
                      <button class="action-btn delete-btn" data-order="${safe(row.Order)}">Delete</button>`;
    tr.appendChild(tdAct);
    tbody.appendChild(tr);
  });
  attachTableEvents();
}

/* ===================== EDIT / SAVE / DELETE ===================== */
function attachTableEvents(){
  document.querySelectorAll(".edit-btn").forEach(btn=>{ btn.onclick=()=>startEdit(btn.dataset.order); });
  document.querySelectorAll(".delete-btn").forEach(btn=>{ btn.onclick=()=>deleteOrder(btn.dataset.order); });
}

function startEdit(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  const tbody=document.querySelector("#output-table tbody"); const tr=tbody.children[idx]; const row=mergedData[idx];
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m))).sort();
  const monthOptions=`<option value="">--Select Month--</option>`+months.map(m=>`<option value="${m}">${m}</option>`).join("");
  tr.innerHTML=`
    <td><input type="text" value="${safe(row.Room)}" data-field="Room"/></td>
    <td><input type="text" value="${safe(row["Order Type"])}" data-field="Order Type"/></td>
    <td>${safe(row.Order)}</td>
    <td><input type="text" value="${safe(row.Description)}" data-field="Description"/></td>
    <td><input type="date" value="${formatDateISO(row["Created On"])}" data-field="Created On"/></td>
    <td><input type="text" value="${safe(row["User Status"])}" data-field="User Status"/></td>
    <td><input type="text" value="${safe(row.MAT)}" data-field="MAT"/></td>
    <td><input type="text" value="${safe(row.CPH)}" data-field="CPH"/></td>
    <td><input type="text" value="${safe(row.Section)}" data-field="Section"/></td>
    <td><input type="text" value="${safe(row["Status Part"])}" data-field="Status Part"/></td>
    <td><input type="text" value="${safe(row.Aging)}" data-field="Aging"/></td>
    <td><select data-field="Month">${monthOptions}</select></td>
    <td><input type="text" value="${safe(row.Cost)}" data-field="Cost" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${safe(row.Reman)}" data-field="Reman"/></td>
    <td><input type="text" value="${safe(row.Include)}" data-field="Include" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${safe(row.Exclude)}" data-field="Exclude" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="date" value="${formatDateISO(row.Planning)}" data-field="Planning"/></td>
    <td><input type="text" value="${safe(row["Status AMT"])}" data-field="Status AMT"/></td>
    <td>
      <button class="action-btn save-btn" data-order="${safe(order)}">Save</button>
      <button class="action-btn cancel-btn" data-order="${safe(order)}">Cancel</button>
    </td>
  `;
  tr.querySelector("select[data-field='Month']").value=row.Month||"";
  tr.querySelector(".save-btn").onclick=()=>saveEdit(order);
  tr.querySelector(".cancel-btn").onclick=()=>cancelEdit();
}

function cancelEdit(){ renderTable(mergedData); }

function saveEdit(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  const tr=document.querySelector("#output-table tbody").children[idx];
  tr.querySelectorAll("input[data-field], select[data-field]").forEach(input=>{
    mergedData[idx][input.dataset.field]=input.value;
  });
  saveUserEdits();
  mergeData();
  renderTable(mergedData);
}

function deleteOrder(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  if(!confirm(`Hapus data order ${order}?`)) return;
  mergedData.splice(idx,1);
  saveUserEdits();
  renderTable(mergedData);
}

/* ===================== FILTER ===================== */
function filterData(){
  const roomF=document.getElementById("filter-room").value.toLowerCase();
  const orderF=document.getElementById("filter-order").value.toLowerCase();
  const cphF=document.getElementById("filter-cph").value.toLowerCase();
  const matF=document.getElementById("filter-mat").value.toLowerCase();
  const secF=document.getElementById("filter-section").value.toLowerCase();
  const monthF=document.getElementById("filter-month").value.toLowerCase();
  const filtered=mergedData.filter(r=>{
    if(roomF && !r.Room.toLowerCase().includes(roomF)) return false;
    if(orderF && !r.Order.toLowerCase().includes(orderF)) return false;
    if(cphF && !r.CPH.toLowerCase().includes(cphF)) return false;
    if(matF && !r.MAT.toLowerCase().includes(matF)) return false;
    if(secF && !r.Section.toLowerCase().includes(secF)) return false;
    if(monthF && r.Month.toLowerCase()!==monthF) return false;
    return true;
  });
  renderTable(filtered);
}

function resetFilters(){
  ["room","order","cph","mat","section","month"].forEach(f=>document.getElementById("filter-"+f).value="");
  renderTable(mergedData);
}

function updateMonthFilterOptions(){
  const sel=document.getElementById("filter-month");
  if(!sel) return;
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m))).sort();
  sel.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDERS ===================== */
function addOrders(){
  const input=document.getElementById("add-order-input");
  const status=document.getElementById("add-order-status");
  const orders=(input.value||"").trim().split(/[\s,]+/).filter(Boolean);
  let added=0;
  orders.forEach(o=>{
    if(!mergedData.some(r=>r.Order===o)){
      mergedData.push({
        Room:"", "Order Type":"", Order:o, Description:"", "Created On":"", "User Status":"",
        MAT:"", CPH:"", Section:"", "Status Part":"", Aging:"", Month:"",
        Cost:"-", Reman:"", Include:"-", Exclude:"-", Planning:"", "Status AMT":""
      }); added++;
    }
  });
  if(added){ saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan.`; }
  else status.textContent="Order sudah ada di data.";
  input.value="";
}

/* ===================== SAVE / LOAD JSON ===================== */
function saveToJSON(){
  if(!mergedData.length){ alert("Tidak ada data."); return; }
  const dataStr=JSON.stringify({mergedData,timestamp:new Date().toISOString()},null,2);
  const blob=new Blob([dataStr],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url; a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

async function loadFromJSON(file){
  try{
    const text=await file.text();
    const obj=JSON.parse(text);
    if(obj.mergedData && Array.isArray(obj.mergedData)){
      mergedData=obj.mergedData; renderTable(mergedData); updateMonthFilterOptions();
      alert("Data berhasil dimuat dari JSON.");
    }else alert("File JSON tidak valid.");
  }catch(e){ alert("Gagal membaca file JSON: "+e.message); }
}

/* ===================== USER EDITS PERSISTENCE ===================== */
function saveUserEdits(){
  try{
    localStorage.setItem(UI_LS_KEY, JSON.stringify({userEdits:mergedData.map(r=>({...r}))}));
  }catch{}
}

/* ===================== CLEAR ALL ===================== */
function clearAllData(){
  if(!confirm("Yakin hapus semua data?")) return;
  iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[]; mergedData=[];
  renderTable([]); document.getElementById("upload-status").textContent="Data dihapus.";
  updateMonthFilterOptions();
}

/* ===================== BUTTONS ===================== */
function setupButtons(){
  const uploadBtn=document.getElementById("upload-btn");
  if(uploadBtn) uploadBtn.onclick=handleUpload;

  const clearBtn=document.getElementById("clear-files-btn");
  if(clearBtn) clearBtn.onclick=clearAllData;

  const refreshBtn=document.getElementById("refresh-btn");
  if(refreshBtn) refreshBtn.onclick=()=>{ mergeData(); renderTable(mergedData); };

  const filterBtn=document.getElementById("filter-btn");
  if(filterBtn) filterBtn.onclick=filterData;

  const resetBtn=document.getElementById("reset-btn");
  if(resetBtn) resetBtn.onclick=resetFilters;

  const saveBtn=document.getElementById("save-btn");
  if(saveBtn) saveBtn.onclick=saveToJSON;

  const loadBtn=document.getElementById("load-btn");
  if(loadBtn){
    loadBtn.onclick=()=>{
      const input=document.createElement("input");
      input.type="file"; input.accept="application/json";
      input.onchange=()=>{ if(input.files.length) loadFromJSON(input.files[0]); };
      input.click();
    };
  }

  const addOrderBtn=document.getElementById("add-order-btn");
  if(addOrderBtn) addOrderBtn.onclick=addOrders;
}

