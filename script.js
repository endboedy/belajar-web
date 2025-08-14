/****************************************************
 * Ndarboe.net - FULL script.js (revisi)
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
  renderTable();        // render awal
  updateMonthFilterOptions();
  setupButtons();
  setupMenu();          // panggil menu
  setupTableEvents();   // event edit/delete/save/cancel
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
        if (sec.id === menuId) sec.classList.add("active");
        else sec.classList.remove("active");
      });

      if(menuId === "menu1"){ renderTable(mergedData); }
    });
  });
}

/* ===================== HELPERS ===================== */
function safe(str){ return str==null?"":str.toString(); }

function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;
  if (anyDate instanceof Date && !isNaN(anyDate)) return anyDate;
  if (typeof anyDate === "number") {
    if (XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const dec = XLSX.SSF.parse_date_code(anyDate);
      if (dec) return new Date(dec.y, (dec.m||1)-1, dec.d||1, dec.H||0, dec.M||0, dec.S||0);
    }
  }
  if (typeof anyDate === "string") {
    const d = new Date(anyDate);
    return isNaN(d)?null:d;
  }
  return null;
}

function formatDateDDMMMYYYY(input) {
  if (!input) return "";
  let d = (typeof input === "number") ? excelDateToJS(input) : new Date(input);
  if (!d || isNaN(d)) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

function formatDateISO(anyDate) {
  const d = toDateObj(anyDate);
  if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

/* ===================== UPLOAD & PARSE ===================== */
async function parseFile(file, jenis) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        let sheetName = workbook.SheetNames.includes(jenis) ? jenis : workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval:"", raw:false });
        resolve(json);
      } catch(err) { reject(err); }
    };
    reader.onerror = err => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

async function handleUpload() {
  const fileSelect = document.getElementById("file-select");
  const fileInput = document.getElementById("file-input");
  const status = document.getElementById("upload-status");
  if (!fileInput.files.length){ alert("Pilih file dulu."); return; }
  const file = fileInput.files[0];
  const jenis = fileSelect.value;
  status.textContent = `Memproses file ${file.name} sebagai ${jenis}...`;
  try {
    const json = await parseFile(file, jenis);
    switch(jenis){
      case "IW39": iw39Data=json; break;
      case "SUM57": sum57Data=json; break;
      case "Planning": planningData=json; break;
      case "Data1": data1Data=json; break;
      case "Data2": data2Data=json; break;
      case "Budget": budgetData=json; break;
    }
    status.textContent = `File ${file.name} berhasil diupload (${json.length} rows)`;
    fileInput.value="";
  } catch(e){
    status.textContent = `Error: ${e.message}`;
  }
}

/* ===================== MERGE ===================== */
function mergeData(){
  if(!iw39Data.length){ alert("Upload data IW39 dulu"); return; }
  mergedData = iw39Data.map(row=>({
    Room: safe(row.Room),
    "Order Type": safe(row["Order Type"]),
    Order: safe(row.Order),
    Description: safe(row.Description),
    "Created On": row["Created On"]||"",
    "User Status": safe(row["User Status"]),
    MAT: safe(row.MAT),
    CPH: "",
    Section: "",
    "Status Part": "",
    Aging: "",
    Month: safe(row.Month),
    Cost: "-",
    Reman: safe(row.Reman),
    Include: "-",
    Exclude: "-",
    Planning: "",
    "Status AMT": ""
  }));

  mergedData.forEach(md=>{
    if((md.Description||"").trim().toUpperCase().startsWith("JR")) md.CPH="External Job";
    else{
      const d2=data2Data.find(d=>safe(d.MAT)===md.MAT.trim());
      md.CPH=d2?safe(d2.CPH):"";
    }
  });

  mergedData.forEach(md=>{
    const d1=data1Data.find(d=>safe(d.Room)===md.Room.trim());
    md.Section=d1?safe(d1.Section):"";
  });

  mergedData.forEach(md=>{
    const s57=sum57Data.find(s=>safe(s.Order)===md.Order);
    if(s57){ md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||""; }
  });

  mergedData.forEach(md=>{
    const pl=planningData.find(p=>safe(p.Order)===md.Order);
    if(pl){ md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||""; }
  });

  mergedData.forEach(md=>{
    const src=iw39Data.find(i=>safe(i.Order)===md.Order);
    if(!src) return;
    const plan=parseFloat((src["Total sum (plan)"]||"").replace(/,/g,""))||0;
    const actual=parseFloat((src["Total sum (actual)"]||"").replace(/,/g,""))||0;
    let cost=(plan-actual)/16500;
    if(!isFinite(cost)||cost<0){ md.Cost="-"; md.Include="-"; md.Exclude="-"; }
    else{
      md.Cost=cost.toFixed(2);
      const isReman=(md.Reman||"").toLowerCase().includes("reman");
      const includeNum=isReman?cost*0.25:cost;
      md.Include=includeNum.toFixed(2);
      md.Exclude=md["Order Type"]==="PM38"? "-":md.Include;
    }
  });

  // restore edits
  try{
    const raw=localStorage.getItem(UI_LS_KEY);
    if(raw){ const saved=JSON.parse(raw); saved.userEdits?.forEach(edit=>{
      const idx=mergedData.findIndex(r=>r.Order===edit.Order);
      if(idx!==-1) mergedData[idx]={...mergedData[idx], ...edit};
    }); }
  }catch{}
  updateMonthFilterOptions();
}

/* ===================== RENDER TABLE ===================== */
function renderTable(data=mergedData){
  const tbody=document.querySelector("#output-table tbody");
  if(!tbody) return;
  tbody.innerHTML="";
  data.forEach((row,idx)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${safe(row.Room)}</td>
      <td>${safe(row["Order Type"])}</td>
      <td>${safe(row.Order)}</td>
      <td>${safe(row.Description)}</td>
      <td>${formatDateDDMMMYYYY(row["Created On"])}</td>
      <td>${asColoredStatusUser(row["User Status"])}</td>
      <td>${safe(row.MAT)}</td>
      <td>${safe(row.CPH)}</td>
      <td>${safe(row.Section)}</td>
      <td>${asColoredStatusPart(row["Status Part"])}</td>
      <td>${safe(row.Aging)}</td>
      <td class="col-month">${safe(row.Month)}</td>
      <td class="col-cost" style="text-align:right;">${safe(row.Cost)}</td>
      <td class="col-reman">${safe(row.Reman)}</td>
      <td style="text-align:right;">${safe(row.Include)}</td>
      <td style="text-align:right;">${safe(row.Exclude)}</td>
      <td>${formatDateDDMMMYYYY(row.Planning)}</td>
      <td>${asColoredStatusAMT(row["Status AMT"])}</td>
      <td>
        <button class="action-btn edit-btn" data-index="${idx}">Edit</button>
        <button class="action-btn delete-btn" data-index="${idx}">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/* ===================== CELL COLORING ===================== */
function asColoredStatusUser(val){
  const v=(val||"").toUpperCase(); let bg="",fg="";
  if(v==="OUTS"){ bg="#ffeb3b"; fg="#000"; }
  else if(v==="RELE"){ bg="#2e7d32"; fg="#fff"; }
  else if(v==="PROG"){ bg="#1976d2"; fg="#fff"; }
  else if(v==="COMP"){ bg="#000"; fg="#fff"; }
  else return safe(val);
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${safe(val)}</span>`;
}

function asColoredStatusPart(val){
  const s=(val||"").toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

function asColoredStatusAMT(val){
  const v=(val||"").toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#ffeb3b;color:#000;">${safe(val)}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

/* ===================== EDIT / DELETE ===================== */
function setupTableEvents(){
  const tbody=document.querySelector("#output-table tbody");
  if(!tbody) return;

  tbody.addEventListener("click", function(e){
    const btn=e.target;
    const tr=btn.closest("tr");
    if(!tr) return;
    const index=btn.dataset.index;

    if(btn.classList.contains("edit-btn")){
      activateEdit(tr,index);
    } else if(btn.classList.contains("delete-btn")){
      if(confirm("Hapus data ini?")){
        mergedData.splice(index,1);
        saveUserEdits();
        renderTable(mergedData);
      }
    } else if(btn.classList.contains("save-btn")){
      const tdMonth=tr.querySelector("td.col-month select");
      const tdCost=tr.querySelector("td.col-cost input");
      const tdReman=tr.querySelector("td.col-reman select");
      mergedData[index].Month=tdMonth.value;
      mergedData[index].Cost=tdCost.value;
      mergedData[index].Reman=tdReman.value;
      saveUserEdits();
      renderTable(mergedData);
    } else if(btn.classList.contains("cancel-btn")){
      renderTable(mergedData);
    }
  });
}

function activateEdit(tr,index){
  const tdMonth=tr.querySelector("td.col-month");
  const tdCost=tr.querySelector("td.col-cost");
  const tdReman=tr.querySelector("td.col-reman");

  const currentMonth=tdMonth.textContent.trim();
  const currentCost=tdCost.textContent.trim();
  const currentReman=tdReman.textContent.trim();

  const monthOptions=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    .map(m=>`<option value="${m}" ${m===currentMonth?"selected":""}>${m}</option>`).join("");

  tdMonth.innerHTML=`<select class="edit-month">${monthOptions}</select>`;
  tdCost.innerHTML=`<input type="text" class="edit-cost" value="${currentCost}" style="width:80px;text-align:right;">`;
  tdReman.innerHTML=`
    <select class="edit-reman">
      <option value="Reman" ${currentReman==="Reman"?"selected":""}>Reman</option>
      <option value="-" ${currentReman==="-"?"selected":""}>-</option>
    </select>`;

  const tdAction=tr.querySelector("td:last-child");
  tdAction.innerHTML=`
    <button class="action-btn save-btn" data-index="${index}">Save</button>
    <button class="action-btn cancel-btn">Cancel</button>`;
}

/* ===================== FILTER ===================== */
function filterData(){
  const room=document.getElementById("filter-room").value.toLowerCase();
  const order=document.getElementById("filter-order").value.toLowerCase();
  const cph=document.getElementById("filter-cph").value.toLowerCase();
  const mat=document.getElementById("filter-mat").value.toLowerCase();
  const section=document.getElementById("filter-section").value.toLowerCase();
  const month=document.getElementById("filter-month").value.toLowerCase();

  const filtered=mergedData.filter(r=>{
    if(room && !r.Room.toLowerCase().includes(room)) return false;
    if(order && !r.Order.toLowerCase().includes(order)) return false;
    if(cph && !r.CPH.toLowerCase().includes(cph)) return false;
    if(mat && !r.MAT.toLowerCase().includes(mat)) return false;
    if(section && !r.Section.toLowerCase().includes(section)) return false;
    if(month && r.Month.toLowerCase()!==month) return false;
    return true;
  });
  renderTable(filtered);
}

function resetFilters(){
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section","filter-month"].forEach(id=>document.getElementById(id).value="");
  renderTable(mergedData);
}

function updateMonthFilterOptions(){
  const sel=document.getElementById("filter-month");
  if(!sel) return;
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m))).sort();
  sel.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDER ===================== */
function addOrders(){
  const input=document.getElementById("add-order-input");
  const status=document.getElementById("add-order-status");
  const text=(input.value||"").trim();
  if(!text){ alert("Masukkan Order"); return; }
  const orders=text.split(/[\s,]+/).filter(Boolean);
  let added=0;
  orders.forEach(o=>{
    if(!mergedData.some(r=>r.Order===o)){
      mergedData.push({
        Room:"", "Order Type":"", Order:o, Description:"", "Created On":"", "User Status":"",
        MAT:"", CPH:"", Section:"", "Status Part":"", Aging:"", Month:"", Cost:"-", Reman:"",
        Include:"-", Exclude:"-", Planning:"", "Status AMT":""
      });
      added++;
    }
  });
  if(added){ saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan`; }
  else status.textContent="Order sudah ada";
  input.value="";
}

/* ===================== SAVE / LOAD JSON ===================== */
function saveToJSON(){
  if(!mergedData.length){ alert("Tidak ada data"); return; }
  const blob=new Blob([JSON.stringify({mergedData,timestamp:new Date().toISOString()},null,2)],{type:"application/json"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
}

async function loadFromJSON(file){
  try{
    const text=await file.text();
    const obj=JSON.parse(text);
    if(obj.mergedData && Array.isArray(obj.mergedData)){
      mergedData=obj.mergedData;
      renderTable(mergedData);
      updateMonthFilterOptions();
      alert("Data berhasil dimuat dari JSON");
    }else alert("File JSON tidak valid");
  }catch(e){ alert("Gagal baca JSON: "+e.message); }
}

/* ===================== USER EDITS ===================== */
function saveUserEdits(){
  try{
    const userEdits=mergedData.map(item=>({
      Order:item.Order, Room:item.Room, "Order Type":item["Order Type"], Description:item.Description,
      "Created On":item["Created On"], "User Status":item["User Status"], MAT:item.MAT, CPH:item.CPH,
      Section:item.Section, "Status Part":item["Status Part"], Aging:item.Aging, Month:item.Month,
      Cost:item.Cost, Reman:item.Reman, Include:item.Include, Exclude:item.Exclude, Planning:item.Planning,
      "Status AMT":item["Status AMT"]
    }));
    localStorage.setItem(UI_LS_KEY,JSON.stringify({userEdits}));
  }catch{}
}

/* ===================== CLEAR ===================== */
function clearAllData(){
  if(!confirm("Hapus semua data?")) return;
  iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[]; mergedData=[];
  renderTable([]); document.getElementById("upload-status").textContent="Data dihapus";
  updateMonthFilterOptions();
}

/* ===================== BUTTONS ===================== */
function setupButtons(){
  const uploadBtn=document.getElementById("upload-btn"); if(uploadBtn) uploadBtn.onclick=handleUpload;
  const clearBtn=document.getElementById("clear-files-btn"); if(clearBtn) clearBtn.onclick=clearAllData;
  const refreshBtn=document.getElementById("refresh-btn"); if(refreshBtn) refreshBtn.onclick=()=>{ mergeData(); renderTable(mergedData); };
  const filterBtn=document.getElementById("filter-btn"); if(filterBtn) filterBtn.onclick=filterData;
  const resetBtn=document.getElementById("reset-btn"); if(resetBtn) resetBtn.onclick=resetFilters;
  const saveBtn=document.getElementById("save-btn"); if(saveBtn) saveBtn.onclick=saveToJSON;
  const loadBtn=document.getElementById("load-btn"); if(loadBtn){
    loadBtn.onclick=()=>{
      const input=document.createElement("input"); input.type="file"; input.accept="application/json";
      input.onchange=()=>{ if(input.files.length) loadFromJSON(input.files[0]); };
      input.click();
    };
  }
  const addOrderBtn=document.getElementById("add-order-btn"); if(addOrderBtn) addOrderBtn.onclick=addOrders;
}
