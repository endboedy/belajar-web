/****************************************************
 * Ndarboe.net - FULL script.js (revisi)
 * --------------------------------------------------
 * - Upload 6 sumber (IW39, SUM57, Planning, Budget, Data1, Data2)
 * - Merge & render Lembar Kerja
 * - Filter, Add Order, Edit/Save/Delete (3 kolom)
 * - Save/Load JSON
 * - Pewarnaan kolom status
 * - Format tanggal dd-MMM-yyyy (Created On, Planning)
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

  const tbody = document.querySelector("#output-table tbody");
  tbody.addEventListener("click", function(e) {
    const btn = e.target;
    const tr = btn.closest("tr");
    if (!tr) return;
    const index = btn.dataset.index;

    // Tombol Edit
    if (btn.classList.contains("edit-btn")) {
      activateEdit(tr, index);
    }

    // Tombol Delete
    else if (btn.classList.contains("delete-btn")) {
      if (confirm("Yakin mau hapus data ini?")) {
        mergedData.splice(index, 1);
        saveUserEdits();
        renderTable();
      }
    }

    // Tombol Save (3 kolom)
    else if (btn.classList.contains("save-btn")) {
      const tdMonth = tr.querySelector("td.col-month select");
      const tdCost  = tr.querySelector("td.col-cost input");
      const tdReman = tr.querySelector("td.col-reman select");

      mergedData[index]["Month"] = tdMonth.value;
      mergedData[index]["Cost"]  = tdCost.value;
      mergedData[index]["Reman"] = tdReman.value;

      saveUserEdits();
      renderTable();
    }

    // Tombol Cancel
    else if (btn.classList.contains("cancel-btn")) {
      renderTable();
    }
  });
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
    });
  });
}

/* ===================== HELPERS: DATE PARSING/FORMATTING ===================== */
function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;

  if (typeof anyDate === "number") {
    const dec = XLSX && XLSX.SSF && XLSX.SSF.parse_date_code
      ? XLSX.SSF.parse_date_code(anyDate)
      : null;
    if (dec && typeof dec === "object") {
      return new Date(dec.y, (dec.m || 1) - 1, dec.d || 1, dec.H || 0, dec.M || 0, dec.S || 0);
    }
  }

  if (anyDate instanceof Date && !isNaN(anyDate)) return anyDate;

  if (typeof anyDate === "string") {
    const s = anyDate.trim();
    if (!s) return null;
    const iso = new Date(s);
    if (!isNaN(iso)) return iso;

    const parts = s.match(/(\d{1,4})/g);
    if (parts && parts.length >= 3) {
      const ampm = /am|pm/i.test(s) ? s.match(/am|pm/i)[0] : "";
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      let year, month, day, hour=0, min=0, sec=0;
      if (p1 <= 12 && p2 <= 31) { month=p1; day=p2; year=p3; }
      else { day=p1; month=p2; year=p3; }
      if (parts.length >= 5) {
        hour=parseInt(parts[3],10); min=parseInt(parts[4],10);
        if (parts.length>=6) sec=parseInt(parts[5],10)||0;
        if (ampm) {
          if (/pm/i.test(ampm)&&hour<12) hour+=12;
          if (/am/i.test(ampm)&&hour===12) hour=0;
        }
      }
      const d=new Date(year,(month||1)-1,day||1,hour,min,sec);
      if (!isNaN(d)) return d;
    }
  }
  return null;
}

function formatDateDDMMMYYYY(input){
  if(!input) return "";
  let d = (typeof input==="number") ? excelDateToJS(input) : new Date(input);
  if(isNaN(d)){ const alt=new Date(String(input).replace(/\//g,"-")); d=isNaN(alt)?null:alt; }
  if(!d) return "";
  const months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

function formatDateISO(anyDate){
  const d=toDateObj(anyDate);
  if(!d||isNaN(d)) return "";
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

/* ===================== UPLOAD & PARSE EXCEL ===================== */
async function parseFile(file, jenis){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const data=new Uint8Array(e.target.result);
        const workbook=XLSX.read(data,{type:"array"});
        let sheetName=workbook.SheetNames.includes(jenis)?jenis:workbook.SheetNames[0];
        const ws=workbook.Sheets[sheetName];
        if(!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan`);
        const json=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});
        resolve(json);
      }catch(err){ reject(err); }
    };
    reader.onerror=err=>reject(err);
    reader.readAsArrayBuffer(file);
  });
}

async function handleUpload(){
  const fileSelect=document.getElementById("file-select");
  const fileInput=document.getElementById("file-input");
  const status=document.getElementById("upload-status");
  if(!fileInput.files.length){ alert("Pilih file terlebih dahulu."); return; }
  const file=fileInput.files[0], jenis=fileSelect.value;
  status.textContent=`Memproses file ${file.name} sebagai ${jenis}...`;
  try{
    const json=await parseFile(file, jenis);
    switch(jenis){
      case "IW39": iw39Data=json; break;
      case "SUM57": sum57Data=json; break;
      case "Planning": planningData=json; break;
      case "Data1": data1Data=json; break;
      case "Data2": data2Data=json; break;
      case "Budget": budgetData=json; break;
      default: break;
    }
    status.textContent=`File ${file.name} berhasil diupload sebagai ${jenis} (rows: ${json.length}).`;
    fileInput.value="";
  }catch(e){ status.textContent=`Error saat membaca file: ${e.message}`; }
}

/* ===================== MERGE ===================== */
function mergeData(){
  if(!iw39Data.length){ alert("Upload data IW39 dulu."); return; }

  mergedData = iw39Data.map(row=>({
    Room:(row.Room||"").toString(),
    "Order Type":(row["Order Type"]||"").toString(),
    Order:(row.Order||"").toString(),
    Description:(row.Description||"").toString(),
    "Created On":row["Created On"]||"",
    "User Status":(row["User Status"]||"").toString(),
    MAT:(row.MAT||"").toString(),
    CPH:"",
    Section:"",
    "Status Part":"",
    Aging:"",
    Month:(row.Month||"").toString(),
    Cost:"-",
    Reman:(row.Reman||"").toString(),
    Include:"-",
    Exclude:"-",
    Planning:"",
    "Status AMT":""
  }));

  mergedData.forEach(md=>{
    if((md.Description||"").trim().toUpperCase().startsWith("JR")) md.CPH="External Job";
    else { const d2=data2Data.find(d=>(d.MAT||"").toString().trim()===md.MAT.trim()); md.CPH=d2?d2.CPH||"":""; }
    const d1=data1Data.find(d=>(d.Room||"").toString().trim()===md.Room.trim()); md.Section=d1?d1.Section||"":""; 
    const s57=sum57Data.find(s=>(s.Order||"").toString()===md.Order); if(s57){ md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||""; }
    const pl=planningData.find(p=>(p.Order||"").toString()===md.Order); if(pl){ md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||""; }

    const src=iw39Data.find(i=>(i.Order||"").toString()===md.Order);
    if(src){
      const plan=parseFloat((src["Total sum (plan)"]||"").toString().replace(/,/g,""))||0;
      const actual=parseFloat((src["Total sum (actual)"]||"").toString().replace(/,/g,""))||0;
      let cost=(plan-actual)/16500;
      if(!isFinite(cost)||cost<0){ md.Cost="-"; md.Include="-"; md.Exclude="-"; }
      else {
        md.Cost=cost.toFixed(2);
        const isReman=(md.Reman||"").toLowerCase().includes("reman");
        md.Include=(isReman?cost*0.25:cost).toFixed(2);
        md.Exclude=md["Order Type"]==="PM38"?"-":md.Include;
      }
    }
  });

  // Restore edits
  try{ const raw=localStorage.getItem(UI_LS_KEY); if(raw){ const saved=JSON.parse(raw); saved.userEdits?.forEach(edit=>{ const idx=mergedData.findIndex(r=>r.Order===edit.Order); if(idx!==-1) mergedData[idx]={...mergedData[idx],...edit}; }); } }catch{}

  updateMonthFilterOptions();
}

/* ===================== ACTIVATE EDIT (3 kolom) ===================== */
function activateEdit(tr,index){
  const tdMonth=tr.querySelector("td.col-month");
  const tdCost=tr.querySelector("td.col-cost");
  const tdReman=tr.querySelector("td.col-reman");
  const tdAction=tr.querySelector("td:last-child");

  const currentMonth=tdMonth.textContent.trim();
  const currentCost=tdCost.textContent.trim();
  const currentReman=tdReman.textContent.trim();

  tdMonth.innerHTML=`<select class="edit-month">${
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    .map(m=>`<option value="${m}" ${m===currentMonth?"selected":""}>${m}</option>`).join("")
  }</select>`;
  tdCost.innerHTML=`<input type="text" class="edit-cost" value="${currentCost}" style="width:80px;text-align:right;">`;
  tdReman.innerHTML=`<select class="edit-reman">
    <option value="Reman" ${currentReman==="Reman"?"selected":""}>Reman</option>
    <option value="-" ${currentReman==="-"?"selected":""}>-</option>
  </select>`;
  tdAction.innerHTML=`<button class="action-btn save-btn" data-index="${index}">Save</button>
                      <button class="action-btn cancel-btn">Cancel</button>`;
}

/* ===================== CELL COLORING ===================== */
function asColoredStatusUser(val){
  const v=(val||"").toString().toUpperCase(); let bg="",fg="";
  if(v==="OUTS"){bg="#ffeb3b";fg="#000";}
  else if(v==="RELE"){bg="#2e7d32";fg="#fff";}
  else if(v==="PROG"){bg="#1976d2";fg="#fff";}
  else if(v==="COMP"){bg="#000";fg="#fff";}
  else return safe(val);
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${safe(val)}</span>`;
}
function asColoredStatusPart(val){
  const s=(val||"").toString().toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}
function asColoredStatusAMT(val){
  const v=(val||"").toString().toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#ffeb3b;color:#000;">${safe(val)}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

/* ===================== FILTERS ===================== */
function filterData(){
  const roomFilter=(document.getElementById("filter-room").value||"").trim().toLowerCase();
  const orderFilter=(document.getElementById("filter-order").value||"").trim().toLowerCase();
  const cphFilter=(document.getElementById("filter-cph").value||"").trim().toLowerCase();
  const matFilter=(document.getElementById("filter-mat").value||"").trim().toLowerCase();
  const sectionFilter=(document.getElementById("filter-section").value||"").trim().toLowerCase();
  const monthFilter=(document.getElementById("filter-month").value||"").trim().toLowerCase();

  const filtered=mergedData.filter(row=>{
    if(roomFilter && !String(row.Room).toLowerCase().includes(roomFilter)) return false;
    if(orderFilter && !String(row.Order).toLowerCase().includes(orderFilter)) return false;
    if(cphFilter && !String(row.CPH).toLowerCase().includes(cphFilter)) return false;
    if(matFilter && !String(row.MAT).toLowerCase().includes(matFilter)) return false;
    if(sectionFilter && !String(row.Section).toLowerCase().includes(sectionFilter)) return false;
    if(monthFilter && String(row.Month).toLowerCase()!==monthFilter) return false;
    return true;
  });
  renderTable(filtered);
}
function resetFilters(){
  ["filter-room","filter-order","filter-cph","filter-mat","filter-section","filter-month"].forEach(id=>{
    document.getElementById(id).value="";
  });
  renderTable(mergedData);
}
function updateMonthFilterOptions(){
  const monthSelect=document.getElementById("filter-month"); if(!monthSelect) return;
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m&&m.trim()!==""))).sort();
  monthSelect.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDERS ===================== */
function addOrders(){
  const input=document.getElementById("add-order-input");
  const status=document.getElementById("add-order-status");
  const text=(input.value||"").trim();
  if(!text){ alert("Masukkan Order terlebih dahulu."); return; }
  const orders=text.split(/[\s,]+/).filter(Boolean);
  let added=0;
  orders.forEach(o=>{
    if(!mergedData.some(r=>r.Order===o)){
      mergedData.push({
        Room:"", "Order Type":"", Order:o, Description:"", "Created On":"",
        "User Status":"", MAT:"", CPH:"", Section:"", "Status Part":"",
        Aging:"", Month:"", Cost:"-", Reman:"", Include:"-", Exclude:"-",
        Planning:"", "Status AMT":""
      });
      added++;
    }
  });
  if(added){ saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan.`; }
  else status.textContent="Order sudah ada di data.";
  input.value="";
}

/* ===================== JSON SAVE / LOAD ===================== */
function saveToJSON(){
  if(!mergedData.length){ alert("Tidak ada data untuk disimpan."); return; }
  const dataStr=JSON.stringify({mergedData,timestamp:new Date().toISOString()},null,2);
  const blob=new Blob([dataStr],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url;
  a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
async function loadFromJSON(file){
  try{
    const text=await file.text();
    const obj=JSON.parse(text);
    if(obj.mergedData && Array.isArray(obj.mergedData)){
      mergedData=obj.mergedData;
      renderTable(mergedData);
      updateMonthFilterOptions();
      alert("Data berhasil dimuat dari JSON.");
    }else alert("File JSON tidak valid.");
  }catch(e){ alert("Gagal membaca file JSON: "+e.message); }
}

/* ===================== SAVE USER EDITS ===================== */
function saveUserEdits(){
  try{
    const userEdits=mergedData.map(item=>({
      Order:item.Order, Room:item.Room, "Order Type":item["Order Type"], Description:item.Description,
      "Created On":item["Created On"], "User Status":item["User Status"], MAT:item.MAT, CPH:item.CPH,
      Section:item.Section, "Status Part":item["Status Part"], Aging:item.Aging, Month:item.Month,
      Cost:item.Cost, Reman:item.Reman, Include:item.Include, Exclude:item.Exclude,
      Planning:item.Planning, "Status AMT":item["Status AMT"]
    }));
    localStorage.setItem(UI_LS_KEY, JSON.stringify({userEdits}));
  }catch{}
}

/* ===================== CLEAR ALL ===================== */
function clearAllData(){
  if(!confirm("Yakin ingin menghapus semua data yang telah diupload?")) return;
  iw39Data=[]; sum57Data=[]; planningData=[]; data1Data=[]; data2Data=[]; budgetData=[];
  mergedData=[];
  renderTable([]);
  document.getElementById("upload-status").textContent="Data dihapus.";
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
      input.onchange=()=>{ if(input.files.length) loadFromJSON(input.files[0]); }; input.click();
    };
  }
  const addOrderBtn=document.getElementById("add-order-btn"); if(addOrderBtn) addOrderBtn.onclick=addOrders;
}

/* ===================== RENDER TABLE ===================== */
function renderTable(data){
  data=data||mergedData;
  const tbody=document.querySelector("#output-table tbody");
  if(!tbody) return;
  tbody.innerHTML="";
  data.forEach((row,i)=>{
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
      <td class="col-cost" style="text-align:right">${safe(row.Cost)}</td>
      <td class="col-reman">${safe(row.Reman)}</td>
      <td style="text-align:right">${safe(row.Include)}</td>
      <td style="text-align:right">${safe(row.Exclude)}</td>
      <td>${formatDateDDMMMYYYY(row.Planning)}</td>
      <td>${asColoredStatusAMT(row["Status AMT"])}</td>
      <td>
        <button class="action-btn edit-btn" data-index="${i}">Edit</button>
        <button class="action-btn delete-btn" data-index="${i}">Delete</button>
      </td>`;
    tbody.appendChild(tr);
  });
}

/* ===================== SAFE HTML ===================== */
function safe(str){ return str==null?"":String(str).replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
