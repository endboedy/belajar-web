/****************************************************
 * Ndarboe.net - FULL script.js (revisi)
 * --------------------------------------------------
 * - Merge & render Lembar Kerja
 * - Filter, Add Order, Edit/Save/Delete
 * - Pewarnaan kolom status
 * - Format tanggal dd-MMM-yyyy
 * - Format angka dolar 1 decimal
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

/* ===================== HELPERS ===================== */
function safe(val) { return val == null ? "" : val; }

function excelDateToJS(serial) {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;
  if (typeof anyDate === "number") return excelDateToJS(anyDate);
  if (anyDate instanceof Date) return anyDate;
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
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

function formatNumber1(val) {
  let num = parseFloat(String(val).replace(/,/g,""));
  if (isNaN(num)) return val;
  return num.toFixed(1);
}

/* ===================== CELL COLORING ===================== */
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

function asColoredStatusPart(val) {
  const s = (val || "").toString().toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

function asColoredStatusAMT(val) {
  const v = (val || "").toString().toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#bdbdbd;color:#000;">${safe(val)}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}

function asColoredAging(val){
  const n=parseInt(val,10);
  if(isNaN(n)) return val;
  if(n>=1 && n<=30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  if(n>30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${val}</span>`;
  return val;
}

/* ===================== UPLOAD ===================== */
async function parseFile(file, sheetName) {
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const data=new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:"array"});
        const wsName=wb.SheetNames.includes(sheetName)?sheetName:wb.SheetNames[0];
        const ws=wb.Sheets[wsName];
        if(!ws) throw new Error("Sheet tidak ditemukan");
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
  if(!fileInput.files.length){ alert("Pilih file terlebih dahulu"); return; }
  const file=fileInput.files[0];
  const jenis=fileSelect.value;
  status.textContent=`Memproses ${file.name} sebagai ${jenis}...`;
  try{
    const data=await parseFile(file,jenis);
    switch(jenis){
      case "IW39": iw39Data=data; break;
      case "SUM57": sum57Data=data; break;
      case "Planning": planningData=data; break;
      case "Data1": data1Data=data; break;
      case "Data2": data2Data=data; break;
      case "Budget": budgetData=data; break;
      default: break;
    }
    status.textContent=`File ${file.name} berhasil diupload (${data.length} baris)`;
    fileInput.value="";
  }catch(e){
    status.textContent=`Error: ${e.message}`;
  }
}

/* ===================== MERGE ===================== */
function mergeData(){
  if(!iw39Data.length){ alert("Upload IW39 dulu"); return; }
  mergedData=iw39Data.map(row=>({
    Room:row.Room||"",
    "Order Type":row["Order Type"]||"",
    Order:row.Order||"",
    Description:row.Description||"",
    "Created On":row["Created On"]||"",
    "User Status":row["User Status"]||"",
    MAT:row.MAT||"",
    CPH:"",
    Section:"",
    "Status Part":"",
    Aging:"",
    Month:row.Month||"",
    Cost:"-",
    Reman:row.Reman||"",
    Include:"-",
    Exclude:"-",
    Planning:"",
    "Status AMT":""
  }));
  // CPH via Data2
  mergedData.forEach(md=>{
    if((md.Description||"").trim().toUpperCase().startsWith("JR")) md.CPH="External Job";
    else{
      const d2=data2Data.find(d=>(d.MAT||"").toString().trim()===md.MAT.trim());
      md.CPH=d2?d2.CPH||"":"";
    }
  });
  // Section via Data1
  mergedData.forEach(md=>{
    const d1=data1Data.find(d=>(d.Room||"").toString().trim()===md.Room.trim());
    md.Section=d1?d1.Section||"":"";
  });
  // SUM57
  mergedData.forEach(md=>{
    const s57=sum57Data.find(s=>(s.Order||"").toString()===md.Order);
    if(s57){ md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||""; }
  });
  // Planning
  mergedData.forEach(md=>{
    const pl=planningData.find(p=>(p.Order||"").toString()===md.Order);
    if(pl){ md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||""; }
  });
  // Cost/Include/Exclude
  mergedData.forEach(md=>{
    const src=iw39Data.find(i=>(i.Order||"").toString()===md.Order);
    if(!src) return;
    const plan=parseFloat((src["Total sum (plan)"]||"").toString().replace(/,/g,""))||0;
    const actual=parseFloat((src["Total sum (actual)"]||"").toString().replace(/,/g,""))||0;
    let cost=(plan-actual)/16500;
    if(!isFinite(cost)||cost<0){ md.Cost="-"; md.Include="-"; md.Exclude="-"; }
    else{
      md.Cost=formatNumber1(cost);
      const isReman=(md.Reman||"").toLowerCase().includes("reman");
      const includeNum=isReman?cost*0.25:cost;
      md.Include=formatNumber1(includeNum);
      md.Exclude=md["Order Type"]==="PM38"? "-":md.Include;
    }
  });
  // Restore user edits
  try{
    const raw=localStorage.getItem(UI_LS_KEY);
    if(raw){
      const saved=JSON.parse(raw);
      if(saved && Array.isArray(saved.userEdits)){
        saved.userEdits.forEach(edit=>{
          const idx=mergedData.findIndex(r=>r.Order===edit.Order);
          if(idx!==-1) mergedData[idx]={...mergedData[idx],...edit};
        });
      }
    }
  }catch{}
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender=mergedData){
  const tbody=document.querySelector("#output-table tbody");
  if(!tbody){ console.warn("Tabel #output-table tidak ditemukan."); return; }
  tbody.innerHTML="";
  dataToRender.forEach((row,index)=>{
    const tr=document.createElement("tr");
    tr.dataset.index=index;
    const cols=["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
    cols.forEach(col=>{
      const td=document.createElement("td");
      let val=row[col]||"";
      if(col==="User Status") td.innerHTML=asColoredStatusUser(val);
      else if(col==="Status Part") td.innerHTML=asColoredStatusPart(val);
      else if(col==="Status AMT") td.innerHTML=asColoredStatusAMT(val);
      else if(col==="Aging") td.innerHTML=asColoredAging(val);
      else if(col==="Created On"||col==="Planning") td.textContent=formatDateDDMMMYYYY(val);
      else if(col==="Cost"||col==="Include"||col==="Exclude"){ td.textContent=formatNumber1(val); td.style.textAlign="right"; }
      else td.textContent=val;
      tr.appendChild(td);
    });
    // Action
    const tdAct=document.createElement("td");
    tdAct.innerHTML=`<button class="action-btn edit-btn" data-order="${safe(row.Order)}">Edit</button>
                      <button class="action-btn delete-btn" data-order="${safe(row.Order)}">Delete</button>`;
    tr.appendChild(tdAct);
    tbody.appendChild(tr);
  });
  attachTableEvents();
}

/* ===================== EDIT / DELETE ===================== */
function attachTableEvents(){
  document.querySelectorAll(".edit-btn").forEach(btn=>btn.onclick=()=>startEdit(btn.dataset.order));
  document.querySelectorAll(".delete-btn").forEach(btn=>btn.onclick=()=>deleteOrder(btn.dataset.order));
}

function startEdit(order){
  const rowIndex=mergedData.findIndex(r=>r.Order===order);
  if(rowIndex===-1) return;
  const tbody=document.querySelector("#output-table tbody");
  const tr=tbody.children[rowIndex];
  const row=mergedData[rowIndex];
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m&&m.trim()!==""))).sort();
  const monthOptions=[`<option value="">--Select Month--</option>`,...months.map(m=>`<option value="${m}">${m}</option>`)].join("");
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
  const rowIndex=mergedData.findIndex(r=>r.Order===order);
  if(rowIndex===-1) return;
  const tbody=document.querySelector("#output-table tbody");
  const tr=tbody.children[rowIndex];
  const inputs=tr.querySelectorAll("input[data-field],select[data-field]");
  inputs.forEach(input=>{
    const field=input.dataset.field;
    mergedData[rowIndex][field]=input.value;
  });
  saveUserEdits();
  mergeData();
  renderTable(mergedData);
}

function deleteOrder(order){
  const idx=mergedData.findIndex(r=>r.Order===order);
  if(idx===-1) return;
  if(!confirm(`Hapus data order ${order} ?`)) return;
  mergedData.splice(idx,1);
  saveUserEdits();
  renderTable(mergedData);
}

/* ===================== FILTER ===================== */
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
    const el=document.getElementById(id); if(el) el.value="";
  });
  renderTable(mergedData);
}

function updateMonthFilterOptions(){
  const monthSelect=document.getElementById("filter-month");
  if(!monthSelect) return;
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m&&m.trim()!==""))).sort();
  monthSelect.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDER ===================== */
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
        Aging:"", Month:"", Cost:"-", Reman:"", Include:"-", Exclude:"-", Planning:"", "Status AMT":""
      });
      added++;
    }
  });
  if(added){ saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan.`; }
  else status.textContent="Order sudah ada di data.";
  input.value="";
}

/* ===================== SAVE / LOAD ===================== */
function saveToJSON(){
  if(!mergedData.length){ alert("Tidak ada data untuk disimpan."); return; }
  const dataStr=JSON.stringify({mergedData,timestamp:new Date().toISOString()},null,2);
  const blob=new Blob([dataStr],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;
  a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
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
    } else alert("File JSON tidak valid.");
  }catch(e){ alert("Gagal membaca file JSON: "+e.message); }
}

/* ===================== USER EDITS ===================== */
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
      input.type="file";
      input.accept="application/json";
      input.onchange=()=>{ if(input.files.length) loadFromJSON(input.files[0]); };
      input.click();
    };
  }

  const addOrderBtn=document.getElementById("add-order-btn");
  if(addOrderBtn) addOrderBtn.onclick=addOrders;
}

/* ===================== INIT ===================== */
document.addEventListener("DOMContentLoaded",()=>{
  setupButtons();
  updateMonthFilterOptions();
});
