/****************************************************
 * Ndarboe.net - FULL script.js REVISI
 * --------------------------------------------------
 * - Merge 6 sumber: IW39, SUM57, Planning, Data1, Data2, Budget
 * - Render & format table, termasuk pewarnaan otomatis
 * - Format tanggal dd-MMM-yyyy
 * - Format angka (Cost, Include, Exclude): 1 desimal, rata kanan
 * - Color rules: User Status, Status Part, Status AMT, Aging
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
  renderTable([]); // kosong dulu
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
        if (sec.id === menuId) sec.classList.add("active");
        else sec.classList.remove("active");
      });
    });
  });
}

/* ===================== HELPERS: DATE & NUMBER ===================== */
function toDateObj(anyDate) {
  if (anyDate == null || anyDate === "") return null;
  if (typeof anyDate === "number") { // Excel serial
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
      let p1 = parseInt(parts[0],10), p2 = parseInt(parts[1],10), p3 = parseInt(parts[2],10);
      let year, month, day, hour=0, min=0, sec=0;
      if (p1<=12 && p2<=31) { month=p1; day=p2; year=p3; } 
      else { day=p1; month=p2; year=p3; }
      if(parts.length>=5){ hour=parseInt(parts[3],10); min=parseInt(parts[4],10);
        if(parts.length>=6) sec=parseInt(parts[5],10)||0;
        if(ampm){ if(/pm/i.test(ampm)&&hour<12) hour+=12; if(/am/i.test(ampm)&&hour===12) hour=0; }
      }
      const d = new Date(year,(month||1)-1,day||1,hour,min,sec);
      if(!isNaN(d)) return d;
    }
  }
  return null;
}
function formatDateDDMMMYYYY(input){
  const d = toDateObj(input);
  if(!d) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
}
function formatDateISO(input){ const d=toDateObj(input); if(!d) return ""; return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
function formatNumber1(val){ const n=parseFloat(val); return isNaN(n)?"-":n.toFixed(1); }

/* ===================== UPLOAD & PARSE EXCEL ===================== */
async function parseFile(file, jenis){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const data=new Uint8Array(e.target.result);
        const workbook=XLSX.read(data,{type:"array"});
        let sheetName = workbook.SheetNames.includes(jenis)? jenis: workbook.SheetNames[0];
        const ws=workbook.Sheets[sheetName];
        if(!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan.`);
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
  const file=fileInput.files[0];
  const jenis=fileSelect.value;
  status.textContent=`Memproses file ${file.name} sebagai ${jenis}...`;
  try{
    const json=await parseFile(file,jenis);
    switch(jenis){
      case "IW39": iw39Data=json; break;
      case "SUM57": sum57Data=json; break;
      case "Planning": planningData=json; break;
      case "Data1": data1Data=json; break;
      case "Data2": data2Data=json; break;
      case "Budget": budgetData=json; break;
    }
    status.textContent=`File ${file.name} berhasil diupload sebagai ${jenis} (rows: ${json.length}).`;
    fileInput.value="";
  }catch(e){ status.textContent=`Error saat membaca file: ${e.message}`; }
}

/* ===================== MERGE DATA ===================== */
function mergeData(){
  if(!iw39Data.length){ alert("Upload data IW39 dulu sebelum refresh."); return; }
  mergedData = iw39Data.map(row=>({
    Room: row.Room||"", "Order Type": row["Order Type"]||"", Order: row.Order||"", Description: row.Description||"",
    "Created On": row["Created On"]||"", "User Status": row["User Status"]||"", MAT: row.MAT||"", CPH:"", Section:"",
    "Status Part":"", Aging:"", Month: row.Month||"", Cost:"-", Reman: row.Reman||"", Include:"-", Exclude:"-", Planning:"", "Status AMT":""
  }));

  // CPH lookup
  mergedData.forEach(md=>{
    if(md.Description.trim().toUpperCase().startsWith("JR")) md.CPH="External Job";
    else { const d2=data2Data.find(d=>(d.MAT||"").trim()===md.MAT.trim()); md.CPH=d2?d2.CPH||"":"";
    }
  });
  // Section lookup
  mergedData.forEach(md=>{ const d1=data1Data.find(d=>(d.Room||"").trim()===md.Room.trim()); md.Section=d1?d1.Section||"":"";
  });
  // SUM57 lookup
  mergedData.forEach(md=>{ const s57=sum57Data.find(s=>(s.Order||"")===md.Order); if(s57){ md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||""; } });
  // Planning lookup
  mergedData.forEach(md=>{ const pl=planningData.find(p=>(p.Order||"")===md.Order); if(pl){ md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||""; } });
  // Hitung Cost/Include/Exclude
  mergedData.forEach(md=>{
    const src=iw39Data.find(i=>(i.Order||"")===md.Order); if(!src) return;
    const plan=parseFloat((src["Total sum (plan)"]||"").replace(/,/g,""))||0;
    const actual=parseFloat((src["Total sum (actual)"]||"").replace(/,/g,""))||0;
    let cost=(plan-actual)/16500;
    if(!isFinite(cost)||cost<0){ md.Cost="-"; md.Include="-"; md.Exclude="-"; }
    else{ md.Cost=formatNumber1(cost); const isReman=(md.Reman||"").toLowerCase().includes("reman"); const includeNum=isReman?cost*0.25:cost; md.Include=formatNumber1(includeNum); md.Exclude=(md["Order Type"]==="PM38")?"-":md.Include; }
  });
  // Restore user edits
  try{ const raw=localStorage.getItem(UI_LS_KEY); if(raw){ const saved=JSON.parse(raw); if(saved && Array.isArray(saved.userEdits)) saved.userEdits.forEach(edit=>{ const idx=mergedData.findIndex(r=>r.Order===edit.Order); if(idx!==-1) mergedData[idx]={...mergedData[idx],...edit}; }); } }catch{}
  updateMonthFilterOptions();
}

/* ===================== CELL COLORING ===================== */
function asColoredStatusUser(val){
  const v=(val||"").toString().toUpperCase(); let bg="",fg="";
  if(v==="OUTS"){ bg="#ffeb3b"; fg="#000"; }
  else if(v==="RELE"){ bg="#2e7d32"; fg="#fff"; }
  else if(v==="PROG"){ bg="#1976d2"; fg="#fff"; }
  else if(v==="COMP"){ bg="#d32f2f"; fg="#fff"; }
  else return safe(val);
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${safe(val)}</span>`;
}
function asColoredStatusPart(val){
  const s=(val||"").toString().toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}
function asColoredStatusAMT(val){
  const v=(val||"").toString().toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#bdbdbd;color:#000;">${safe(val)}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${safe(val)}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${safe(val)}</span>`;
  return safe(val);
}
function asColoredAging(val){
  const n=parseInt(val); if(isNaN(n)) return safe(val);
  if(n>=30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${val}</span>`;
  if(n>0) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  return safe(val);
}

/* ===================== SAFE ESCAPE ===================== */
function safe(str){ return String(str||"").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender = mergedData) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) {
    console.warn("Tabel #output-table tidak ditemukan.");
    return;
  }

  tbody.innerHTML = ""; // reset tabel

  dataToRender.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index;

    const cols = [
      "Room","Order Type","Order","Description","Created On",
      "User Status","MAT","CPH","Section","Status Part","Aging",
      "Month","Cost","Reman","Include","Exclude","Planning","Status AMT"
    ];

    cols.forEach(col => {
      const td = document.createElement("td");
      let val = row[col] || "";

      if (col === "User Status") {
        td.innerHTML = asColoredStatusUser(val);
      } else if (col === "Status Part") {
        td.innerHTML = asColoredStatusPart(val);
      } else if (col === "Status AMT") {
        td.innerHTML = asColoredStatusAMT(val);
      } else if (col === "Aging") {
        td.innerHTML = asColoredAging(val);
      } else if (col === "Created On" || col === "Planning") {
        td.textContent = formatDateDDMMMYYYY(val);
      } else if (col === "Cost" || col === "Include" || col === "Exclude") {
        td.textContent = formatNumber1(val);
        td.style.textAlign = "right";
      } else {
        td.textContent = val;
      }

      tr.appendChild(td);
    });

    // Kolom Action
    const tdAct = document.createElement("td");
    tdAct.innerHTML = `
      <button class="action-btn edit-btn" data-order="${safe(row.Order)}">Edit</button>
      <button class="action-btn delete-btn" data-order="${safe(row.Order)}">Delete</button>
    `;
    tr.appendChild(tdAct);

    tbody.appendChild(tr);
  });

  attachTableEvents();
}

/* ===================== EDIT / DELETE ===================== */
function attachTableEvents(){
  document.querySelectorAll(".edit-btn").forEach(btn=>{ btn.onclick=()=>startEdit(btn.dataset.order); });
  document.querySelectorAll(".delete-btn").forEach(btn=>{ btn.onclick=()=>deleteOrder(btn.dataset.order); });
}
function startEdit(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  const tbody=document.querySelector("#output-table tbody"); const tr=tbody.children[idx]; const row=mergedData[idx];
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m))).sort();
  const monthOptions=`<option value="">--Select--</option>${months.map(m=>`<option value="${m}" ${m===row.Month?"selected":""}>${m}</option>`).join("")}`;
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
    <td><input type="text" value="${formatNumber1(row.Cost)}" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${safe(row.Reman)}" data-field="Reman"/></td>
    <td><input type="text" value="${formatNumber1(row.Include)}" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="text" value="${formatNumber1(row.Exclude)}" readonly style="text-align:right;background:#eee;"/></td>
    <td><input type="date" value="${formatDateISO(row.Planning)}" data-field="Planning"/></td>
    <td><input type="text" value="${safe(row["Status AMT"])}" data-field="Status AMT"/></td>
    <td><button class="action-btn save-btn" data-order="${safe(order)}">Save</button>
        <button class="action-btn cancel-btn" data-order="${safe(order)}">Cancel</button></td>
  `;
  tr.querySelector(".save-btn").onclick=()=>saveEdit(order);
  tr.querySelector(".cancel-btn").onclick=()=>renderTable(mergedData);
}
function saveEdit(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  const tr=document.querySelector("#output-table tbody").children[idx];
  const inputs=tr.querySelectorAll("input[data-field], select[data-field]");
  inputs.forEach(input=>{ mergedData[idx][input.dataset.field]=input.value; });
  saveUserEdits();
  mergeData(); renderTable(mergedData);
}
function deleteOrder(order){
  const idx=mergedData.findIndex(r=>r.Order===order); if(idx===-1) return;
  if(!confirm(`Hapus data order ${order}?`)) return;
  mergedData.splice(idx,1); saveUserEdits(); renderTable(mergedData);
}

/* ===================== FILTER ===================== */
function filterData(){
  const room=document.getElementById("filter-room").value.toLowerCase().trim();
  const order=document.getElementById("filter-order").value.toLowerCase().trim();
  const cph=document.getElementById("filter-cph").value.toLowerCase().trim();
  const mat=document.getElementById("filter-mat").value.toLowerCase().trim();
  const section=document.getElementById("filter-section").value.toLowerCase().trim();
  const month=document.getElementById("filter-month").value.toLowerCase().trim();
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
  const sel=document.getElementById("filter-month"); if(!sel) return;
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m))).sort();
  sel.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDERS ===================== */
function addOrders(){
  const input=document.getElementById("add-order-input"); const status=document.getElementById("add-order-status");
  const text=(input.value||"").trim(); if(!text){ alert("Masukkan Order terlebih dahulu."); return; }
  const orders=text.split(/[\s,]+/).filter(Boolean); let added=0;
  orders.forEach(o=>{ if(!mergedData.some(r=>r.Order===o)){ mergedData.push({ Room:"", "Order Type":"", Order:o, Description:"", "Created On":"", "User Status":"", MAT:"", CPH:"", Section:"", "Status Part":"", Aging:"", Month:"", Cost:"-", Reman:"", Include:"-", Exclude:"-", Planning:"", "Status AMT":"" }); added++; } });
  if(added){ saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan.`; } else status.textContent="Order sudah ada di data.";
  input.value="";
}

/* ===================== SAVE / LOAD JSON ===================== */
function saveToJSON(){
  if(!mergedData.length){ alert("Tidak ada data untuk disimpan."); return; }
  const blob=new Blob([JSON.stringify({ mergedData,timestamp:new Date().toISOString() },null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url; a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
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
  } catch(e){ alert("Gagal membaca file JSON: "+e.message); }
}

/* ===================== USER EDITS PERSISTENCE ===================== */
function saveUserEdits(){
  try{ const userEdits=mergedData.map(r=>({...r})); localStorage.setItem(UI_LS_KEY,JSON.stringify({userEdits})); } catch{}
}

/* ===================== CLEAR ALL ===================== */
function clearAllData(){
  if(!confirm("Yakin ingin menghapus semua data yang telah diupload?")) return;
  iw39Data=sum57Data=planningData=data1Data=data2Data=budgetData=mergedData=[];
  renderTable([]); document.getElementById("upload-status").textContent="Data dihapus.";
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
  const loadBtn=document.getElementById("load-btn"); if(loadBtn){ loadBtn.onclick=()=>{
    const input=document.createElement("input"); input.type="file"; input.accept="application/json";
    input.onchange=()=>{ if(input.files.length) loadFromJSON(input.files[0]); }; input.click();
  }; }
  const addOrderBtn=document.getElementById("add-order-btn"); if(addOrderBtn) addOrderBtn.onclick=addOrders;
}

