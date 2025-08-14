/****************************************************
 * Ndarboe.net - FULL script.js (all menus, features)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
let iw39Data=[], sum57Data=[], planningData=[], data1Data=[], data2Data=[], budgetData=[], mergedData=[];
const UI_LS_KEY="ndarboe_ui_edits_v2";

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded",()=>{
  setupButtons();
  setupMenu();
  renderTable([]);
  updateMonthFilterOptions();
});

/* ===================== MENU HANDLER ===================== */
function setupMenu(){
  const menuItems=document.querySelectorAll(".sidebar .menu-item");
  const contentSections=document.querySelectorAll(".content-section");
  menuItems.forEach(item=>{
    item.addEventListener("click",()=>{
      menuItems.forEach(i=>i.classList.remove("active"));
      item.classList.add("active");
      const menuId=item.dataset.menu;
      contentSections.forEach(sec=>{
        if(sec.id===menuId) sec.classList.add("active");
        else sec.classList.remove("active");
      });
    });
  });
}

/* ===================== HELPERS ===================== */
function toDateObj(anyDate){
  if(anyDate==null||anyDate==="") return null;
  if(typeof anyDate==="number"){ // Excel serial
    const dec=XLSX?.SSF?.parse_date_code(anyDate);
    if(dec) return new Date(dec.y,(dec.m||1)-1,dec.d||1,dec.H||0,dec.M||0,dec.S||0);
  }
  if(anyDate instanceof Date&&!isNaN(anyDate)) return anyDate;
  if(typeof anyDate==="string"){
    const s=anyDate.trim(); if(!s) return null;
    const iso=new Date(s); if(!isNaN(iso)) return iso;
    const parts=s.match(/(\d{1,4})/g);
    if(parts&&parts.length>=3){
      const ampm=/am|pm/i.test(s)?s.match(/am|pm/i)[0]:"";
      const p1=parseInt(parts[0],10),p2=parseInt(parts[1],10),p3=parseInt(parts[2],10);
      let year,month,day,hour=0,min=0,sec=0;
      if(p1<=12&&p2<=31){month=p1;day=p2;year=p3;} else {day=p1;month=p2;year=p3;}
      if(parts.length>=5){hour=parseInt(parts[3],10);min=parseInt(parts[4],10); if(parts.length>=6) sec=parseInt(parts[5],10)||0;
        if(ampm){if(/pm/i.test(ampm)&&hour<12) hour+=12; if(/am/i.test(ampm)&&hour===12) hour=0;}
      }
      const d=new Date(year,(month||1)-1,day||1,hour,min,sec);
      if(!isNaN(d)) return d;
    }
  }
  return null;
}
function formatDateDDMMMYYYY(input){
  if(input==null||input==="") return "";
  let d=(typeof input==="number")?excelDateToJS(input):new Date(input);
  if(isNaN(d)) return "";
  const months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
}
function formatDateISO(input){
  const d=toDateObj(input); if(!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function formatNumber1(val){
  let num=parseFloat(String(val).replace(/,/g,""));
  if(isNaN(num)) return val;
  return num.toFixed(1);
}

/* ===================== CELL COLORING ===================== */
function asColoredStatusUser(val){
  const v=(val||"").toString().toUpperCase(); let bg="",fg="";
  if(v==="OUTS"){bg="#ffeb3b";fg="#000";}
  else if(v==="RELE"){bg="#2e7d32";fg="#fff";}
  else if(v==="PROG"){bg="#1976d2";fg="#fff";}
  else if(v==="COMP"){bg="#d32f2f";fg="#fff";}
  else return val;
  return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:${bg};color:${fg};">${val}</span>`;
}
function asColoredStatusPart(val){
  const s=(val||"").toLowerCase();
  if(s==="complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  if(s==="not complete") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${val}</span>`;
  return val;
}
function asColoredStatusAMT(val){
  const v=(val||"").toUpperCase();
  if(v==="O") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#ffeb3b;color:#000;">${val}</span>`;
  if(v==="IP") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#1976d2;color:#fff;">${val}</span>`;
  if(v==="YTS") return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  return val;
}
function asColoredAging(val){
  const num=parseFloat(val); if(isNaN(num)) return val;
  if(num>=1&&num<30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#2e7d32;color:#fff;">${val}</span>`;
  if(num>=30) return `<span style="display:inline-block;padding:2px 6px;border-radius:4px;background:#d32f2f;color:#fff;">${val}</span>`;
  return val;
}

/* ===================== UPLOAD & PARSE ===================== */
async function handleUpload(){
  const fileInput=document.getElementById("file-input");
  const fileSelect=document.getElementById("file-select");
  const status=document.getElementById("upload-status");
  if(!fileInput.files.length){alert("Pilih file!");return;}
  const file=fileInput.files[0],jenis=fileSelect.value;
  status.textContent=`Memproses ${file.name} sebagai ${jenis}...`;
  try{
    const json=await parseFile(file,jenis);
    if(jenis==="IW39") iw39Data=json;
    else if(jenis==="SUM57") sum57Data=json;
    else if(jenis==="Planning") planningData=json;
    else if(jenis==="Data1") data1Data=json;
    else if(jenis==="Data2") data2Data=json;
    else if(jenis==="Budget") budgetData=json;
    status.textContent=`File ${file.name} berhasil diupload (${json.length} rows).`;
    fileInput.value="";
  }catch(e){status.textContent=`Error: ${e.message}`;}
}
async function parseFile(file,jenis){return new Promise((resolve,reject)=>{
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const data=new Uint8Array(e.target.result);
      const workbook=XLSX.read(data,{type:"array"});
      let sheetName="";
      if(workbook.SheetNames.includes(jenis)) sheetName=jenis;
      else sheetName=workbook.SheetNames[0];
      const ws=workbook.Sheets[sheetName]; if(!ws) throw new Error(`Sheet "${sheetName}" tidak ditemukan`);
      const json=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});
      resolve(json);
    }catch(err){reject(err);}
  };
  reader.onerror=err=>reject(err);
  reader.readAsArrayBuffer(file);
});}

/* ===================== MERGE DATA ===================== */
function mergeData(){
  if(!iw39Data.length){alert("Upload IW39 dulu"); return;}
  mergedData=iw39Data.map(row=>({
    Room:row.Room||"", "Order Type":row["Order Type"]||"", Order:row.Order||"", Description:row.Description||"",
    "Created On":row["Created On"]||"", "User Status":row["User Status"]||"", MAT:row.MAT||"", CPH:"", Section:"", 
    "Status Part":"", Aging:"", Month:row.Month||"", Cost:"-", Reman:row.Reman||"", Include:"-", Exclude:"-", 
    Planning:"", "Status AMT":""
  }));
  mergedData.forEach(md=>{
    if((md.Description||"").toUpperCase().startsWith("JR")) md.CPH="External Job";
    else{
      const d2=data2Data.find(d=>(d.MAT||"").trim()===md.MAT.trim());
      md.CPH=d2?d2.CPH||"":"";
    }
    const d1=data1Data.find(d=>(d.Room||"").trim()===md.Room.trim());
    md.Section=d1?d1.Section||"":"";
    const s57=sum57Data.find(s=>(s.Order||"")==md.Order);
    if(s57){md.Aging=s57.Aging||""; md["Status Part"]=s57["Part Complete"]||"";}
    const pl=planningData.find(p=>(p.Order||"")==md.Order);
    if(pl){md.Planning=pl["Event Start"]||""; md["Status AMT"]=pl.Status||"";}
    const src=iw39Data.find(i=>(i.Order||"")==md.Order);
    if(src){
      const plan=parseFloat((src["Total sum (plan)"]||"").replace(/,/g,""))||0;
      const actual=parseFloat((src["Total sum (actual)"]||"").replace(/,/g,""))||0;
      let cost=(plan-actual)/16500;
      if(!isFinite(cost)||cost<0){md.Cost="-";md.Include="-";md.Exclude="-";}
      else{
        md.Cost=cost.toFixed(1);
        const isReman=(md.Reman||"").toLowerCase().includes("reman");
        const inc=isReman?cost*0.25:cost;
        md.Include=inc.toFixed(1); md.Exclude=md.Include;
      }
    }
  });
  try{
    const raw=localStorage.getItem(UI_LS_KEY);
    if(raw){const saved=JSON.parse(raw);
      saved.userEdits?.forEach(edit=>{
        const idx=mergedData.findIndex(r=>r.Order===edit.Order);
        if(idx!==-1) mergedData[idx]={...mergedData[idx],...edit};
      });
    }
  }catch{}
  updateMonthFilterOptions();
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender = mergedData) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) { console.warn("Tabel #output-table tidak ditemukan."); return; }
  tbody.innerHTML = "";

  dataToRender.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index; // simpan index row

    const cols = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];

    cols.forEach(col => {
      const td = document.createElement("td");
      let val = row[col] || "";

      // Pewarnaan
      if(col === "User Status") td.innerHTML = asColoredStatusUser(val);
      else if(col === "Status Part") td.innerHTML = asColoredStatusPart(val);
      else if(col === "Status AMT") td.innerHTML = asColoredStatusAMT(val);
      else if(col === "Aging") td.innerHTML = asColoredAging(val);

      // Tanggal
      else if(col === "Created On" || col === "Planning") td.textContent = formatDateDDMMMYYYY(val);

      // Angka / cost / include / exclude
      else if(col === "Cost" || col === "Include" || col === "Exclude") {
        td.textContent = formatNumber1(val);
        td.style.textAlign = "right";
      }
      else td.textContent = val;

      tr.appendChild(td);
    });

    // Action
    const tdAct = document.createElement("td");
    tdAct.innerHTML = `
      <button class="action-btn edit-btn" data-index="${index}">Edit</button>
      <button class="action-btn delete-btn" data-index="${index}">Delete</button>
    `;
    tr.appendChild(tdAct);

    tbody.appendChild(tr);
  });

  attachTableEvents(); // pastikan tombol berfungsi
}

/* ===================== START EDIT (HANYA Month, Cost, Reman) ===================== */
function startEdit(index) {
  const row = mergedData[index];
  if (!row) return;

  const tbody = document.querySelector("#output-table tbody");
function startEdit(index) {
  const row = mergedData[index];
  if (!row) return;

  const tbody = document.querySelector("#output-table tbody");
  const tr = tbody.querySelector(`tr[data-index='${index}']`);
  if (!tr) return;

  const months = Array.from(new Set(mergedData.map(d => d.Month).filter(m => m && m.trim() !== ""))).sort();
  const monthOptions = [`<option value="">--Select Month--</option>`].concat(months.map(m => `<option value="${m}">${m}</option>`)).join("");

  tr.innerHTML = `
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
    <td>${asColoredAging(row.Aging)}</td>
    <td><select data-field="Month">${monthOptions}</select></td>
    <td><input type="text" value="${safe(row.Cost)}" data-field="Cost" style="text-align:right;"/></td>
    <td>
      <select data-field="Reman">
        <option value="Reman" ${row.Reman==="Reman"?"selected":""}>Reman</option>
        <option value="-" ${row.Reman==="-"?"selected":""}>-</option>
      </select>
    </td>
    <td>${formatNumber1(row.Include)}</td>
    <td>${formatNumber1(row.Exclude)}</td>
    <td>${formatDateDDMMMYYYY(row.Planning)}</td>
    <td>${asColoredStatusAMT(row["Status AMT"])}</td>
    <td>
      <button class="action-btn save-btn" data-index="${index}">Save</button>
      <button class="action-btn cancel-btn" data-index="${index}">Cancel</button>
    </td>
  `;

  // set selected month
  tr.querySelector("select[data-field='Month']").value = row.Month || "";

  // tombol save/cancel
  tr.querySelector(".save-btn").onclick = () => saveEdit(index);
  tr.querySelector(".cancel-btn").onclick = () => cancelEdit();
}

/* ===================== SAVE EDIT ===================== */
function saveEdit(index) {
  const tr = document.querySelector(`#output-table tbody tr[data-index='${index}']`);
  if (!tr) return;

  const month = tr.querySelector("select[data-field='Month']").value;
  const cost = tr.querySelector("input[data-field='Cost']").value;
  const reman = tr.querySelector("select[data-field='Reman']").value;

  // update data
  mergedData[index].Month = month;
  mergedData[index].Cost = cost;
  mergedData[index].Reman = reman;

  // render ulang table
  renderTable();
}

/* ===================== CANCEL EDIT ===================== */
function cancelEdit() {
  renderTable();
}

/* ===================== ATTACH TABLE EVENTS ===================== */
function attachTableEvents() {
  const editButtons = document.querySelectorAll(".edit-btn");
  editButtons.forEach(btn => {
    btn.addEventListener("click", () => startEdit(parseInt(btn.dataset.index)));
  });

  const deleteButtons = document.querySelectorAll(".delete-btn");
  deleteButtons.forEach(btn => {
    btn.addEventListener("click", () => deleteRow(parseInt(btn.dataset.index)));
  });
}


/* ===================== FILTERS ===================== */
function filterData(){
  const room=document.getElementById("filter-room").value.trim().toLowerCase();
  const order=document.getElementById("filter-order").value.trim().toLowerCase();
  const cph=document.getElementById("filter-cph").value.trim().toLowerCase();
  const mat=document.getElementById("filter-mat").value.trim().toLowerCase();
  const section=document.getElementById("filter-section").value.trim().toLowerCase();
  const month=document.getElementById("filter-month").value.trim().toLowerCase();
  const filtered=mergedData.filter(row=>{
    if(room&&!(row.Room||"").toLowerCase().includes(room)) return false;
    if(order&&!(row.Order||"").toLowerCase().includes(order)) return false;
    if(cph&&!(row.CPH||"").toLowerCase().includes(cph)) return false;
    if(mat&&!(row.MAT||"").toLowerCase().includes(mat)) return false;
    if(section&&!(row.Section||"").toLowerCase().includes(section)) return false;
    if(month&&(row.Month||"").toLowerCase()!==month) return false;
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
  const months=Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m&&m.trim()!=""))).sort();
  sel.innerHTML=`<option value="">-- All --</option>`+months.map(m=>`<option value="${m.toLowerCase()}">${m}</option>`).join("");
}

/* ===================== ADD ORDERS ===================== */
function addOrders(){
  const input=document.getElementById("add-order-input");
  const status=document.getElementById("add-order-status");
  const orders=(input.value||"").trim().split(/[\s,]+/).filter(Boolean);
  let added=0; orders.forEach(o=>{
    if(!mergedData.some(r=>r.Order===o)){
      mergedData.push({Room:"", "Order Type":"", Order:o, Description:"", "Created On":"", "User Status":"", MAT:"", CPH:"", Section:"", "Status Part":"", Aging:"", Month:"", Cost:"-", Reman:"", Include:"-", Exclude:"-", Planning:"", "Status AMT":""});
      added++;
    }
  });
  if(added) { saveUserEdits(); renderTable(mergedData); status.textContent=`${added} Order berhasil ditambahkan.`; }
  else status.textContent="Order sudah ada di data.";
  input.value="";
}

/* ===================== SAVE / LOAD JSON ===================== */
function saveToJSON(){
  if(!mergedData.length){alert("Tidak ada data"); return;}
  const dataStr=JSON.stringify({mergedData,timestamp:new Date().toISOString()},null,2);
  const blob=new Blob([dataStr],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url;
  a.download=`ndarboe_data_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}
async function loadFromJSON(file){
  try{
    const text=await file.text(); const obj=JSON.parse(text);
    if(obj.mergedData&&Array.isArray(obj.mergedData)){mergedData=obj.mergedData; renderTable(mergedData); updateMonthFilterOptions(); alert("Data berhasil dimuat");}
    else alert("File JSON tidak valid");
  }catch(e){alert("Gagal membaca JSON: "+e.message);}
}

/* ===================== PERSISTENCE ===================== */
function saveUserEdits(){
  try{localStorage.setItem(UI_LS_KEY,JSON.stringify({userEdits:mergedData.map(r=>({...r}))}));}catch{}
}

/* ===================== CLEAR ALL ===================== */
function clearAllData(){
  if(!confirm("Yakin ingin hapus semua data?")) return;
  [iw39Data,sum57Data,planningData,data1Data,data2Data,budgetData,mergedData]=[[],[],[],[],[],[],[]];
  renderTable([]); document.getElementById("upload-status").textContent="Data dihapus";
  updateMonthFilterOptions();
}

/* ===================== BUTTON SETUP ===================== */
function setupButtons(){
  const ids={upload:"upload-btn",clear:"clear-files-btn",refresh:"refresh-btn",filter:"filter-btn",reset:"reset-btn",save:"save-btn",load:"load-btn",add:"add-order-btn"};
  if(document.getElementById(ids.upload)) document.getElementById(ids.upload).onclick=handleUpload;
  if(document.getElementById(ids.clear)) document.getElementById(ids.clear).onclick=clearAllData;
  if(document.getElementById(ids.refresh)) document.getElementById(ids.refresh).onclick=()=>{mergeData(); renderTable(mergedData);};
  if(document.getElementById(ids.filter)) document.getElementById(ids.filter).onclick=filterData;
  if(document.getElementById(ids.reset)) document.getElementById(ids.reset).onclick=resetFilters;
  if(document.getElementById(ids.save)) document.getElementById(ids.save).onclick=saveToJSON;
  if(document.getElementById(ids.load)) document.getElementById(ids.load).onclick=()=>{
    const input=document.createElement("input"); input.type="file"; input.accept="application/json";
    input.onchange=()=>{if(input.files.length) loadFromJSON(input.files[0]);}; input.click();
  };
  if(document.getElementById(ids.add)) document.getElementById(ids.add).onclick=addOrders;
}




