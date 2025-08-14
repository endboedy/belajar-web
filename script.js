/****************************************************
 * Ndarboe.net - FULL script.js (all menus, features)
 ****************************************************/

/* ===================== GLOBAL STATE ===================== */
let iw39Data=[], sum57Data=[], planningData=[], data1Data=[], data2Data=[], budgetData=[], mergedData=[];
const UI_LS_KEY="ndarboe_ui_edits_v2";

/* ===================== HELPERS ===================== */
function safe(val){return val==null?"":val;}
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
      let year,month,day,hour=0,min=0,sec=0;
      const p1=parseInt(parts[0],10),p2=parseInt(parts[1],10),p3=parseInt(parts[2],10);
      if(p1<=12&&p2<=31){month=p1;day=p2;year=p3;} else {day=p1;month=p2;year=p3;}
      if(parts.length>=5){hour=parseInt(parts[3],10);min=parseInt(parts[4],10); if(parts.length>=6) sec=parseInt(parts[5],10)||0;}
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

/* ===================== ATTACH TABLE EVENTS ===================== */
function attachTableEvents() {
  document.querySelectorAll(".edit-btn").forEach(btn => {
    btn.addEventListener("click", () => startEdit(parseInt(btn.dataset.index)));
  });
  document.querySelectorAll(".delete-btn").forEach(btn => {
    btn.addEventListener("click", () => deleteRow(parseInt(btn.dataset.index)));
  });
}

/* ===================== RENDER TABLE ===================== */
function renderTable(dataToRender = mergedData) {
  const tbody = document.querySelector("#output-table tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  dataToRender.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index;

    const cols = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT"];
    cols.forEach(col => {
      const td = document.createElement("td");
      let val = row[col] || "";
      if(col==="User Status") td.innerHTML=asColoredStatusUser(val);
      else if(col==="Status Part") td.innerHTML=asColoredStatusPart(val);
      else if(col==="Status AMT") td.innerHTML=asColoredStatusAMT(val);
      else if(col==="Aging") td.innerHTML=asColoredAging(val);
      else if(col==="Created On"||col==="Planning") td.textContent=formatDateDDMMMYYYY(val);
      else if(col==="Cost"||col==="Include"||col==="Exclude") {td.textContent=formatNumber1(val); td.style.textAlign="right";}
      else td.textContent=val;
      tr.appendChild(td);
    });

    const tdAct=document.createElement("td");
    tdAct.innerHTML=`
      <button class="action-btn edit-btn" data-index="${index}">Edit</button>
      <button class="action-btn delete-btn" data-index="${index}">Delete</button>
    `;
    tr.appendChild(tdAct);

    tbody.appendChild(tr);
  });

  attachTableEvents();
}

/* ===================== START EDIT ===================== */
function startEdit(index){
  const row = mergedData[index];
  if(!row) return;
  const tbody=document.querySelector("#output-table tbody");
  const tr=tbody.querySelector(`tr[data-index='${index}']`);
  if(!tr) return;

  const months = Array.from(new Set(mergedData.map(d=>d.Month).filter(m=>m && m.trim()!=""))).sort();
  const monthOptions = [`<option value="">--Select Month--</option>`].concat(months.map(m=>`<option value="${m}">${m}</option>`)).join("");

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

  tr.querySelector("select[data-field='Month']").value=row.Month||"";
  tr.querySelector(".save-btn").onclick=()=>saveEdit(index);
  tr.querySelector(".cancel-btn").onclick=()=>cancelEdit();
}

/* ===================== SAVE EDIT ===================== */
function saveEdit(index){
  const tr=document.querySelector(`#output-table tbody tr[data-index='${index}']`);
  if(!tr) return;

  const month=tr.querySelector("select[data-field='Month']").value;
  const cost=tr.querySelector("input[data-field='Cost']").value;
  const reman=tr.querySelector("select[data-field='Reman']").value;

  mergedData[index].Month=month;
  mergedData[index].Cost=cost;
  mergedData[index].Reman=reman;

  renderTable();
}

/* ===================== CANCEL EDIT ===================== */
function cancelEdit(){renderTable();}

/* ===================== DOM READY ===================== */
document.addEventListener("DOMContentLoaded",()=>{
  setupButtons();
  setupMenu();
  renderTable([]); // aman karena attachTableEvents sudah ada
  updateMonthFilterOptions();
});
/* ===================== BUTTON SETUP ===================== */
function setupButtons(){
  const ids={
    upload:"upload-btn",
    clear:"clear-files-btn",
    refresh:"refresh-btn",
    filter:"filter-btn",
    reset:"reset-btn",
    save:"save-btn",
    load:"load-btn",
    add:"add-order-btn"
  };
  if(document.getElementById(ids.upload)) document.getElementById(ids.upload).onclick=handleUpload;
  if(document.getElementById(ids.clear)) document.getElementById(ids.clear).onclick=clearAllData;
  if(document.getElementById(ids.refresh)) document.getElementById(ids.refresh).onclick=()=>{mergeData(); renderTable(mergedData);};
  if(document.getElementById(ids.filter)) document.getElementById(ids.filter).onclick=filterData;
  if(document.getElementById(ids.reset)) document.getElementById(ids.reset).onclick=resetFilters;
  if(document.getElementById(ids.save)) document.getElementById(ids.save).onclick=saveToJSON;
  if(document.getElementById(ids.load)) document.getElementById(ids.load).onclick=()=>{ 
    const input=document.createElement("input"); 
    input.type="file"; input.accept="application/json";
    input.onchange=()=>{if(input.files.length) loadFromJSON(input.files[0]);}; 
    input.click();
  };
  if(document.getElementById(ids.add)) document.getElementById(ids.add).onclick=addOrders;
}
