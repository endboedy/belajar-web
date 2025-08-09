let iwData = [];
let data1 = [];
let data2 = [];
let sum57 = [];
let planning = [];
let budget = [];
let merged = [];

function normalizeKey(k){
  return String(k||"").toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'');
}

function normalizeRows(rows){
  return (rows||[]).map(r=>{
    const o={};
    Object.keys(r||{}).forEach(k=>{
      o[normalizeKey(k)] = r[k];
    });
    return o;
  });
}

function asNumber(v){
  if(v===undefined || v===null || v==="" ) return 0;
  const num = Number(String(v).toString().replace(/[^0-9.\-]/g,''));
  return isNaN(num) ? 0 : num;
}

function readExcelFile(file){
  return new Promise((resolve)=>{
    if(!file) return resolve([]);
    const reader = new FileReader();
    reader.onload = e=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data,{type:'array'});
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

function buildMergedRow(origNorm, orderKey, monthVal, remanVal, d1Norm, d2Norm, sumNorm, planNorm){
  const get = (o, names) => {
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

  let CreatedOn = "";
  if(CreatedOnRaw){
    const d = new Date(CreatedOnRaw);
    if(!isNaN(d)) {
      CreatedOn = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    } else {
      CreatedOn = CreatedOnRaw;
    }
  }

  const d1row = (d1Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const d2row = (d2Norm || []).find(r => Object.values(r).some(v => String(v).trim() === String(MAT).trim())) || {};
  const sumrow = (sumNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};
  const planrow = (planNorm || []).find(r => Object.values(r).some(v => String(v).trim() === String(Order).trim())) || {};

  let Description = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) Description = "JR";
  else Description = get(d1row, ['description','desc','keterangan','shortdesc']) || "";

  let CPH = "";
  if(String(DescriptionRaw).toUpperCase().startsWith("JR")) CPH = "JR";
  else CPH = get(d2row, ['cph','costperhour','cphvalue']) || get(d1row, ['cph','costperhour','cphvalue']) || "";

  const Section = get(d1row, ['section','dept','deptcode','department']);
  const StatusPart = get(sumrow, ['statuspart','status','status_part']) || "";

  let AgingRaw = get(sumrow, ['aging','age']);
  let Aging = "";
  if(AgingRaw !== "" && AgingRaw !== undefined && AgingRaw !== null){
    Aging = parseInt(Number(AgingRaw));
  }

  let PlanningRaw = get(planrow, ['planning','eventstart','description']);
  let Planning = "";
  if(PlanningRaw){
    const d = new Date(PlanningRaw);
    if(!isNaN(d)) {
      Planning = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    } else {
      Planning = PlanningRaw;
    }
  }
  const StatusAMT = get(planrow, ['statusamt','status']) || "";

  const planKey = Object.keys(origNorm||{}).find(k => k.includes('plan')) || null;
  const actualKey = Object.keys(origNorm||{}).find(k => k.includes('actual')) || null;
  const planVal = planKey ? asNumber(origNorm[planKey]) : 0;
  const actualVal = actualKey ? asNumber(origNorm[actualKey]) : 0;
  let Cost = "-";
  const rawCost = (planVal - actualVal) / 16500;
  if(!isNaN(rawCost) && isFinite(rawCost) && rawCost >= 0) Cost = Number(rawCost.toFixed(2));

  let Include = "-";
  if(typeof Cost === "number"){
    Include = (String(remanVal).toLowerCase().includes("reman") ? Number((Cost * 0.25).toFixed(2)) : Number(Cost.toFixed(2)));
  }
  let Exclude = (String(OrderType).toUpperCase() === "PM38") ? "-" : Include;

  return {
    "Room": Room||"",
    "Order Type": OrderType||"",
    "Order": Order||"",
    "Description": Description||"",
    "Created On": CreatedOn||"",
    "User Status": UserStatus||"",
    "MAT": MAT||"",
    "CPH": CPH||"",
    "Section": Section||"",
    "Status Part": StatusPart||"",
    "Aging": Aging||"",
    "Month": monthVal||"",
    "Cost": Cost,
    "Reman": remanVal||"",
    "Include": Include,
    "Exclude": Exclude,
    "Planning": Planning||"",
    "Status AMT": StatusAMT||""
  };
}

const DISPLAY = ["Room","Order Type","Order","Description","Created On","User Status","MAT","CPH","Section","Status Part","Aging","Month","Cost","Reman","Include","Exclude","Planning","Status AMT","Action"];

function renderTable(data){
  const container = document.getElementById('tableContainer');
  if(!data || data.length===0){
    container.innerHTML = `<div class="hint" style="padding:18px;text-align:center">Tidak ada data. Tambahkan order atau load files dulu.</div>`;
    return;
  }
  container.style.overflowX = "auto";
  container.style.width = "100vw";

  let html = "<table>";
  html += "<thead><tr>";
  DISPLAY.forEach(col => {
    html += `<th>${col}</th>`;
  });
  html += "</tr></thead><tbody>";

  data.forEach((row, idx) => {
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col => {
      let v = row[col];
      if(col === "Order"){
        if(typeof v === "string") v = v.replace(/\./g, "").split(",")[0];
      }
      if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
      html += `<td>${v !== undefined && v !== null ? v : ""}</td>`;
    });
    html += `<td>
      <button onclick="editRow(${idx})">Edit</button>
      <button onclick="deleteRow(${idx})">Delete</button>
    </td>`;
    html += "</tr>";
  });
  html += "</tbody></table>";
  container.innerHTML = html;
}

function editRow(idx){
  const row = merged[idx];
  if(!row) return;
  const container = document.getElementById('tableContainer');
  let html = "<table><thead><tr>";
  DISPLAY.forEach(col => html += `<th>${col}</th>`);
  html += "</tr></thead><tbody>";

  merged.forEach((r,i)=>{
    html += "<tr>";
    const cols = DISPLAY.filter(c=>c!=="Action");
    cols.forEach(col=>{
      if(i === idx){
        if(col === "Month"){
          html += `<td>
            <select id="edit_month">
              <option value="">--Pilih--</option>
              <option>Jan</option><option>Feb</option><option>Mar</option>
              <option>Apr</option><option>Mei</option><option>Jun</option>
              <option>Jul</option><option>Agu</option><option>Sep</option>
              <option>Okt</option><option>Nov</option><option>Des</option>
            </select>
          </td>`;
        } else if(col === "Reman"){
          html += `<td><input id="edit_reman" type="text" value="${r.Reman||""}" /></td>`;
        } else {
          let v = r[col];
          if(typeof v === "number") v = v.toLocaleString();
          html += `<td>${v !== undefined && v !== null ? v : ""}</td>`;
        }
      } else {
        let v = r[col];
        if(typeof v === "number") v = v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
        html += `<td>${v !== undefined && v !== null ? v : ""}</td>`;
      }
    });

    if(i === idx){
      html += `<td>
        <button onclick="saveEdit(${i})">Save</button>
        <button onclick="cancelEdit()">Cancel</button>
      </td>`;
    } else {
      html += `<td>
        <button onclick="editRow(${i})">Edit</button>
        <button onclick="deleteRow(${i})">Delete</button>
      </td>`;
    }
    html += "</tr>";
  });
  html += "</tbody></table>";
  container.innerHTML = html;

  const sel = document.getElementById('edit_month');
  if(sel) sel.value = row.Month || "";
  const rin = document.getElementById('edit_reman');
  if(rin) rin.value = row.Reman || "";
}

function saveEdit(idx){
  const month = document.getElementById('edit_month') ? document.getElementById('edit_month').value : "";
  const reman = document.getElementById('edit_reman') ? document.getElementById('edit_reman').value : "";
  if(merged[idx]){
    merged[idx].Month = month;
    merged[idx].Reman = reman;
    const cost = merged[idx].Cost;
    if(typeof cost === "number"){
      merged[idx].Include = String(reman).toLowerCase().includes("reman") ? Number((cost * 0.25).toFixed(2)) : Number(cost.toFixed(2));
    } else {
      merged[idx].Include = "-";
    }
    merged[idx].Exclude = String(merged[idx]["Order Type"]).toUpperCase() === "PM38" ? "-" : merged[idx].Include;
  }
  renderTable(merged);
  const lmMsg = document.getElementById('lmMsg');
  if(lmMsg) lmMsg.textContent = "Perubahan disimpan (sementara). Jangan lupa klik Save Lembar Kerja untuk persist.";
}

function cancelEdit(){
  renderTable(merged);
}

function deleteRow(idx){
  if(!confirm("Hapus baris ini?")) return;
  merged.splice(idx,1);
  renderTable(merged);
  const lmMsg = document.getElementById('lmMsg');
  if(lmMsg) lmMsg.textContent = "Baris dihapus (sementara). Klik Save Lembar Kerja untuk persist.";
}

document.addEventListener('DOMContentLoaded', () => {
  const btnLoad = document.getElementById('btnLoad');
  if(btnLoad) btnLoad.addEventListener('click', async () => {
    const loadMsg = document.getElementById('loadMsg');
    if(loadMsg) loadMsg.textContent = "Loading files...";

    const fIW = document.getElementById('fileIW39')?.files[0];
    const fSUM = document.getElementById('fileSUM57')?.files[0];
    const fPlan = document.getElementById('filePlanning')?.files[0];
    const fBud = document.getElementById('fileBudget')?.files[0];
    const fD1 = document.getElementById('fileData1')?.files[0];
    const fD2 = document.getElementById('fileData2')?.files[0];

    iwData = fIW ? await readExcelFile(fIW) : [];
    sum57 = fSUM ? await readExcelFile(fSUM) : [];
    planning = fPlan ? await readExcelFile(fPlan) : [];
    budget = fBud ? await readExcelFile(fBud) : [];
    data1 = fD1 ? await readExcelFile(fD1) : [];
    data2 = fD2 ? await readExcelFile(fD2) : [];

    const iwNorm = normalizeRows(iwData);
    const d1Norm = normalizeRows(data1);
    const d2Norm = normalizeRows(data2);
    const sumNorm = normalizeRows(sum57);
    const planNorm = normalizeRows(planning);

    window.__nd = {iw:iwNorm,d1:d1Norm,d2:d2Norm,sum:sumNorm,plan:planNorm};

    if(loadMsg) loadMsg.textContent = `Selesai load. IW39 baris: ${iwData.length}`;
  });

  const btnAddOrders = document.getElementById('btnAddOrders');
  if(btnAddOrders) btnAddOrders.addEventListener('click', ()=>{
    const raw = document.getElementById('inputOrders')?.value || "";
    if(raw.trim()===""){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Tidak ada order yang dimasukkan.";
      return;
    }
    if(!window.__nd || !window.__nd.iw){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "File IW39 belum di-load.";
      return;
    }
    const orders = raw.split(/[\r\n]+/).map(s=>s.trim()).filter(s=>s.length>0);

    const iwNorm = window.__nd.iw;
    const d1Norm = window.__nd.d1;
    const d2Norm = window.__nd.d2;
    const sumNorm = window.__nd.sum;
    const planNorm = window.__nd.plan;

    orders.forEach(orderKey => {
      const row = iwNorm.find(r => {
        const v = r['order'] || r['orderno'] || r['no'] || "";
        return String(v).trim() === orderKey.trim();
      });
      if(row){
        const m = "";
        const rmn = "";
        const newRow = buildMergedRow(row, orderKey, m, rmn, d1Norm, d2Norm, sumNorm, planNorm);
        merged.push(newRow);
      }
    });
    renderTable(merged);
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = `${orders.length} order berhasil ditambahkan.`;
  });

  const btnSave = document.getElementById('btnSave');
  if(btnSave) btnSave.addEventListener('click', ()=>{
    localStorage.setItem('ndarboe_merged', JSON.stringify(merged));
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "Lembar Kerja tersimpan di localStorage.";
  });

  function loadSaved(){
    const raw = localStorage.getItem('ndarboe_merged');
    if(raw){
      try{
        merged = JSON.parse(raw);
        renderTable(merged);
        const lmMsg = document.getElementById('lmMsg');
        if(lmMsg) lmMsg.textContent = "Lembar Kerja dimuat dari localStorage.";
      }catch(e){
        console.error(e);
        const lmMsg = document.getElementById('lmMsg');
        if(lmMsg) lmMsg.textContent = "Gagal memuat Lembar Kerja dari localStorage.";
      }
    }
  }
  loadSaved();

  const btnClear = document.getElementById('btnClear');
  if(btnClear) btnClear.addEventListener('click', ()=>{
    if(!confirm("Yakin ingin hapus semua Lembar Kerja?")) return;
    merged = [];
    renderTable(merged);
    localStorage.removeItem('ndarboe_merged');
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "Lembar Kerja dihapus.";
  });

  const inputOrders = document.getElementById('inputOrders');
  if(inputOrders) inputOrders.addEventListener('focus', ()=>{
    const lmMsg = document.getElementById('lmMsg');
    if(lmMsg) lmMsg.textContent = "";
  });

  const btnDownload = document.getElementById('btnDownload');
  if(btnDownload) btnDownload.addEventListener('click', ()=>{
    if(!merged || merged.length===0){
      const lmMsg = document.getElementById('lmMsg');
      if(lmMsg) lmMsg.textContent = "Tidak ada data untuk diunduh.";
      return;
    }
    const ws = XLSX.utils.json_to_sheet(merged);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "LembarKerja");
    XLSX.writeFile(wb, "LembarKerja.xlsx");
  });
});

// Show / Hide pages
function showPage(pageId){
  document.querySelectorAll('.page').forEach(p => {
    p.style.display = (p.id === pageId) ? "block" : "none";
  });
}

// Expose functions globally
window.editRow = editRow;
window.deleteRow = deleteRow;
window.saveEdit = saveEdit;
window.cancelEdit = cancelEdit;
