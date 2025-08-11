// script.js
// Pastikan xlsx.full.min.js (SheetJS) dimuat sebelum file ini.
// Full: upload (IW39, Data1, Data2, SUM57, Planning), add order, inline edit (Month/Reman), refresh lookup (overlay), save/load.

'use strict';

window.addEventListener('DOMContentLoaded', () => {

  // ---------- Helpers ----------
  const safeStr = v => (v === null || v === undefined) ? '' : String(v).trim();
  const safeLower = v => safeStr(v).toLowerCase();
  const isNumeric = v => v !== '' && !isNaN(Number(v));
  const monthAbbr = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  function excelDateToJS(n) {
    if (!isFinite(n)) return null;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const days = Math.floor(Number(n));
    const ms = days * 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + ms);
  }

  function formatToDDMMMYYYY(val) {
    if (val === null || val === undefined || val === '') return '';
    if (isNumeric(val)) {
      const d = excelDateToJS(Number(val));
      if (d && !isNaN(d.getTime())) return `${String(d.getDate()).padStart(2,'0')}-${monthAbbr[d.getMonth()]}-${d.getFullYear()}`;
      return String(val);
    }
    const tryDate = new Date(val);
    if (!isNaN(tryDate.getTime())) return `${String(tryDate.getDate()).padStart(2,'0')}-${monthAbbr[tryDate.getMonth()]}-${tryDate.getFullYear()}`;
    return String(val);
  }

  function mapRowKeys(row) {
    const mapped = {};
    Object.keys(row).forEach(k => {
      const lk = String(k).trim().toLowerCase();
      const v = row[k];
      if (lk === 'order' || lk === 'order no' || (lk.includes('order') && !lk.includes('order type'))) mapped.Order = safeStr(v);
      else if (lk === 'room' || lk.includes('room')) mapped.Room = safeStr(v);
      else if (lk === 'order type' || lk === 'ordertype' || lk.includes('order type')) mapped.OrderType = safeStr(v);
      else if (lk === 'description' || lk.includes('desc')) mapped.Description = safeStr(v);
      else if (lk === 'created on' || lk.includes('created')) mapped.CreatedOn = v;
      else if (lk === 'user status' || lk.includes('user status')) mapped.UserStatus = safeStr(v);
      else if (lk === 'mat' || lk.includes('mat')) mapped.MAT = safeStr(v);
      else if (lk === 'totalplan' || lk.includes('total plan')) mapped.TotalPlan = v;
      else if (lk === 'totalactual' || lk.includes('total actual')) mapped.TotalActual = v;
      else if (lk === 'section' || lk.includes('section')) mapped.Section = safeStr(v);
      else if (lk === 'cph' || lk.includes('cph')) mapped.CPH = safeStr(v);
      else if (lk === 'status part' || lk.includes('statuspart')) mapped.StatusPart = safeStr(v);
      else if (lk === 'aging' || lk.includes('aging')) mapped.Aging = safeStr(v);
      else if (lk === 'planning' || lk.includes('event start') || lk.includes('event_start')) mapped.Planning = v;
      else if (lk === 'statusamt' || lk.includes('status amt') || lk.includes('status_amt')) mapped.StatusAMT = safeStr(v);
      else mapped[k] = v;
    });
    return mapped;
  }

  // ---------- Global datasets ----------
  let IW39 = [];     // canonical array of IW39 rows
  let Data1 = {};    // MAT or Order -> Section (we auto-detect key when loading)
  let Data2 = {};    // MAT -> CPH
  let SUM57 = {};    // MAT or Order -> {StatusPart, Aging}
  let Planning = {}; // Order -> {Planning (Event Start), StatusAMT}

  // Master Lembar Kerja rows
  let dataLembarKerja = [];

  // ---------- DOM refs ----------
  const $ = id => document.getElementById(id);
  const fileSelect = $('file-select'); // select input to choose which dataset file maps to (IW39/Data1/Data2/SUM57/Planning/Budget)
  const fileInput = $('file-input'); // file input element
  const uploadBtn = $('upload-btn');
  const refreshBtn = $('refresh-btn'); // new refresh button
  const progressContainer = $('progress-container');
  const uploadProgress = $('upload-progress');
  const progressText = $('progress-text');
  const uploadStatus = $('upload-status');

  const addOrderInput = $('add-order-input');
  const addOrderBtn = $('add-order-btn');
  const addOrderStatus = $('add-order-status');

  const filterRoom = $('filter-room');
  const filterOrder = $('filter-order');
  const filterCPH = $('filter-cph');
  const filterMAT = $('filter-mat');
  const filterSection = $('filter-section');
  const filterBtn = $('filter-btn');
  const resetBtn = $('reset-btn');

  const saveBtn = $('save-btn');
  const loadBtn = $('load-btn');

  // loading overlay (create if not exists)
  let loadingOverlay = $('loading-overlay');
  if (!loadingOverlay) {
    loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'loading-overlay';
    loadingOverlay.style.position = 'fixed';
    loadingOverlay.style.left = '0';
    loadingOverlay.style.top = '0';
    loadingOverlay.style.width = '100%';
    loadingOverlay.style.height = '100%';
    loadingOverlay.style.background = 'rgba(0,0,0,0.45)';
    loadingOverlay.style.display = 'none';
    loadingOverlay.style.alignItems = 'center';
    loadingOverlay.style.justifyContent = 'center';
    loadingOverlay.style.zIndex = 9999;
    loadingOverlay.innerHTML = '<div style="background:#fff;padding:18px 26px;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,0.2);font-family:Arial, sans-serif;">Loading... Please wait</div>';
    document.body.appendChild(loadingOverlay);
  }

  function showLoading(msg) {
    if (loadingOverlay) {
      loadingOverlay.querySelector('div').textContent = msg || 'Loading... Please wait';
      loadingOverlay.style.display = 'flex';
    }
  }
  function hideLoading() {
    if (loadingOverlay) loadingOverlay.style.display = 'none';
  }

  // Ensure output tbody exists
  let outputTableBody = document.querySelector('#output-table tbody');
  if (!outputTableBody) {
    console.warn('#output-table tbody not found — creating fallback table at end of body');
    const t = document.createElement('table'); t.id = 'output-table';
    const tb = document.createElement('tbody'); t.appendChild(tb); document.body.appendChild(t); outputTableBody = tb;
  }

  // ---------- Parse Excel ----------
  function parseExcelFileToJSON(file) {
    return new Promise((resolve, reject) => {
      if (!file) return reject(new Error('No file provided'));
      if (typeof FileReader === 'undefined') return reject(new Error('FileReader not supported'));
      const reader = new FileReader();
      reader.onload = e => {
        try {
          if (typeof XLSX === 'undefined') throw new Error('XLSX not found. Include xlsx.full.min.js before this script.');
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          resolve({ sheetName, json });
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = err => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  // ---------- Upload handler ----------
  if (uploadBtn) {
    uploadBtn.addEventListener('click', async () => {
      if (!fileInput || !fileSelect) { alert('Elemen input file atau selector tidak ditemukan'); return; }
      const file = fileInput.files && fileInput.files[0];
      const selectedType = (fileSelect.value || '').trim();
      if (!file) { alert('Pilih file dulu bro!'); return; }

      if (progressContainer) progressContainer.classList.remove('hidden');
      if (uploadProgress) uploadProgress.value = 0; if (progressText) progressText.textContent = '0%';
      if (uploadStatus) uploadStatus.textContent = '';

      try {
        const { sheetName, json } = await parseExcelFileToJSON(file);
        // normalize
        const normalized = (Array.isArray(json) ? json : []).map(r => mapRowKeys(r));

        // Route based on selectedType
        if (selectedType === 'IW39') {
          IW39 = normalized.map(r => ({
            Order: safeStr(r.Order),
            Room: safeStr(r.Room),
            OrderType: safeStr(r.OrderType),
            Description: safeStr(r.Description),
            CreatedOn: (r.CreatedOn !== undefined) ? r.CreatedOn : safeStr(r.CreatedOn),
            UserStatus: safeStr(r.UserStatus),
            MAT: safeStr(r.MAT),
            TotalPlan: isNumeric(r.TotalPlan) ? Number(r.TotalPlan) : (r.TotalPlan || 0),
            TotalActual: isNumeric(r.TotalActual) ? Number(r.TotalActual) : (r.TotalActual || 0)
          }));
          // If no master rows yet -> create base rows for each IW39 order (Order kept)
          if (!dataLembarKerja.length) {
            dataLembarKerja = IW39.map(i => ({
              Order: safeStr(i.Order),
              Room: i.Room || '',
              OrderType: i.OrderType || '',
              Description: i.Description || '',
              CreatedOn: i.CreatedOn || '',
              UserStatus: i.UserStatus || '',
              MAT: i.MAT || '',
              CPH: '',
              Section: '',
              StatusPart: '',
              Aging: '',
              Month: '',
              Cost: '-',
              Reman: '',
              Include: '-',
              Exclude: '-',
              Planning: '',
              StatusAMT: ''
            }));
          }
          // Note: do NOT automatically refresh lookups here. User will press Refresh to re-run lookups.
        } else if (selectedType === 'Data1') {
          // We'll map by MAT if present; fallback to Order if MAT missing
          Data1 = {};
          normalized.forEach(r => {
            const key = safeStr(r.MAT) || safeStr(r.Order);
            if (key) Data1[key] = safeStr(r.Section || r.Section || '');
          });
        } else if (selectedType === 'Data2') {
          Data2 = {};
          normalized.forEach(r => {
            const key = safeStr(r.MAT) || '';
            if (key) Data2[key] = safeStr(r.CPH || '');
          });
        } else if (selectedType === 'SUM57') {
          SUM57 = {};
          normalized.forEach(r => {
            const key = safeStr(r.MAT) || safeStr(r.Order);
            if (key) SUM57[key] = { StatusPart: safeStr(r.StatusPart), Aging: safeStr(r.Aging) };
          });
        } else if (selectedType === 'Planning') {
          Planning = {};
          normalized.forEach(r => {
            const key = safeStr(r.Order);
            if (key) Planning[key] = { Planning: r.Planning || r['Event Start'] || '', StatusAMT: r.StatusAMT || '' };
          });
        } else if (selectedType.toLowerCase() === 'budget' || (sheetName && sheetName.toLowerCase().includes('budget'))) {
          // skip budgets gracefully
          if (uploadStatus) { uploadStatus.style.color = 'green'; uploadStatus.textContent = `File "${file.name}" (Budget) di-skip.`; }
        } else {
          console.log('Unhandled upload type:', selectedType);
        }

        if (uploadProgress) uploadProgress.value = 100; if (progressText) progressText.textContent = '100%';
        if (uploadStatus) { uploadStatus.style.color = 'green'; uploadStatus.textContent = `File "${file.name}" uploaded as ${selectedType}. (refresh to apply)`; }

        // save new datasets to localStorage for persistence (optional)
        try {
          localStorage.setItem('dataSources_IW39', JSON.stringify(IW39));
          localStorage.setItem('dataSources_Data1', JSON.stringify(Data1));
          localStorage.setItem('dataSources_Data2', JSON.stringify(Data2));
          localStorage.setItem('dataSources_SUM57', JSON.stringify(SUM57));
          localStorage.setItem('dataSources_Planning', JSON.stringify(Planning));
        } catch (e) { /* ignore */ }

        // Do not refresh lookups here; user will press Refresh (or call refreshLookup() manually)
      } catch (err) {
        if (uploadStatus) { uploadStatus.style.color = 'red'; uploadStatus.textContent = `Error saat memproses file: ${err && err.message ? err.message : err}`; }
        console.error(err);
      } finally {
        setTimeout(() => { if (progressContainer) progressContainer.classList.add('hidden'); }, 600);
        if (fileInput) fileInput.value = '';
      }
    });
  } else {
    console.warn('uploadBtn not found — upload feature disabled.');
  }

  // ---------- Refresh / Lookup (MAIN) ----------
  async function refreshLookup() {
    showLoading('Refreshing data — rebuilding Menu 2 (please wait)...');

    // We will iterate dataLembarKerja and for each row, lookup fields using the loaded datasets.
    // Preserve Order, Month, Reman that user may have edited.
    await new Promise(resolve => setTimeout(resolve, 50)); // small tick to let UI update

    dataLembarKerja = dataLembarKerja.map(row => {
      const preservedOrder = safeStr(row.Order);
      const preservedMonth = row.Month || '';
      const preservedReman = row.Reman || '';

      // Lookup IW39 by Order
      const iw = (Array.isArray(IW39) ? IW39.find(i => safeLower(i.Order) === safeLower(preservedOrder)) : undefined) || {};

      // Basic fields from IW39 (overwrite except Order)
      const Room = iw.Room || row.Room || '';
      const OrderType = iw.OrderType || row.OrderType || '';
      const Description = iw.Description || row.Description || '';
      const CreatedOn = iw.CreatedOn !== undefined ? iw.CreatedOn : (row.CreatedOn || '');
      const UserStatus = iw.UserStatus || row.UserStatus || '';
      const MAT = iw.MAT || row.MAT || '';

      // CPH logic: if description starts with JR (case-insensitive) => JR, else lookup Data2 by MAT
      let CPH = row.CPH || '';
      if (safeLower((Description || '').substring(0,2)) === 'jr') {
        CPH = 'JR';
      } else if (MAT && Data2 && Data2[safeStr(MAT)]) {
        CPH = Data2[safeStr(MAT)];
      } else {
        CPH = row.CPH || '';
      }

      // Section: lookup Data1 by MAT first, then by Order fallback
      let Section = '';
      if (MAT && Data1 && Data1[safeStr(MAT)]) Section = Data1[safeStr(MAT)];
      else if (Data1 && Data1[safeStr(preservedOrder)]) Section = Data1[safeStr(preservedOrder)];
      else Section = row.Section || '';

      // SUM57: use MAT key preferred, else Order fallback
      let StatusPart = row.StatusPart || '';
      let Aging = row.Aging || '';
      if (MAT && SUM57 && SUM57[safeStr(MAT)]) {
        StatusPart = SUM57[safeStr(MAT)].StatusPart || '';
        Aging = SUM57[safeStr(MAT)].Aging || '';
      } else if (SUM57 && SUM57[safeStr(preservedOrder)]) {
        StatusPart = SUM57[safeStr(preservedOrder)].StatusPart || '';
        Aging = SUM57[safeStr(preservedOrder)].Aging || '';
      }

      // Cost: need numeric TotalPlan & TotalActual from IW39
      let Cost = row.Cost || '-';
      if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined && isFinite(iw.TotalPlan) && isFinite(iw.TotalActual)) {
        const costCalc = (Number(iw.TotalPlan) - Number(iw.TotalActual)) / 16500;
        Cost = costCalc < 0 ? '-' : Number(costCalc);
      } else {
        Cost = row.Cost || '-';
      }

      // Include: if Reman contains 'reman' (case-insensitive) -> Cost*0.25, else Cost
      const RemanNow = preservedReman;
      let Include = row.Include || '-';
      if (safeLower(RemanNow).includes('reman') && typeof Cost === 'number') Include = Number(Cost * 0.25);
      else Include = Cost;

      // Exclude: if OrderType == 'PM38' -> '-', else same as Include
      let Exclude = (safeLower(OrderType) === 'pm38') ? '-' : Include;

      // Planning & StatusAMT lookup by Order
      let PlanningVal = '';
      let StatusAMTVal = '';
      if (Planning && Planning[safeStr(preservedOrder)]) {
        PlanningVal = Planning[safeStr(preservedOrder)].Planning || '';
        StatusAMTVal = Planning[safeStr(preservedOrder)].StatusAMT || '';
      }

      // Return rebuilt row (Order preserved)
      return {
        Order: preservedOrder,
        Room,
        OrderType,
        Description,
        CreatedOn,
        UserStatus,
        MAT,
        CPH,
        Section,
        StatusPart,
        Aging,
        Month: preservedMonth,
        Cost,
        Reman: RemanNow,
        Include,
        Exclude,
        Planning: PlanningVal,
        StatusAMT: StatusAMTVal
      };
    });

    // render finished table
    renderTable(dataLembarKerja);
    hideLoading();
  }

  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      refreshLookup();
    });
  }

  // ---------- BuildDataLembarKerja (initial small helper) ----------
  // Only used to ensure data rows exist and basic fields from IW39 fill initial master when first IW39 upload happens.
  function buildInitialMaster() {
    if (!dataLembarKerja.length && Array.isArray(IW39) && IW39.length) {
      dataLembarKerja = IW39.map(i => ({
        Order: safeStr(i.Order),
        Room: i.Room || '',
        OrderType: i.OrderType || '',
        Description: i.Description || '',
        CreatedOn: i.CreatedOn || '',
        UserStatus: i.UserStatus || '',
        MAT: i.MAT || '',
        CPH: '',
        Section: '',
        StatusPart: '',
        Aging: '',
        Month: '',
        Cost: '-',
        Reman: '',
        Include: '-',
        Exclude: '-',
        Planning: '',
        StatusAMT: ''
      }));
    }
  }

  // ---------- Render table ----------
  function renderTable(data) {
    const dt = Array.isArray(data) ? data : dataLembarKerja;
    const ordersLower = dt.map(d => safeLower(d.Order || ''));
    const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

    if (!outputTableBody) { console.error('#output-table tbody missing; cannot render table'); return; }
    outputTableBody.innerHTML = '';

    if (!dt.length) {
      outputTableBody.innerHTML = `<tr><td colspan="19" style="text-align:center;color:#666;">Tidak ada data.</td></tr>`;
      return;
    }

    dt.forEach(row => {
      const tr = document.createElement('tr');
      if (duplicates.includes(safeLower(row.Order))) tr.classList.add('duplicate');

      function mkCell(val) { const td = document.createElement('td'); td.textContent = (val === null || val === undefined) ? '' : String(val); return td; }

      tr.appendChild(mkCell(row.Room));
      tr.appendChild(mkCell(row.OrderType));
      tr.appendChild(mkCell(row.Order));
      tr.appendChild(mkCell(row.Description));
      tr.appendChild(mkCell(formatToDDMMMYYYY(row.CreatedOn)));
      tr.appendChild(mkCell(row.UserStatus));
      tr.appendChild(mkCell(row.MAT));
      tr.appendChild(mkCell(row.CPH));
      tr.appendChild(mkCell(row.Section));
      tr.appendChild(mkCell(row.StatusPart));
      tr.appendChild(mkCell(row.Aging));

      // Month editable cell
      const tdMonth = document.createElement('td');
      tdMonth.classList.add('editable');
      tdMonth.textContent = row.Month || '';
      tdMonth.title = 'Klik untuk edit Month';
      tdMonth.addEventListener('click', () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost (right)
      const tdCost = document.createElement('td');
      tdCost.classList.add('cost');
      tdCost.textContent = (typeof row.Cost === 'number') ? Number(row.Cost).toFixed(1) : row.Cost;
      tdCost.style.textAlign = 'right';
      tr.appendChild(tdCost);

      // Reman editable
      const tdReman = document.createElement('td');
      tdReman.classList.add('editable');
      tdReman.textContent = row.Reman || '';
      tdReman.title = 'Klik untuk edit Reman';
      tdReman.addEventListener('click', () => editReman(tdReman, row));
      tr.appendChild(tdReman);

      // Include
      const tdInclude = document.createElement('td');
      tdInclude.classList.add('include');
      tdInclude.textContent = (typeof row.Include === 'number') ? Number(row.Include).toFixed(1) : row.Include;
      tdInclude.style.textAlign = 'right';
      tr.appendChild(tdInclude);

      // Exclude
      const tdExclude = document.createElement('td');
      tdExclude.classList.add('exclude');
      tdExclude.textContent = (typeof row.Exclude === 'number') ? Number(row.Exclude).toFixed(1) : row.Exclude;
      tdExclude.style.textAlign = 'right';
      tr.appendChild(tdExclude);

      // Planning & StatusAMT
      tr.appendChild(mkCell(formatToDDMMMYYYY(row.Planning)));
      tr.appendChild(mkCell(row.StatusAMT));

      // Actions
      const tdAction = document.createElement('td');
      const btnEdit = document.createElement('button'); btnEdit.textContent = 'Edit'; btnEdit.classList.add('btn-action','btn-edit');
      btnEdit.addEventListener('click', () => openInlineEditorsForRow(row));
      const btnDelete = document.createElement('button'); btnDelete.textContent = 'Delete'; btnDelete.classList.add('btn-action','btn-delete');
      btnDelete.addEventListener('click', () => {
        if (confirm(`Hapus order ${row.Order}?`)) {
          dataLembarKerja = dataLembarKerja.filter(d => safeLower(d.Order) !== safeLower(row.Order));
          renderTable(dataLembarKerja);
          saveDataToLocalStorage();
        }
      });
      tdAction.appendChild(btnEdit); tdAction.appendChild(btnDelete);
      tr.appendChild(tdAction);

      outputTableBody.appendChild(tr);
    });
  }

  // ---------- Inline edit controls ----------
  function editMonth(td, row) {
    const sel = document.createElement('select');
    monthAbbr.forEach(m => {
      const opt = document.createElement('option'); opt.value = m; opt.textContent = m;
      if (row.Month === m) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', () => {
      row.Month = sel.value;
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    sel.addEventListener('blur', () => renderTable(dataLembarKerja));
    td.textContent = ''; td.appendChild(sel); sel.focus();
  }

  function editReman(td, row) {
    const inp = document.createElement('input'); inp.type = 'text'; inp.value = row.Reman || '';
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        row.Reman = inp.value.trim();
        // recalc include/exclude locally
        if (safeLower(row.Reman).includes('reman') && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
        else row.Include = row.Cost;
        row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === 'Escape') renderTable(dataLembarKerja);
    });
    inp.addEventListener('blur', () => {
      row.Reman = inp.value.trim();
      if (safeLower(row.Reman).includes('reman') && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
      else row.Include = row.Cost;
      row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    td.textContent = ''; td.appendChild(inp); inp.focus();
  }

  function editCost(td, row) {
    const inp = document.createElement('input'); inp.type = 'number'; inp.step = '0.1';
    inp.value = (typeof row.Cost === 'number') ? Number(row.Cost).toFixed(1) : (row.Cost === '-' ? '' : row.Cost);
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        const v = inp.value.trim();
        row.Cost = v === '' ? '-' : Number(v);
        if (safeLower(row.Reman).includes('reman') && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
        else row.Include = row.Cost;
        row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === 'Escape') renderTable(dataLembarKerja);
    });
    inp.addEventListener('blur', () => {
      const v = inp.value.trim();
      row.Cost = v === '' ? '-' : Number(v);
      if (safeLower(row.Reman).includes('reman') && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
      else row.Include = row.Cost;
      row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    td.textContent = ''; td.appendChild(inp); inp.focus();
  }

  function openInlineEditorsForRow(row) {
    const trs = Array.from(outputTableBody.querySelectorAll('tr'));
    for (const tr of trs) {
      const orderCell = tr.children[2]; // Order column index in renderTable
      if (orderCell && safeLower(orderCell.textContent) === safeLower(row.Order)) {
        const monthTd = tr.children[11];
        const costTd = tr.children[12];
        const remanTd = tr.children[13];
        if (monthTd) editMonth(monthTd, row);
        if (costTd) editCost(costTd, row);
        if (remanTd) editReman(remanTd, row);
        break;
      }
    }
  }

  // ---------- Add Order ----------
  if (addOrderBtn && addOrderInput && addOrderStatus) {
    addOrderBtn.addEventListener('click', () => {
      const raw = addOrderInput.value.trim();
      if (!raw) {
        addOrderStatus.style.color = 'red'; addOrderStatus.textContent = 'Masukkan minimal 1 order.';
        return;
      }
      const orders = raw.split(/[\s,]+/).map(s => s.trim()).filter(Boolean);
      let added = 0, skipped = [], invalid = [];
      for (const o of orders) {
        if (!isValidOrder(o)) { invalid.push(o); continue; }
        if (dataLembarKerja.some(d => safeLower(d.Order) === safeLower(o))) { skipped.push(o); continue; }
        dataLembarKerja.push({
          Order: safeStr(o),
          Room: '', OrderType: '', Description: '', CreatedOn: '', UserStatus: '',
          MAT: '', CPH: '', Section: '', StatusPart: '', Aging: '', Month: '', Cost: '-', Reman: '', Include: '-', Exclude: '-', Planning: '', StatusAMT: ''
        });
        added++;
      }

      // Do not auto-refresh lookups — user must press Refresh to rebuild all columns
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();

      let msg = `${added} order ditambahkan.`;
      if (skipped.length) msg += ` Sudah ada: ${skipped.join(', ')}.`;
      if (invalid.length) msg += ` Invalid: ${invalid.join(', ')}.`;
      addOrderStatus.style.color = added ? 'green' : 'red'; addOrderStatus.textContent = msg;
      addOrderInput.value = '';
    });
  } else {
    console.warn('addOrder elements missing — Add disabled');
  }

  // ---------- Filters ----------
  if (filterBtn) {
    filterBtn.addEventListener('click', () => {
      let filtered = dataLembarKerja;
      if (filterRoom && filterRoom.value.trim()) filtered = filtered.filter(d => (d.Room || '').toLowerCase().includes(filterRoom.value.trim().toLowerCase()));
      if (filterOrder && filterOrder.value.trim()) filtered = filtered.filter(d => (d.Order || '').toLowerCase().includes(filterOrder.value.trim().toLowerCase()));
      if (filterCPH && filterCPH.value.trim()) filtered = filtered.filter(d => (d.CPH || '').toLowerCase().includes(filterCPH.value.trim().toLowerCase()));
      if (filterMAT && filterMAT.value.trim()) filtered = filtered.filter(d => (d.MAT || '').toLowerCase().includes(filterMAT.value.trim().toLowerCase()));
      if (filterSection && filterSection.value.trim()) filtered = filtered.filter(d => (d.Section || '').toLowerCase().includes(filterSection.value.trim().toLowerCase()));
      renderTable(filtered);
    });
  }

  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      if (filterRoom) filterRoom.value = ''; if (filterOrder) filterOrder.value = '';
      if (filterCPH) filterCPH.value = ''; if (filterMAT) filterMAT.value = '';
      if (filterSection) filterSection.value = '';
      renderTable(dataLembarKerja);
    });
  }

  // ---------- Save / Load ----------
  function saveDataToLocalStorage() {
    try { localStorage.setItem('lembarKerjaData', JSON.stringify(dataLembarKerja)); }
    catch (e) { console.warn('Gagal menyimpan:', e); }
  }

  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      saveDataToLocalStorage(); alert('Data disimpan di browser.');
    });
  }

  if (loadBtn) {
    loadBtn.addEventListener('click', () => {
      const saved = localStorage.getItem('lembarKerjaData');
      if (!saved) { alert('Tidak ada data tersimpan.'); return; }
      try {
        dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) }));
        renderTable(dataLembarKerja); alert('Data dimuat.');
      } catch (e) { alert('Gagal muat data: ' + (e && e.message ? e.message : e)); }
    });
  }

  function isValidOrder(order) { return !/[.,]/.test(order); }

  // ---------- Init menu (safe) ----------
  (function initMenu() {
    const menuItems = document.querySelectorAll('.menu-item');
    const contentSections = document.querySelectorAll('.content-section');
    menuItems.forEach(item => {
      item.addEventListener('click', () => {
        menuItems.forEach(i => i.classList.remove('active')); contentSections.forEach(s => s.classList.remove('active'));
        item.classList.add('active'); const target = item.getAttribute('data-menu'); const sec = document.getElementById(target);
        if (sec) sec.classList.add('active');
      });
    });
  })();

  // ---------- Init load saved / initial master ----------
  (function init() {
    // Try to restore previous data sources (optional)
    try {
      const sIW = localStorage.getItem('dataSources_IW39'); if (sIW) IW39 = JSON.parse(sIW);
      const sD1 = localStorage.getItem('dataSources_Data1'); if (sD1) Data1 = JSON.parse(sD1);
      const sD2 = localStorage.getItem('dataSources_Data2'); if (sD2) Data2 = JSON.parse(sD2);
      const sS5 = localStorage.getItem('dataSources_SUM57'); if (sS5) SUM57 = JSON.parse(sS5);
      const sPl = localStorage.getItem('dataSources_Planning'); if (sPl) Planning = JSON.parse(sPl);
    } catch (e) { /* ignore */ }

    const saved = localStorage.getItem('lembarKerjaData');
    if (saved) {
      try { dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) })); }
      catch (e) { dataLembarKerja = []; }
    } else {
      dataLembarKerja = [];
    }

    // If IW39 already loaded but master empty -> init basic master
    buildInitialMaster();
    renderTable(dataLembarKerja);
  })();

});
