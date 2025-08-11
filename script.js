// script.js
// Copy-paste seluruh file ini. Pastikan xlsx.full.min.js (SheetJS) sudah dimuat sebelum script ini.
// This script:
// - load multiple Excel files via file input (IW39, Data1, Data2, SUM57, Planning)
// - normalize headers
// - build master "Lembar Kerja" using IW39 as base and lookups from Data1/Data2/SUM57/Planning
// - render table with fixed column order
// - allow inline edit: Month (select), Cost (number), Reman (text)
// - Add Order (manual), Filter, Save/Load (localStorage)
// - hardened: checks for missing DOM elements to avoid null errors

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
      else if (lk === 'created on' || lk.includes('created')) mapped.CreatedOn = v; // keep raw (could be number)
      else if (lk === 'user status' || lk.includes('user status')) mapped.UserStatus = safeStr(v);
      else if (lk === 'mat' || lk.includes('mat')) mapped.MAT = safeStr(v);
      else if (lk === 'totalplan' || lk.includes('total plan')) mapped.TotalPlan = v;
      else if (lk === 'totalactual' || lk.includes('total actual')) mapped.TotalActual = v;
      else if (lk === 'section' || lk.includes('section')) mapped.Section = safeStr(v);
      else if (lk === 'cph' || lk.includes('cph')) mapped.CPH = safeStr(v);
      else if (lk === 'status part' || lk.includes('statuspart')) mapped.StatusPart = safeStr(v);
      else if (lk === 'aging' || lk.includes('aging')) mapped.Aging = safeStr(v);
      else if (lk === 'planning' || lk.includes('event start') || lk.includes('event_start')) mapped.Planning = v;
      else if (lk === 'statusamt' || lk.includes('status amt')) mapped.StatusAMT = safeStr(v);
      else mapped[k] = v;
    });
    return mapped;
  }

  // ---------- Global datasets ----------
  let IW39 = [];     // canonical array of IW39 rows
  let Data1 = {};    // Order -> Section
  let Data2 = {};    // MAT -> CPH
  let SUM57 = {};    // Order -> {StatusPart, Aging}
  let Planning = {}; // Order -> {Planning, StatusAMT}

  // Master Lembar Kerja rows
  let dataLembarKerja = [];

  // ---------- DOM refs (safe) ----------
  const $ = id => document.getElementById(id);
  const fileSelect = $('file-select');
  const fileInput = $('file-input');
  const uploadBtn = $('upload-btn');
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

  // Ensure output tbody exists (or create fallback)
  let outputTableBody = document.querySelector('#output-table tbody');
  if (!outputTableBody) {
    console.warn('#output-table tbody not found, creating fallback table at document.body end');
    const t = document.createElement('table'); t.id = 'output-table'; t.style.width = '100%';
    const tb = document.createElement('tbody'); t.appendChild(tb); document.body.appendChild(t); outputTableBody = tb;
  }

  // Fixed column indices (used by inline edit lookup)
  const COL = {
    ROOM: 0, ORDER_TYPE: 1, ORDER: 2, DESCRIPTION: 3, CREATED_ON: 4, USER_STATUS: 5,
    MAT: 6, CPH: 7, SECTION: 8, STATUS_PART: 9, AGING: 10, MONTH: 11,
    COST: 12, REMAN: 13, INCLUDE: 14, EXCLUDE: 15, PLANNING: 16, STATUS_AMT: 17, ACTION: 18
  };

  // ---------- Excel parsing ----------
  function parseExcelFileToJSON(file) {
    return new Promise((resolve, reject) => {
      if (!file) return reject(new Error('No file'));
      if (typeof FileReader === 'undefined') return reject(new Error('FileReader unsupported'));
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
      if (!fileInput || !fileSelect) { alert('Elemen file input atau file select tidak ditemukan'); return; }
      const file = fileInput.files && fileInput.files[0];
      const selectedType = (fileSelect.value || '').trim();

      if (!file) { alert('Pilih file dulu bro!'); return; }
      if (progressContainer) progressContainer.classList.remove('hidden');
      if (uploadProgress) uploadProgress.value = 0; if (progressText) progressText.textContent = '0%';
      if (uploadStatus) uploadStatus.textContent = '';

      try {
        const { sheetName, json } = await parseExcelFileToJSON(file);

        // skip "budget"
        if (selectedType.toLowerCase() === 'budget' || (sheetName && sheetName.toLowerCase().includes('budget'))) {
          if (uploadProgress) uploadProgress.value = 100; if (progressText) progressText.textContent = '100%';
          if (uploadStatus) { uploadStatus.style.color = 'green'; uploadStatus.textContent = `File "${file.name}" (Budget) di-skip.`; }
          setTimeout(()=> { if (progressContainer) progressContainer.classList.add('hidden'); }, 600);
          fileInput.value = ''; return;
        }

        const normalized = (Array.isArray(json) ? json : []).map(r => mapRowKeys(r));

        if (selectedType === 'IW39') {
          IW39 = normalized.map(r => ({
            Order: safeStr(r.Order), Room: safeStr(r.Room), OrderType: safeStr(r.OrderType),
            Description: safeStr(r.Description), CreatedOn: r.CreatedOn !== undefined ? r.CreatedOn : safeStr(r.CreatedOn),
            UserStatus: safeStr(r.UserStatus), MAT: safeStr(r.MAT),
            TotalPlan: isNumeric(r.TotalPlan) ? Number(r.TotalPlan) : (r.TotalPlan || 0),
            TotalActual: isNumeric(r.TotalActual) ? Number(r.TotalActual) : (r.TotalActual || 0)
          }));
          // If master empty initialize with IW39 orders
          if (!dataLembarKerja.length) {
            dataLembarKerja = IW39.map(i => ({
              Order: safeStr(i.Order), Room: i.Room || '', OrderType: i.OrderType || '', Description: i.Description || '',
              CreatedOn: i.CreatedOn || '', UserStatus: i.UserStatus || '', MAT: i.MAT || '', CPH: '',
              Section: '', StatusPart: '', Aging: '', Month: '', Cost: '-', Reman: '', Include: '-', Exclude: '-',
              Planning: '', StatusAMT: ''
            }));
          } else {
            // update existing master rows
            dataLembarKerja = dataLembarKerja.map(row => {
              const match = IW39.find(i => safeLower(i.Order) === safeLower(row.Order));
              if (match) {
                return {
                  ...row,
                  Room: match.Room || row.Room,
                  OrderType: match.OrderType || row.OrderType,
                  Description: match.Description || row.Description,
                  CreatedOn: match.CreatedOn || row.CreatedOn,
                  UserStatus: match.UserStatus || row.UserStatus,
                  MAT: match.MAT || row.MAT
                };
              }
              return row;
            });
          }
        } else if (selectedType === 'Data1') {
          Data1 = {};
          normalized.forEach(r => { if (r.Order) Data1[safeStr(r.Order)] = safeStr(r.Section || ''); });
        } else if (selectedType === 'Data2') {
          Data2 = {};
          normalized.forEach(r => { if (r.MAT) Data2[safeStr(r.MAT)] = safeStr(r.CPH || ''); });
        } else if (selectedType === 'SUM57') {
          SUM57 = {};
          normalized.forEach(r => { if (r.Order) SUM57[safeStr(r.Order)] = { StatusPart: safeStr(r.StatusPart), Aging: safeStr(r.Aging) }; });
        } else if (selectedType === 'Planning') {
          Planning = {};
          normalized.forEach(r => { if (r.Order) Planning[safeStr(r.Order)] = { Planning: r.Planning || '', StatusAMT: r.StatusAMT || '' }; });
        } else {
          console.log('Unhandled upload type:', selectedType);
        }

        if (uploadProgress) uploadProgress.value = 100; if (progressText) progressText.textContent = '100%';
        if (uploadStatus) { uploadStatus.style.color = 'green'; uploadStatus.textContent = `File "${file.name}" berhasil diupload untuk ${selectedType}.`; }

        buildDataLembarKerja();
        renderTable(dataLembarKerja);

      } catch (err) {
        if (uploadStatus) { uploadStatus.style.color = 'red'; uploadStatus.textContent = `Error: ${err && err.message ? err.message : err}`; }
        console.error(err);
      } finally {
        setTimeout(()=> { if (progressContainer) progressContainer.classList.add('hidden'); }, 600);
        if (fileInput) fileInput.value = '';
      }
    });
  } else {
    console.warn('uploadBtn not found — upload disabled.');
  }

  // ---------- Build / lookup / compute ----------
  function buildDataLembarKerja() {
    dataLembarKerja = Array.isArray(dataLembarKerja) ? dataLembarKerja.map(r => ({ ...r, Order: safeStr(r.Order) })) : [];

    dataLembarKerja = dataLembarKerja.map(row => {
      const iw = (Array.isArray(IW39) ? IW39.find(i => safeLower(i.Order) === safeLower(row.Order)) : undefined) || {};

      row.Room = iw.Room || row.Room || '';
      row.OrderType = iw.OrderType || row.OrderType || '';
      row.Description = iw.Description || row.Description || '';
      row.CreatedOn = iw.CreatedOn !== undefined ? iw.CreatedOn : row.CreatedOn || '';
      row.UserStatus = iw.UserStatus || row.UserStatus || '';
      row.MAT = iw.MAT || row.MAT || '';

      // CPH logic
      if (safeLower((row.Description || '').substring(0,2)) === 'jr') row.CPH = 'JR';
      else row.CPH = Data2[row.MAT] || row.CPH || '';

      // section
      row.Section = Data1[row.Order] || row.Section || '';

      // SUM57 lookups
      if (SUM57[row.Order]) {
        row.StatusPart = SUM57[row.Order].StatusPart || '';
        row.Aging = SUM57[row.Order].Aging || '';
      } else {
        row.StatusPart = row.StatusPart || '';
        row.Aging = row.Aging || '';
      }

      // Cost calc
      if (iw.TotalPlan !== undefined && iw.TotalActual !== undefined && isFinite(iw.TotalPlan) && isFinite(iw.TotalActual)) {
        const costCalc = (Number(iw.TotalPlan) - Number(iw.TotalActual)) / 16500;
        row.Cost = costCalc < 0 ? '-' : Number(costCalc);
      } else {
        row.Cost = row.Cost || '-';
      }

      // Include
      if (safeLower(row.Reman) === 'reman') row.Include = (typeof row.Cost === 'number') ? Number(row.Cost * 0.25) : '-';
      else row.Include = row.Cost;

      // Exclude
      if (safeLower(row.OrderType) === 'pm38') row.Exclude = '-';
      else row.Exclude = row.Include;

      // Planning
      if (Planning[row.Order]) {
        row.Planning = Planning[row.Order].Planning || '';
        row.StatusAMT = Planning[row.Order].StatusAMT || '';
      } else {
        row.Planning = row.Planning || '';
        row.StatusAMT = row.StatusAMT || '';
      }

      return row;
    });
  }

  // ---------- Render table ----------
  function renderTable(data) {
    const dt = Array.isArray(data) ? data : dataLembarKerja;
    const ordersLower = dt.map(d => safeLower(d.Order));
    const duplicates = ordersLower.filter((item, idx) => ordersLower.indexOf(item) !== idx);

    if (!outputTableBody) { console.error('#output-table tbody missing'); return; }
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

      // Month (editable)
      const tdMonth = document.createElement('td');
      tdMonth.classList.add('editable');
      tdMonth.textContent = row.Month || '';
      tdMonth.title = 'Klik untuk edit Month';
      tdMonth.addEventListener('click', () => editMonth(tdMonth, row));
      tr.appendChild(tdMonth);

      // Cost
      const tdCost = document.createElement('td');
      tdCost.classList.add('cost');
      tdCost.textContent = (typeof row.Cost === 'number') ? Number(row.Cost).toFixed(1) : row.Cost;
      tdCost.style.textAlign = 'right';
      tr.appendChild(tdCost);

      // Reman (editable)
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

      // Action buttons
      const tdAction = document.createElement('td');
      const btnEdit = document.createElement('button'); btnEdit.textContent = 'Edit'; btnEdit.classList.add('btn-action','btn-edit');
      btnEdit.addEventListener('click', () => openInlineEditorsForRow(row));
      const btnDelete = document.createElement('button'); btnDelete.textContent = 'Delete'; btnDelete.classList.add('btn-action','btn-delete');
      btnDelete.addEventListener('click', () => {
        if (confirm(`Hapus order ${row.Order}?`)) {
          dataLembarKerja = dataLembarKerja.filter(d => safeLower(d.Order) !== safeLower(row.Order));
          buildDataLembarKerja();
          renderTable(dataLembarKerja);
          saveDataToLocalStorage();
        }
      });
      tdAction.appendChild(btnEdit); tdAction.appendChild(btnDelete);
      tr.appendChild(tdAction);

      outputTableBody.appendChild(tr);
    });
  }

  // ---------- Inline editing ----------
  function editMonth(td, row) {
    const sel = document.createElement('select');
    monthAbbr.forEach(m => {
      const opt = document.createElement('option'); opt.value = m; opt.textContent = m;
      if (row.Month === m) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', () => {
      row.Month = sel.value;
      buildDataLembarKerja();
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
        // recalc include/exclude
        if (safeLower(row.Reman) === 'reman' && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
        else row.Include = row.Cost;
        row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === 'Escape') renderTable(dataLembarKerja);
    });
    inp.addEventListener('blur', () => {
      row.Reman = inp.value.trim();
      if (safeLower(row.Reman) === 'reman' && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
      else row.Include = row.Cost;
      row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
      buildDataLembarKerja();
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
        if (safeLower(row.Reman) === 'reman' && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
        else row.Include = row.Cost;
        row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
        buildDataLembarKerja();
        renderTable(dataLembarKerja);
        saveDataToLocalStorage();
      } else if (e.key === 'Escape') renderTable(dataLembarKerja);
    });
    inp.addEventListener('blur', () => {
      const v = inp.value.trim();
      row.Cost = v === '' ? '-' : Number(v);
      if (safeLower(row.Reman) === 'reman' && typeof row.Cost === 'number') row.Include = Number(row.Cost * 0.25);
      else row.Include = row.Cost;
      row.Exclude = safeLower(row.OrderType) === 'pm38' ? '-' : row.Include;
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
    });
    td.textContent = ''; td.appendChild(inp); inp.focus();
  }

  function openInlineEditorsForRow(row) {
    // find table row element by order
    const trs = Array.from(outputTableBody.querySelectorAll('tr'));
    for (const tr of trs) {
      const orderCell = tr.children[COL.ORDER];
      if (orderCell && safeLower(orderCell.textContent) === safeLower(row.Order)) {
        const monthTd = tr.children[COL.MONTH];
        const costTd = tr.children[COL.COST];
        const remanTd = tr.children[COL.REMAN];
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
          Order: safeStr(o), Room: '', OrderType: '', Description: '', CreatedOn: '', UserStatus: '',
          MAT: '', CPH: '', Section: '', StatusPart: '', Aging: '', Month: '', Cost: '-', Reman: '', Include: '-', Exclude: '-', Planning: '', StatusAMT: ''
        });
        added++;
      }
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
      saveDataToLocalStorage();
      let msg = `${added} order ditambahkan.`;
      if (skipped.length) msg += ` Sudah ada: ${skipped.join(', ')}.`;
      if (invalid.length) msg += ` Invalid: ${invalid.join(', ')}.`;
      addOrderStatus.style.color = added ? 'green' : 'red'; addOrderStatus.textContent = msg;
      addOrderInput.value = '';
    });
  } else console.warn('addOrder elements missing — Add disabled');

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
        buildDataLembarKerja(); renderTable(dataLembarKerja); alert('Data dimuat.');
      } catch (e) { alert('Gagal muat data: ' + (e && e.message ? e.message : e)); }
    });
  }

  function isValidOrder(order) { return !/[.,]/.test(order); }

  // ---------- Menu init (safe) ----------
  (function initMenu() {
    const menuItems = document.querySelectorAll('.menu-item'); const contentSections = document.querySelectorAll('.content-section');
    menuItems.forEach(item => {
      item.addEventListener('click', () => {
        menuItems.forEach(i => i.classList.remove('active')); contentSections.forEach(s => s.classList.remove('active'));
        item.classList.add('active'); const target = item.getAttribute('data-menu'); const sec = document.getElementById(target);
        if (sec) sec.classList.add('active');
      });
    });
  })();

  // ---------- Init load saved ----------
  (function init() {
    const saved = localStorage.getItem('lembarKerjaData');
    if (saved) {
      try { dataLembarKerja = JSON.parse(saved).map(r => ({ ...r, Order: safeStr(r.Order) })); }
      catch (e) { dataLembarKerja = []; }
    } else dataLembarKerja = [];
    buildDataLembarKerja(); renderTable(dataLembarKerja);
  })();

}); // end DOMContentLoaded
