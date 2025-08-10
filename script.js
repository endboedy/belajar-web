let tableData = [];

function showMenu(num) {
  document.querySelectorAll('.menu-section').forEach(sec => sec.classList.add('hidden'));
  document.getElementById(`menu${num}`).classList.remove('hidden');
}

document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    tableData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    renderTable();
    showMenu(2);
  };
  reader.readAsArrayBuffer(file);
}

function renderTable() {
  if (tableData.length === 0) return;
  const headerRow = document.getElementById('headerRow');
  const filterRow = document.getElementById('filterRow');
  const tbody = document.querySelector('#dataTable tbody');

  headerRow.innerHTML = '';
  filterRow.innerHTML = '';
  tbody.innerHTML = '';

  // Header
  tableData[0].forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headerRow.appendChild(th);

    const td = document.createElement('td');
    const input = document.createElement('input');
    input.classList.add('table-filter');
    input.dataset.col = h;
    input.addEventListener('input', applyFilters);
    td.appendChild(input);
    filterRow.appendChild(td);
  });

  // Body
  const orderSet = new Set();
  for (let i = 1; i < tableData.length; i++) {
    const row = tableData[i];
    const tr = document.createElement('tr');

    row.forEach((cell, idx) => {
      const td = document.createElement('td');
      td.textContent = cell;

      // Highlight duplicate in column "Order"
      if (tableData[0][idx].toLowerCase() === 'order') {
        if (orderSet.has(cell)) {
          td.classList.add('duplicate');
        } else {
          orderSet.add(cell);
        }
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
}

function applyFilters() {
  const filters = Array.from(document.querySelectorAll('.table-filter')).map(inp => inp.value.toLowerCase());
  const tbody = document.querySelector('#dataTable tbody');
  tbody.innerHTML = '';

  const orderSet = new Set();
  for (let i = 1; i < tableData.length; i++) {
    const row = tableData[i];
    let match = true;

    row.forEach((cell, idx) => {
      if (filters[idx] && !String(cell).toLowerCase().includes(filters[idx])) {
        match = false;
      }
    });

    if (match) {
      const tr = document.createElement('tr');
      row.forEach((cell, idx) => {
        const td = document.createElement('td');
        td.textContent = cell;
        if (tableData[0][idx].toLowerCase() === 'order') {
          if (orderSet.has(cell)) {
            td.classList.add('duplicate');
          } else {
            orderSet.add(cell);
          }
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
  }
}
