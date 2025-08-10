// Data global
let iwData = [];
let data1 = [];
let data2 = [];
let dataPlanning = [];

// Ambil data dari file Excel
async function loadData() {
    const [iwResp, data1Resp, data2Resp, planningResp] = await Promise.all([
        fetch('IW39.json'),
        fetch('Data1.json'),
        fetch('SUM57.json'),
        fetch('Planning.json')
    ]);

    iwData = await iwResp.json();
    data1 = await data1Resp.json();
    data2 = await data2Resp.json();
    dataPlanning = await planningResp.json();

    renderTable(iwData);
}

// Render tabel
function renderTable(data) {
    const tableContainer = document.getElementById('table-container');
    tableContainer.innerHTML = '';

    let table = document.createElement('table');
    table.classList.add('data-table');
    table.style.width = '100%';

    // Header
    const headerRow = document.createElement('tr');
    const headers = [
        "Room", "Order", "Order Type", "Order Description", "Created On", "User Status", 
        "MAT", "CPH", "Section", "Status Part", "Aging", "Month", "Cost", "Reman", 
        "Include", "Exclude", "Planning", "Status AMT"
    ];

    headers.forEach(h => {
        let th = document.createElement('th');
        th.textContent = h;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Filter row
    const filterRow = document.createElement('tr');
    headers.forEach((_, i) => {
        let th = document.createElement('th');
        let input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Filter...';
        input.style.width = '100%';
        input.addEventListener('input', () => filterTable());
        th.appendChild(input);
        filterRow.appendChild(th);
    });
    table.appendChild(filterRow);

    // Isi data
    const orderCounts = {};
    data.forEach(row => {
        orderCounts[row.Order] = (orderCounts[row.Order] || 0) + 1;
    });

    data.forEach(row => {
        const tr = document.createElement('tr');

        headers.forEach(h => {
            let td = document.createElement('td');

            let value = row[h.replace(/\s+/g, '')] || '';

            // Isi kolom tambahan dari sumber lain
            if (h === "Section") {
                let d1 = data1.find(d => d.Room === row.Room);
                value = d1 ? d1.Section : '';
            }
            if (h === "Status Part") {
                let d2 = data2.find(d => d.Order === row.Order);
                value = d2 ? d2.StatusPart : '';
            }
            if (h === "Status AMT") {
                let plan = dataPlanning.find(d => d.Order === row.Order);
                value = plan ? plan.StatusAMT : '';
            }

            td.textContent = value;

            // Warna merah font putih jika order duplikat
            if (h === "Order" && orderCounts[row.Order] > 1) {
                td.style.backgroundColor = 'red';
                td.style.color = 'white';
                td.style.fontWeight = 'bold';
            }

            // Format angka rata kanan
            if (["Cost", "Include", "Exclude"].includes(h)) {
                td.style.textAlign = 'right';
            }

            tr.appendChild(td);
        });

        table.appendChild(tr);
    });

    tableContainer.appendChild(table);
}

// Filter tabel real-time
function filterTable() {
    const table = document.querySelector('.data-table');
    const filters = Array.from(table.querySelectorAll('tr:nth-child(2) input')).map(input => input.value.toLowerCase());

    Array.from(table.querySelectorAll('tr')).slice(2).forEach(row => {
        let cells = row.querySelectorAll('td');
        let match = true;

        cells.forEach((cell, i) => {
            if (filters[i] && !cell.textContent.toLowerCase().includes(filters[i])) {
                match = false;
            }
        });

        row.style.display = match ? '' : 'none';
    });
}

// Sticky header & filter
document.addEventListener('scroll', () => {
    const ths = document.querySelectorAll('.data-table th');
    ths.forEach(th => th.style.position = 'sticky');
});

// Load data awal
document.addEventListener('DOMContentLoaded', loadData);
