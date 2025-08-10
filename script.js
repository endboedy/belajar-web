// Data global
let iwData = [];
let data1 = [];
let data2 = [];
let planningData = [];
let currentMenu = null;

// Jalankan setelah DOM siap
document.addEventListener('DOMContentLoaded', () => {
    // Ambil elemen tombol menu
    const menu1Btn = document.getElementById('menu1Btn');
    const menu2Btn = document.getElementById('menu2Btn');

    // Menu 1
    if (menu1Btn) {
        menu1Btn.addEventListener('click', () => {
            currentMenu = 'menu1';
            document.getElementById('menu1').style.display = 'block';
            document.getElementById('menu2').style.display = 'none';
        });
    }

    // Menu 2
    if (menu2Btn) {
        menu2Btn.addEventListener('click', () => {
            currentMenu = 'menu2';
            document.getElementById('menu1').style.display = 'none';
            document.getElementById('menu2').style.display = 'block';
        });
    }

    // Upload file IW39
    const iw39Input = document.getElementById('iw39File');
    if (iw39Input) {
        iw39Input.addEventListener('change', (e) => {
            handleFileUpload(e.target.files[0], iwData);
        });
    }

    // Upload file Data1
    const data1Input = document.getElementById('data1File');
    if (data1Input) {
        data1Input.addEventListener('change', (e) => {
            handleFileUpload(e.target.files[0], data1);
        });
    }

    // Upload file Data2
    const data2Input = document.getElementById('data2File');
    if (data2Input) {
        data2Input.addEventListener('change', (e) => {
            handleFileUpload(e.target.files[0], data2);
        });
    }

    // Upload file Planning
    const planningInput = document.getElementById('planningFile');
    if (planningInput) {
        planningInput.addEventListener('change', (e) => {
            handleFileUpload(e.target.files[0], planningData);
        });
    }

    // Tombol Cari Order (Menu 2)
    const orderSearchBtn = document.getElementById('orderSearchBtn');
    if (orderSearchBtn) {
        orderSearchBtn.addEventListener('click', () => {
            const orderVal = document.getElementById('orderInput').value.trim();
            if (orderVal === '') {
                alert('Masukkan nomor Order');
                return;
            }
            applyExternalFilters(orderVal);
        });
    }
});

// Fungsi upload file (read Excel)
function handleFileUpload(file, targetArray) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Ambil sheet pertama
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Konversi ke array objek
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        // Simpan data
        targetArray.length = 0;
        targetArray.push(...jsonData);
        console.log('Data loaded:', targetArray);
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi filter & render tabel (Menu 2)
function applyExternalFilters(orderNumber) {
    // Cari data IW39 sesuai Order
    const iwFiltered = iwData.filter(row => String(row.Order || '').toLowerCase() === orderNumber.toLowerCase());

    // Gabungkan data lain berdasarkan key
    const mergedData = iwFiltered.map(row => {
        const sectionInfo = data1.find(d => d.Room === row.Room) || {};
        const statusPartInfo = data2.find(d => d.Order === row.Order) || {};
        const planningInfo = planningData.find(d => d.Order === row.Order) || {};

        return {
            Room: row.Room || '',
            Order: row.Order || '',
            OrderType: row['Order Type'] || '',
            Description: row.Description || '',
            CreatedOn: row['Created On'] || '',
            UserStatus: row['User Status'] || '',
            MAT: row.MAT || '',
            CPH: row.CPH || '',
            Section: sectionInfo.Section || '',
            StatusPart: statusPartInfo['Status Part'] || '',
            Aging: row.Aging || '',
            Month: planningInfo.Month || '',
            Cost: planningInfo.Cost || '',
            Reman: planningInfo.Reman || '',
            Include: planningInfo.Include || '',
            Exclude: planningInfo.Exclude || '',
            Planning: planningInfo.Planning || '',
            StatusAMT: planningInfo['Status AMT'] || ''
        };
    });

    renderTable(mergedData);
}

// Fungsi render tabel ke HTML
function renderTable(data) {
    let container = document.getElementById('tableContainer');

    // Jika container tidak ada, buat otomatis
    if (!container) {
        const menu2Section = document.getElementById('menu2');
        container = document.createElement('div');
        container.id = 'tableContainer';
        menu2Section.appendChild(container);
    }

    if (!data.length) {
        container.innerHTML = '<p style="color:red">Tidak ada data ditemukan.</p>';
        return;
    }

    // Buat tabel HTML
    let html = '<table class="data-table"><thead><tr>';
    html += '<th>Room</th><th>Order</th><th>Order Type</th><th>Description</th><th>Created On</th><th>User Status</th><th>MAT</th><th>CPH</th><th>Section</th><th>Status Part</th><th>Aging</th><th>Month</th><th>Cost</th><th>Reman</th><th>Include</th><th>Exclude</th><th>Planning</th><th>Status AMT</th>';
    html += '</tr></thead><tbody>';

    data.forEach(row => {
        html += '<tr>';
        html += `<td>${row.Room}</td>`;
        html += `<td>${row.Order}</td>`;
        html += `<td>${row.OrderType}</td>`;
        html += `<td>${row.Description}</td>`;
        html += `<td>${row.CreatedOn}</td>`;
        html += `<td>${row.UserStatus}</td>`;
        html += `<td>${row.MAT}</td>`;
        html += `<td>${row.CPH}</td>`;
        html += `<td>${row.Section}</td>`;
        html += `<td>${row.StatusPart}</td>`;
        html += `<td>${row.Aging}</td>`;
        html += `<td>${row.Month}</td>`;
        html += `<td style="text-align:right">${row.Cost}</td>`;
        html += `<td>${row.Reman}</td>`;
        html += `<td style="text-align:right">${row.Include}</td>`;
        html += `<td style="text-align:right">${row.Exclude}</td>`;
        html += `<td>${row.Planning}</td>`;
        html += `<td>${row.StatusAMT}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}
