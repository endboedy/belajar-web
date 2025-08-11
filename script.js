// =============================
// Variabel Global
// =============================
let iw39Data = [];
let sum57Data = [];
let planningData = [];
let data1 = [];
let data2 = [];
let data3 = []; // kalau memang ada

let uploadedFiles = {}; // tracking file yang sudah diupload

// =============================
// Helper: Baca file Excel
// =============================
function readExcelFile(file, callback) {
    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet);
        callback(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// =============================
// Upload File Handler
// =============================
document.getElementById("fileUpload").addEventListener("change", function(event) {
    let files = event.target.files;
    Array.from(files).forEach(file => {
        let name = file.name.toLowerCase();
        if (name.includes("iw39")) {
            readExcelFile(file, (data) => {
                iw39Data = data;
                uploadedFiles.iw39 = true;
                console.log("IW39 Loaded", iw39Data);
            });
        } else if (name.includes("sum57")) {
            readExcelFile(file, (data) => {
                sum57Data = data;
                uploadedFiles.sum57 = true;
                console.log("SUM57 Loaded", sum57Data);
            });
        } else if (name.includes("planning")) {
            readExcelFile(file, (data) => {
                planningData = data;
                uploadedFiles.planning = true;
                console.log("Planning Loaded", planningData);
            });
        } else if (name.includes("data1")) {
            readExcelFile(file, (data) => {
                data1 = data;
                uploadedFiles.data1 = true;
            });
        } else if (name.includes("data2")) {
            readExcelFile(file, (data) => {
                data2 = data;
                uploadedFiles.data2 = true;
            });
        } else if (name.includes("data3")) {
            readExcelFile(file, (data) => {
                data3 = data;
                uploadedFiles.data3 = true;
            });
        }
    });
});

// =============================
// Lookup & Bangun Tabel Menu 2
// =============================
function buildDataLembarKerja(orderInput) {
    let orderKey = String(orderInput || "").toLowerCase();

    let result = [];

    if (uploadedFiles.iw39) {
        let iwMatch = iw39Data.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (iwMatch) result.push({ source: "IW39", ...iwMatch });
    }
    if (uploadedFiles.sum57) {
        let sumMatch = sum57Data.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (sumMatch) result.push({ source: "SUM57", ...sumMatch });
    }
    if (uploadedFiles.planning) {
        let planMatch = planningData.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (planMatch) result.push({ source: "Planning", ...planMatch });
    }
    if (uploadedFiles.data1) {
        let match1 = data1.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (match1) result.push({ source: "Data1", ...match1 });
    }
    if (uploadedFiles.data2) {
        let match2 = data2.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (match2) result.push({ source: "Data2", ...match2 });
    }
    if (uploadedFiles.data3) {
        let match3 = data3.find(i => String(i.Order || "").toLowerCase() === orderKey);
        if (match3) result.push({ source: "Data3", ...match3 });
    }

    // Render ke tabel
    let tbody = document.querySelector("#lembarKerjaTable tbody");
    tbody.innerHTML = "";

    result.forEach((row, index) => {
        let tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${row.source}</td>
            <td>${row.Order || ""}</td>
            <td>${row.Description || ""}</td>
            <td>${row.Date || ""}</td>
            <td>
                <button onclick="editRow(${index})">Edit</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// =============================
// Edit Row
// =============================
function editRow(index) {
    alert("Edit row index: " + index);
    // bisa tambahkan modal/form untuk update data
}

// =============================
// Load Data dari Storage (opsional)
// =============================
function loadDataFromStorage() {
    let orderInput = document.getElementById("orderInput").value;
    if (!orderInput) {
        alert("Masukkan nomor Order");
        return;
    }
    buildDataLembarKerja(orderInput);
}

document.getElementById("btnLookup").addEventListener("click", loadDataFromStorage);
