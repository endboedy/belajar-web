// =======================
// script.js FINAL (Menu 1 & 2)
// =======================

// Global data store
let excelData = {};
let mergedData = [];
let currentMenu = 1;

// =======================
// Helper: Read Excel
// =======================
function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        workbook.SheetNames.forEach(sheetName => {
            if (sheetName.toLowerCase() === "budget") return; // skip budget

            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            excelData[sheetName] = jsonData;
        });

        callback();
    };
    reader.readAsArrayBuffer(file);
}

// =======================
// Helper: Merge Data
// =======================
function mergeData() {
    const IW39 = excelData["IW39"] || [];
    const SUM57 = excelData["SUM57"] || [];
    const Data1 = excelData["Data1"] || [];
    const Data2 = excelData["Data2"] || [];
    const Planning = excelData["Planning"] || [];

    mergedData = IW39.map(row => {
        const orderId = row["Order"]?.toString().trim();

        const sum57Row = SUM57.find(r => r["Order"]?.toString().trim() === orderId) || {};
        const data1Row = Data1.find(r => r["Order"]?.toString().trim() === orderId) || {};
        const data2Row = Data2.find(r => r["Order"]?.toString().trim() === orderId) || {};
        const planningRow = Planning.find(r => r["Order"]?.toString().trim() === orderId) || {};

        const cost = parseFloat(sum57Row["Cost"] || 0);
        const include = data1Row["Include"] || 0;
        const exclude = data2Row["Exclude"] || 0;

        return {
            Order: orderId,
            "Order Type": row["Order Type"] || "",
            Description: row["Description"] || "",
            "Created On": formatDate(row["Created On"]),
            Month: planningRow["Month"] || "",
            Reman: planningRow["Reman"] || "",
            Cost: cost,
            Include: include,
            Exclude: exclude,
            "Status AMT": planningRow["Status AMT"] || sum57Row["Status AMT"] || ""
        };
    });
}

// =======================
// Helper: Date Formatter
// =======================
function formatDate(value) {
    if (!value) return "";
    const date = new Date(value);
    if (isNaN(date)) return value;
    return date.toLocaleDateString("en-GB", { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');
}

// =======================
// Render Table (Menu 2)
// =======================
function renderTable() {
    const tableBody = document.querySelector("#outputTable tbody");
    tableBody.innerHTML = "";

    mergedData.forEach((row, index) => {
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td>${row["Order"]}</td>
            <td>${row["Order Type"]}</td>
            <td>${row["Description"]}</td>
            <td>${row["Created On"]}</td>
            <td>
                <select data-index="${index}" data-field="Month" class="edit-input">
                    ${["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                        .map(m => `<option value="${m}" ${row.Month === m ? "selected" : ""}>${m}</option>`).join("")}
                </select>
            </td>
            <td><input type="text" data-index="${index}" data-field="Reman" value="${row.Reman}" class="edit-input"></td>
            <td><input type="number" step="0.01" data-index="${index}" data-field="Cost" value="${row.Cost}" class="edit-input"></td>
            <td>${row.Include}</td>
            <td>${row.Exclude}</td>
            <td>${row["Status AMT"]}</td>
        `;

        tableBody.appendChild(tr);
    });
}

// =======================
// Event: Edit Table Inline
// =======================
document.addEventListener("input", function (e) {
    if (e.target.classList.contains("edit-input")) {
        const index = e.target.getAttribute("data-index");
        const field = e.target.getAttribute("data-field");
        let value = e.target.value;

        if (field === "Cost") {
            value = parseFloat(value) || 0;
        }

        mergedData[index][field] = value;
    }
});

// =======================
// Menu 1: Upload
// =======================
document.querySelector("#uploadExcel").addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;
    readExcel(file, () => {
        mergeData();
        if (currentMenu === 2) {
            renderTable();
        }
        alert("Excel uploaded and processed successfully!");
    });
});

// =======================
// Menu 2: Show Table
// =======================
document.querySelector("#menu2").addEventListener("click", function () {
    currentMenu = 2;
    mergeData();
    renderTable();
});
