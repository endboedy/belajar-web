let dataIW39, data1, data2, dataSUM57, dataPlanning, mergedData = [];

function readExcel(fileInput, callback) {
    const file = fileInput.files[0];
    if (!file) return callback([]);
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        callback(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsArrayBuffer(file);
}

function processData() {
    readExcel(document.getElementById('fileIW39'), (iw39) => {
        dataIW39 = iw39;
        readExcel(document.getElementById('fileData1'), (d1) => {
            data1 = d1;
            readExcel(document.getElementById('fileData2'), (d2) => {
                data2 = d2;
                readExcel(document.getElementById('fileSUM57'), (sum) => {
                    dataSUM57 = sum;
                    readExcel(document.getElementById('filePlanning'), (plan) => {
                        dataPlanning = plan;
                        mergeData();
                    });
                });
            });
        });
    });
}

function mergeData() {
    const month = document.getElementById('monthInput').value;
    const reman = document.getElementById('remanInput').value;
    mergedData = dataIW39.map(row => {
        const MAT = row.MAT;
        const Order = row.Order;

        // Lookup Section dari Data1
        const secRow = data1.find(r => r.MAT === MAT);
        const section = secRow ? secRow.Section : "";

        // Lookup CPH dari Data2
        const cphRow = data2.find(r => r.MAT === MAT);
        const cph = cphRow ? cphRow.CPH : "";

        // Lookup Status Part & Aging dari SUM57
        const sumRow = dataSUM57.find(r => r.Order === Order);
        const statusPart = sumRow ? sumRow["Status Part"] : "";
        const aging = sumRow ? sumRow.Aging : "";

        // Lookup Planning & Status AMT
        const planRow = dataPlanning.find(r => r.Order === Order);
        const planning = planRow ? planRow.Planning : "";
        const statusAMT = planRow ? planRow["Status AMT"] : "";

        // Description rule
        let description = "";
        if (row.Description && row.Description.startsWith("JR")) {
            description = "JR";
        } else {
            description = secRow ? secRow.Description : "";
        }

        // Cost calculation
        let cost = (row.TotalPlan - row.TotalActual) / 16500;
        if (cost < 0) cost = "-";

        // Include
        let include = reman.includes("Reman") ? (typeof cost === "number" ? cost * 0.25 : cost) : cost;

        // Exclude
        let exclude = row["Order Type"] === "PM38" ? "-" : include;

        return {
            Room: row.Room,
            "Order Type": row["Order Type"],
            Order,
            Description: description,
            "Created On": row["Created On"],
            "User Status": row["User Status"],
            MAT,
            CPH: cph,
            Section: section,
            "Status Part": statusPart,
            Aging: aging,
            Month: month,
            Cost: cost,
            Reman: reman,
            Include: include,
            Exclude: exclude,
            Planning: planning,
            "Status AMT": statusAMT
        };
    });

    renderTable(mergedData, "tableContainer");
    renderSummary();
}

function renderTable(data, containerId) {
    let html = "<table><tr>";
    Object.keys(data[0]).forEach(col => html += `<th>${col}</th>`);
    html += "</tr>";
    data.forEach(row => {
        html += "<tr>";
        Object.values(row).forEach(val => html += `<td>${val}</td>`);
        html += "</tr>";
    });
    html += "</table>";
    document.getElementById(containerId).innerHTML = html;
}

function renderSummary() {
    let totalCost = mergedData.reduce((sum, row) => typeof row.Cost === "number" ? sum + row.Cost : sum, 0);
    document.getElementById("summaryContainer").innerHTML = `<p>Total Cost: ${totalCost.toFixed(2)}</p>`;
}

function showPage(pageId) {
    document.querySelectorAll(".page").forEach(p => p.style.display = "none");
    document.getElementById(pageId).style.display = "block";
}

function downloadExcel() {
    const ws = XLSX.utils.json_to_sheet(mergedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    XLSX.writeFile(wb, "Hasil_Merge.xlsx");
}
function filterTable() {
    const roomVal = document.getElementById('filterRoom').value.toLowerCase();
    const orderVal = document.getElementById('filterOrder').value.toLowerCase();
    const matVal = document.getElementById('filterMAT').value.toLowerCase();
    const sectionVal = document.getElementById('filterSection').value.toLowerCase();
    const cphVal = document.getElementById('filterCPH').value.toLowerCase();

    const filtered = mergedData.filter(row => 
        String(row.Room || "").toLowerCase().includes(roomVal) &&
        String(row.Order || "").toLowerCase().includes(orderVal) &&
        String(row.MAT || "").toLowerCase().includes(matVal) &&
        String(row.Section || "").toLowerCase().includes(sectionVal) &&
        String(row.CPH || "").toLowerCase().includes(cphVal)
    );

    renderTable(filtered, "tableContainer");
