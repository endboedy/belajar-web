
// Include xlsx.js via HTML: <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

let IW39 = [], SUM57 = [], Data1 = [], Data2 = [], Planning = [];
let dataLembarKerja = [];

function parseExcel(file, callback) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    callback(jsonData);
  };
  reader.readAsArrayBuffer(file);
}

function updateDataFromUpload(fileName, fileObj) {
  const lowerName = fileName.toLowerCase();
  if (lowerName.includes('iw39')) {
    parseExcel(fileObj, (data) => {
      IW39 = data;
      buildDataLembarKerja();
      renderTable(dataLembarKerja);
    });
  } else if (lowerName.includes('sum57')) {
    parseExcel(fileObj, (data) => { SUM57 = data; });
  } else if (lowerName.includes('planning')) {
    parseExcel(fileObj, (data) => { Planning = data; });
  } else if (lowerName.includes('data1')) {
    parseExcel(fileObj, (data) => { Data1 = data; });
  } else if (lowerName.includes('data2')) {
    parseExcel(fileObj, (data) => { Data2 = data; });
  }
}

function addManualOrder(order) {
  dataLembarKerja.push({ Order: order });
  buildDataLembarKerja();
  renderTable(dataLembarKerja);
}

function buildDataLembarKerja() {
  dataLembarKerja = dataLembarKerja.map(row => {
    const iw = IW39.find(i => i.Order?.toString().toLowerCase() === row.Order?.toString().toLowerCase()) || {};
    const sum = SUM57.find(s => s.Order?.toString().toLowerCase() === row.Order?.toString().toLowerCase()) || {};
    const plan = Planning.find(p => p.Order?.toString().toLowerCase() === row.Order?.toString().toLowerCase()) || {};
    const matKey = iw.MAT || '';
    const data2 = Data2.find(d => d.MAT?.toString().toLowerCase() === matKey.toLowerCase()) || {};
    const data1 = Data1.find(d => d.Order?.toString().toLowerCase() === row.Order?.toString().toLowerCase()) || {};

    const cph = iw.Description?.startsWith("JR") ? "JR" : (data2.CPH || "");
    const cost = ((iw.TotalPlan || 0) - (iw.TotalActual || 0)) / 16500;
    const finalCost = cost < 0 ? "-" : cost;
    const include = row.Reman === "Reman" ? cost * 0.25 : cost;
    const exclude = iw.OrderType === "PM38" ? "-" : include;

    return {
      ...row,
      Room: iw.Room || "",
      OrderType: iw.OrderType || "",
      Description: iw.Description || "",
      CreatedOn: iw.CreatedOn || "",
      UserStatus: iw.UserStatus || "",
      MAT: iw.MAT || "",
      CPH: cph,
      Section: data1.Section || "",
      StatusPart: sum.StatusPart || "",
      Aging: sum.Aging || "",
      Month: row.Month || "",
      Cost: finalCost,
      Reman: row.Reman || "",
      Include: include,
      Exclude: exclude,
      Planning: plan.EventStart || "",
      StatusAMT: plan.StatusAMT || ""
    };
  });
}

function renderTable(data) {
  const table = document.getElementById("lembarKerjaTable");
  table.innerHTML = "";
  if (data.length === 0) return;

  const headers = Object.keys(data[0]);
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  data.forEach((row, idx) => {
    const tr = document.createElement("tr");
    headers.forEach(h => {
      const td = document.createElement("td");
      td.textContent = row[h];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

function saveToLocalStorage() {
  localStorage.setItem("lembarKerjaData", JSON.stringify(dataLembarKerja));
}

function loadFromLocalStorage() {
  const saved = localStorage.getItem("lembarKerjaData");
  if (saved) {
    dataLembarKerja = JSON.parse(saved);
    renderTable(dataLembarKerja);
  }
}

function filterTable(keyword) {
  const filtered = dataLembarKerja.filter(row =>
    Object.values(row).some(val => val?.toString().toLowerCase().includes(keyword.toLowerCase()))
  );
  renderTable(filtered);
}
