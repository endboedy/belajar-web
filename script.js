
// Include xlsx.js in your HTML before this script:
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

let IW39 = [];
let dataLembarKerja = [];

// Function to parse Excel file and populate IW39
function parseExcelToIW39(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    IW39.length = 0;
    jsonData.forEach(row => {
      IW39.push({
        Room: row.Room || "",
        OrderType: row.OrderType || "",
        Order: row.Order || "",
        Description: row.Description || "",
        CreatedOn: row.CreatedOn || "",
        UserStatus: row.UserStatus || "",
        MAT: row.MAT || "",
        TotalPlan: row.TotalPlan || 0,
        TotalActual: row.TotalActual || 0
      });
    });

    console.log("IW39 loaded:", IW39);
    buildDataLembarKerja();
    renderTable(dataLembarKerja);
  };
  reader.readAsArrayBuffer(file);
}

// Function to handle file upload
function updateDataFromUpload(fileName, fileObj) {
  if (fileName.toLowerCase().includes('iw39')) {
    parseExcelToIW39(fileObj);
  }
}

// Dummy manual input orders
let manualOrders = [
  { Order: "4700706921", Month: "Jan", Reman: "Reman" },
  { Order: "4700712211", Month: "Feb", Reman: "" },
  { Order: "4700719015", Month: "Mar", Reman: "Reman" }
];

// Function to build data for Lembar Kerja
function buildDataLembarKerja() {
  dataLembarKerja = manualOrders.map(row => {
    const iw = IW39.find(i => i.Order.toString().toLowerCase() === row.Order.toString().toLowerCase()) || {};

    const cost = ((iw.TotalPlan || 0) - (iw.TotalActual || 0)) / 16500;
    const finalCost = cost < 0 ? "-" : cost;
    const include = row.Reman === "Reman" ? (typeof cost === "number" ? cost * 0.25 : cost) : cost;
    const exclude = iw.OrderType === "PM38" ? "-" : include;

    return {
      Order: row.Order,
      Room: iw.Room || "",
      OrderType: iw.OrderType || "",
      Description: iw.Description || "",
      CreatedOn: iw.CreatedOn || "",
      UserStatus: iw.UserStatus || "",
      MAT: iw.MAT || "",
      CPH: iw.Description && iw.Description.startsWith("JR") ? "JR" : "", // Simplified logic
      Section: "", // Placeholder for lookup from Data1
      StatusPart: "", // Placeholder for lookup from SUM57
      Aging: "", // Placeholder for lookup from SUM57
      Month: row.Month,
      Cost: finalCost,
      Reman: row.Reman,
      Include: include,
      Exclude: exclude,
      Planning: "", // Placeholder for lookup from Planning
      StatusAMT: "" // Placeholder for lookup from Planning
    };
  });
}

// Function to render table (simplified)
function renderTable(data) {
  const table = document.getElementById("lembarKerjaTable");
  table.innerHTML = "";

  const header = "<tr>" + Object.keys(data[0] || {}).map(k => `<th>${k}</th>`).join("") + "</tr>";
  const rows = data.map(row => "<tr>" + Object.values(row).map(v => `<td>${v}</td>`).join("") + "</tr>").join("");

  table.innerHTML = header + rows;
}
