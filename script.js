
function showTab(tabId) {
  const tabs = document.querySelectorAll('.tab-content');
  tabs.forEach(tab => tab.style.display = 'none');
  document.getElementById(tabId).style.display = 'block';
}

async function loadExcelFiles() {
  const files = ['IW39.xlsx', 'SUM57.xlsx', 'Planning.xlsx', 'Budget.xlsx', 'Data1.xlsx', 'Data2.xlsx'];
  const baseURL = 'https://raw.githubusercontent.com/endboedy/belajar-web/main/excel/';
  const previewDiv = document.getElementById('excel-preview');
  previewDiv.innerHTML = '';

  for (const file of files) {
    const url = baseURL + file;
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Status ${response.status}`);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const html = XLSX.utils.sheet_to_html(sheet);
      const section = document.createElement('div');
      section.innerHTML = `<h3>${file}</h3>` + html;
      previewDiv.appendChild(section);
    } catch (error) {
      const errorMsg = document.createElement('p');
      errorMsg.textContent = `Gagal memuat ${file}: ${error}`;
      previewDiv.appendChild(errorMsg);
    }
  }
}

function handleUpload(fileKey) {
  const input = document.querySelector(`input[name="${fileKey}"]`);
  const file = input.files[0];
  const status = document.getElementById(`status-${fileKey}`);

  if (!file) {
    status.textContent = `${fileKey}: file belum dipilih`;
    status.style.color = 'red';
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const html = XLSX.utils.sheet_to_html(sheet);
      const previewDiv = document.getElementById('excel-preview');
      const section = document.createElement('div');
      section.innerHTML = `<h3>${fileKey}</h3>` + html;
      previewDiv.appendChild(section);

      status.textContent = `${fileKey}: sukses upload`;
      status.style.color = fileKey === 'Budget' ? '#004080' : 'green';
    } catch (err) {
      status.textContent = `${fileKey}: gagal upload`;
      status.style.color = 'red';
    }
  };
  reader.readAsArrayBuffer(file);
}

function uploadAll() {
  const fileKeys = ['IW39', 'SUM57', 'Planning', 'Budget', 'Data1', 'Data2'];
  fileKeys.forEach(key => handleUpload(key));
}
