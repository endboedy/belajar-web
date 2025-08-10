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
