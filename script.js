//====================
// Navigasi Menu
//====================
            document.getElementById('btnUpload').addEventListener('click', ()=> showPage('pageUpload'));
            document.getElementById('btnLembar').addEventListener('click', ()=> showPage('pageLembar'));
            document.getElementById('btnSummary').addEventListener('click', ()=> showPage('pageSummary'));
            document.getElementById('btnDownload').addEventListener('click', downloadExcel);
            document.getElementById('btnProcess').addEventListener('click', processData);
            showPage('pageUpload');
                showPage('pageLembar');
                                      function showPage(id){

//====================
// Upload dan Baca File Excel
//====================
      function readExcelFile(file) {
              const reader = new FileReader();
                dataIW39 = await readExcelFile(document.getElementById('fileIW39').files[0]);
                data1 = await readExcelFile(document.getElementById('fileData1').files[0]);
                data2 = await readExcelFile(document.getElementById('fileData2').files[0]);
                dataSUM57 = await readExcelFile(document.getElementById('fileSUM57').files[0]);
                dataPlanning = await readExcelFile(document.getElementById('filePlanning').files[0]);

//====================
// Proses dan Merge Data
//====================
            // Main: processData -> mergeData -> render
            async function processData(){
                mergeData(iwNorm, d1Norm, d2Norm, sumNorm, planNorm);
            function mergeData(iw, d1, d2, sum, plan){

//====================
// Render Tabel dan Filter
//====================
                    renderTable([]);
                                    renderTable(mergedData);
                                function renderTable(data){
                                    function filterTable(){
                                          renderTable(filtered);

//====================
// Summary Perhitungan
//====================
                    renderSummary();
                                    renderSummary();
                                      function renderSummary(){

//====================
// Download ke Excel
//====================
                                      function downloadExcel(){

//====================
// Utilities
//====================
function normalizeKey(k){
            obj[normalizeKey(k)] = r[k];
    function findRowByKey(rowsNorm, keyName, value){
            if(r[normalizeKey(keyName)] && String(r[normalizeKey(keyName)]) === nv) return true;
      function asNumber(v){
                    const get = (k) => orig[normalizeKey(k)] ?? "";
                              const planVal = planKey ? asNumber(orig[planKey]) : asNumber(orig['totalsumplan'] || orig['plan'] || 0);
                              const actualVal = actualKey ? asNumber(orig[actualKey]) : asNumber(orig['totalactual'] || 0);

