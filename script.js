// script.js
document.addEventListener("DOMContentLoaded", () => {
    let iw39Data = [];
    let sum57Data = [];
    let planningData = [];
    let data1Data = [];
    let data2Data = [];

    // Input file elements
    const iw39Input = document.getElementById("iw39File");
    const sum57Input = document.getElementById("sum57File");
    const planningInput = document.getElementById("planningFile");
    const data1Input = document.getElementById("data1File");
    const data2Input = document.getElementById("data2File");

    const orderSearchInput = document.getElementById("orderSearch");
    const searchBtn = document.getElementById("searchBtn");
    const menu2TableBody = document.getElementById("menu2TableBody");

    // Fungsi baca Excel
    function readExcel(file, callback) {
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheet];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                callback(jsonData);
            } catch (err) {
                console.error("Gagal memproses file:", err);
                alert(`Error saat memproses file: ${file.name}`);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Event upload
    iw39Input?.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (data) => {
            iw39Data = data;
            console.log("IW39 loaded:", iw39Data.length);
        });
    });

    sum57Input?.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (data) => {
            sum57Data = data;
            console.log("SUM57 loaded:", sum57Data.length);
        });
    });

    planningInput?.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (data) => {
            planningData = data;
            console.log("Planning loaded:", planningData.length);
        });
    });

    data1Input?.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (data) => {
            data1Data = data;
            console.log("Data1 loaded:", data1Data.length);
        });
    });

    data2Input?.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (data) => {
            data2Data = data;
            console.log("Data2 loaded:", data2Data.length);
        });
    });

    // Lookup Menu 2
    searchBtn?.addEventListener("click", () => {
        const searchOrder = orderSearchInput.value.trim().toLowerCase();
        if (!searchOrder) {
            alert("Masukkan nomor order untuk mencari.");
            return;
        }

        // Gabungkan semua data
        const allData = [...iw39Data, ...sum57Data, ...planningData, ...data1Data, ...data2Data];

        // Filter by Order
        const filtered = allData.filter(i => {
            const orderValue = i.Order ? String(i.Order).toLowerCase() : "";
            return orderValue.includes(searchOrder);
        });

        // Tampilkan hasil
        renderMenu2Table(filtered);
    });

    function renderMenu2Table(data) {
        menu2TableBody.innerHTML = "";
        if (data.length === 0) {
            menu2TableBody.innerHTML = `<tr><td colspan="5" style="text-align:center;">Tidak ada data ditemukan</td></tr>`;
            return;
        }

        data.forEach(row => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${row.Order || ""}</td>
                <td>${row.Description || ""}</td>
                <td>${row.Status || ""}</td>
                <td>${row.Date || ""}</td>
                <td>${row.Remarks || ""}</td>
            `;
            menu2TableBody.appendChild(tr);
        });
    }
});
