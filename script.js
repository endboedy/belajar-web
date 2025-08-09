document.addEventListener("DOMContentLoaded", function() {
  const content = document.getElementById("content");

  const pages = {
    home: `<h2>🏠 Halaman Home</h2><p>Ini adalah halaman utama.</p>`,
    submenu1: `<h2>📄 Submenu 1</h2><p>Konten untuk submenu 1.</p>`,
    subsubmenu1: `<h2>📄 Sub Submenu 1</h2><p>Konten untuk sub submenu 1.</p>`,
    subsubmenu2: `<h2>📄 Sub Submenu 2</h2><p>Konten untuk sub submenu 2.</p>`,
    menu2: `<h2>📄 Menu 2</h2><p>Konten untuk menu 2.</p>`
  };

  document.querySelectorAll("a[data-page]").forEach(link => {
    link.addEventListener("click", function(e) {
      e.preventDefault();
      const page = this.getAttribute("data-page");
      content.innerHTML = pages[page] || `<h2>404</h2><p>Halaman tidak ditemukan.</p>`;
    });
  });
});
