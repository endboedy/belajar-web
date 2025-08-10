// Sidebar menu switching
const menuItems = document.querySelectorAll('.menu-item');
const sections = document.querySelectorAll('.content-section');

menuItems.forEach(item => {
  item.addEventListener('click', () => {
    // remove active class from all
    menuItems.forEach(i => i.classList.remove('active'));
    sections.forEach(s => s.classList.remove('active'));

    // add active to clicked and corresponding section
    item.classList.add('active');
    const target = item.dataset.menu;
    document.getElementById(target).classList.add('active');
  });
});

// Upload simulation
const uploadBtn = document.getElementById('upload-btn');
const fileInput = document.getElementById('file-input');
const fileSelect = document.getElementById('file-select');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('upload-progress');
const progressText = document.getElementById('progress-text');
const uploadStatus = document.getElementById('upload-status');

uploadBtn.addEventListener('click', () => {
  const file = fileInput.files[0];
  const selectedFileType = fileSelect.value;

  if (!file) {
    alert('Pilih file terlebih dahulu ya bro!');
    return;
  }

  // Disable button saat upload
  uploadBtn.disabled = true;
  uploadStatus.textContent = '';
  progressBar.value = 0;
  progressText.textContent = '0%';
  progressContainer.classList.remove('hidden');

  // Simulasi upload dengan progress
  let progress = 0;
  const interval = setInterval(() => {
    progress += Math.floor(Math.random() * 15) + 5; // naikin progress acak 5-20%
    if (progress >= 100) {
      progress = 100;
      clearInterval(interval);
      uploadStatus.textContent = `File "${file.name}" untuk kategori ${selectedFileType} berhasil diupload! 🎉`;
      uploadBtn.disabled = false;
      fileInput.value = '';
      progressContainer.classList.add('hidden');
    }
    progressBar.value = progress;
    progressText.textContent = progress + '%';
  }, 300);
});

// Filter button dummy (belum ada data nyata)
const filterBtn = document.getElementById('filter-btn');
filterBtn.addEventListener('click', () => {
  alert('Fitur filter sedang dalam pengembangan bro...');
});
