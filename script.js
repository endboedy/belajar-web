// Simple interactions: menu switching, file name preview, and table filtering
document.addEventListener('DOMContentLoaded', () => {
  const menuItems = document.querySelectorAll('.menu-item');
  const panels = document.querySelectorAll('.panel');
  const panelMap = {};
  panels.forEach(p => panelMap[p.id] = p);

  menuItems.forEach(mi => {
    mi.addEventListener('click', () => {
      menuItems.forEach(x=>x.classList.remove('active'));
      mi.classList.add('active');
      const target = mi.dataset.target;
      // show/hide panels (simple)
      panels.forEach(p => p.style.display = (p.id === target ? 'block' : 'none'));
      // ensure first hero after header scrolls into view
      const hero = document.querySelector('.hero');
      if (hero) hero.scrollIntoView({behavior:'smooth'});
    });
  });

  // initialize: hide all except first (upload)
  panels.forEach(p => p.style.display = (p.id === 'upload' ? 'block' : 'none'));

  // file input change => show file name (append small span)
  document.querySelectorAll('.file-input').forEach(input => {
    const wrapper = document.createElement('span');
    wrapper.className = 'file-name';
    wrapper.style.marginLeft = '12px';
    wrapper.style.fontSize = '12px';
    wrapper.style.opacity = '.9';
    input.parentNode.appendChild(wrapper);

    input.addEventListener('change', (e) => {
      const file = e.target.files[0];
      wrapper.textContent = file ? file.name : '';
    });
  });

  // table filtering (simple contains on several columns)
  const table = document.getElementById('data-table');
  const filters = {
    room: document.getElementById('f-room'),
    order: document.getElementById('f-order'),
    mat: document.getElementById('f-mat'),
    section: document.getElementById('f-section'),
    cph: document.getElementById('f-cph')
  };

  function applyFilters(){
    const rows = table.querySelectorAll('tbody tr');
    const vals = {
      room: filters.room.value.trim().toLowerCase(),
      order: filters.order.value.trim().toLowerCase(),
      mat: filters.mat.value.trim().toLowerCase(),
      section: filters.section.value.trim().toLowerCase(),
      cph: filters.cph.value.trim().toLowerCase()
    };
    rows.forEach(row => {
      const cells = Array.from(row.children).map(td => td.textContent.toLowerCase());
      const matches =
        (!vals.room || cells[0].includes(vals.room)) &&
        (!vals.order || cells[2].includes(vals.order)) &&
        (!vals.mat || cells[6].includes(vals.mat)) &&
        (!vals.section || cells[8].includes(vals.section)) &&
        (!vals.cph || cells[7].includes(vals.cph));
      row.style.display = matches ? '' : 'none';
    });
  }

  Object.values(filters).forEach(inp => inp.addEventListener('input', applyFilters));
  document.getElementById('clear-filters').addEventListener('click', () => {
    Object.values(filters).forEach(i=>i.value='');
    applyFilters();
  });
});
