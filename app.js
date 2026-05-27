let uploadedFiles = [];
let aggregatedData = [];
let filteredData = [];
let selectedKeys = new Set();
let isCalculating = false;
let dayChangeTimeout = null;
const FREQ_KEY = 'uab_product_frequencies_v1';
const STOCK_KEY = 'uab_stock_entries_v1';
let stockEntries = [];

// DOM Elements
const fileInput = document.getElementById('fileInput');
const fileStatus = document.getElementById('fileStatus');
const daysSelector = document.getElementById('daysSelector');
const daysCheckboxes = document.getElementById('daysCheckboxes');
const calcBtn = document.getElementById('calcBtn');
const resultsSection = document.getElementById('resultsSection');
const resultsTableBody = document.querySelector('#resultsTable tbody');
const selectAll = document.getElementById('selectAll');
const refreshBtn = document.getElementById('refreshBtn');
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');
const orderTableBody = document.querySelector('#orderTable tbody');
const emptyOrder = document.getElementById('emptyOrder');
const orderFooter = document.getElementById('orderFooter');
const exportBtn = document.getElementById('exportBtn');
const stockBtn = document.getElementById('stockBtn');
const stockModal = document.getElementById('stockModal');
const closeStockModal = document.getElementById('closeStockModal');
const stockProduct = document.getElementById('stockProduct');
const stockUnit = document.getElementById('stockUnit');
const stockQty = document.getElementById('stockQty');
const addStockBtn = document.getElementById('addStockBtn');
const stockTableBody = document.querySelector('#stockTable tbody');
const emptyStock = document.getElementById('emptyStock');
const clearStockBtn = document.getElementById('clearStockBtn');
const stockCount = document.getElementById('stockCount');
const freqBtn = document.getElementById('freqBtn');
const freqModal = document.getElementById('freqModal');
const closeFreqModal = document.getElementById('closeFreqModal');
const clearFreqBtn = document.getElementById('clearFreqBtn');
const freqTableBody = document.querySelector('#freqTable tbody');
const emptyFreq = document.getElementById('emptyFreq');

// Utilidad: Formato Título
function toTitleCase(str) {
  return String(str).toLowerCase().replace(/(?:^|\s)\S/g, a => a.toUpperCase());
}

// Función para limpiar decimales
function formatQty(val) {
  if (val === undefined || val === null || isNaN(val)) return '0';
  const num = parseFloat(val);
  return num % 1 === 0 ? num.toString() : num.toFixed(2).replace(/\.?0+$/, '');
}

function normalize(str) { 
  return String(str).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " "); 
}

// 1. Carga múltiple
fileInput.addEventListener('change', (e) => {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  fileStatus.textContent = '⏳ Cargando archivos...';
  uploadedFiles = files.map(f => ({ name: f.name, workbook: null }));
  selectedKeys.clear();

  let loaded = 0;
  files.forEach((file, idx) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try { uploadedFiles[idx].workbook = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' }); } 
      catch { console.warn(`Error en ${file.name}`); }
      if (++loaded === files.length) finalizeLoad(files);
    };
    reader.readAsArrayBuffer(file);
  });
});

function finalizeLoad(files) {
  uploadedFiles = uploadedFiles.filter(f => f.workbook);
  fileStatus.textContent = uploadedFiles.length ? `✅ ${uploadedFiles.length} archivo(s) cargados` : '❌ Ningún archivo válido';
  loadStockEntries(); // Cargar stock guardado
  populateDays();
  setupEventListeners();
  daysSelector.classList.remove('hidden');
  resultsSection.classList.add('hidden');
  emptyOrder.classList.remove('hidden');
  orderFooter.classList.add('hidden');
}

// 2. Poblar días
function populateDays() {
  daysCheckboxes.innerHTML = '';
  const days = new Set();
  uploadedFiles.forEach(wb => wb.workbook.SheetNames.forEach(s => days.add(s.trim().toUpperCase())));
  const order = ['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO','DOMINGO'];
  [...days].sort((a,b) => (order.indexOf(a)+1) - (order.indexOf(b)+1) || a.localeCompare(b)).forEach(day => {
    daysCheckboxes.innerHTML += `<div><input type="checkbox" id="day-${day}" value="${day}" checked><label for="day-${day}">${day}</label></div>`;
  });
}

// Configurar listeners
function setupEventListeners() {
  // Cambio de días con debounce
  daysCheckboxes.addEventListener('change', () => {
    clearTimeout(dayChangeTimeout);
    dayChangeTimeout = setTimeout(() => {
      if (!isCalculating) handleDayChange();
    }, 50);
  });

  // Búsqueda en tiempo real
  searchInput.addEventListener('input', (e) => {
    const searchTerm = normalize(e.target.value);
    filterProducts(searchTerm);
  });

  // Checkboxes de productos
  resultsTableBody.addEventListener('change', (e) => {
    if (e.target.type === 'checkbox' && e.target.dataset.key) {
      const key = e.target.dataset.key;
      if (selectedKeys.has(key)) selectedKeys.delete(key); else selectedKeys.add(key);
      updateOrderPanel();
      updateSelectAllCheckbox();
    }
  });

  // Seleccionar todo
  selectAll.addEventListener('change', () => {
    const checkboxes = resultsTableBody.querySelectorAll('input[type="checkbox"][data-key]');
    checkboxes.forEach(cb => {
      cb.checked = selectAll.checked;
      const key = cb.dataset.key;
      selectAll.checked ? selectedKeys.add(key) : selectedKeys.delete(key);
    });
    updateOrderPanel();
  });

  // Botón Calcular
  calcBtn.addEventListener('click', () => {
    const selectedDays = [...document.querySelectorAll('#daysCheckboxes input:checked')].map(c => c.value);
    if (!selectedDays.length) return alert('Selecciona al menos un día.');
    
    calcBtn.disabled = true; calcBtn.textContent = '⏳ Procesando...';
    setTimeout(() => {
      isCalculating = true;
      calculate(selectedDays);
      renderTable();
      resultsSection.classList.remove('hidden');
      setTimeout(() => { isCalculating = false; calcBtn.disabled = false; calcBtn.textContent = '️ Calcular Consolidado'; }, 100);
    }, 50);
  });

  // Limpiar selección
  if (refreshBtn) refreshBtn.addEventListener('click', clearSelections);

  // Modal Stock
  if (stockBtn) stockBtn.addEventListener('click', () => { renderStockModal(); stockModal.classList.remove('hidden'); });
  if (closeStockModal) closeStockModal.addEventListener('click', () => stockModal.classList.add('hidden'));
  if (addStockBtn) addStockBtn.addEventListener('click', addStockEntry);
  if (clearStockBtn) clearStockBtn.addEventListener('click', clearStockEntries);

  // Modal Frecuencia
  if (freqBtn) freqBtn.addEventListener('click', () => { renderFreqModal(); freqModal.classList.remove('hidden'); });
  if (closeFreqModal) closeFreqModal.addEventListener('click', () => freqModal.classList.add('hidden'));
  if (clearFreqBtn) clearFreqBtn.addEventListener('click', () => {
    if (confirm('¿Seguro que deseas borrar todo el historial de frecuencias?')) {
      localStorage.removeItem(FREQ_KEY);
      renderFreqModal();
    }
  });
  if (freqModal) freqModal.addEventListener('click', (e) => { if (e.target === freqModal) freqModal.classList.add('hidden'); });
  
  // Exportar
  if (exportBtn) exportBtn.addEventListener('click', handleExport);
}

// Filtrar productos
function filterProducts(searchTerm) {
  if (!searchTerm) {
    filteredData = [...aggregatedData];
    searchResults.classList.add('hidden');
  } else {
    filteredData = aggregatedData.filter(item => 
      normalize(item.product).includes(searchTerm) || normalize(item.unit).includes(searchTerm)
    );
    if (filteredData.length > 0) {
      searchResults.textContent = `🔍 Encontrados ${filteredData.length} producto(s) con "${searchInput.value}"`;
      searchResults.classList.remove('hidden');
    } else {
      searchResults.textContent = `❌ No se encontraron productos con "${searchInput.value}"`;
      searchResults.classList.remove('hidden');
    }
  }
  renderTable();
}

// Manejar cambio de días
function handleDayChange() {
  const selectedDays = [...document.querySelectorAll('#daysCheckboxes input:checked')].map(c => c.value);
  if (selectedDays.length === 0) {
    resultsSection.classList.add('hidden');
    aggregatedData = [];
    filteredData = [];
    updateOrderPanel();
    return;
  }
  isCalculating = true;
  calculate(selectedDays);
  if (searchInput) searchInput.value = '';
  filteredData = [...aggregatedData];
  renderTable();
  resultsSection.classList.remove('hidden');
  requestAnimationFrame(() => { isCalculating = false; });
}

// 3. Cálculo con correcciones de unidades
function calculate(selectedDays) {
  const map = {};
  const footerKw = ['recibi', 'entregue', 'encargado', 'vo bo', 'chef', 'nutricion', 'shirley', 'fecha'];

  selectedDays.forEach(day => {
    uploadedFiles.forEach(({ name, workbook }) => {
      const sn = workbook.SheetNames.find(s => s.trim().toUpperCase() === day);
      if (!sn) return;
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sn], { header: 1, defval: "" });
      let inBlock = false, cp=1, cu=2, cq=3;

      for (let r=0; r<rows.length; r++) {
        const row = rows[r].map(c => String(c||"").trim());
        if (!inBlock) {
          const pi = row.findIndex(c => /^producto$/i.test(c));
          if (pi !== -1) { 
            cp = pi; 
            cu = row.findIndex(c=>/unid/i.test(c)) !== -1 ? row.findIndex(c=>/unid/i.test(c)) : cp+1; 
            cq = row.findIndex(c=>/cantidad/i.test(c) && !/devuelta|final/i.test(c)) !== -1 ? row.findIndex(c=>/cantidad/i.test(c) && !/devuelta|final/i.test(c)) : cp+2; 
            inBlock = true; 
          }
          continue;
        }
        const txt = row.join(' ').toLowerCase();
        if (footerKw.some(k=>txt.includes(k)) || (txt==="" && r>rows.length-5)) { inBlock=false; continue; }

        const prod = row[cp], unit = row[cu], qtyV = row[cq];
        if (!prod || !unit || !qtyV) continue;
        const qty = parseFloat(qtyV);
        if (isNaN(qty) || qty<=0) continue;

        let u = normalize(unit);
        let p = normalize(prod);
        let displayUnit = unit;

        // Correcciones específicas UAB
        if (p === 'leche' && u === 'kg') { u = 'lt'; displayUnit = 'LT'; }
        if (p === 'leche de soya' && u === 'lt') { u = 'bolsa'; displayUnit = 'BOLSA'; }

        const key = `${p}|${u}`;
        if (!map[key]) map[key] = { product: prod, unit: displayUnit, qty: 0 };
        map[key].qty += qty;
      }
    });
  });

  aggregatedData = Object.values(map).sort((a,b) => a.product.localeCompare(b.product, 'es'));
  filteredData = [...aggregatedData];
  
  // Limpiar selecciones huérfanas
  const currentKeys = new Set(aggregatedData.map(i => `${normalize(i.product)}|${normalize(i.unit)}`));
  selectedKeys.forEach(k => { if (!currentKeys.has(k)) selectedKeys.delete(k); });
}

// Obtener solo productos que NO están en stock
function getNonStockedData() {
  return filteredData.filter(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    return !stockEntries.some(s => s.key === key);
  });
}

// 4. Renderizado (Solo productos sin stock)
function renderTable() {
  resultsTableBody.innerHTML = '';
  const displayData = getNonStockedData();
  updateSelectAllCheckbox(displayData);
  
  if (displayData.length === 0) {
    resultsTableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding:20px; color:var(--text-secondary);">No hay productos para mostrar</td></tr>';
    return;
  }
  
  const searchTerm = normalize(searchInput.value);
  
  displayData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    const isChecked = selectedKeys.has(key);
    
    let displayProduct = toTitleCase(item.product);
    if (searchTerm) {
      const regex = new RegExp(`(${searchTerm})`, 'gi');
      displayProduct = displayProduct.replace(regex, '<span class="highlight">$1</span>');
    }
    
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="check-col"><input type="checkbox" data-key="${key}" ${isChecked ? 'checked' : ''}></td>
      <td>${displayProduct}</td>
      <td class="num">${formatQty(item.qty)}</td>
      <td>${item.unit}</td>
    `;
    resultsTableBody.appendChild(tr);
  });
  updateOrderPanel();
}

function updateSelectAllCheckbox(displayData) {
  const list = displayData || getNonStockedData();
  if (list.length === 0) { selectAll.checked = false; return; }
  selectAll.checked = list.every(item => 
    selectedKeys.has(`${normalize(item.product)}|${normalize(item.unit)}`)
  );
}

// 5. Panel de Pedido (Solo seleccionados y sin stock)
function updateOrderPanel() {
  orderTableBody.innerHTML = '';
  const nonStockedSelected = aggregatedData.filter(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    return selectedKeys.has(key) && !stockEntries.some(s => s.key === key);
  });

  if (nonStockedSelected.length === 0) {
    emptyOrder.classList.remove('hidden');
    orderFooter.classList.add('hidden');
    return;
  }
  emptyOrder.classList.add('hidden');
  orderFooter.classList.remove('hidden');

  nonStockedSelected.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${toTitleCase(item.product)}</td><td class="num">${formatQty(item.qty)}</td><td>${item.unit}</td>`;
    orderTableBody.appendChild(tr);
  });
}

function clearSelections() {
  selectedKeys.clear();
  selectAll.checked = false;
  document.querySelectorAll('#resultsTable input[type="checkbox"][data-key]').forEach(cb => cb.checked = false);
  updateOrderPanel();
}

// 6. Gestión de Stock
function loadStockEntries() {
  try {
    stockEntries = JSON.parse(localStorage.getItem(STOCK_KEY) || '[]');
  } catch {
    stockEntries = [];
  }
}

function saveStockEntries() {
  localStorage.setItem(STOCK_KEY, JSON.stringify(stockEntries));
  // Actualizar vistas principales al cambiar stock
  renderTable();
  updateOrderPanel();
}

function renderStockModal() {
  stockTableBody.innerHTML = '';
  
  if (stockEntries.length === 0) {
    emptyStock.style.display = 'block';
    document.querySelector('.stock-table-wrapper').style.display = 'none';
  } else {
    emptyStock.style.display = 'none';
    document.querySelector('.stock-table-wrapper').style.display = 'block';
    
    stockEntries.forEach((entry, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${toTitleCase(entry.product)}</td>
        <td>${entry.unit}</td>
        <td class="num">${formatQty(entry.qty)}</td>
        <td class="action-col"><button class="delete-btn" data-idx="${idx}">🗑️</button></td>
      `;
      stockTableBody.appendChild(tr);
    });

    stockTableBody.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.idx);
        stockEntries.splice(idx, 1);
        saveStockEntries();
        renderStockModal();
      });
    });
  }
  stockCount.textContent = `${stockEntries.length} producto(s)`;
}

function addStockEntry() {
  const product = stockProduct.value.trim();
  const unit = stockUnit.value.trim();
  const qty = parseFloat(stockQty.value);

  if (!product || !unit || isNaN(qty) || qty <= 0) {
    return alert('Por favor completa todos los campos correctamente.');
  }

  const key = `${normalize(product)}|${normalize(unit)}`;
  
  // Actualizar si ya existe, o agregar nuevo
  const existingIdx = stockEntries.findIndex(e => e.key === key);
  if (existingIdx !== -1) {
    stockEntries[existingIdx].product = product;
    stockEntries[existingIdx].unit = unit;
    stockEntries[existingIdx].qty += qty;
  } else {
    stockEntries.push({ key, product, unit, qty });
  }

  saveStockEntries();
  renderStockModal();

  // Limpiar formulario
  stockProduct.value = '';
  stockUnit.value = '';
  stockQty.value = '';
  stockProduct.focus();
}

function clearStockEntries() {
  if (stockEntries.length === 0) return;
  if (confirm('¿Seguro que deseas eliminar todo el stock registrado?')) {
    stockEntries = [];
    saveStockEntries();
    renderStockModal();
  }
}

// 7. Historial de Frecuencias
function getFreqData() { return JSON.parse(localStorage.getItem(FREQ_KEY) || '{}'); }
function saveFreqData(data) { localStorage.setItem(FREQ_KEY, JSON.stringify(data)); }

function updateFrequencyHistory() {
  let data = getFreqData();
  aggregatedData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    if (!selectedKeys.has(key)) return;
    // Solo registrar si no está en stock
    if (stockEntries.some(s => s.key === key)) return;
    
    if (!data[key]) data[key] = { name: item.product, unit: item.unit, count: 0, totalQty: 0 };
    data[key].count += 1;
    data[key].totalQty += item.qty;
  });
  saveFreqData(data);
}

function renderFreqModal() {
  const data = getFreqData();
  freqTableBody.innerHTML = '';
  const entries = Object.values(data).sort((a,b) => b.count - a.count);
  if (!entries.length) { emptyFreq.style.display = 'block'; return; }
  emptyFreq.style.display = 'none';

  entries.forEach(e => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${toTitleCase(e.name)} <span class="freq-badge">Top ${Math.ceil((entries.indexOf(e)+1)/entries.length*10)}%</span></td>
      <td>${e.unit}</td>
      <td class="num">${e.count}</td>
      <td class="num">${formatQty(e.totalQty)}</td>
    `;
    freqTableBody.appendChild(tr);
  });
}

// 8. Exportación (Solo productos sin stock, sin columna Stock Actual)
function handleExport() {
  const nonStockedSelected = aggregatedData.filter(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    return selectedKeys.has(key) && !stockEntries.some(s => s.key === key);
  });

  if (nonStockedSelected.length === 0) return alert('Selecciona al menos un producto.');
  updateFrequencyHistory();
  
  let counter = 1;
  const data = [];
  nonStockedSelected.forEach(item => {
    data.push({
      "Items": counter++,
      "Producto": toTitleCase(item.product),
      "Unidad de Medida": item.unit,
      "Cantidad Solicitada": formatQty(item.qty),
      "Costo": "",
      "Especificaciones": ""
    });
  });

  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [
    { wch: 8 },   // Items
    { wch: 25 },  // Producto
    { wch: 18 },  // Unidad de Medida
    { wch: 22 },  // Cantidad Solicitada
    { wch: 12 },  // Costo
    { wch: 30 }   // Especificaciones
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pedido Seleccionado");
  XLSX.writeFile(wb, `Pedido_UAB_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// Cargar stock al inicio
loadStockEntries();