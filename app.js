let uploadedFiles = [];
let aggregatedData = [];
let selectedKeys = new Set();
const FREQ_KEY = 'uab_product_frequencies_v1';

const fileInput = document.getElementById('fileInput');
const fileStatus = document.getElementById('fileStatus');
const daysSelector = document.getElementById('daysSelector');
const daysCheckboxes = document.getElementById('daysCheckboxes');
const calcBtn = document.getElementById('calcBtn');
const resultsSection = document.getElementById('resultsSection');
const resultsTableBody = document.querySelector('#resultsTable tbody');
const selectAll = document.getElementById('selectAll');
const orderTableBody = document.querySelector('#orderTable tbody');
const emptyOrder = document.getElementById('emptyOrder');
const orderFooter = document.getElementById('orderFooter');
const exportBtn = document.getElementById('exportBtn');
const freqBtn = document.getElementById('freqBtn');
const freqModal = document.getElementById('freqModal');
const closeFreqModal = document.getElementById('closeFreqModal');
const clearFreqBtn = document.getElementById('clearFreqBtn');
const freqTableBody = document.querySelector('#freqTable tbody');
const emptyFreq = document.getElementById('emptyFreq');

// Utilidad: Formato Título (Primera letra mayúscula por palabra)
function toTitleCase(str) {
  return String(str).toLowerCase().replace(/(?:^|\s)\S/g, a => a.toUpperCase());
}

// NUEVO: Función para limpiar decimales (14.00 -> 14, 14.50 -> 14.5)
function formatQty(val) {
  if (val === undefined || val === null || isNaN(val)) return '0';
  const num = parseFloat(val);
  return num % 1 === 0 ? num.toString() : num.toFixed(2).replace(/\.?0+$/, '');
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
  populateDays();
  daysSelector.classList.remove('hidden');
  resultsSection.classList.add('hidden');
  emptyOrder.classList.remove('hidden');
  orderFooter.classList.add('hidden');
}

function populateDays() {
  daysCheckboxes.innerHTML = '';
  const days = new Set();
  uploadedFiles.forEach(wb => wb.workbook.SheetNames.forEach(s => days.add(s.trim().toUpperCase())));
  const order = ['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO','DOMINGO'];
  [...days].sort((a,b) => (order.indexOf(a)+1) - (order.indexOf(b)+1) || a.localeCompare(b)).forEach(day => {
    daysCheckboxes.innerHTML += `<div><input type="checkbox" id="day-${day}" value="${day}" checked><label for="day-${day}">${day}</label></div>`;
  });
}

calcBtn.addEventListener('click', () => {
  const selected = [...document.querySelectorAll('#daysCheckboxes input:checked')].map(c => c.value);
  if (!selected.length) return alert('Selecciona al menos un día.');
  calcBtn.disabled = true; calcBtn.textContent = '⏳ Procesando...';
  setTimeout(() => { 
    calculate(selected); 
    renderTable(); 
    resultsSection.classList.remove('hidden'); 
    calcBtn.disabled = false; calcBtn.textContent = '⚙️ Calcular Consolidado'; 
  }, 100);
});

function normalize(str) { return String(str).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " "); }

// 2. Cálculo y correcciones
function calculate(selectedDays) {
  aggregatedData = [];
  selectedKeys.clear();
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
          if (pi !== -1) { cp=pi; cu = row.findIndex(c=>/unid/i.test(c))>-1 ? row.findIndex(c=>/unid/i.test(c)) : cp+1; cq = row.findIndex(c=>/cantidad/i.test(c) && !/devuelta|final/i.test(c))>-1 ? row.findIndex(c=>/cantidad/i.test(c) && !/devuelta|final/i.test(c)) : cp+2; inBlock=true; }
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

        // Correcciones de unidades
        if (p === 'leche' && u === 'kg') { u = 'lt'; displayUnit = 'LT'; }
        if (p === 'leche de soya' && u === 'lt') { u = 'bolsa'; displayUnit = 'BOLSA'; }

        const key = `${p}|${u}`;
        if (!map[key]) map[key] = { product: prod, unit: displayUnit, qty: 0 };
        map[key].qty += qty;
      }
    });
  });

  aggregatedData = Object.values(map).sort((a,b) => a.product.localeCompare(b.product, 'es'));
}

// 3. Renderizado de tabla principal (Con formato limpio de cantidad)
function renderTable() {
  resultsTableBody.innerHTML = '';
  selectAll.checked = false;
  aggregatedData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="check-col"><input type="checkbox" data-key="${key}"></td>
      <td>${toTitleCase(item.product)}</td>
      <td class="num">${formatQty(item.qty)}</td>
      <td>${item.unit}</td>
    `;
    resultsTableBody.appendChild(tr);
  });

  resultsTableBody.addEventListener('change', (e) => {
    if (e.target.type === 'checkbox') {
      const key = e.target.dataset.key;
      if (selectedKeys.has(key)) selectedKeys.delete(key); else selectedKeys.add(key);
      updateOrderPanel();
    }
  });

  selectAll.addEventListener('change', () => {
    const checkboxes = resultsTableBody.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(cb => {
      cb.checked = selectAll.checked;
      const key = cb.dataset.key;
      selectAll.checked ? selectedKeys.add(key) : selectedKeys.delete(key);
    });
    updateOrderPanel();
  });
}

// 4. Panel de Pedido (Con formato limpio de cantidad)
function updateOrderPanel() {
  orderTableBody.innerHTML = '';
  if (selectedKeys.size === 0) {
    emptyOrder.classList.remove('hidden');
    orderFooter.classList.add('hidden');
    return;
  }
  emptyOrder.classList.add('hidden');
  orderFooter.classList.remove('hidden');

  aggregatedData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    if (!selectedKeys.has(key)) return;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${toTitleCase(item.product)}</td><td class="num">${formatQty(item.qty)}</td><td>${item.unit}</td>`;
    orderTableBody.appendChild(tr);
  });
}

// 5. Historial / Frecuencia (localStorage)
function getFreqData() { return JSON.parse(localStorage.getItem(FREQ_KEY) || '{}'); }
function saveFreqData(data) { localStorage.setItem(FREQ_KEY, JSON.stringify(data)); }

function updateFrequencyHistory() {
  let data = getFreqData();
  aggregatedData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    if (!selectedKeys.has(key)) return;
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

freqBtn.addEventListener('click', () => { renderFreqModal(); freqModal.classList.remove('hidden'); });
closeFreqModal.addEventListener('click', () => freqModal.classList.add('hidden'));
clearFreqBtn.addEventListener('click', () => {
  if (confirm('¿Seguro que deseas borrar todo el historial de frecuencias?')) {
    localStorage.removeItem(FREQ_KEY);
    renderFreqModal();
  }
});

// 6. Exportación con formato solicitado y cantidades limpias
exportBtn.addEventListener('click', () => {
  if (selectedKeys.size === 0) return alert('Selecciona al menos un producto.');
  updateFrequencyHistory();
  
  let counter = 1;
  const data = [];
  aggregatedData.forEach(item => {
    const key = `${normalize(item.product)}|${normalize(item.unit)}`;
    if (!selectedKeys.has(key)) return;
    data.push({
      "Items": counter++,
      "Producto": toTitleCase(item.product),
      "Cantidad Solicitada": formatQty(item.qty), // Ahora exporta 14 en vez de 14.00
      "Costo": "",
      "Especificaciones": ""
    });
  });

  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [
    { wch: 8 },  // Items
    { wch: 30 }, // Producto
    { wch: 22 }, // Cantidad Solicitada
    { wch: 12 }, // Costo
    { wch: 35 }  // Especificaciones
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pedido Seleccionado");
  XLSX.writeFile(wb, `Pedido_UAB_${new Date().toISOString().slice(0,10)}.xlsx`);
});

freqModal.addEventListener('click', (e) => { if (e.target === freqModal) freqModal.classList.add('hidden'); });