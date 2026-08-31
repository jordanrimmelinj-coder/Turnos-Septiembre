/* ============================================================
   Consulta de Horarios — app.js
   Lee un archivo Excel (.xlsx) con hojas por mes y bloques de
   semanas, y permite a cualquier promotor consultar su horario
   desde el celular sin login, sin registro y sin instalar nada.
   ============================================================ */

const CONFIG = {
  // Nombre del archivo Excel que debe estar en la misma carpeta
  // que index.html. Cambia esto si tu archivo tiene otro nombre.
  excelFile: 'horarios.xlsx',
};

// Indexado por Date#getDay() (0=Domingo)
const DAY_NAMES_BY_INDEX = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

// ---------- Estado global ----------
let weekBlocks = [];          // todos los bloques de semana encontrados en el Excel
let promotorIndex = {};       // nombre -> [indices de weekBlocks donde aparece], ordenado por fecha
let promotorNames = [];       // lista de nombres únicos, ordenada alfabéticamente
let currentPromotor = null;
let currentBlockPointer = 0;  // posición dentro de promotorIndex[currentPromotor]

// ---------- Elementos del DOM ----------
const selectEl = document.getElementById('promotorSelect');
const statusMsg = document.getElementById('statusMsg');
const resultSection = document.getElementById('resultSection');
const promotorNameEl = document.getElementById('promotorName');
const storeInfoEl = document.getElementById('storeInfo');
const weekLabelEl = document.getElementById('weekLabel');
const dateRangeEl = document.getElementById('dateRange');
const scheduleBodyEl = document.getElementById('scheduleBody');
const prevWeekBtn = document.getElementById('prevWeekBtn');
const nextWeekBtn = document.getElementById('nextWeekBtn');

// ---------- Utilidades ----------
function pad2(n) { return String(n).padStart(2, '0'); }

// IMPORTANTE: usamos siempre getters LOCALES (no getUTC*) para leer las
// fechas/horas que vienen del Excel. SheetJS construye estos objetos Date
// de forma que su hora/fecha "local" reproduce el valor original de la
// celda; su representación UTC interna puede quedar desplazada por un
// desfase histórico de la zona horaria (por ejemplo, para el año base
// 1899 que Excel usa para las celdas de solo-hora). Por eso getUTC* puede
// devolver un valor incorrecto aunque get* (local) sea siempre correcto.
function formatDateShort(date) {
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}`;
}

function dayNameOf(date) {
  return DAY_NAMES_BY_INDEX[date.getDay()];
}

function isSameDay(dateA, dateB) {
  return dateA.getFullYear() === dateB.getFullYear() &&
    dateA.getMonth() === dateB.getMonth() &&
    dateA.getDate() === dateB.getDate();
}

// "Hoy" representado como fecha local pura (sin hora), para comparar
// de forma justa contra las fechas que vienen del Excel.
function todayAtMidnight() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function formatCellValue(value) {
  if (value === null || value === undefined || value === '') return { text: '-', libre: false, feriado: false };
  if (value instanceof Date) {
    return { text: `${pad2(value.getHours())}:${pad2(value.getMinutes())}`, libre: false, feriado: false };
  }
  const text = String(value).trim();
  return { text: text || '-', libre: /libre/i.test(text), feriado: /feriado/i.test(text) };
}

// ---------- Carga y parseo del Excel ----------
async function loadSchedule() {
  try {
    const response = await fetch(CONFIG.excelFile, { cache: 'no-store' });
    if (!response.ok) throw new Error('No se pudo descargar el archivo de horarios.');
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

    weekBlocks = [];
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      // raw:true + cellDates:true (definido en XLSX.read) hace que las celdas
      // de fecha/hora lleguen como objetos Date reales, no como texto formateado.
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
      parseSheetIntoBlocks(sheetName, rows);
    });

    if (weekBlocks.length === 0) {
      throw new Error('No se encontraron semanas en el archivo Excel.');
    }

    buildPromotorIndex();
    populateSelect();
    statusMsg.textContent = '';
    selectEl.disabled = false;
  } catch (err) {
    console.error(err);
    statusMsg.textContent = 'No se pudo cargar el archivo de horarios. Intenta recargar la página.';
    statusMsg.classList.add('error');
    selectEl.innerHTML = '<option value="">No disponible</option>';
  }
}

function isSemanaLabel(cell) {
  return typeof cell === 'string' && /semana/i.test(cell.trim());
}

function parseSheetIntoBlocks(sheetName, rows) {
  let i = 0;
  while (i < rows.length) {
    const row = rows[i] || [];
    const weekLabelCell = row.find((c) => isSemanaLabel(c));
    if (weekLabelCell) {
      const weekLabel = weekLabelCell.trim();
      const headerRow = rows[i + 1] || [];
      const subHeaderRowIndex = i + 2;

      // columnas donde empieza cada día (contienen un objeto Date)
      const dayCols = [];
      headerRow.forEach((cellValue, colIndex) => {
        const parsed = parsePossibleDate(cellValue);
        if (parsed) dayCols.push({ col: colIndex, date: parsed });
      });

      // columna opcional "Tienda" (si existe en el Excel)
      let storeCol = -1;
      headerRow.forEach((cellValue, colIndex) => {
        if (typeof cellValue === 'string' && /tienda|local|sucursal/i.test(cellValue)) {
          storeCol = colIndex;
        }
      });

      if (dayCols.length > 0) {
        // filas de datos: empiezan después del sub-encabezado Entrada/Salida
        const dataStart = subHeaderRowIndex + 1;
        const promotores = [];
        let j = dataStart;
        while (j < rows.length) {
          const dataRow = rows[j] || [];
          const nameCell = dataRow[0];
          if (!nameCell || isSemanaLabel(dataRow.find((c) => isSemanaLabel(c)))) break;
          const name = String(nameCell).trim();
          if (!name) break;

          const days = dayCols.map(({ col, date }) => {
            const entrada = formatCellValue(parsePossibleValue(dataRow[col]));
            const salida = formatCellValue(parsePossibleValue(dataRow[col + 1]));
            const feriado = entrada.feriado || salida.feriado;
            const libre = entrada.libre || salida.libre;
            return { date, entrada, salida, feriado, libre };
          });

          const store = storeCol >= 0 ? (dataRow[storeCol] ? String(dataRow[storeCol]).trim() : '') : '';

          promotores.push({ name, store, days });
          j += 1;
        }

        if (promotores.length > 0) {
          const dates = dayCols.map((d) => d.date);
          weekBlocks.push({
            sheetName,
            weekLabel,
            startDate: dates[0],
            endDate: dates[dates.length - 1],
            promotores,
          });
        }
        i = j > i ? j : i + 1;
        continue;
      }
    }
    i += 1;
  }
}

// Respaldo defensivo por si alguna celda de fecha llega como texto ISO
// en vez de objeto Date (no debería ocurrir con raw:true, pero por si acaso).
function parsePossibleDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) {
    const d = new Date(value); // fechas ISO "yyyy-mm-dd" se interpretan como UTC
    if (!isNaN(d.getTime())) return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  return null;
}

// Entrada/Salida pueden venir como Date (hora) o como texto ("Libre").
// Respaldo defensivo por si alguna celda de hora llega como texto "HH:MM".
function parsePossibleValue(value) {
  if (value === null || value === undefined) return value;
  if (typeof value === 'string' && /^\d{2}:\d{2}(:\d{2})?$/.test(value.trim())) {
    const [h, m] = value.trim().split(':');
    return new Date(1899, 11, 30, Number(h), Number(m));
  }
  return value;
}

// ---------- Índice de promotores ----------
function buildPromotorIndex() {
  promotorIndex = {};
  weekBlocks.forEach((block, blockIndex) => {
    block.promotores.forEach((p) => {
      if (!promotorIndex[p.name]) promotorIndex[p.name] = [];
      promotorIndex[p.name].push(blockIndex);
    });
  });

  // ordenar los bloques de cada promotor cronológicamente
  Object.keys(promotorIndex).forEach((name) => {
    promotorIndex[name].sort((a, b) => weekBlocks[a].startDate - weekBlocks[b].startDate);
  });

  promotorNames = Object.keys(promotorIndex).sort((a, b) => a.localeCompare(b, 'es'));
}

function populateSelect() {
  selectEl.innerHTML = '<option value="">Selecciona tu nombre</option>' +
    promotorNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(toTitleCase(name))}</option>`).join('');
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function toTitleCase(str) {
  return str.toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}

// ---------- Selección y render ----------
function findDefaultBlockPointer(blockIndices) {
  const today = todayAtMidnight();
  // 1) semana que contiene hoy
  for (let p = 0; p < blockIndices.length; p += 1) {
    const block = weekBlocks[blockIndices[p]];
    if (today >= block.startDate && today <= block.endDate) return p;
  }
  // 2) próxima semana futura
  for (let p = 0; p < blockIndices.length; p += 1) {
    const block = weekBlocks[blockIndices[p]];
    if (block.startDate > today) return p;
  }
  // 3) última semana disponible (pasada)
  return blockIndices.length - 1;
}

function renderForPromotor(name) {
  currentPromotor = name;
  const blockIndices = promotorIndex[name] || [];
  currentBlockPointer = findDefaultBlockPointer(blockIndices);
  renderCurrentBlock();
}

function renderCurrentBlock() {
  const blockIndices = promotorIndex[currentPromotor] || [];
  if (blockIndices.length === 0) {
    resultSection.classList.add('hidden');
    return;
  }
  const block = weekBlocks[blockIndices[currentBlockPointer]];
  const promotorData = block.promotores.find((p) => p.name === currentPromotor);

  promotorNameEl.textContent = toTitleCase(currentPromotor);

  if (promotorData.store) {
    storeInfoEl.textContent = `Tienda: ${promotorData.store}`;
    storeInfoEl.classList.remove('hidden');
  } else {
    storeInfoEl.classList.add('hidden');
  }

  weekLabelEl.textContent = block.weekLabel;
  dateRangeEl.textContent = `${formatDateShort(block.startDate)} – ${formatDateShort(block.endDate)}`;

  const today = todayAtMidnight();
  scheduleBodyEl.innerHTML = promotorData.days.map((day) => {
    const rowClass = isSameDay(day.date, today) ? ' class="is-today"' : '';
    const entradaCell = day.entrada.feriado
      ? `<td class="feriado-cell">${escapeHtml(day.entrada.text)}</td>`
      : day.entrada.libre
      ? `<td class="libre-cell">${escapeHtml(day.entrada.text === '-' ? 'Libre' : day.entrada.text)}</td>`
      : `<td>${escapeHtml(day.entrada.text)}</td>`;
    const salidaCell = day.salida.feriado
      ? `<td class="feriado-cell">${escapeHtml(day.salida.text)}</td>`
      : day.salida.libre
      ? `<td class="libre-cell">${escapeHtml(day.salida.text === '-' ? 'Libre' : day.salida.text)}</td>`
      : `<td>${escapeHtml(day.salida.text)}</td>`;
    const dayCellClass = day.feriado ? ' day-cell feriado-cell' : day.libre ? ' day-cell libre-day-cell' : ' day-cell';
    return `<tr${rowClass}>
      <td class="${dayCellClass.trim()}">${dayNameOf(day.date)} ${day.date.getDate()}</td>
      ${entradaCell}
      ${salidaCell}
    </tr>`;
  }).join('');

  prevWeekBtn.disabled = currentBlockPointer <= 0;
  nextWeekBtn.disabled = currentBlockPointer >= blockIndices.length - 1;

  resultSection.classList.remove('hidden');
}

// ---------- Eventos ----------
selectEl.addEventListener('change', (e) => {
  const name = e.target.value;
  if (!name) {
    resultSection.classList.add('hidden');
    return;
  }
  renderForPromotor(name);
});

prevWeekBtn.addEventListener('click', () => {
  if (currentBlockPointer > 0) {
    currentBlockPointer -= 1;
    renderCurrentBlock();
  }
});

nextWeekBtn.addEventListener('click', () => {
  const blockIndices = promotorIndex[currentPromotor] || [];
  if (currentBlockPointer < blockIndices.length - 1) {
    currentBlockPointer += 1;
    renderCurrentBlock();
  }
});

// ---------- Inicio ----------
loadSchedule();
