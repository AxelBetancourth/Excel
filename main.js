// Elementos de la cinta y barra lateral
const tabs = document.querySelectorAll('.ribbon-tabs button');
const submenus = document.querySelectorAll('.submenu');
const submenuContainer = document.getElementById('submenu-container');
const sidebar = document.getElementById('archivo-sidebar');
const sidebarClose = document.getElementById('sidebar-close');

// Elementos de la barra de fórmulas y hojas
const formulaInput = document.getElementById('formula-input');
const activeCellDisplay = document.getElementById('active-cell');
const sheetsContainer = document.querySelector('.sheets-container');
const addSheetButton = document.getElementById('add-sheet');

/**
 * Activa la pestaña indicada. Si es "archivo" muestra la barra lateral.
 */
function activateTab(tabName) {
  if (tabName === 'archivo') {
    sidebar.classList.add('active');
    submenuContainer.style.display = 'none';
  } else {
    sidebar.classList.remove('active');
    submenuContainer.style.display = 'block';
    tabs.forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === tabName);
    });
    submenus.forEach(menu => {
      menu.classList.toggle('active', menu.id === 'submenu-' + tabName);
    });
    localStorage.setItem('activeTab', tabName);
  }
}

// Eventos en los botones de la cinta
tabs.forEach(btn => {
  btn.addEventListener('click', () => {
    activateTab(btn.dataset.tab);
  });
});

// Evento para cerrar la barra lateral "Archivo"
sidebarClose.addEventListener('click', () => {
  sidebar.classList.remove('active');
  submenuContainer.style.display = 'block';
});

// Al cargar la página, se activa la pestaña guardada (o "inicio" por defecto)
document.addEventListener('DOMContentLoaded', () => {
  const activeTab = localStorage.getItem('activeTab') || 'inicio';
  activateTab(activeTab);
});

// ================================
// Código de la Hoja de Cálculo
// ================================

const $ = el => document.querySelector(el);
const $$ = el => document.querySelectorAll(el);

const $table = $('table');
const $head = $('thead');
const $body = $('tbody');

const rows = 30;
const cols = 26;
const letras = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

let selectedColumn = null;
let selectedRow = null;
let activeCell = { x: 0, y: 0 };
let currentSheetIndex = 0;
let sheets = [
  { name: 'Hoja1', data: [] },
  { name: 'Hoja2', data: [] },
  { name: 'Hoja3', data: [] }
];

const range = length => Array.from({ length }, (_, i) => i);
const rangeChar = length => letras.slice(0, length);

// Inicializar datos para cada hoja
sheets.forEach(sheet => {
  sheet.data = range(cols).map(() =>
    range(rows).map(() => ({ computedValue: "", value: "" }))
  );
});

let state = sheets[currentSheetIndex].data;

/**
 * Renderiza la hoja de excel.
 */

const renderSpreadsheet = () => {
  const headerHTML = `<tr>
    <th></th>
    ${rangeChar(cols).map(letra => `<th>${letra}</th>`).join('')}
  </tr>`;
  $head.innerHTML = headerHTML;
  
  const bodyHTML = range(rows).map(row => `
    <tr>
      <th class="row-header">${row + 1}</th>
      ${range(cols).map(col => `
        <td data-x="${col}" data-y="${row}">
          <input type="text" value="${state[col][row].value}" />
          <span>${state[col][row].computedValue || ""}</span>
        </td>
      `).join('')}
    </tr>
  `).join('');
  $body.innerHTML = bodyHTML;
  
  $table.innerHTML = "";
  $table.appendChild($head);
  $table.appendChild($body);
};

/**
 * Convierte una referencia de celda en coordenadas [columna, fila].
 */
function getCellCoords(ref) {
  ref = ref.toUpperCase();
  const letter = ref.match(/[A-Z]+/)[0];
  const number = ref.match(/\d+/)[0];
  return [letras.indexOf(letter), parseInt(number) - 1];
}

/**
 * Obtiene el valor calculado de una celda mediante su referencia.
 */
function getCellValue(ref) {
  try {
    const [x, y] = getCellCoords(ref);
    return state[x]?.[y]?.computedValue || "";
  } catch (error) {
    return `!ERROR: ${error.message}`;
  }
}

/**
 * Transforma los nombres de las funciones a mayúsculas (solo fuera de comillas).
 */
function transformFunctionNames(formula) {
  const parts = formula.split(/(".*?")/);
  for (let i = 0; i < parts.length; i++) {
    if (i % 2 === 0) {
      parts[i] = parts[i].replace(/\b(si|o|y|concatenar)\b/gi, match => match.toUpperCase());
    }
  }
  return parts.join('');
}

/**
 * Devuelve el índice de la llave de cierre que hace match con la apertura en startIndex.
 */

function findMatchingParen(str, startIndex) {
  let count = 0;
  for (let i = startIndex; i < str.length; i++) {
    if (str[i] === '(') count++;
    else if (str[i] === ')') {
      count--;
      if (count === 0) return i;
    }
  }
  return -1;
}

/**
 * Extrae los argumentos de una función ignorando comas dentro de paréntesis o comillas.
 */

function extractFunctionArguments(str) {
  let args = [];
  let current = "";
  let count = 0;
  let inQuote = false;
  for (let i = 0; i < str.length; i++) {
    let char = str[i];
    if (char === '"' && (i === 0 || str[i - 1] !== '\\')) {
      inQuote = !inQuote;
    }
    if (!inQuote && char === ',' && count === 0) {
      args.push(current.trim());
      current = "";
      continue;
    }
    if (!inQuote) {
      if (char === '(') count++;
      if (char === ')') count--;
    }
    current += char;
  }
  if (current.trim() !== "") {
    args.push(current.trim());
  }
  return args;
}

/**
 * Reemplaza las llamadas a MOD() por su equivalente usando el operador de módulo (%).
 */
function replaceMOD(formula) {
  let index = formula.indexOf("MOD(");
  while (index !== -1) {
    let end = findMatchingParen(formula, index + 3);
    if (end === -1) break;
    let inside = formula.substring(index + 4, end);
    let args = extractFunctionArguments(inside);
    if (args.length !== 2) break;
    args = args.map(arg => replaceFunctions(arg));
    let replacement = `(${args[0]}) % (${args[1]})`;
    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
    index = formula.indexOf("MOD(");
  }
  return formula;
}

/**
 * Reemplaza las llamadas a SI() por su equivalente ternario.
 */
function replaceSI(formula) {
  let index = formula.indexOf("SI(");
  while (index !== -1) {
    let end = findMatchingParen(formula, index + 2);
    if (end === -1) break;
    let inside = formula.substring(index + 3, end);
    let args = extractFunctionArguments(inside);
    if (args.length !== 3) break;
    args = args.map(arg => replaceFunctions(arg));
    let replacement = `(${args[0]}) ? (${args[1]}) : (${args[2]})`;
    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
    index = formula.indexOf("SI(");
  }
  return formula;
}

/**
 * Reemplaza las llamadas a Y() por una expresión con &&.
 */
function replaceY(formula) {
  let index = formula.indexOf("Y(");
  while (index !== -1) {
    let end = findMatchingParen(formula, index + 1);
    if (end === -1) break;
    let inside = formula.substring(index + 2, end);
    let args = extractFunctionArguments(inside);
    args = args.map(arg => replaceFunctions(arg));
    let replacement = args.map(a => `(${a})`).join(' && ');
    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
    index = formula.indexOf("Y(");
  }
  return formula;
}

/**
 * Reemplaza las llamadas a O() por una expresión con ||.
 */
function replaceO(formula) {
  let index = formula.indexOf("O(");
  while (index !== -1) {
    // Para O, el paréntesis de apertura está en index + 1
    let end = findMatchingParen(formula, index + 1);
    if (end === -1) break;
    let inside = formula.substring(index + 2, end);
    let args = extractFunctionArguments(inside);
    args = args.map(arg => replaceFunctions(arg));
    let replacement = args.map(a => `(${a})`).join(' || ');
    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
    index = formula.indexOf("O(");
  }
  return formula;
}

/**
 * Reemplaza recursivamente las funciones MOD, SI, Y y O.
 */
function replaceFunctions(formula) {
  let prev;
  do {
    prev = formula;
    formula = replaceMOD(formula);
    formula = replaceSI(formula);
    formula = replaceY(formula);
    formula = replaceO(formula);
  } while (formula !== prev);
  return formula;
}

/**
 * Calcula el valor computado de una celda a partir de su contenido.
 * Se eliminan espacios extra luego del signo "=" y entre el nombre de la función y el paréntesis.
 */
function calculateComputedValue(value) {
  const strValue = String(value).trim();
  if (!strValue) return "";
  
  if (!strValue.startsWith('=')) {
    return isNaN(strValue) ? strValue : Number(strValue);
  }
  
  let formula = strValue.slice(1).trim();
  formula = formula.replace(/([A-Za-z]+)\s+\(/g, "$1(");
  formula = transformFunctionNames(formula);
  
  const concatRegex = /CONCATENAR\(([^)]+)\)/gi;
  formula = formula.replace(concatRegex, (match, args) => {
    const values = args.split(',').map(arg => {
      if (/^[A-Za-z]+\d+$/.test(arg.trim())) {
        return `"${getCellValue(arg.trim())}"`;
      } else if (/^".*"$/.test(arg.trim())) {
        return arg.trim();
      } else {
        return `"${arg.trim()}"`;
      }
    });
    return values.join(' + ');
  });
  

  formula = replaceFunctions(formula);
  
  // Reemplazar referencias de celdas
  formula = formula.replace(/([A-Za-z]+\d+)/gi, match => {
    const val = getCellValue(match);
    if (val === "") return 0;
    if (isNaN(val)) return `"${val}"`;
    return val;
  });
  
  try {
    const result = eval(formula);
    return result === 0 ? "" : result;
  } catch (error) {
    return `!ERROR: ${error.message}`;
  }
}

/**
 * Actualiza una celda en la posición (x, y) con el nuevo valor y recalcula la fórmula.
 */
function updateCell(x, y, value) {
  x = parseInt(x);
  y = parseInt(y);
  
  const newState = structuredClone(state);
  newState[x][y] = {
    value: value,
    computedValue: value.trim() === '' ? '' : calculateComputedValue(value)
  };
  
  state = newState;
  sheets[currentSheetIndex].data = state;
  updateAllDependentCells();
  
  const selector = `td[data-x="${x}"][data-y="${y}"]`;
  const cell = $(selector);
  if (cell) {
    cell.querySelector('input').value = value;
    cell.querySelector('span').textContent = state[x][y].computedValue;
  }
}

/**
 * Recalcula todas las celdas que tengan fórmulas.
 */
function updateAllDependentCells() {
  state.forEach((col, x) => {
    col.forEach((cell, y) => {
      if (cell.value.startsWith('=')) {
        // Recalcular valor
        const newValue = calculateComputedValue(cell.value);
        state[x][y].computedValue = newValue;
        
        // Actualizar DOM
        const selector = `td[data-x="${x}"][data-y="${y}"]`;
        const cellElement = document.querySelector(selector);
        if (cellElement) {
          cellElement.querySelector('span').textContent = newValue;
        }
      }
    });
  });
}

/**
 * Limpia las selecciones (columnas, filas y celdas resaltadas).
 */
function clearSelections() {
  $$('.selected-column').forEach(el => el.classList.remove('selected-column'));
  $$('.selected-row').forEach(el => el.classList.remove('selected-row'));
  $$('.selected-header').forEach(el => el.classList.remove('selected-header'));
  $$('.celdailumida').forEach(el => el.classList.remove('celdailumida'));
  selectedColumn = null;
  selectedRow = null;
}

/**
 * Actualiza la celda activa.
 */
function updateActiveCellDisplay(x, y) {
  activeCell = { x, y };
  const columnLetter = letras[x];
  const rowNumber = y + 1;
  activeCellDisplay.textContent = `${columnLetter}${rowNumber}`;
  formulaInput.value = state[x][y].value;
}

// Eventos en el encabezado
$head.addEventListener('click', event => {
  const th = event.target.closest('th:not(:first-child)');
  if (!th) return;
  clearSelections();
  const x = [...th.parentElement.children].indexOf(th) - 1;
  selectedColumn = x;
  th.classList.add('selected-header');
  $$(`tr td:nth-child(${x + 2})`).forEach(td => td.classList.add('selected-column'));
});

// Eventos en el cuerpo
$body.addEventListener('click', event => {
  const th = event.target.closest('th.row-header');
  const td = event.target.closest('td');
  
  if (th) {
    clearSelections();
    const y = [...th.parentElement.parentElement.children].indexOf(th.parentElement);
    selectedRow = y;
    th.classList.add('selected-header');
    $$(`tr:nth-child(${y + 1}) td`).forEach(td => td.classList.add('selected-row'));
  }
  
  if (td) {
    clearSelections();
    $$('.celdailumida').forEach(celda => celda.classList.remove('celdailumida'));
    td.classList.add('celdailumida');
    const x = parseInt(td.dataset.x);
    const y = parseInt(td.dataset.y);
    updateActiveCellDisplay(x, y);
  }
});

// Manejo de copiar
document.addEventListener('copy', (e) => {
  let data = [];
  if (selectedColumn !== null) {
    data = state[selectedColumn]
      .map(cell => cell.computedValue)
      .filter(value => value !== "");
  } else if (selectedRow !== null) {
    data = state.map(col => col[selectedRow].computedValue)
      .filter(value => value !== "");
  }
  if (data.length > 0) {
    e.clipboardData.setData('text/plain', data.join('\n'));
    e.preventDefault();
  }
});

// Manejo de pegar
document.addEventListener('paste', (e) => {
  if(e.target === formulaInput) return;
  const clipboardData = e.clipboardData.getData('text/plain');
  if (!clipboardData) return;
  const values = clipboardData.split('\n').filter(v => v !== '');
  
  if (selectedColumn !== null) {
    values.forEach((value, y) => {
      if (y < rows) updateCell(selectedColumn, y, value);
    });
  } else if (selectedRow !== null) {
    values.forEach((value, x) => {
      if (x < cols) updateCell(x, selectedRow, value);
    });
  }
  e.preventDefault();
});

// Doble clic para editar la celda
$body.addEventListener('dblclick', event => {
  const td = event.target.closest('td');
  if (!td) return;
  clearSelections();
  const { x, y } = td.dataset;
  const input = td.querySelector('input');
  input.focus();
  input.setSelectionRange(input.value.length, input.value.length);
  
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') input.blur();
  });
  
  input.addEventListener('blur', () => {
    updateCell(parseInt(x), parseInt(y), input.value.trim());
  });
});

// Borrar contenido con Delete o Backspace
document.addEventListener('keydown', (e) => {
  if (e.key === 'Delete' || e.key === 'Backspace') {
    if (selectedColumn !== null) {
      range(rows).forEach(y => updateCell(selectedColumn, y, ''));
    } else if (selectedRow !== null) {
      range(cols).forEach(x => updateCell(x, selectedRow, ''));
    }
    $$('.celdailumida').forEach(td => {
      const x = parseInt(td.dataset.x);
      const y = parseInt(td.dataset.y);
      updateCell(x, y, '');
    });
  }
});

/**
 * Cambia a la hoja especificada por índice.
 */
function switchSheet(index) {
  sheets[currentSheetIndex].data = state;
  currentSheetIndex = index;
  state = sheets[currentSheetIndex].data;
  document.querySelectorAll('.sheet').forEach((sheet, i) => {
    sheet.classList.toggle('active', i === index);
  });
  renderSpreadsheet();
}

// Eventos para cambiar de hoja
document.querySelectorAll('.sheet').forEach((sheet, index) => {
  sheet.addEventListener('click', () => {
    switchSheet(index);
  });
});

// Agregar nueva hoja
addSheetButton.addEventListener('click', () => {
  const newIndex = sheets.length;
  const newSheetName = `Hoja${newIndex + 1}`;
  sheets.push({
    name: newSheetName,
    data: range(cols).map(() =>
      range(rows).map(() => ({ computedValue: "", value: "" }))
    )
  });
  const newSheet = document.createElement('div');
  newSheet.className = 'sheet';
  newSheet.dataset.sheet = newSheetName.toLowerCase();
  newSheet.textContent = newSheetName;
  newSheet.addEventListener('click', () => {
    switchSheet(newIndex);
  });
  sheetsContainer.appendChild(newSheet);
  switchSheet(newIndex);
});

// ================================
// Eventos en la Barra de Fórmulas
// ================================

// Actualizar celda al presionar Enter en el input de fórmulas
formulaInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    const { x, y } = activeCell;
    updateCell(x, y, formulaInput.value);
    formulaInput.blur();
  }
});

// Agregar el signo "=" al hacer clic en el prefijo
document.querySelector('.formula-prefix').addEventListener('click', () => {
  formulaInput.value = '=';
  formulaInput.focus();
});

renderSpreadsheet();

// ================================
// Sección para guardar archivos (Claudio Tejeda)
// ================================

// Aquí irían las funciones para "Nuevo", "Descargar" y "Exportar a PDF"


// ================================
// Sección para guardar archivos (Coloca tu nombre aqui)
// ================================



// ================================
// Sección para guardar archivos (Claudio Tejeda)
// ================================



// ================================
// Sección para guardar archivos (Claudio Tejeda)
// ================================