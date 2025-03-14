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
	
	//anadido por CLAUDIO para --> archivo
	
	document.addEventListener("DOMContentLoaded", function () {
	  const buttons = document.querySelectorAll(".nav-btn");
	  const sections = document.querySelectorAll(".contenido-seccion");
	
	  function activarSeccion(target) {
	    // Ocultar todas las secciones
	    sections.forEach(section => section.classList.remove("activo"));
	
	    // Quitar la clase 'activo' de todos los botones
	    buttons.forEach(button => button.classList.remove("activo"));
	
	    // Mostrar la sección correspondiente
	    document.getElementById(target).classList.add("activo");
	
	    // Activar el botón correspondiente
	    document.querySelector(`.nav-btn[data-target="${target}"]`).classList.add("activo");
	  }
	
	  buttons.forEach(button => {
	    button.addEventListener("click", function () {
	      const target = this.getAttribute("data-target");
	      activarSeccion(target);
	    });
	  });
	
	  // Activar la sección de inicio por defecto
	  activarSeccion("inicio");
	});
	
	//Ventana modal
	
	document.getElementById('openSaveModal').addEventListener('click', function() {
	  document.getElementById('saveModal').style.display = 'flex';  // Cambié 'block' a 'flex' para mejor posicionamiento
	});
	
	// Función para cerrar la ventana modal
	document.getElementById('cancelSave').addEventListener('click', function() {
	  document.getElementById('saveModal').style.display = 'none';
	});
	
	// Función para guardar el archivo cuando el usuario haga clic en "Guardar"
	document.getElementById('saveFile').addEventListener('click', function() {
	  const nombreArchivo = document.getElementById('filename').value.trim() || "datos_guardados";
	  const tabla = document.getElementById('spreadsheet');
	
	  // Crear array de datos con formato correcto para Excel
	  const datos = [];
	
	  // Extraer los datos de la tabla
	  for (let y = 0; y < rows; y++) {
	    const filaDatos = [];
	    for (let x = 0; x < cols; x++) {
	      const celda = state[x][y];
	      filaDatos.push(celda.computedValue.toString());
	    }
	    datos.push(filaDatos);
	  }
	
	    // Crear el libro de trabajo de Excel
	    const wb = XLSX.utils.book_new();
	    const ws = XLSX.utils.aoa_to_sheet(datos);
	    XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
	  
	    // Convertir el libro en binario
	    const wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
	  
	    // Convertir el binario a Blob
	    const blob = new Blob([stringToArrayBuffer(wbout)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
	  
	    // Crear un enlace invisible para la descarga
	    const enlace = document.createElement('a');
	    enlace.href = URL.createObjectURL(blob);
	    enlace.download = nombreArchivo + ".xlsx";
	  
	    // Simular un clic en el enlace para descargar el archivo
	    document.body.appendChild(enlace);
	    enlace.click();
	    document.body.removeChild(enlace);
	  
	    // Guardar en archivos recientes
	    agregarArchivoReciente(nombreArchivo + ".xlsx");
	  
	    // Cerrar la ventana modal
	    document.getElementById('saveModal').style.display = 'none';
	});
	
	// Función para convertir string binario a ArrayBuffer
	function stringToArrayBuffer(s) {
	  const buf = new ArrayBuffer(s.length);
	  const view = new Uint8Array(buf);
	  for (let i = 0; i < s.length; i++) {
	    view[i] = s.charCodeAt(i) & 0xFF;
	  }
	  return buf;
	}
	
	// Función para agregar archivos recientes
	function agregarArchivoReciente(nombreArchivo) {
	  let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
	  
	  // Evitar duplicados
	  const index = archivos.indexOf(nombreArchivo);
	  if (index !== -1) {
	    // Si ya existe, quítalo primero para ponerlo al inicio
	    archivos.splice(index, 1);
	  }
	  
	  archivos.unshift(nombreArchivo); // Agregar al inicio
	  if (archivos.length > 5) archivos.pop(); // Máximo 5 recientes
	  localStorage.setItem('file_' + nombreArchivo + ".xlsx", JSON.stringify(state));
	  
	  mostrarArchivosRecientes();
	}
	
	// Función para mostrar archivos recientes con mejor formato
	function mostrarArchivosRecientes() {
	  let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
	  let lista = document.getElementById("lista-recientes");
	  lista.innerHTML = ""; // Limpiar antes de actualizar
	
	  if (archivos.length === 0) {
	    let li = document.createElement("li");
	    li.className = "no-recent-files";
	    li.textContent = "No hay archivos recientes";
	    lista.appendChild(li);
	    return;
	  }
	
	  archivos.forEach(archivo => {
	    let li = document.createElement("li");
	    li.className = "recent-file-item";
	    
	    // Crear icono
	    let icon = document.createElement("img");
	    icon.src = "images/logoexcel.png"; // Asegúrate de tener este icono
	    icon.alt = "Excel";
	    icon.className = "file-icon";
	    
	    // Crear nombre de archivo
	    let nameSpan = document.createElement("span");
	    nameSpan.textContent = archivo;
	    nameSpan.className = "file-name";
	    
	    // Añadir elementos al li
	    li.appendChild(icon);
	    li.appendChild(nameSpan);
	    
	    // Añadir evento de clic para abrir el archivo
	    li.addEventListener("click", function() {
	      // Guarda el estado actual antes de abrir el archivo
	      localStorage.setItem("tempState", JSON.stringify(state));
	      
	      // Abre el archivo seleccionado
	      window.location.href = `main.html?archivo=${encodeURIComponent(archivo)}`;
	    });
	    
	    lista.appendChild(li);
	  });
	}
	
	// Llamar a la función al cargar la página
	document.addEventListener("DOMContentLoaded", mostrarArchivosRecientes);
	
	//script para poder abrir una nueva pestana
	
	document.getElementById("btnNuevo").addEventListener("click", function () {
	  // Crear un nombre de archivo único
	  let nombreArchivo = "Archivo_" + new Date().toISOString().replace(/[:.-]/g, "_") + ".xlsx";
	
	  // Almacenar en archivos recientes
	  agregarArchivoReciente(nombreArchivo);
	
	  // Abrir nueva pestaña con main.html pasando el nombre del archivo como parámetro
	  window.open(`main.html?archivo=${encodeURIComponent(nombreArchivo)}`, "_blank");
	});
	
	//anadido por CLAUDIO para --> archivo
	
	//AÑADIR MAS HOJAS DE CALCULO, aqui debe de ir el codigo que te eh pedido
	

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
	
	//mostrar mensaje de bienvenida inicio
	
	document.addEventListener("DOMContentLoaded", function () {
	  const mensajeBienvenida = document.getElementById("mensaje-bienvenida");
	  const hora = new Date().getHours();
	
	  if (hora >= 6 && hora < 12) {
	      mensajeBienvenida.textContent = "Buenos días";
	  } else if (hora >= 12 && hora < 18) {
	      mensajeBienvenida.textContent = "Buenas tardes";
	  } else {
	      mensajeBienvenida.textContent = "Buenas noches";
	  }
	});
	
	//Funcion para el boton abrir
	
	function abrirExcel() {
	  let fileInput = document.getElementById('fileInput');
	  let file = fileInput.files[0]; // Obtener el archivo seleccionado
	
	  if (!file) {
	      alert("Por favor, selecciona un archivo de Excel.");
	      return;
	  }
	
	  let reader = new FileReader();
	
	  reader.onload = function (e) {
	      let data = new Uint8Array(e.target.result);
	      let workbook = XLSX.read(data, { type: 'array' });
	
	      let sheetName = workbook.SheetNames[0]; // Tomar la primera hoja
	      let sheet = workbook.Sheets[sheetName];
	
	      let htmlTable = XLSX.utils.sheet_to_html(sheet); // Convertir la hoja a HTML
	      document.getElementById('excelTable').innerHTML = htmlTable;
	  };
	
	  reader.readAsArrayBuffer(file); // Leer el archivo como ArrayBuffer
	}
	
	// ================================
	// Código de la Hoja de Cálculo
	// ================================
	
	const $ = el => document.querySelector(el);
	const $$ = el => document.querySelectorAll(el);
	
	const $table = $('table');
	const $head = $('thead');
	const $body = $('tbody');
	
	const rows = 100;
	const cols = 26;
	const letras = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
	
	let selectedColumn = null;
	let selectedRow = null;
	let activeCell = { x: 0, y: 0 };
	let currentSheetIndex = 0;
	let sheets = [
    { name: 'Hoja1', data: [] }
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
	
	document.addEventListener("DOMContentLoaded", function() {
	  const params = new URLSearchParams(window.location.search);
	  const archivo = params.get("archivo");
	  if (archivo) {
	    const savedState = localStorage.getItem('file_' + archivo);
	    if (savedState) {
	      state = JSON.parse(savedState);
	      sheets[currentSheetIndex].data = state;
	      renderSpreadsheet();
	    }
	  }
	});
	
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
	  let index = formula.indexOf("RESIDUO(");
	  while (index !== -1) {
	    let end = findMatchingParen(formula, index + 7);
	    if (end === -1) break;
	    let inside = formula.substring(index + 8, end);
	    let args = extractFunctionArguments(inside);
	    if (args.length !== 2) break;
	    args = args.map(arg => replaceFunctions(arg));
	    let replacement = `(${args[0]}) % (${args[1]})`;
	    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
	    index = formula.indexOf("RESIDUO(");
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
      formula = replaceSUMA(formula);
      formula = replacePROMEDIO(formula);
	  } while (formula !== prev);
	  return formula;
	}

  /**
 * Reemplaza las llamadas a SUMA()
 */
function replaceSUMA(formula) {
  let index = formula.indexOf("SUMA(");
  while (index !== -1) {
    let end = findMatchingParen(formula, index + 4);
    if (end === -1) break;
    let inside = formula.substring(index + 5, end);
    let args = extractFunctionArguments(inside);
    let sumExpr = args.map(arg => {
      arg = arg.trim();
      if (arg.includes(":")) {
        // Si es un rango (por ejemplo, A1:A10)
        let parts = arg.split(":");
        if (parts.length === 2) {
          let startCoords = getCellCoords(parts[0].trim());
          let endCoords = getCellCoords(parts[1].trim());
          let startX = Math.min(startCoords[0], endCoords[0]);
          let endX = Math.max(startCoords[0], endCoords[0]);
          let startY = Math.min(startCoords[1], endCoords[1]);
          let endY = Math.max(startCoords[1], endCoords[1]);
          let cells = [];
          for (let x = startX; x <= endX; x++) {
            for (let y = startY; y <= endY; y++) {
              let cellVal = state[x][y].computedValue;
              if (cellVal === "") cellVal = 0;
              cells.push(`(${cellVal})`);
            }
          }
          return cells.join(" + ");
        }
      } else {
        // Si es una referencia simple o un número
        if (/^[A-Za-z]+\d+$/.test(arg)) {
          let val = getCellValue(arg);
          if (val === "") val = 0;
          return `(${val})`;
        }
        return `(${arg})`;
      }
    }).join(" + ");
    let replacement = `(${sumExpr})`;
    formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
    index = formula.indexOf("SUMA(");
  }
  return formula;
}

/**
 * Reemplaza las llamadas a PROMEDIO()
 */
function replacePROMEDIO(formula) {
  let index = formula.indexOf("PROMEDIO(");
  if (index === -1) {
    index = formula.indexOf("PROMEDI(");
  }

  while (index !== -1) {
    let funcName = formula.substring(index, formula.indexOf("(", index));
    let startParenIndex = index + funcName.length;
    let end = findMatchingParen(formula, startParenIndex);
    
    if (end === -1) break;
    
    let inside = formula.substring(startParenIndex + 1, end);
    let args = extractFunctionArguments(inside);
    
    let valores = [];
    
    for (let arg of args) {
      arg = arg.trim();
      
      // Caso de rango con dos puntos
      if (arg.includes(":")) {
        let parts = arg.split(":");
        if (parts.length === 2) {
          if (/^\d+(\.\d+)?$/.test(parts[0]) && /^\d+(\.\d+)?$/.test(parts[1])) {
            valores.push(Number(parts[0]));
            valores.push(Number(parts[1]));
          } 

          else {
            try {
              let startCoords = getCellCoords(parts[0]);
              let endCoords = getCellCoords(parts[1]);
              
              let startX = Math.min(startCoords[0], endCoords[0]);
              let endX = Math.max(startCoords[0], endCoords[0]);
              let startY = Math.min(startCoords[1], endCoords[1]);
              let endY = Math.max(startCoords[1], endCoords[1]);
              
              for (let x = startX; x <= endX; x++) {
                for (let y = startY; y <= endY; y++) {
                  if (state[x] && state[x][y]) {
                    let cellVal = state[x][y].computedValue;
                    if (cellVal !== "" && !isNaN(Number(cellVal))) {
                      valores.push(Number(cellVal));
                    } else if (cellVal === "") {
                      valores.push(0);
                    }
                  }
                }
              }
            } catch (error) {
              console.error("Error al procesar rango de celdas:", error);
            }
          }
        }
      } 

      else if (/^[A-Za-z]+\d+$/.test(arg)) {
        let val = getCellValue(arg);
        if (val !== "" && !isNaN(Number(val))) {
          valores.push(Number(val));
        } else if (val === "") {
          valores.push(0);
        }
      } 
      // Si es un número literal
      else if (!isNaN(Number(arg))) {
        valores.push(Number(arg));
      }
    }
    
    // Calcular el promedio
    let suma = 0;
    let cantidad = valores.length;
    
    if (cantidad > 0) {
      suma = valores.reduce((total, valor) => total + valor, 0);
      let promedio = suma / cantidad;
      formula = formula.substring(0, index) + promedio + formula.substring(end + 1);
    } else {
      formula = formula.substring(0, index) + "#¡DIV/0!" + formula.substring(end + 1);
    }
    
    // Buscar la siguiente ocurrencia
    index = formula.indexOf("PROMEDIO(");
  }
  
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
	  
    // Depuración: ver la fórmula resultante antes de eval()
  console.log("Evaluating formula:", formula);

	  try {
      const result = eval(formula);
      if (result === Infinity || result === -Infinity) return "#!DiV/0!";
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
    
    // Cambiar a nueva hoja
    currentSheetIndex = index;
    state = sheets[currentSheetIndex].data;
    
    // Actualizar clases active
    document.querySelectorAll('.sheet').forEach((sheet, i) => {
      sheet.classList.toggle('active', i === index);
    });
    
    renderSpreadsheet();
  }


	// Eventos para cambiar de hoja
	document.querySelector('.sheets-container').addEventListener('click', (e) => {
    if (e.target.classList.contains('sheet')) {
      const index = Array.from(document.querySelectorAll('.sheet')).indexOf(e.target);
      switchSheet(index);
    }
  });
	
	// Agregar nueva hoja
	document.addEventListener("DOMContentLoaded", function () {
	  const sheetsContainer = document.querySelector(".sheets-container");
	  const addSheetBtn = document.getElementById("add-sheet");
	  let sheetCount = 1;
	  let contextMenu;
	
	  function addNewSheet() {
      sheetCount++;
      
      // Añadir nueva hoja al array sheets
      const newSheetData = {
        name: `Hoja${sheetCount}`,
        data: range(cols).map(() => 
          range(rows).map(() => ({ computedValue: "", value: "" }))
      )};
      sheets.push(newSheetData);
    
      // Crear elemento en el DOM
      let newSheetElement = document.createElement("div");
      newSheetElement.classList.add("sheet");
      newSheetElement.textContent = `Hoja${sheetCount}`;
      sheetsContainer.appendChild(newSheetElement);
      
      // Activar la nueva hoja automáticamente
      switchSheet(sheets.length - 1); // <- ¡Esto es lo que faltaba!
    }
	
	  addSheetBtn.addEventListener("click", addNewSheet);
	
	  sheetsContainer.addEventListener("click", function (e) {
	      if (e.target.classList.contains("sheet")) {
	          activateSheet(e.target);
	      }
	  });
	
	  sheetsContainer.addEventListener("contextmenu", function (e) {
	      e.preventDefault();
	
	      if (e.target.classList.contains("sheet")) {
	          if (contextMenu) contextMenu.remove();
	
	          const selectedSheet = e.target;
	          contextMenu = document.createElement("div");
	          contextMenu.classList.add("context-menu");
	          contextMenu.innerHTML = `
	              <button id="rename-sheet">Editar Nombre</button>
	              <button id="delete-sheet">Eliminar</button>
	          `;
	
	          document.body.appendChild(contextMenu);
	
	          // Obtener la posición exacta de la hoja
	          const rect = selectedSheet.getBoundingClientRect();
	          const menuHeight = contextMenu.offsetHeight;
	
	          let topPosition = rect.top - menuHeight - 5; 
	          if (topPosition < 0) {
	              topPosition = rect.bottom + 5; // Si no cabe arriba, poner debajo
	          }
	
	          contextMenu.style.position = "absolute";
	          contextMenu.style.top = `${topPosition}px`;
	          contextMenu.style.left = `${rect.left}px`;
	
	          // Editar nombre
	          document.getElementById("rename-sheet").addEventListener("click", function () {
	              let newName = prompt("Nuevo nombre de la hoja:", selectedSheet.textContent);
	              if (newName) selectedSheet.textContent = newName;
	              contextMenu.remove();
	          });
	
	          // Eliminar hoja
	          document.getElementById("delete-sheet").addEventListener("click", function () {
              if (document.querySelectorAll(".sheet").length > 1) {
                const sheetIndex = Array.from(sheetsContainer.children).indexOf(selectedSheet);
                
                // Eliminar del array
                sheets.splice(sheetIndex, 1);
                
                // Eliminar del DOM
                selectedSheet.remove();
                
                // Cambiar a hoja 0 si era la actual
                if (currentSheetIndex === sheetIndex) {
                  switchSheet(0);
                }
              } else {
                alert("Debe haber al menos una hoja.");
              }
              contextMenu.remove();
            });
	
	          // Cerrar menú al hacer clic afuera
	          document.addEventListener("click", function closeMenu(e) {
	              if (!contextMenu.contains(e.target)) {
	                  contextMenu.remove();
	                  document.removeEventListener("click", closeMenu);
	              }
	          });
	      }
	  });
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
	// Sección para Inicio (Coloca tu nombre aqui)
	// ================================
	

	// ================================
	// Sección para Insertar (Coloca tu nombre aqui)
	// ================================
	

	// ================================
	// Sección para Formulas (Coloca tu nombre aqui)
	// ================================
	
	// ================================
	// Sección para Datos (Coloca tu nombre aqui)
	// ================================
	