// ===== CONSTANTES Y SELECTORES =====
const $ = el => document.querySelector(el);
const $$ = el => document.querySelectorAll(el);

const $table = $('table');
const $head = $('thead');
const $body = $('tbody');

const rows = 100;
const cols = 26;
const letras = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

// ===== VARIABLES DE ESTADO =====
let selectedColumn = null;
let selectedRow = null;
let activeCell = { x: 0, y: 0 };
let currentSheetIndex = 0;
let sheets = [
	{ 
		name: 'Hoja1', 
		data: [],
		textBoxes: []
	}
];

let isSelectingRange = false;
let selectionStart = null;
let selectionEnd = null;

// ===== FUNCIONES UTILITARIAS =====
const range = length => Array.from({ length }, (_, i) => i);
const rangeChar = length => letras.slice(0, length);

// Inicializar datos para cada hoja
sheets.forEach(sheet => {
	sheet.data = range(cols).map(() =>
		range(rows).map(() => ({ computedValue: "", value: "" }))
	);
});

let state = sheets[currentSheetIndex].data;

// ===== INICIALIZACIÓN DE LA HOJA =====
document.addEventListener("DOMContentLoaded", function () {
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

// Renderización de la hoja de cálculo
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

// ===== FUNCIONES DE REFERENCIA DE CELDAS =====

// Convertir referencia a coordenadas
function getCellCoords(ref) {
	ref = ref.toUpperCase();
	const letter = ref.match(/[A-Z]+/)[0];
	const number = ref.match(/\d+/)[0];
	return [letras.indexOf(letter), parseInt(number) - 1];
}

// Obtener valor de celda
function getCellValue(ref) {
	try {
		const [x, y] = getCellCoords(ref);
		return state[x]?.[y]?.computedValue || "";
	} catch (error) {
		return `!ERROR: ${error.message}`;
	}
}

// ===== FUNCIONES PARA ANÁLISIS DE FÓRMULAS =====

// Transformar nombres a mayúsculas
function transformFunctionNames(formula) {
	const parts = formula.split(/(".*?")/);
	for (let i = 0; i < parts.length; i++) {
		if (i % 2 === 0) {
			parts[i] = parts[i].replace(/\b(si|o|y|concatenar)\b/gi, match => match.toUpperCase());
		}
	}
	return parts.join('');
}

// Encontrar paréntesis coincidente
function findMatchingParen(str, startIndex) {
	let count = 1;
	for (let i = startIndex + 1; i < str.length; i++) {
	  if (str[i] === '(') count++;
	  else if (str[i] === ')') {
		count--;
		if (count === 0) return i;
	  }
	}
	return -1;
}

// Extraer argumentos de función
function extractFunctionArguments(str) {
	const delimiter = str.includes(";") ? ";" : ",";
	let args = [];
	let current = "";
	let count = 0;
	let inQuote = false;
	
	for (let i = 0; i < str.length; i++) {
	  let char = str[i];
	  
	  if (char === '"' && (i === 0 || str[i - 1] !== '\\')) {
		inQuote = !inQuote;
	  }
	  
	  if (!inQuote && char === delimiter && count === 0) {
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

// ===== FUNCIONES MATEMÁTICAS =====

// Función MOD/RESIDUO
function replaceMOD(formula) {
	let index = formula.indexOf("RESIDUO(");
	
	while (index !== -1) {
		let funcName = "";
		if (formula.substring(index, index + 8) === "RESIDUO(") {
			funcName = "RESIDUO";
		}
		
		let startParenIndex = index + funcName.length;
		let end = findMatchingParen(formula, startParenIndex);
		
		if (end === -1) break;
		
		let inside = formula.substring(startParenIndex + 1, end);
		let args = extractFunctionArguments(inside);
		
		if (args.length !== 2) {
			formula = formula.substring(0, index) + "#¡ERROR!" + formula.substring(end + 1);
			// Buscar la siguiente aparición
			let nextIndex = Math.min(
				formula.indexOf("RESIDUO(", index + 1) !== -1 ? formula.indexOf("RESIDUO(", index + 1) : Infinity,
				formula.indexOf("RESIDU(", index + 1) !== -1 ? formula.indexOf("RESIDU(", index + 1) : Infinity,
				formula.indexOf("MOD(", index + 1) !== -1 ? formula.indexOf("MOD(", index + 1) : Infinity
			);
			index = nextIndex === Infinity ? -1 : nextIndex;
			continue;
		}
		
		// Procesar argumentos
		let arg1 = args[0].trim();
		let arg2 = args[1].trim();
		
		// Manejar referencias de celda
		if (/^[A-Za-z]+\d+$/.test(arg1)) {
			arg1 = getCellValue(arg1);
			if (arg1 === "") arg1 = 0;
		}
		
		if (/^[A-Za-z]+\d+$/.test(arg2)) {
			arg2 = getCellValue(arg2);
			if (arg2 === "") arg2 = 0;
		}
		
		// Verificar si el divisor es cero
		if (arg2 == 0) {
			formula = formula.substring(0, index) + "#¡DIV/0!" + formula.substring(end + 1);
		} else {
			// Usar el operador de módulo nativo de JavaScript
			let replacement = `(Number(${arg1}) % Number(${arg2}))`;
			formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
		}
		
		// Buscar la siguiente aparición
		let nextIndex = Math.min(
			formula.indexOf("RESIDUO(", index + 1) !== -1 ? formula.indexOf("RESIDUO(", index + 1) : Infinity,
			formula.indexOf("RESIDU(", index + 1) !== -1 ? formula.indexOf("RESIDU(", index + 1) : Infinity,
			formula.indexOf("MOD(", index + 1) !== -1 ? formula.indexOf("MOD(", index + 1) : Infinity
		);
		index = nextIndex === Infinity ? -1 : nextIndex;
	}
	
	return formula;
}

// Función SUMA
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

// Función PROMEDIO
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

			if (arg.includes(":")) {
				let parts = arg.split(":");
				if (parts.length === 2) {
					if (/^\d+(\.\d+)?$/.test(parts[0]) && /^\d+(\.\d+)?$/.test(parts[1])) {
						valores.push(Number(parts[0]));
						valores.push(Number(parts[1]));
					} else {
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
			} else if (/^[A-Za-z]+\d+$/.test(arg)) {
				let val = getCellValue(arg);
				if (val !== "" && !isNaN(Number(val))) {
					valores.push(Number(val));
				} else if (val === "") {
					valores.push(0);
				}
			} else if (!isNaN(Number(arg))) {
				valores.push(Number(arg));
			}
		}

		let suma = 0;
		let cantidad = valores.length;

		if (cantidad > 0) {
			suma = valores.reduce((total, valor) => total + valor, 0);
			let promedio = suma / cantidad;
			formula = formula.substring(0, index) + promedio + formula.substring(end + 1);
		} else {
			formula = formula.substring(0, index) + "#¡DIV/0!" + formula.substring(end + 1);
		}

		index = formula.indexOf("PROMEDIO(");
	}
	return formula;
}

// Función MAX
function replaceMAX(formula) {
	let index = formula.indexOf("MAX(");
	while (index !== -1) {
	  let end = findMatchingParen(formula, index + 3);
	  if (end === -1) break;
	  let inside = formula.substring(index + 4, end);
	  let args = extractFunctionArguments(inside);
	  let values = [];
	  
	  for (let arg of args) {
		arg = arg.trim();
		if (arg.includes(":")) {
		  let parts = arg.split(":");
		  if (parts.length === 2) {
			let startCoords = getCellCoords(parts[0].trim());
			let endCoords = getCellCoords(parts[1].trim());
			let startX = Math.min(startCoords[0], endCoords[0]);
			let endX = Math.max(startCoords[0], endCoords[0]);
			let startY = Math.min(startCoords[1], endCoords[1]);
			let endY = Math.max(startCoords[1], endCoords[1]);
			for (let x = startX; x <= endX; x++) {
			  for (let y = startY; y <= endY; y++) {
				if (state[x] && state[x][y]) {
				  let cellVal = state[x][y].computedValue;
				  if (cellVal !== "" && !isNaN(Number(cellVal))) {
					values.push(Number(cellVal));
				  } else if (cellVal === "") {
					values.push(0);
				  }
				}
			  }
			}
		  }
		} else if (/^[A-Za-z]+\d+$/.test(arg)) {
		  let val = getCellValue(arg);
		  if (val !== "" && !isNaN(Number(val))) {
			values.push(Number(val));
		  } else if (val === "") {
			values.push(0);
		  }
		} else if (!isNaN(Number(arg))) {
		  values.push(Number(arg));
		}
	  }
	  
	  let replacement;
	  if (values.length > 0) {
		replacement = Math.max(...values);
	  } else {
		replacement = "#¡DIV/0!";
	  }
	  
	  formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
	  index = formula.indexOf("MÁX(");
	}
	return formula;
}
  
// Función MIN
function replaceMIN(formula) {
	let index = formula.indexOf("MIN(");
	while (index !== -1) {
	  let end = findMatchingParen(formula, index + 3);
	  if (end === -1) break;
	  let inside = formula.substring(index + 4, end);
	  let args = extractFunctionArguments(inside);
	  let values = [];
	  
	  for (let arg of args) {
		arg = arg.trim();
		if (arg.includes(":")) {
		  let parts = arg.split(":");
		  if (parts.length === 2) {
			let startCoords = getCellCoords(parts[0].trim());
			let endCoords = getCellCoords(parts[1].trim());
			let startX = Math.min(startCoords[0], endCoords[0]);
			let endX = Math.max(startCoords[0], endCoords[0]);
			let startY = Math.min(startCoords[1], endCoords[1]);
			let endY = Math.max(startCoords[1], endCoords[1]);
			for (let x = startX; x <= endX; x++) {
			  for (let y = startY; y <= endY; y++) {
				if (state[x] && state[x][y]) {
				  let cellVal = state[x][y].computedValue;
				  if (cellVal !== "" && !isNaN(Number(cellVal))) {
					values.push(Number(cellVal));
				  } else if (cellVal === "") {
					values.push(0);
				  }
				}
			  }
			}
		  }
		} else if (/^[A-Za-z]+\d+$/.test(arg)) {
		  let val = getCellValue(arg);
		  if (val !== "" && !isNaN(Number(val))) {
			values.push(Number(val));
		  } else if (val === "") {
			values.push(0);
		  }
		} else if (!isNaN(Number(arg))) {
		  values.push(Number(arg));
		}
	  }
	  
	  let replacement;
	  if (values.length > 0) {
		replacement = Math.min(...values);
	  } else {
		replacement = "#¡DIV/0!";
	  }
	  
	  formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
	  index = formula.indexOf("MÍN(");
	}
	return formula;
}

// Función CONTAR
function replaceCONTAR(formula) {
    let index = formula.indexOf("CONTAR(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 6);
        if (end === -1) break;
        
        let inside = formula.substring(index + 7, end);
        let args = extractFunctionArguments(inside);
        let conteo = 0;
        
        for (let arg of args) {
            arg = arg.trim();
            if (arg.includes(":")) {
                let parts = arg.split(":");
                if (parts.length === 2) {
                    let startCoords = getCellCoords(parts[0].trim());
                    let endCoords = getCellCoords(parts[1].trim());
                    let startX = Math.min(startCoords[0], endCoords[0]);
                    let endX = Math.max(startCoords[0], endCoords[0]);
                    let startY = Math.min(startCoords[1], endCoords[1]);
                    let endY = Math.max(startCoords[1], endCoords[1]);
                    
                    for (let x = startX; x <= endX; x++) {
                        for (let y = startY; y <= endY; y++) {
                            if (state[x] && state[x][y]) {
                                let cellVal = state[x][y].computedValue;
                                if (cellVal !== "" && !isNaN(Number(cellVal))) {
                                    conteo++;
                                }
                            }
                        }
                    }
                }
            } else if (/^[A-Za-z]+\d+$/.test(arg)) {
                let val = getCellValue(arg);
                if (val !== "" && !isNaN(Number(val))) {
                    conteo++;
                }
            } else if (!isNaN(Number(arg))) {
                conteo++;
            }
        }
        
        formula = formula.substring(0, index) + conteo + formula.substring(end + 1);
        index = formula.indexOf("CONTAR(", index + 1);
    }
    return formula;
}

// ===== FUNCIONES LÓGICAS =====

// Función SI
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

// Función Y
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

// Función O
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

// Función CAMBIAR
function replaceCAMBIAR(formula) {
	let index = formula.indexOf("CAMBIAR(");
	while (index !== -1) {
		let end = findMatchingParen(formula, index + 7);
		if (end === -1) break;
		
		let inside = formula.substring(index + 8, end);
		let args = extractFunctionArguments(inside);
		
		if (args.length !== 3) {
			formula = formula.substring(0, index) + "#¡ERROR!" + formula.substring(end + 1);
			index = formula.indexOf("CAMBIAR(", index + 1);
			continue;
		}
		
		let [condicion, valorVerdadero, valorFalso] = args.map(arg => arg.trim());
		
		let condicionValor;
		if (condicion.toUpperCase() === "VERDADERO()") {
			condicionValor = true;
		} else if (condicion.toUpperCase() === "FALSO()") {
			condicionValor = false;
		} else if (/^[A-Za-z]+\d+$/.test(condicion)) {
			let valor = getCellValue(condicion);
			condicionValor = valor === "VERDADERO" || valor === "TRUE" || valor === "1";
		} else {
			try {
				condicionValor = eval(replaceFunctions(condicion));
			} catch (e) {
				condicionValor = false;
			}
		}
		
		let replacement = condicionValor ? valorVerdadero : valorFalso;
		formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
		index = formula.indexOf("CAMBIAR(", index + 1);
	}
	return formula;
}

// Función FALSO
function replaceFALSO(formula) {
	let index = formula.indexOf("FALSO()");
	while (index !== -1) {
		formula = formula.substring(0, index) + "false" + formula.substring(index + 7);
		index = formula.indexOf("FALSO()", index + 1);
	}
	return formula;
}

// Función NO
function replaceNO(formula) {
	let index = formula.indexOf("NO(");
	while (index !== -1) {
		let end = findMatchingParen(formula, index + 2);
		if (end === -1) break;
		
		let inside = formula.substring(index + 3, end);
		let argumento = inside.trim();
		let valor;
		
		if (argumento.toUpperCase() === "VERDADERO()") {
			valor = false;
		} else if (argumento.toUpperCase() === "FALSO()") {
			valor = true;
		} else if (/^[A-Za-z]+\d+$/.test(argumento)) {
			let celdaValor = getCellValue(argumento);
			if (celdaValor === "VERDADERO" || celdaValor === "TRUE" || celdaValor === "1" || celdaValor === true) {
				valor = false;
			} else if (celdaValor === "FALSO" || celdaValor === "FALSE" || celdaValor === "0" || celdaValor === false) {
				valor = true;
			} else {
				valor = true;
			}
		} else {
			try {
				let evalResult = eval(replaceFunctions(argumento));
				valor = !evalResult;
			} catch (e) {
				valor = true;
			}
		}
		
		formula = formula.substring(0, index) + valor + formula.substring(end + 1);
		index = formula.indexOf("NO(", index + 1);
	}
	return formula;
}

// Función VERDADERO
function replaceVERDADERO(formula) {
	let index = formula.indexOf("VERDADERO()");
	while (index !== -1) {
		formula = formula.substring(0, index) + "true" + formula.substring(index + 11);
		index = formula.indexOf("VERDADERO()", index + 1);
	}
	return formula;
}

// ===== FUNCIONES DE TEXTO =====

// Función EXTRAE
function replaceEXTRAE(formula) {
	let index = formula.indexOf("EXTRAE(");
	while (index !== -1) {
	  let end = findMatchingParen(formula, index + 6);
	  if (end === -1) break;
	  
	  let inside = formula.substring(index + 7, end);
	  let args = extractFunctionArguments(inside);
	  
	  if (args.length < 3) {
		formula = formula.substring(0, index) + "#¡ERROR!" + formula.substring(end + 1);
		index = formula.indexOf("EXTRAE(", index + 1);
		continue;
	  }
	  
	  let textArg = args[0].trim();
	  let startArg = args[1].trim();
	  let numCharsArg = args[2].trim();
	  
	  let textVal = "";
	  
	  if (/^[A-Za-z]+\d+$/.test(textArg)) {
		textVal = getCellValue(textArg);
		if (textVal === undefined || textVal === null) textVal = "";
	  } else if (textArg.startsWith('"') && textArg.endsWith('"')) {
		textVal = textArg.slice(1, -1);
	  } else {
		try {
		  textVal = String(eval(replaceFunctions(textArg)));
		} catch (e) {
		  textVal = textArg;
		}
	  }
	  
	  let startIndex, lengthVal;
	  
	  if (/^[A-Za-z]+\d+$/.test(startArg)) {
		startIndex = Number(getCellValue(startArg));
	  } else {
		startIndex = Number(startArg);
	  }
	  
	  if (/^[A-Za-z]+\d+$/.test(numCharsArg)) {
		lengthVal = Number(getCellValue(numCharsArg));
	  } else {
		lengthVal = Number(numCharsArg);
	  }
	  
	  let replacement;
	  if (!isNaN(startIndex) && !isNaN(lengthVal) && startIndex > 0) {
		try {
		  replacement = textVal.substring(startIndex - 1, startIndex - 1 + lengthVal);
		  replacement = `"${replacement}"`;
		} catch (e) {
		  replacement = "#¡ERROR!";
		}
	  } else {
		replacement = "#¡ERROR!";
	  }
	  
	  formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
	  index = formula.indexOf("EXTRAE(", index + 1);
	}
	return formula;
}

// ===== FUNCIONES DE FECHA Y HORA =====

// Función AHORA
function replaceAHORA(formula) {
    let index = formula.indexOf("AHORA()");
    while (index !== -1) {
        let fechaHora = new Date().toLocaleString();
        formula = formula.substring(0, index) + `"${fechaHora}"` + formula.substring(index + 7);
        index = formula.indexOf("AHORA()", index + 1);
    }
    return formula;
}

// Analizar cadena de fecha en diferentes formatos
function parseDateString(dateStr) {
    if (!dateStr) return null;

    if (typeof dateStr === 'string') {
        dateStr = dateStr.replace(/^"|"$/g, '');
    }

    const formats = [
        /(\d{1,2})\/(\d{1,2})\/(\d{4}),\s*(\d{1,2}):(\d{2}):(\d{2})\s*(A|P)\.M\./i,
        /(\d{1,2})\/(\d{1,2})\/(\d{4}),\s*(\d{1,2}):(\d{2}):(\d{2})\s*(A|P)M/i,
        /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
        /(\d{4})-(\d{2})-(\d{2})/
    ];

    for (const format of formats) {
        const match = dateStr.match(format);
        if (match) {
            if (match.length === 8) {
                const [_, day, month, year, hour, minute, second, ampm] = match;
                let h = parseInt(hour);
                if (ampm.toUpperCase() === 'P' && h !== 12) h += 12;
                if (ampm.toUpperCase() === 'A' && h === 12) h = 0;
                return new Date(year, month - 1, day, h, minute, second);
            } else if (match.length === 4) {
                const [_, day, month, year] = match;
                return new Date(year, month - 1, day);
            }
        }
    }

    const date = new Date(dateStr);
    if (!isNaN(date.getTime())) return date;

    return null;
}

// Función AÑO
function replaceAÑO(formula) { 
    let index = formula.indexOf("AÑO(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 3);
        if (end === -1) break;
    
        let inside = formula.substring(index + 4, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            replacement = fecha.getFullYear();
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("AÑO(");
    }
    return formula;
}

// Función MES
function replaceMES(formula) {
    let index = formula.indexOf("MES(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 3);
        if (end === -1) break;
        
        let inside = formula.substring(index + 4, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            replacement = fecha.getMonth() + 1;
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("MES(");
    }
    return formula;
}

// Función DIA
function replaceDIA(formula) {
    let index = formula.indexOf("DIA(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 3);
        if (end === -1) break;
        
        let inside = formula.substring(index + 4, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            replacement = fecha.getDate();
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("DIA(");
    }
    return formula;
}

// Función HORA
function replaceHORA(formula) {
    let index = formula.indexOf("HORA(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 4);
        if (end === -1) break;
        
        let inside = formula.substring(index + 5, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            let hora = fecha.getHours();
            if (hora === 0) hora = 24;
            replacement = hora;
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("HORA(");
    }
    return formula;
}

// Función MINUTO
function replaceMINUTO(formula) {
    let index = formula.indexOf("MINUTO(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 6);
        if (end === -1) break;
        
        let inside = formula.substring(index + 7, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            let minutos = fecha.getMinutes();
            if (minutos < 10) minutos = "0" + minutos;
            replacement = minutos;
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("MINUTO(");
    }
    return formula;
}

// Función SEGUNDO
function replaceSEGUNDO(formula) {
    let index = formula.indexOf("SEGUNDO(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 7);
        if (end === -1) break;
        
        let inside = formula.substring(index + 8, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                fecha = parseDateString(valor);
                if (!fecha) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fecha = parseDateString(computedVal);
                }
            }
        } else {
            fecha = parseDateString(argumento);
        }
        
        let replacement;
        if (fecha && !isNaN(fecha.getTime())) {
            let segundos = fecha.getSeconds();
            if (segundos < 10) segundos = "0" + segundos;
            replacement = segundos;
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("SEGUNDO(");
    }
    return formula;
}

// Función FECHA
function replaceFECHA(formula) {
    let index = formula.indexOf("FECHA(");
    while (index !== -1) {
        let end = findMatchingParen(formula, index + 5);
        if (end === -1) break;
        
        let inside = formula.substring(index + 6, end);
        let argumento = inside.trim();
        let fecha;
        
        if (argumento.toUpperCase() === "AHORA()") {
            fecha = new Date().toLocaleDateString();
        } else if (/^[A-Za-z]+\d+$/.test(argumento)) {
            let valor = getCellValue(argumento);
            if (valor) {
                let fechaObj = parseDateString(valor);
                if (!fechaObj) {
                    let computedVal = state[letras.indexOf(argumento.match(/[A-Za-z]+/)[0])][parseInt(argumento.match(/\d+/)[0]) - 1].computedValue;
                    fechaObj = parseDateString(computedVal);
                }
                if (fechaObj && !isNaN(fechaObj.getTime())) {
                    fecha = fechaObj.toLocaleDateString();
                }
            }
        } else {
            let fechaObj = parseDateString(argumento);
            if (fechaObj && !isNaN(fechaObj.getTime())) {
                fecha = fechaObj.toLocaleDateString();
            }
        }
        
        let replacement;
        if (fecha) {
            replacement = `"${fecha}"`;
        } else {
            replacement = "#¡ERROR!";
        }
        
        formula = formula.substring(0, index) + replacement + formula.substring(end + 1);
        index = formula.indexOf("FECHA(");
    }
    return formula;
}

// Función HOY
function replaceHOY(formula) {
    let index = formula.indexOf("HOY()");
    while (index !== -1) {
        let fecha = new Date().toLocaleDateString();
        formula = formula.substring(0, index) + `"${fecha}"` + formula.substring(index + 5);
        index = formula.indexOf("HOY()", index + 1);
    }
    return formula;
}

// ===== PROCESAMIENTO DE FÓRMULAS =====

// Reemplazar funciones reconocidas en la fórmula
function replaceFunctions(formula) {
	// Manejar primero RESIDUO antes de que pueda ser alterado por otras funciones
	if (formula.includes("RESIDUO(") || formula.includes("RESIDU(") || formula.includes("MOD(")) {
		formula = replaceMOD(formula);
	}
	
	let prev;
	do {
		prev = formula;
		// Primero evaluamos las funciones de fecha y hora
		formula = replaceAHORA(formula);
		formula = replaceHOY(formula);
		// Luego las funciones que dependen de fechas
		formula = replaceAÑO(formula);
		formula = replaceMES(formula);
		formula = replaceDIA(formula);
		formula = replaceHORA(formula);
		formula = replaceMINUTO(formula);
		formula = replaceSEGUNDO(formula);
		formula = replaceFECHA(formula);
		// Luego las funciones lógicas básicas
		formula = replaceVERDADERO(formula);
		formula = replaceFALSO(formula);
		formula = replaceNO(formula);
		// Luego las funciones lógicas compuestas
		formula = replaceO(formula);
		formula = replaceY(formula);
		formula = replaceSI(formula);
		formula = replaceCAMBIAR(formula);
		// Finalmente las demás funciones
		formula = replaceSUMA(formula);
		formula = replacePROMEDIO(formula);
		formula = replaceMAX(formula);
		formula = replaceMIN(formula);
		formula = replaceEXTRAE(formula);
		formula = replaceCONTAR(formula);
	} while (formula !== prev);
	
	return formula;
}

// Calcular el valor computado a partir de una fórmula
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
		if (result === Infinity || result === -Infinity) return "#!DiV/0!";
		return result === 0 ? "" : result;
	} catch (error) {
		return `!ERROR: ${error.message}`;
	}
}

// ===== ACTUALIZACIÓN DE CELDAS =====

// Actualizar una celda específica
function updateCell(x, y, value) {
	x = parseInt(x);
	y = parseInt(y);
	
	const newState = structuredClone(state);

	newState[x][y] = {
		value: value,
		computedValue: value.startsWith('=') ? value : calculateComputedValue(value)
	};

	state = newState;
	sheets[currentSheetIndex].data = state;
	updateAllDependentCells();

	const selector = `td[data-x="${x}"][data-y="${y}"]`;
	const cell = document.querySelector(selector);
	if (cell) {
		cell.querySelector('input').value = value;
		cell.querySelector('span').textContent = state[x][y].computedValue;
	}
}

// Recalcular todas las celdas con fórmulas
function updateAllDependentCells() {
	state.forEach((col, x) => {
		col.forEach((cell, y) => {
			if (cell.value.startsWith('=')) {
				const newValue = calculateComputedValue(cell.value);
				state[x][y].computedValue = newValue;

				const selector = `td[data-x="${x}"][data-y="${y}"]`;
				const cellElement = document.querySelector(selector);
				if (cellElement) {
					cellElement.querySelector('span').textContent = newValue;
				}
			}
		});
	});
}

// ===== FUNCIONES DE SELECCIÓN Y UI =====

// Limpiar selecciones de celdas
function clearSelections() {
	$$('.selected-column').forEach(el => el.classList.remove('selected-column'));
	$$('.selected-row').forEach(el => el.classList.remove('selected-row'));
	$$('.selected-header').forEach(el => el.classList.remove('selected-header'));
	$$('.celdailumida').forEach(el => el.classList.remove('celdailumida'));
	$$('.selected-range').forEach(el => el.classList.remove('selected-range'));
	selectedColumn = null;
	selectedRow = null;
}

// Actualizar visualización de celda activa
function updateActiveCellDisplay(x, y) {
	activeCell = { x, y };
	const columnLetter = letras[x];
	const rowNumber = y + 1;
	activeCellDisplay.textContent = `${columnLetter}${rowNumber}`;
	formulaInput.value = state[x][y].value;
}

// Limpiar selección de rango
function clearRangeSelection() {
	$$('.selected-range').forEach(el => el.classList.remove('selected-range'));
}

// Resaltar un rango
function highlightRange(start, end) {
	clearRangeSelection();
	const startX = Math.min(start.x, end.x);
	const endX = Math.max(start.x, end.x);
	const startY = Math.min(start.y, end.y);
	const endY = Math.max(start.y, end.y);
	for (let x = startX; x <= endX; x++) {
	  for (let y = startY; y <= endY; y++) {
		const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
		if (cell) {
		  cell.classList.add('selected-range');
		}
	  }
	}
}

// Actualizar visualización de rango seleccionado
function updateActiveCellDisplayRange(start, end) {
	const startX = Math.min(start.x, end.x);
	const endX = Math.max(start.x, end.x);
	const startY = Math.min(start.y, end.y);
	const endY = Math.max(start.y, end.y);
	const startRef = letras[startX] + (startY + 1);
	const endRef = letras[endX] + (endY + 1);
	activeCellDisplay.textContent = (startRef === endRef) ? startRef : `${startRef}:${endRef}`;
}

// ===== EVENTOS DE INTERACCIÓN =====

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

// Eventos en el cuerpo de la tabla
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

// Eventos para selección de rango
$body.addEventListener('mousedown', event => {
	const td = event.target.closest('td');
	if (!td) return;
	clearSelections();
	isSelectingRange = true;
	const x = parseInt(td.dataset.x);
	const y = parseInt(td.dataset.y);
	selectionStart = { x, y };
	selectionEnd = { x, y };
	highlightRange(selectionStart, selectionEnd);
	updateActiveCellDisplayRange(selectionStart, selectionEnd);
});
  
$body.addEventListener('mouseover', event => {
	if (!isSelectingRange) return;
	const td = event.target.closest('td');
	if (!td) return;
	const x = parseInt(td.dataset.x);
	const y = parseInt(td.dataset.y);
	selectionEnd = { x, y };
	highlightRange(selectionStart, selectionEnd);
	updateActiveCellDisplayRange(selectionStart, selectionEnd);
});
  
document.addEventListener('mouseup', event => {
	if (isSelectingRange) {
		isSelectingRange = false;
	}
});

// Manejo de portapapeles - copiar
document.addEventListener('copy', (e) => {
	let data = [];
	const selectedRangeCells = document.querySelectorAll('td.selected-range');
	
	if (selectedRangeCells.length > 0) {
		let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
		selectedRangeCells.forEach(cell => {
			let x = parseInt(cell.dataset.x);
			let y = parseInt(cell.dataset.y);
			if (x < minX) minX = x;
			if (x > maxX) maxX = x;
			if (y < minY) minY = y;
			if (y > maxY) maxY = y;
		});
		
		for (let y = minY; y <= maxY; y++) {
			let rowData = [];
			for (let x = minX; x <= maxX; x++) {
				const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
				if (cell && cell.classList.contains('selected-range')) {
					rowData.push(cell.querySelector('span').textContent);
				} else {
					rowData.push("");
				}
			}
			data.push(rowData.join("\t"));
		}
	} else if (selectedColumn !== null) {
		data = state[selectedColumn]
			.map(cell => cell.computedValue)
			.filter(value => value !== "");
	} else if (selectedRow !== null) {
		data = state.map(col => col[selectedRow].computedValue)
			.filter(value => value !== "");
	} else {
		const selectedCell = document.querySelector('td.celdailumida');
		if (selectedCell) {
			const x = parseInt(selectedCell.dataset.x);
			const y = parseInt(selectedCell.dataset.y);
			data.push(state[x][y].computedValue);
		}
	}
	
	if (data.length > 0) {
		e.clipboardData.setData('text/plain', data.join('\n'));
		e.preventDefault();
	}
});

// Manejo de portapapeles - pegar
document.addEventListener('paste', (e) => {
	if (e.target === formulaInput) return;
	const clipboardData = e.clipboardData.getData('text/plain');
	if (!clipboardData) return;

	const rowsData = clipboardData.split('\n').filter(v => v !== '');
	const isSingleValue = (rowsData.length === 1 && rowsData[0].indexOf('\t') === -1);

	const selectedRangeCells = document.querySelectorAll('td.selected-range');
	if (selectedRangeCells.length > 0) {
		let minX = Infinity, minY = Infinity;
		selectedRangeCells.forEach(cell => {
			const x = parseInt(cell.dataset.x);
			const y = parseInt(cell.dataset.y);
			if (x < minX) minX = x;
			if (y < minY) minY = y;
		});
		
		if (isSingleValue) {
			selectedRangeCells.forEach(cell => {
				const x = parseInt(cell.dataset.x);
				const y = parseInt(cell.dataset.y);
				updateCell(x, y, clipboardData);
			});
		} else {
			rowsData.forEach((row, rowIndex) => {
				const cells = row.split('\t');
				cells.forEach((cellValue, colIndex) => {
					const targetX = minX + colIndex;
					const targetY = minY + rowIndex;
					if (targetX < cols && targetY < rows) {
						updateCell(targetX, targetY, cellValue);
					}
				});
			});
		}
		e.preventDefault();
		return;
	} else if (selectedColumn !== null) {
		rowsData.forEach((value, y) => {
			if (y < rows) updateCell(selectedColumn, y, value);
		});
		e.preventDefault();
		return;
	} else if (selectedRow !== null) {
		rowsData.forEach((row, x) => {
			if (x < cols) updateCell(x, selectedRow, row);
		});
		e.preventDefault();
		return;
	} else {
		if (activeCell) {
			updateCell(activeCell.x, activeCell.y, clipboardData);
		}
	}
	e.preventDefault();
});

// Edición de celdas
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

// Borrado de contenido
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

document.addEventListener('keydown', (e) => {
	if (e.key === 'Delete' || e.key === 'Backspace') {
		const selectedRangeCells = document.querySelectorAll('td.selected-range');
		if (selectedRangeCells.length > 0) {
			selectedRangeCells.forEach(cell => {
				const x = parseInt(cell.dataset.x);
				const y = parseInt(cell.dataset.y);
				updateCell(x, y, "");
			});
			e.preventDefault();
		} else if (selectedColumn !== null) {
			state[selectedColumn].forEach((cell, y) => {
				updateCell(selectedColumn, y, "");
			});
			e.preventDefault();
		} else if (selectedRow !== null) {
			state.forEach((col, x) => {
				updateCell(x, selectedRow, "");
			});
			e.preventDefault();
		}
	}
});

// ===== GESTIÓN DE HOJAS =====

// Cambiar a una hoja específica
function switchSheet(index) {
    sheets[currentSheetIndex].data = state;
    currentSheetIndex = index;
    state = sheets[currentSheetIndex].data;

    document.querySelectorAll('.sheet').forEach((sheet, i) => {
        sheet.classList.toggle('active', i === index);
    });

    const event = new CustomEvent('sheetChanged', { 
        detail: { 
            previousIndex: currentSheetIndex,
            newIndex: index 
        } 
    });
    document.dispatchEvent(event);

    renderSpreadsheet();
}

// Eventos para cambiar de hoja
document.querySelector('.sheets-container').addEventListener('click', (e) => {
	if (e.target.classList.contains('sheet')) {
		const index = Array.from(document.querySelectorAll('.sheet')).indexOf(e.target);
		switchSheet(index);
	}
});

// Agregar y gestionar hojas
document.addEventListener("DOMContentLoaded", function () {
	const sheetsContainer = document.querySelector(".sheets-container");
	const addSheetBtn = document.getElementById("add-sheet");
	let sheetCount = 1;
	let contextMenu;

	function addNewSheet() {
		sheetCount++;

		const newSheetData = {
			name: `Hoja${sheetCount}`,
			data: range(cols).map(() =>
				range(rows).map(() => ({ computedValue: "", value: "" }))
			),
			textBoxes: []
		};
		sheets.push(newSheetData);

		let newSheetElement = document.createElement("div");
		newSheetElement.classList.add("sheet");
		newSheetElement.textContent = `Hoja${sheetCount}`;
		sheetsContainer.appendChild(newSheetElement);

		switchSheet(sheets.length - 1);
	}

	addSheetBtn.addEventListener("click", addNewSheet);

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

			const rect = selectedSheet.getBoundingClientRect();
			const menuHeight = contextMenu.offsetHeight;

			let topPosition = rect.top - menuHeight - 5;
			if (topPosition < 0) {
				topPosition = rect.bottom + 5;
			}

			contextMenu.style.position = "absolute";
			contextMenu.style.top = `${topPosition}px`;
			contextMenu.style.left = `${rect.left}px`;

			document.getElementById("rename-sheet").addEventListener("click", function () {
                let newName = prompt("Nuevo nombre de la hoja:", selectedSheet.textContent);
                if (newName) {
                    selectedSheet.textContent = newName;
                    let sheetIndex = Array.from(document.querySelectorAll(".sheet")).indexOf(selectedSheet);
                    sheets[sheetIndex].name = newName;
                }
                contextMenu.remove();
            });

			document.getElementById("delete-sheet").addEventListener("click", function () {
				if (document.querySelectorAll(".sheet").length > 1) {
					const sheetIndex = Array.from(sheetsContainer.children).indexOf(selectedSheet);
					sheets.splice(sheetIndex, 1);
					selectedSheet.remove();

					if (currentSheetIndex === sheetIndex) {
						switchSheet(0);
					}
				} else {
					alert("Debe haber al menos una hoja.");
				}
				contextMenu.remove();
			});

			document.addEventListener("click", function closeMenu(e) {
				if (!contextMenu.contains(e.target)) {
					contextMenu.remove();
					document.removeEventListener("click", closeMenu);
				}
			});
		}
	});
});

// ===== BARRA DE FÓRMULAS =====

// Actualización de la barra de fórmulas
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

// Inicializar la hoja de cálculo
renderSpreadsheet();
