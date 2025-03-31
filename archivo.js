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
  const nombreInput = document.getElementById('filename').value.trim();
  
  // Verificar si el campo está vacío
  if (!nombreInput) {
    alert("Por favor, introduce un nombre para el archivo");
    return; // Detener la ejecución si no hay nombre
  }
  
  const nombreArchivo = nombreInput;

  // Actualizar el título del archivo en la interfaz
  const tituloElemento = document.querySelector('.ribbon-title h1');
  if (tituloElemento) {
      tituloElemento.textContent = nombreArchivo;
  }
  
  // Crear el libro de trabajo de Excel
  const wb = XLSX.utils.book_new();

  // Extraer correctamente los nombres de las hojas desde el DOM
  document.querySelectorAll(".sheet").forEach((sheetElement, index) => {
    sheets[index].name = sheetElement.textContent.trim(); // Guardamos el nombre correcto en sheets
  });

  // Guardar primero la hoja actual en la estructura de datos
  sheets[currentSheetIndex].data = state;

  // Procesar cada hoja
  sheets.forEach((sheet, sheetIndex) => {
    const datos = [];
    const sheetName = sheet.name || `Hoja${sheetIndex + 1}`;
    const formulas = {}; // Objeto para almacenar las fórmulas

    // Limpiar caracteres inválidos en nombres de hojas de Excel
    let nombreValido = sheetName.replace(/[\[\]:\*\?\/\\]/g, "").trim();
    if (nombreValido.length > 31) nombreValido = nombreValido.substring(0, 31);

    // Asegurar que la hoja tiene datos válidos
    if (!sheet.data || !Array.isArray(sheet.data)) return;

    for (let y = 0; y < rows; y++) {
      const filaDatos = [];
      for (let x = 0; x < cols; x++) {
        const celda = sheet.data[x]?.[y] || { value: "", computedValue: "" };

        if (typeof celda.value === "string" && celda.value.startsWith('=')) {
          filaDatos.push(null); // Dejar la celda vacía para Excel
          const cellRef = XLSX.utils.encode_cell({ c: x, r: y }); 
          formulas[cellRef] = celda.value.substring(1); // Guardar la fórmula sin '='
        } else {
          filaDatos.push(celda.computedValue || "");
        }
      }
      datos.push(filaDatos);
    }

    // Crear hoja de Excel con los datos
    const ws = XLSX.utils.aoa_to_sheet(datos);

    // Aplicar manualmente las fórmulas en las celdas correctas
    Object.keys(formulas).forEach(cellRef => {
      if (!ws[cellRef]) ws[cellRef] = {}; // Asegurar que la celda existe
      ws[cellRef].f = formulas[cellRef]; // Asignar la fórmula correctamente
    });

    // Añadir la hoja al libro de trabajo con su nombre correcto
    XLSX.utils.book_append_sheet(wb, ws, nombreValido);
  });

  // Convertir el libro en binario
  const wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'binary' });

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
    console.log("Agregando archivo reciente:", nombreArchivo); // Para depuración
    
    // Obtener la lista actual de archivos recientes
    let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
    
    // Evitar duplicados
    const index = archivos.indexOf(nombreArchivo);
    if (index !== -1) {
        archivos.splice(index, 1);
    }
    
    // Agregar al inicio del array
    archivos.unshift(nombreArchivo);
    
    // Mantener solo los últimos 10 archivos
    archivos = archivos.slice(0, 10);
    
    // Guardar la lista actualizada
    localStorage.setItem("archivosRecientes", JSON.stringify(archivos));
    
    // Guardar el estado actual del archivo
    localStorage.setItem('file_' + nombreArchivo, JSON.stringify(state));
    
    // Actualizar la visualización
    mostrarArchivosRecientes();
}

// Función para mostrar archivos recientes
function mostrarArchivosRecientes() {
    console.log("Actualizando lista de recientes"); // Para depuración
    
    let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
    let lista = document.getElementById("lista-recientes");
    
    // Limpiar la lista actual
    lista.innerHTML = "";
    
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
        
        // Contenedor principal
        let fileContainer = document.createElement("div");
        fileContainer.className = "file-container";
        
        // Icono del archivo
        let icon = document.createElement("img");
        icon.src = "images/logoexcel.png";
        icon.alt = "Excel";
        icon.className = "file-icon";
        
        // Nombre del archivo
        let nameSpan = document.createElement("span");
        nameSpan.textContent = archivo;
        nameSpan.className = "file-name";
        
        // Botón de eliminar
        let deleteBtn = document.createElement("button");
        deleteBtn.className = "delete-file-btn";
        deleteBtn.innerHTML = '<i class="fas fa-trash"></i>';
        deleteBtn.title = "Eliminar de recientes";
        
        // Agregar elementos al contenedor
        fileContainer.appendChild(icon);
        fileContainer.appendChild(nameSpan);
        li.appendChild(fileContainer);
        li.appendChild(deleteBtn);
        
        // Evento para eliminar archivo
        deleteBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            eliminarArchivoReciente(archivo);
        });
        
        // Evento para abrir archivo
        fileContainer.addEventListener("click", () => {
            abrirArchivoReciente(archivo);
        });
        
        lista.appendChild(li);
    });
}

// Función para eliminar archivo de recientes
function eliminarArchivoReciente(nombreArchivo) {
    let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
    const index = archivos.indexOf(nombreArchivo);
    
    if (index !== -1) {
        archivos.splice(index, 1);
        localStorage.setItem("archivosRecientes", JSON.stringify(archivos));
        localStorage.removeItem('file_' + nombreArchivo); // Eliminar también el estado guardado
        mostrarArchivosRecientes();
    }
}

// Función para abrir archivo reciente
function abrirArchivoReciente(nombreArchivo) {
    const savedState = localStorage.getItem('file_' + nombreArchivo);
    if (savedState) {
        // Abrir en nueva pestaña
        window.open(`main.html?archivo=${encodeURIComponent(nombreArchivo)}`, '_blank');
    } else {
        alert("No se encontró el archivo en el almacenamiento local.");
    }
}

// Asegurarse de que la lista se actualice cuando se carga la página
document.addEventListener("DOMContentLoaded", function() {
    mostrarArchivosRecientes();
    
    // Verificar si hay un archivo en la URL
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

// Función para exportar la hoja de cálculo a diferentes formatos
function exportarHojaCalculo(formato) {
  // Obtener el nombre del archivo sin extensión
  const nombreInput = document.getElementById('nombre-exportacion').value.trim();
  
  // Verificar si el campo está vacío
  if (!nombreInput) {
    alert("Por favor, introduce un nombre para el archivo");
    return; // Detener la ejecución si no hay nombre
  }
  
  // Exportar según el formato seleccionado
  switch (formato) {
    case 'csv':
      exportarCSV(nombreInput);
      break;
    case 'txt':
      exportarTXT(nombreInput);
      break;
    case 'pdf':
      exportarPDF(nombreInput);
      break;
    default:
      alert("Formato no válido");
  }
}

// Función para exportar a CSV con separador de punto y coma y codificación UTF-8 con BOM
function exportarCSV(nombreArchivo) {
    // Verificar si el campo de nombre del archivo está vacío
    if (!nombreArchivo) {
        alert("Por favor, introduce un nombre para el archivo");
        return; // Detener la ejecución si no hay nombre
    }

    let contenidoCSV = [];

    // Garantizar que la hoja actual está guardada en la estructura de datos
    sheets[currentSheetIndex].data = state;

    // Obtener los datos de la hoja actual
    const hojaActual = sheets[currentSheetIndex].data;

    // Recorrer las filas y columnas de la hoja actual
    for (let y = 0; y < rows; y++) {
        let fila = [];
        for (let x = 0; x < cols; x++) {
            // Agregar el valor computado o un texto vacío
            fila.push(`"${hojaActual[x][y].computedValue || ""}"`); // Escapar valores con comillas
        }
        contenidoCSV.push(fila.join(';')); // Usar punto y coma como separador
    }

    // Convertir a string y agregar el BOM al inicio
    const csvString = '\uFEFF' + contenidoCSV.join('\n'); // \uFEFF es el BOM

    // Crear un blob con la codificación UTF-8
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });

    // Crear enlace de descarga
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${nombreArchivo}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Añadir a archivos recientes
    agregarArchivoReciente(`${nombreArchivo}.csv`);
}

// Función para exportar a TXT
function exportarTXT(nombreArchivo) {
    let contenidoTXT = [];

    // Garantizar que la hoja actual está guardada en la estructura de datos
    sheets[currentSheetIndex].data = state;

    // Obtener los datos de la hoja actual
    const hojaActual = sheets[currentSheetIndex].data;

    // Recorrer las filas y columnas de la hoja actual
    for (let y = 0; y < rows; y++) {
        let fila = [];
        for (let x = 0; x < cols; x++) {
            // Agregar el valor computado o un texto vacío
            fila.push(hojaActual[x][y].computedValue || "");
        }
        contenidoTXT.push(fila.join('\t')); // Usar tabulación como separador
    }

    // Convertir a string y crear blob
    const txtString = contenidoTXT.join('\n');
    const blob = new Blob([txtString], { type: 'text/plain;charset=utf-8;' });

    // Crear enlace de descarga
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${nombreArchivo}.txt`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Añadir a archivos recientes
    agregarArchivoReciente(`${nombreArchivo}.txt`);
}

// Función para exportar a PDF
function exportarPDF(nombreArchivo) {
    // Verificar si el campo de nombre del archivo está vacío
    if (!nombreArchivo) {
        alert("Por favor, introduce un nombre para el archivo");
        return; // Detener la ejecución si no hay nombre
    }

    // Garantizar que la hoja actual está guardada en la estructura de datos
    sheets[currentSheetIndex].data = state;

    // Obtener los datos de la hoja actual
    const hojaActual = sheets[currentSheetIndex].data;

    // Determinar el rango de celdas con datos
    let maxRow = 0;
    let maxCol = 0;

    for (let y = 0; y < rows; y++) {
        for (let x = 0; x < cols; x++) {
            const celda = hojaActual[x]?.[y];
            if (celda && (celda.value || celda.computedValue)) {
                if (y > maxRow) maxRow = y;
                if (x > maxCol) maxCol = x;
            }
        }
    }

    // Crear un elemento HTML temporal para generar el PDF
    const elemento = document.createElement('div');
    elemento.style.width = '100%';
    elemento.style.padding = '20px';

    // Crear tabla para el PDF
    const tabla = document.createElement('table');
    tabla.style.width = '100%';
    tabla.style.borderCollapse = 'collapse';

    // Generar encabezados de columnas
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    // Celda vacía para la esquina superior izquierda
    headerRow.appendChild(document.createElement('th'));

    // Agregar letras de columnas
    for (let x = 0; x <= maxCol; x++) {
        const th = document.createElement('th');
        th.textContent = letras[x];
        th.style.border = '1px solid #ccc';
        th.style.padding = '8px';
        th.style.backgroundColor = '#f2f2f2';
        headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    tabla.appendChild(thead);

    // Generar cuerpo de la tabla
    const tbody = document.createElement('tbody');
    for (let y = 0; y <= maxRow; y++) {
        const tr = document.createElement('tr');

        // Agregar número de fila
        const thRow = document.createElement('th');
        thRow.textContent = y + 1;
        thRow.style.border = '1px solid #ccc';
        thRow.style.padding = '8px';
        thRow.style.backgroundColor = '#f2f2f2';
        tr.appendChild(thRow);

        // Agregar celdas
        for (let x = 0; x <= maxCol; x++) {
            const td = document.createElement('td');
            const celda = hojaActual[x]?.[y];
            td.textContent = celda?.computedValue || "";
            td.style.border = '1px solid #ccc';
            td.style.padding = '8px';
            td.style.wordWrap = 'break-word'; // Evitar que el texto se corte
            td.style.maxWidth = '200px'; // Ajustar el ancho máximo de las celdas
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
    tabla.appendChild(tbody);
    elemento.appendChild(tabla);

    // Agregar título
    const titulo = document.createElement('h2');
    titulo.textContent = `${nombreArchivo}`;
    titulo.style.textAlign = 'center';
    elemento.insertBefore(titulo, elemento.firstChild);

    // Insertar el elemento en el documento
    document.body.appendChild(elemento);

    // Configuración para html2pdf
    const opciones = {
        margin: 10,
        filename: `${nombreArchivo}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
    };

    // Generar el PDF
    html2pdf().from(elemento).set(opciones).save().then(() => {
        // Eliminar el elemento temporal
        document.body.removeChild(elemento);

        // Añadir a archivos recientes
        agregarArchivoReciente(`${nombreArchivo}.pdf`);
    });
}

// Configurar el modal de exportación
document.addEventListener('DOMContentLoaded', function() {
  // Evento para el botón de abrir modal de exportación
  const openExportModalBtn = document.getElementById('openExportModal');
  if (openExportModalBtn) {
    openExportModalBtn.addEventListener('click', function() {
      // Asegurarse de que el valor del campo de texto está vacío al abrir
      document.getElementById('nombre-exportacion').value = '';
      document.getElementById('exportModal').style.display = 'flex';
    });
  }
  
  // Eventos para los botones de exportación
  document.getElementById('exportCSV')?.addEventListener('click', function() {
    exportarCSV(document.getElementById('nombre-exportacion').value.trim());
    document.getElementById('exportModal').style.display = 'none';
  });
  
  document.getElementById('exportTXT')?.addEventListener('click', function() {
    exportarTXT(document.getElementById('nombre-exportacion').value.trim());
    document.getElementById('exportModal').style.display = 'none';
  });
  
  document.getElementById('exportPDF')?.addEventListener('click', function() {
    exportarPDF(document.getElementById('nombre-exportacion').value.trim());
    document.getElementById('exportModal').style.display = 'none';
  });
  
  document.getElementById('cancelExport')?.addEventListener('click', function() {
    document.getElementById('exportModal').style.display = 'none';
  });
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

        // Actualizar el título del archivo en la interfaz
        const tituloElemento = document.querySelector('.ribbon-title h1');
        if (tituloElemento) {
            tituloElemento.textContent = file.name.replace(/\.[^/.]+$/, ""); // Quitar la extensión del archivo
        }

        // Crear nuevas hojas basadas en el archivo Excel
        sheets = [];
        
        // Procesar cada hoja del archivo Excel
        workbook.SheetNames.forEach((sheetName, index) => {
            let excelSheet = workbook.Sheets[sheetName];
            let jsonData = XLSX.utils.sheet_to_json(excelSheet, { header: 1 });
            
            // Crear estructura de datos para la hoja
            let sheetData = range(cols).map(() =>
                range(rows).map(() => ({ computedValue: "", value: "" }))
            );
            
            // Llenar los datos de la hoja
            jsonData.forEach((row, rowIndex) => {
                if (rowIndex < rows) { // Asegurar que no exceda el máximo de filas
                    row.forEach((cellValue, colIndex) => {
                        if (colIndex < cols) { // Asegurar que no exceda el máximo de columnas
                            // Verificar si es una fórmula en Excel
                            let cellRef = XLSX.utils.encode_cell({r: rowIndex, c: colIndex});
                            let cell = excelSheet[cellRef];
                            
                            if (cell && cell.f) { // Si tiene fórmula
                                let formula = '=' + cell.f;
                                sheetData[colIndex][rowIndex] = {
                                    value: formula,
                                    computedValue: calculateComputedValue(formula)
                                };
                            } else if (cell && cell.v !== undefined) { // Si tiene valor calculado
                                sheetData[colIndex][rowIndex] = {
                                    value: String(cell.v),
                                    computedValue: cell.v
                                };
                            } else {
                                sheetData[colIndex][rowIndex] = {
                                    value: cellValue !== undefined ? String(cellValue) : "",
                                    computedValue: cellValue
                                };
                            }
                        }
                    });
                }
            });
            
            // Añadir la hoja procesada
            sheets.push({
                name: sheetName,
                data: sheetData
            });
        });
        
        // Actualizar la UI para reflejar las nuevas hojas
        const sheetsContainer = document.querySelector(".sheets-container");
        sheetsContainer.innerHTML = '';
        
        sheets.forEach((sheet, index) => {
            let sheetElement = document.createElement("div");
            sheetElement.classList.add("sheet");
            if (index === 0) sheetElement.classList.add("active");
            sheetElement.textContent = sheet.name;
            sheetsContainer.appendChild(sheetElement);
        });
        
        // Cambiar a la primera hoja
        currentSheetIndex = 0;
        state = sheets[0].data;
        renderSpreadsheet();
        
        // Actualizar la visualización de la celda activa
        updateActiveCellDisplay(0, 0);
        
        // Cerrar la barra lateral de archivo
        document.getElementById('archivo-sidebar').classList.remove('active');
        
        // Guardar el nombre del archivo abierto en localStorage
        agregarArchivoReciente(file.name);
    };

    reader.readAsArrayBuffer(file); // Leer el archivo como ArrayBuffer
}

//PARA ABRIR
if (cell && cell.f) { // Si la celda tiene fórmula
  // Si existe un valor calculado (cell.v) se utiliza, sino se calcula manualmente
  const computedVal = (cell.v !== undefined) ? cell.v : calculateComputedValue('=' + cell.f);
  sheetData[colIndex][rowIndex] = {
      value: '=' + cell.f,
      computedValue: computedVal
  };
} else if (cell && cell.v !== undefined) { // Si la celda tiene valor y no es fórmula
  sheetData[colIndex][rowIndex] = {
      value: String(cell.v),
      computedValue: cell.v
  };
} else {
  sheetData[colIndex][rowIndex] = {
      value: cellValue !== undefined ? String(cellValue) : "",
      computedValue: cellValue
  };
}

// Actualizar el nombre del archivo seleccionado
document.addEventListener('DOMContentLoaded', function() {
  const fileInput = document.getElementById('fileInput');
  const filePathDisplay = document.getElementById('file-path-display');
  
  if (fileInput && filePathDisplay) {
    fileInput.addEventListener('change', function() {
      if (this.files && this.files.length > 0) {
        filePathDisplay.textContent = this.files[0].name;
      } else {
        filePathDisplay.textContent = 'Ningún archivo seleccionado';
      }
    });
  }
});
