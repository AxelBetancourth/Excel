// Script para manejar la funcionalidad del menú de Datos

document.addEventListener('DOMContentLoaded', function() {
    // ========================= MENÚS Y BOTONES ===========================
    const botonesNivel1 = document.querySelectorAll('#submenu-datos .menu-btn-nivel1');
    
    botonesNivel1.forEach(boton => {
        boton.addEventListener('click', function(e) {
            e.stopPropagation();
            
            const submenu = this.nextElementSibling;
            
            // Verificar si el submenú existe
            if (submenu && submenu.tagName === 'UL') {
                // Si ya está visible, ocultarlo
                if (submenu.classList.contains('show')) {
                    submenu.classList.remove('show');
                } else {
                    // Primero, ocultar todos los submenús abiertos
                    document.querySelectorAll('#submenu-datos ul.show').forEach(menu => {
                        menu.classList.remove('show');
                    });
                    
                    // Mostrar este submenú
                    submenu.classList.add('show');
                }
            }
        });
    });
    
    // Agregar event listeners a los botones de nivel 2
    const botonesNivel2 = document.querySelectorAll('#submenu-datos .menu-btn-nivel2');
    
    botonesNivel2.forEach(boton => {
        boton.addEventListener('click', function(e) {
            e.stopPropagation();
            
            const submenu = this.nextElementSibling;
            
            if (submenu && submenu.tagName === 'UL') {
                if (submenu.classList.contains('show')) {
                    submenu.classList.remove('show');
                } else {
                    const submenusNivel3 = document.querySelectorAll('#submenu-datos .menu-btn-nivel2 + ul');
                    submenusNivel3.forEach(menu => {
                        menu.classList.remove('show');
                    });
                    
                    submenu.classList.add('show');
                }
            }
        });
    });
    
    const btnPDF = document.getElementById('btn-pdf');
    const btnExcel = document.getElementById('btn-excel');
    
    if (btnPDF) {
        btnPDF.addEventListener('click', function(e) {
            e.stopPropagation();
            document.getElementById('archivoPDF').click();
        });
    }
    
    if (btnExcel) {
        btnExcel.addEventListener('click', function(e) {
            e.stopPropagation();
            document.getElementById('archivoExcel').click();
        });
    }
    
    // Cerrar submenús al hacer clic fuera de ellos
    document.addEventListener('click', function(e) {
        if (!e.target.closest('#submenu-datos .container')) {
            const submenus = document.querySelectorAll('#submenu-datos ul.show');
            submenus.forEach(menu => {
                menu.classList.remove('show');
            });
        }
    });

    // ========================= PROCESAR PDF ===========================
    function processPDF(file) {
        // Verificar si la biblioteca PDF.js está disponible
        if (typeof pdfjsLib === 'undefined') {
            console.error('PDF.js no está cargado. Por favor, incluya la biblioteca PDF.js.');
            alert('La biblioteca PDF.js no está disponible. No se puede procesar el archivo PDF.');
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            const typedarray = new Uint8Array(e.target.result);
            
            // Mostrar indicador de carga
            const loadingMessage = document.createElement('div');
            loadingMessage.textContent = 'Procesando PDF, por favor espere...';
            loadingMessage.style.position = 'fixed';
            loadingMessage.style.top = '50%';
            loadingMessage.style.left = '50%';
            loadingMessage.style.transform = 'translate(-50%, -50%)';
            loadingMessage.style.backgroundColor = 'rgba(0,0,0,0.7)';
            loadingMessage.style.color = 'white';
            loadingMessage.style.padding = '20px';
            loadingMessage.style.borderRadius = '5px';
            loadingMessage.style.zIndex = '9999';
            document.body.appendChild(loadingMessage);
            
            pdfjsLib.getDocument(typedarray).promise.then(async function (pdf) {
                console.log("PDF cargado");
                let allTextContent = [];

                try {
                    // Extraer texto de todas las páginas
                    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                        const page = await pdf.getPage(pageNum);
                        const textContent = await page.getTextContent();

                        const textItems = textContent.items;
                        
                        // Agrupar por líneas (elementos con posición Y similar)
                        const lines = {};
                        textItems.forEach(item => {
                            // Redondear la posición Y para agrupar líneas cercanas
                            const yPos = Math.round(item.transform[5]);
                            if (!lines[yPos]) lines[yPos] = [];
                            lines[yPos].push({
                                text: item.str,
                                x: item.transform[4] // posición X para ordenar dentro de una línea
                            });
                        });
                        
                        // Ordenar por posición Y (de arriba a abajo en la página)
                        const sortedYPositions = Object.keys(lines).map(Number).sort((a, b) => b - a);
                        
                        // Para cada línea, ordenar elementos por posición X y concatenar
                        sortedYPositions.forEach(yPos => {
                            lines[yPos].sort((a, b) => a.x - b.x);
                            const lineText = lines[yPos].map(item => item.text).join(' ').trim();
                            if (lineText) {
                                allTextContent.push(lineText);
                            }
                        });
                    }
                    
                    // Convertir a formato de tabla (filas y columnas)
                    // Cada línea es una fila, y dividimos por espacios o tabulaciones para columnas
                    const tableData = allTextContent.map(line => {
                        // Dividir la línea en columnas (por espacios múltiples o tabulaciones)
                        return line.split(/\s{2,}|\t/).filter(cell => cell.trim() !== '');
                    }).filter(row => row.length > 0); // Filtrar filas vacías
                    
                    // Asegurar el tamaño mínimo de la tabla y mostrarla
                    ensureMinGridSize(tableData);
                    
                    // Ajustar el ancho de las columnas
                    adjustColumnWidths();
                    
                    // Mostrar mensaje de éxito
                    alert(`PDF procesado correctamente. Se han extraído ${tableData.length} filas de datos.`);
                    
                } catch (error) {
                    console.error("Error al procesar el PDF:", error);
                    alert("Error al procesar el PDF: " + error.message);
                } finally {
                    // Eliminar indicador de carga
                    document.body.removeChild(loadingMessage);
                }
            }).catch(error => {
                document.body.removeChild(loadingMessage);
                console.error("Error al cargar el PDF:", error);
                alert("Error al cargar el PDF. Por favor, verifique que el archivo sea válido.");
            });
        };
        reader.readAsArrayBuffer(file);
    }

    // ========================= PROCESAR EXCEL ===========================
    function processExcel(file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: "array" });

                // Crear selector de hojas si hay más de una
                if (workbook.SheetNames.length > 1) {
                    let tableSelector = document.createElement('select');
                    tableSelector.id = "table-selector";
                    tableSelector.innerHTML = workbook.SheetNames
                        .map(sheet => `<option value="${sheet}">${sheet}</option>`)
                        .join("");
                    
                    // Insertar el selector antes de la tabla
                    const spreadsheetContainer = document.querySelector('.spreadsheet-container');
                    if (spreadsheetContainer) {
                        spreadsheetContainer.insertBefore(tableSelector, spreadsheetContainer.firstChild);
                    }
                    
                    tableSelector.onchange = function () {
                        loadSelectedTable(workbook, this.value);
                    };
                }

                // Cargar la primera hoja automáticamente
                loadSelectedTable(workbook, workbook.SheetNames[0]);
            } catch (error) {
                console.error("Error al procesar el Excel:", error);
                alert("Error al procesar el archivo Excel. Consulte la consola para más detalles.");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function loadSelectedTable(workbook, sheetName) {
        try {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });

            if (!jsonData || jsonData.length === 0) {
                console.warn("La hoja seleccionada está vacía");
                return;
            }

            ensureMinGridSize(jsonData);
            // Ajustar el ancho de las columnas después de mostrar los datos
            adjustColumnWidths();
        } catch (error) {
            console.error("Error al cargar la hoja de Excel:", error);
        }
    }

    // ====================== AJUSTAR ANCHO DE CELDAS =======================
    function adjustColumnWidths() {
        const table = document.getElementById("spreadsheet");
        if (!table) return;
        
        const rows = table.getElementsByTagName("tr");
        if (rows.length === 0) return;

        const columnWidths = [];

        // Determinar el ancho máximo de cada columna según el texto
        for (let row of rows) {
            const cells = row.getElementsByTagName("td");
            for (let i = 0; i < cells.length; i++) {
                const textLength = cells[i].textContent.trim().length;
                const cellWidth = Math.max(textLength * 10, 60); // Mínimo 60px, o más según contenido
                columnWidths[i] = Math.max(columnWidths[i] || 0, cellWidth);
            }
        }

        // Aplicar los anchos calculados a las columnas
        for (let row of rows) {
            const cells = row.getElementsByTagName("td");
            for (let i = 0; i < cells.length && i < columnWidths.length; i++) {
                cells[i].style.minWidth = `${columnWidths[i]}px`;
            }
        }
    }

    // ========================= MANEJAR ARCHIVOS ===========================
    function handleFileUpload(event) {
        const file = event.target.files[0];

        if (!file) return;

        // Procesar archivo PDF
        if (file.type === "application/pdf") {
            processPDF(file);
        } 
        // Procesar archivo Excel
        else if (file.type.includes("spreadsheet") || file.name.endsWith(".xls") || file.name.endsWith(".xlsx")) {
            processExcel(file);
        }  
        else {
            console.error("Tipo de archivo no soportado");
            alert("Tipo de archivo no soportado. Por favor, suba un archivo PDF o Excel.");
        }
    }

    // ========================= MOSTRAR DATOS EN LA TABLA ===========================
    function displayTableData(data) {
        const table = document.getElementById("spreadsheet");
        if (!table) {
            console.error("No se encontró la tabla en el HTML.");
            return;
        }

        const thead = table.querySelector("thead");
        const tbody = table.querySelector("tbody");
        thead.innerHTML = "";
        tbody.innerHTML = "";

        if (data.length === 0) return;

        let maxRows = data.length;
        let maxCols = Math.max(...data.map(row => row.length));

        // Crear encabezados de columna (A, B, C...)
        let headerRow = document.createElement("tr");

        // Crear celda vacía en la esquina superior izquierda
        let emptyHeaderCell = document.createElement("th");
        headerRow.appendChild(emptyHeaderCell);

        for (let i = 0; i < maxCols; i++) {
            const th = document.createElement("th");
            th.textContent = String.fromCharCode(65 + i); // A, B, C...
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);

        // Insertar filas con datos
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
            const tr = document.createElement("tr");

            // Agregar número de fila (1, 2, 3...)
            const rowLabelCell = document.createElement("th");
            rowLabelCell.textContent = rowIndex + 1; // 1, 2, 3...
            tr.appendChild(rowLabelCell);

            // Insertar datos para cada columna
            for (let colIndex = 0; colIndex < maxCols; colIndex++) {
                const td = document.createElement("td");
                td.contentEditable = true; // Hacer la celda editable
                
                // Asignar el valor de la celda, si existe
                if (data[rowIndex] && colIndex < data[rowIndex].length) {
                    td.textContent = data[rowIndex][colIndex] !== undefined && data[rowIndex][colIndex] !== null
                        ? data[rowIndex][colIndex].toString().trim()
                        : "";
                }
                
                // Agregar evento de clic para mostrar la celda seleccionada
                td.addEventListener("click", function() {
                    // Actualizar el indicador de celda activa
                    const activeCell = document.getElementById("active-cell");
                    if (activeCell) {
                        activeCell.textContent = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                    }
                    
                    // Actualizar el contenido del input de fórmula
                    const formulaInput = document.getElementById("formula-input");
                    if (formulaInput) {
                        formulaInput.value = this.textContent;
                    }
                });
                
                tr.appendChild(td);
            }

            tbody.appendChild(tr);
        }
    }

    // ========================= AJUSTAR FILAS Y COLUMNAS ===========================
    function ensureMinGridSize(data) {
        let maxRows = Math.max(data.length, 20); // Mínimo 20 filas
        let maxColumns = Math.max(...data.map(row => row.length || 0), 10); // Mínimo 10 columnas (A-J)

        let structuredData = Array.from({ length: maxRows }, (_, rowIndex) => {
            if (rowIndex < data.length) {
                // Si la fila existe en los datos originales, la usamos
                const row = data[rowIndex] || [];
                // Aseguramos que tenga el número correcto de columnas
                return Array.from({ length: maxColumns }, (_, colIndex) => 
                    colIndex < row.length ? row[colIndex] : "");
            } else {
                // Si es una fila adicional, creamos una fila vacía
                return Array(maxColumns).fill("");
            }
        });

        displayTableData(structuredData);
    }

    // ===================== EVENTOS PARA INPUTS DE ARCHIVO ======================
    // Agregar listeners para los inputs de archivo
    const archivoPDF = document.getElementById("archivoPDF");
    const archivoExcel = document.getElementById("archivoExcel");
    
    if (archivoPDF) {
        archivoPDF.addEventListener("change", handleFileUpload);
    }
    
    if (archivoExcel) {
        archivoExcel.addEventListener("change", handleFileUpload);
    }

    // Botones de ordenación
    const btnOrdenarAZ = document.getElementById("OrdenarAZ");
    const btnOrdenarZA = document.getElementById("OrdenarZA");
});

