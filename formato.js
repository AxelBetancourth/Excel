/* ================================
// Funcionalidades de formato
// ================================*/

document.addEventListener("DOMContentLoaded", () => {
    const $ = el => document.querySelector(el);
    const $$ = el => document.querySelectorAll(el);

    // Aplicar formato directamente al DOM
    const applyFormat = (cell, format) => {
        if (!cell) {
            console.error("Celda no encontrada para aplicar formato:", format);
            return;
        }
        console.log("Aplicando formato a", cell.dataset.x, cell.dataset.y, format);
        if (format.border) {
            // Limpiar bordes previos
            cell.style.border = "none";
            switch (format.border) {
                case "all": 
                    cell.style.border = "1px solid #000";
                    console.log("Borde aplicado: all");
                    break;
                case "top": 
                    cell.style.borderTop = "1px solid #000";
                    console.log("Borde aplicado: top");
                    break;
                case "bottom": 
                    cell.style.borderBottom = "1px solid #000";
                    console.log("Borde aplicado: bottom");
                    break;
                case "left": 
                    cell.style.borderLeft = "1px solid #000";
                    console.log("Borde aplicado: left");
                    break;
                case "right": 
                    cell.style.borderRight = "1px solid #000";
                    console.log("Borde aplicado: right");
                    break;
                case "none": 
                    console.log("Borde removido");
                    break;
            }
        }
        if (format.fontFamily) cell.style.fontFamily = format.fontFamily;
        if (format.fontSize) cell.style.fontSize = format.fontSize;
        if (format.fontWeight) cell.style.fontWeight = format.fontWeight;
        if (format.fontStyle) cell.style.fontStyle = format.fontStyle;
        if (format.textDecoration) cell.style.textDecoration = format.textDecoration;
        if (format.backgroundColor) {
            cell.style.backgroundColor = format.backgroundColor;
            console.log("Color de fondo aplicado:", cell.style.backgroundColor);
        }
        if (format.color) {
            cell.style.color = format.color;
            console.log("Color de letra aplicado:", cell.style.color);
        }
        if (format.textAlign) cell.style.textAlign = format.textAlign;
        if (format.paddingLeft) cell.style.paddingLeft = format.paddingLeft;
        if (format.writingMode) cell.style.writingMode = format.writingMode;
    };

    // Actualizar celda sin depender de state
    const updateCellValue = (cell, value = null, format = null) => {
        const x = parseInt(cell.dataset.x);
        const y = parseInt(cell.dataset.y);
        if (typeof updateCell === "function" && value !== null) {
            updateCell(x, y, value); // Solo para valores, no formatos
        }
        if (format !== null) {
            applyFormat(cell, format); // Aplicar formato directamente
        }
    };

    // Enfocar la celda activa
    const focusActiveCell = () => {
        const activeCell = $(".celdailumida");
        if (activeCell) {
            const input = activeCell.querySelector("input");
            if (input) input.focus();
        } else {
            console.log("No hay celda activa con .celdailumida");
        }
    };

    const waitForUpdateCell = (callback) => {
        if (typeof updateCell === "function") {
            callback();
        } else {
            console.log("Esperando a que updateCell esté disponible...");
            setTimeout(() => waitForUpdateCell(callback), 100);
        }
    };

    waitForUpdateCell(() => {
        console.log("formato.js cargado y updateCell disponible");

        // Obtener referencias a los botones usando los atributos title
        const btnCortar = $(".ribbon-buttons button[title='Cortar']");
        const btnCopiar = $(".ribbon-buttons button[title='Copiar']");
        const btnPegar = $(".ribbon-buttons button[title='Pegar']");
        const btnCopiarFormato = $(".ribbon-buttons button[title='Copiar formato']");
        
        const btnNegrita = $(".ribbon-buttons button[title='Negrita']");
        const btnCursiva = $(".ribbon-buttons button[title='Cursiva']");
        const btnSubrayado = $(".ribbon-buttons button[title='Subrayado']");
        const btnBordes = $(".ribbon-buttons button[title='Bordes']");
        const btnColorRelleno = $(".ribbon-buttons button[title='Color de relleno']");
        const btnColorFuente = $(".ribbon-buttons button[title='Color de fuente']");

        // Selectores
        const selectorFuente = $("#submenu-inicio .fila.selectores select:first-child");
        const selectorTamaño = $("#submenu-inicio .fila.selectores select:nth-child(2)");
        
        // Botones de alineación
        const btnAlinearIzq = $(".ribbon-buttons button[title='Alinear izquierda']");
        const btnCentrar = $(".ribbon-buttons button[title='Centrar']");
        const btnAlinearDer = $(".ribbon-buttons button[title='Alinear derecha']");
        const btnJustificar = $(".ribbon-buttons button[title='Justificar']");
        const btnOrientacion = $(".ribbon-buttons button[title='Orientación']");
        const btnSangria = $(".ribbon-buttons button[title='Sangría']");

        // Número
        const selectorNumero = $("#submenu-inicio .ribbon-group:nth-child(4) .fila.selectores select");
        
        let clipboardContent = null; // Para texto
        let clipboardFormat = null;  // Para formato

        // Debugging
        console.log("Botones inicializados:");
        console.log("- Negrita:", btnNegrita ? "✓" : "✗");
        console.log("- Cursiva:", btnCursiva ? "✓" : "✗");
        console.log("- Subrayado:", btnSubrayado ? "✓" : "✗");
        console.log("- Selector de fuente:", selectorFuente ? "✓" : "✗");
        console.log("- Selector de tamaño:", selectorTamaño ? "✓" : "✗");

        // ===== PORTAPAPELES =====
        
        // Cortar
        if (btnCortar) {
            btnCortar.addEventListener("click", () => {
                console.log("Botón Cortar clicado");
                const activeCell = $(".celdailumida");
                if (activeCell) {
                    const input = activeCell.querySelector("input");
                    if (input) {
                        clipboardContent = input.value;
                        updateCellValue(activeCell, "");
                        focusActiveCell();
                    }
                }
            });
        }

        // Copiar
        if (btnCopiar) {
            btnCopiar.addEventListener("click", () => {
                console.log("Botón Copiar clicado");
                const activeCell = $(".celdailumida");
                if (activeCell) {
                    const input = activeCell.querySelector("input");
                    if (input) {
                        clipboardContent = input.value;
                        focusActiveCell();
                    }
                }
            });
        }

        // Pegar
        if (btnPegar) {
            btnPegar.addEventListener("click", () => {
                console.log("Botón Pegar clicado", clipboardContent);
                const selectedCells = $$(".celdailumida");
                if (selectedCells.length > 0 && clipboardContent !== null) {
                    selectedCells.forEach(cell => {
                        updateCellValue(cell, clipboardContent);
                        if (clipboardFormat !== null) {
                            updateCellValue(cell, null, clipboardFormat);
                        }
                    });
                    focusActiveCell();
                }
            });
        }

        // Copiar formato
        if (btnCopiarFormato) {
            btnCopiarFormato.addEventListener("click", () => {
                console.log("Botón Copiar formato clicado");
                const activeCell = $(".celdailumida");
                if (activeCell) {
                    const style = window.getComputedStyle(activeCell);
                    clipboardFormat = {
                        fontFamily: style.fontFamily,
                        fontSize: style.fontSize,
                        fontWeight: style.fontWeight,
                        fontStyle: style.fontStyle,
                        textDecoration: style.textDecoration,
                        backgroundColor: style.backgroundColor,
                        color: style.color,
                        border: style.borderWidth !== "0px" ? "all" : "none",
                        textAlign: style.textAlign,
                        paddingLeft: style.paddingLeft,
                        writingMode: style.writingMode
                    };
                    focusActiveCell();
                }
            });
        }

        // ===== FUENTE =====

        // Fuente - familia
        if (selectorFuente) {
            selectorFuente.addEventListener("change", (e) => {
                console.log("Cambio de familia de fuente:", e.target.value);
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { fontFamily: e.target.value });
                });
                focusActiveCell();
            });
        }

        // Fuente - tamaño
        if (selectorTamaño) {
            selectorTamaño.addEventListener("change", (e) => {
                console.log("Cambio de tamaño de fuente:", e.target.value);
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { fontSize: `${e.target.value}px` });
                });
                focusActiveCell();
            });
        }

        // Negrita
        if (btnNegrita) {
            btnNegrita.addEventListener("click", () => {
                console.log("Botón Negrita clicado");
                $$(".celdailumida").forEach(cell => {
                    const isBold = cell.style.fontWeight === "bold";
                    updateCellValue(cell, null, { fontWeight: isBold ? "normal" : "bold" });
                });
                focusActiveCell();
            });
        }

        // Cursiva
        if (btnCursiva) {
            btnCursiva.addEventListener("click", () => {
                console.log("Botón Cursiva clicado");
                $$(".celdailumida").forEach(cell => {
                    const isItalic = cell.style.fontStyle === "italic";
                    updateCellValue(cell, null, { fontStyle: isItalic ? "normal" : "italic" });
                });
                focusActiveCell();
            });
        }

        // Subrayado
        if (btnSubrayado) {
            btnSubrayado.addEventListener("click", () => {
                console.log("Botón Subrayado clicado");
                $$(".celdailumida").forEach(cell => {
                    const isUnderlined = cell.style.textDecoration === "underline";
                    updateCellValue(cell, null, { textDecoration: isUnderlined ? "none" : "underline" });
                });
                focusActiveCell();
            });
        }

        // Color de relleno
        if (btnColorRelleno) {
            btnColorRelleno.addEventListener("click", () => {
                console.log("Botón de color de relleno clicado");
                const fillColorInput = $("#fill-color");
                if (fillColorInput) {
                    fillColorInput.click();
                } else {
                    console.error("No se encontró #fill-color en el DOM");
                }
            });
        }

        // Color de fuente
        if (btnColorFuente) {
            btnColorFuente.addEventListener("click", () => {
                console.log("Botón de color de fuente clicado");
                const fontColorInput = $("#font-color");
                if (fontColorInput) {
                    fontColorInput.click();
                } else {
                    console.error("No se encontró #font-color en el DOM");
                }
            });
        }

        // Evento de color de relleno
        $("#fill-color")?.addEventListener("change", (e) => {
            console.log("Color de celda seleccionado:", e.target.value);
            const selectedCells = $$(".celdailumida");
            if (selectedCells.length === 0) {
                console.warn("No hay celdas con .celdailumida seleccionadas");
                return;
            }
            selectedCells.forEach(cell => {
                updateCellValue(cell, null, { backgroundColor: e.target.value });
                // Forzar reaplicación
                setTimeout(() => {
                    const updatedCell = document.querySelector(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                    if (updatedCell) {
                        updatedCell.style.backgroundColor = e.target.value;
                        console.log("Color de fondo reaplicado a", cell.dataset.x, cell.dataset.y, e.target.value);
                    }
                }, 50);
            });
            focusActiveCell();
        });

        // Evento de color de fuente
        $("#font-color")?.addEventListener("change", (e) => {
            console.log("Color de letra seleccionado:", e.target.value);
            const selectedCells = $$(".celdailumida");
            if (selectedCells.length === 0) {
                console.warn("No hay celdas con .celdailumida seleccionadas");
                return;
            }
            selectedCells.forEach(cell => {
                updateCellValue(cell, null, { color: e.target.value });
                // Forzar reaplicación
                setTimeout(() => {
                    const updatedCell = document.querySelector(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                    if (updatedCell) {
                        updatedCell.style.color = e.target.value;
                        console.log("Color de letra reaplicado a", cell.dataset.x, cell.dataset.y, e.target.value);
                    }
                }, 50);
            });
            focusActiveCell();
        });

        // ===== ALINEACIÓN =====
        
        // Alinear izquierda
        if (btnAlinearIzq) {
            btnAlinearIzq.addEventListener("click", () => {
                console.log("Botón Alinear izquierda clicado");
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { textAlign: "left" });
                });
                focusActiveCell();
            });
        }

        // Centrar
        if (btnCentrar) {
            btnCentrar.addEventListener("click", () => {
                console.log("Botón Centrar clicado");
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { textAlign: "center" });
                });
                focusActiveCell();
            });
        }

        // Alinear derecha
        if (btnAlinearDer) {
            btnAlinearDer.addEventListener("click", () => {
                console.log("Botón Alinear derecha clicado");
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { textAlign: "right" });
                });
                focusActiveCell();
            });
        }

        // Justificar
        if (btnJustificar) {
            btnJustificar.addEventListener("click", () => {
                console.log("Botón Justificar clicado");
                $$(".celdailumida").forEach(cell => {
                    updateCellValue(cell, null, { textAlign: "justify" });
                });
                focusActiveCell();
            });
        }

        // Orientación
        if (btnOrientacion) {
            btnOrientacion.addEventListener("click", () => {
                console.log("Botón Orientación clicado");
                $$(".celdailumida").forEach(cell => {
                    const isVertical = cell.style.writingMode === "vertical-rl";
                    updateCellValue(cell, null, { writingMode: isVertical ? "horizontal-tb" : "vertical-rl" });
                });
                focusActiveCell();
            });
        }

        // Sangría
        if (btnSangria) {
            btnSangria.addEventListener("click", () => {
                console.log("Botón Sangría clicado");
                $$(".celdailumida").forEach(cell => {
                    const hasIndent = cell.style.paddingLeft === "20px";
                    updateCellValue(cell, null, { paddingLeft: hasIndent ? "0px" : "20px" });
                });
                focusActiveCell();
            });
        }

        // ===== NÚMERO =====

        // Función para obtener el valor numérico
        const getBaseNumber = (value) => {
            return parseFloat(value.replace(/[^0-9.-]/g, "")) || value;
        };

        // Selector de formato de número
        if (selectorNumero) {
            selectorNumero.addEventListener("change", (e) => {
                console.log("Formato numérico seleccionado:", e.target.value);
                const selectedCells = $$(".celdailumida");
                if (selectedCells.length === 0) {
                    console.warn("No hay celdas con .celdailumida seleccionadas");
                    return;
                }
                selectedCells.forEach(cell => {
                    const input = cell.querySelector("input");
                    if (!input) {
                        console.warn("No se encontró <input> en celda:", cell.dataset.x, cell.dataset.y);
                        return;
                    }
                    const baseValue = getBaseNumber(input.value);
                    let newValue;
                    if (e.target.value === "Número") {
                        newValue = baseValue.toString();
                        updateCellValue(cell, newValue, { textAlign: "right" });
                    } else if (e.target.value === "Moneda") {
                        newValue = `$${baseValue}`;
                        updateCellValue(cell, newValue);
                    } else if (e.target.value === "Porcentaje") {
                        newValue = `${baseValue}%`;
                        updateCellValue(cell, newValue);
                    } else if (e.target.value === "General") {
                        newValue = baseValue.toString();
                        updateCellValue(cell, newValue);
                    }
                    console.log(`Aplicando ${e.target.value} a (${cell.dataset.x},${cell.dataset.y}): "${newValue}"`);

                    setTimeout(() => {
                        const updatedCell = $(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                        if (updatedCell && updatedCell.querySelector("input")) {
                            updatedCell.querySelector("input").value = newValue;
                            if (e.target.value === "Número") {
                                updatedCell.style.textAlign = "right";
                            } else {
                                updatedCell.style.textAlign = "";
                            }
                            console.log(`Reaplicado ${e.target.value} a (${cell.dataset.x},${cell.dataset.y}): "${newValue}"`);
                        }
                    }, 50);
                });
                focusActiveCell();
            });
        }

        // Inicializar eventos en la tabla
        const table = $("table");
        if (table) {
            table.addEventListener("click", (e) => {
                const td = e.target.closest("td");
                if (td && td.querySelector("input")) {
                    // Aquí podemos hacer acciones cuando se haga clic en una celda
                    console.log("Celda seleccionada:", td.dataset.x, td.dataset.y);
                }
            });
        }
    });
});