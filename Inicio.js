document.addEventListener('DOMContentLoaded', () => {
    // ===== SELECTORES Y CONSTANTES GLOBALES =====
    const $ = el => document.querySelector(el);
    const $$ = el => document.querySelectorAll(el);
    
    // Constantes para bordes
    const borderSelect = document.getElementById('border-select');
    const borderColor = document.getElementById('border-color');
    const defaultBorderColor = '#e2e2e2';
    const defaultBorderWidth = '1px';
    
    // Botones de formato numérico
    const selectorNumero = $("#submenu-inicio .ribbon-group:nth-child(4) .fila.selectores select");
    const btnFormatoMoneda = $("button[title='Formato de número']");
    const btnPorcentaje = $("button[title='Porcentaje']");
    const btnSeparadorMiles = $("button[title='Separador de miles']");
    const btnAumentarDecimales = $("button[title='Aumentar decimales']");
    const btnDisminuirDecimales = $("button[title='Disminuir decimales']");
    
    // Botones de portapapeles
    const btnCortar = $(".ribbon-buttons button[title='Cortar']");
    const btnCopiar = $(".ribbon-buttons button[title='Copiar']");
    const btnPegar = $(".ribbon-buttons button[title='Pegar']");
    const btnCopiarFormato = $(".ribbon-buttons button[title='Copiar formato']");
    
    // Botones de formato de texto
    const btnNegrita = $(".ribbon-buttons button[title='Negrita']");
    const btnCursiva = $(".ribbon-buttons button[title='Cursiva']");
    const btnSubrayado = $(".ribbon-buttons button[title='Subrayado']");
    const btnColorRelleno = $(".ribbon-buttons button[title='Color de relleno']");
    const btnColorFuente = $(".ribbon-buttons button[title='Color de fuente']");

    // Selectores de fuente
    const selectorFuente = $("#submenu-inicio .fila.selectores select:first-child");
    const selectorTamaño = $("#submenu-inicio .fila.selectores select:nth-child(2)");
    
    // Botones de alineación
    const btnAlinearIzq = $(".ribbon-buttons button[title='Alinear izquierda']");
    const btnCentrar = $(".ribbon-buttons button[title='Centrar']");
    const btnAlinearDer = $(".ribbon-buttons button[title='Alinear derecha']");
    const btnJustificar = $(".ribbon-buttons button[title='Justificar']");
    const btnCombinarCeldas = $("button[title='Combinar celdas']");
    const btnOrientacion = $(".ribbon-buttons button[title='Orientación']");
    const btnSangria = $(".ribbon-buttons button[title='Sangría']");
    
    // Botones de búsqueda y reemplazo
    const btnBuscar = $(".ribbon-buttons button[title='Buscar']");
    const btnReemplazar = $(".ribbon-buttons button[title='Reemplazar']");
    
    // Variables para portapapeles
    let clipboardContent = null;
    let clipboardFormat = null;
    
    // Variables para búsqueda
    let searchResults = [];
    let currentSearchIndex = -1;
    
    // ===== FUNCIONES UTILITARIAS =====
    
    // Obtener el valor numérico base a partir de un valor con formato
    function getBaseNumber(value) {
        if (!value && value !== 0) return 0;
        
        if (typeof value === 'string') {
            // Detectamos si es un valor con formato de moneda (comienza con L)
            const isMoneda = value.trim().startsWith('L');
            
            // Limpiamos el valor de cualquier formato
            let cleanValue = String(value)
                .replace(/[L$]/g, '')
                .replace(/%/g, '')
                .replace(/[()]/g, '')
                .replace(/,/g, '');
            
            const num = parseFloat(cleanValue);
            return isNaN(num) ? 0 : num;
        }
        
        return typeof value === 'number' ? value : 0;
    }
    
    // Formatear número con separador de miles y decimales específicos
    function formatNumber(number, decimals) {
        let num = parseFloat(number);
        if (isNaN(num)) num = 0;
        
        let formatted = num.toFixed(decimals);
        let parts = formatted.split('.');
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        
        return parts.join('.');
    }
    
    // Función simplificada para obtener valor numérico
    const getSimpleBaseNumber = (value) => {
        return parseFloat(value.replace(/[^0-9.-]/g, "")) || value;
    };
    
    // Guardar el estado global
    function saveState() {
        if (window.updateAllDependentCells) {
            window.updateAllDependentCells();
        }
    }
    
    // Aplicar formato directamente al DOM
    const applyFormat = (cell, format) => {
        if (!cell) {
            console.error("Celda no encontrada para aplicar formato");
            return;
        }
        
        if (format.border) {
            cell.style.border = "none";
            switch (format.border) {
                case "all": 
                    cell.style.border = "1px solid #000";
                    break;
                case "top": 
                    cell.style.borderTop = "1px solid #000";
                    break;
                case "bottom": 
                    cell.style.borderBottom = "1px solid #000";
                    break;
                case "left": 
                    cell.style.borderLeft = "1px solid #000";
                    break;
                case "right": 
                    cell.style.borderRight = "1px solid #000";
                    break;
            }
        }
        
        if (format.fontFamily) cell.style.fontFamily = format.fontFamily;
        if (format.fontSize) cell.style.fontSize = format.fontSize;
        if (format.fontWeight) cell.style.fontWeight = format.fontWeight;
        if (format.fontStyle) cell.style.fontStyle = format.fontStyle;
        if (format.textDecoration) cell.style.textDecoration = format.textDecoration;
        if (format.backgroundColor) cell.style.backgroundColor = format.backgroundColor;
        if (format.color) cell.style.color = format.color;
        if (format.textAlign) cell.style.textAlign = format.textAlign;
        if (format.paddingLeft) cell.style.paddingLeft = format.paddingLeft;
        if (format.writingMode) cell.style.writingMode = format.writingMode;
    };
    
    // Actualizar celda sin depender de state
    const updateCellValue = (cell, value = null, format = null) => {
        const x = parseInt(cell.dataset.x);
        const y = parseInt(cell.dataset.y);
        if (typeof updateCell === "function" && value !== null) {
            updateCell(x, y, value);
        }
        if (format !== null) {
            applyFormat(cell, format);
        }
    };
    
    // Enfocar la celda activa
    const focusActiveCell = () => {
        const activeCell = $(".celdailumida");
        if (activeCell) {
            const input = activeCell.querySelector("input");
            if (input) input.focus();
        }
    };
    
    // ===== FUNCIONES DE FORMATO =====
    
    // Aplicar bordes a la celda
    function applyBorders(cell, borderType, color) {
        cell.style.outline = '';
        cell.style.borderTop = '';
        cell.style.borderRight = '';
        cell.style.borderBottom = '';
        cell.style.borderLeft = '';
        cell.classList.remove('cell-with-border');
        
        switch(borderType) {
            case 'all':
                cell.style.outline = `1px solid ${color}`;
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                break;
            case 'top':
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                cell.setAttribute('style', cell.getAttribute('style') + `;border-top: 2px solid ${color} !important;`);
                cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                break;
            case 'right':
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.setAttribute('style', cell.getAttribute('style') + `;border-right: 2px solid ${color} !important;`);
                cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                break;
            case 'bottom':
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.setAttribute('style', cell.getAttribute('style') + `;border-bottom: 2px solid ${color} !important;`);
                cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                break;
            case 'left':
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.setAttribute('style', cell.getAttribute('style') + `;border-left: 2px solid ${color} !important;`);
                break;
            case 'none':
                cell.style.position = '';
                cell.style.zIndex = '';
                cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                break;
        }
    }
    
    // Aplicar formato numérico a la celda
    function applyNumberFormat(cell, format, decimals = 2) {
        const input = cell.querySelector('input');
        const span = cell.querySelector('span');
        if (!input || !span) return;
        
        let rawValue = input.value !== "" ? input.value : span.textContent;
        rawValue = rawValue || '0';
        
        const baseValue = getBaseNumber(rawValue);
        let originalValue = baseValue;
        let formattedValue = '';
        
        const previousFormat = cell.dataset.format || 'General';
        
        switch (format) {
            case 'General':
                formattedValue = baseValue.toString();
                break;
            case 'Número':
                formattedValue = formatNumber(baseValue, decimals);
                break;
            case 'Moneda':
                formattedValue = 'L' + formatNumber(baseValue, decimals);
                break;
            case 'Fecha':
                try {
                    const date = new Date(baseValue);
                    if (!isNaN(date.getTime())) {
                        formattedValue = date.toLocaleDateString();
                    } else {
                        formattedValue = baseValue.toString();
                    }
                } catch (e) {
                    formattedValue = baseValue.toString();
                }
                break;
            case 'Porcentaje':
                // Si venimos de un formato moneda o número, mantenemos el valor tal cual
                if (previousFormat === 'Moneda' || previousFormat === 'Número' || previousFormat === 'General') {
                    formattedValue = formatNumber(baseValue, decimals) + '%';
                    originalValue = baseValue / 100; // Guardamos el valor original como decimal
                } 
                // Si ya estaba en porcentaje, multiplicamos por 100 para mostrar
                else if (previousFormat === 'Porcentaje') {
                    formattedValue = formatNumber(baseValue * 100, decimals) + '%';
                    originalValue = baseValue;
                } 
                // Para cualquier otro formato, comportamiento estándar
                else {
                    formattedValue = formatNumber(baseValue, decimals) + '%';
                    originalValue = baseValue / 100;
                }
                break;
            case 'SeparadorMiles':
                formattedValue = formatNumber(baseValue, decimals);
                break;
            case 'AumentarDecimales':
                if (!cell.dataset.format) {
                    cell.dataset.format = 'Número';
                    cell.dataset.decimals = '0';
                }
                
                let decActuales = parseInt(cell.dataset.decimals || '0');
                decActuales++;
                cell.dataset.decimals = decActuales;
                
                let valorBase;
                if (cell.dataset.originalValue) {
                    valorBase = parseFloat(cell.dataset.originalValue);
                } else {
                    valorBase = parseFloat(rawValue.replace(/[^0-9.-]/g, '')) || 0;
                    cell.dataset.originalValue = valorBase.toString();
                }
                
                formattedValue = formatNumber(valorBase, decActuales);
                
                input.value = formattedValue;
                span.textContent = formattedValue;
                
                const xAumentar = parseInt(cell.dataset.x);
                const yAumentar = parseInt(cell.dataset.y);
                if (window.updateCell) {
                    window.updateCell(xAumentar, yAumentar, formattedValue, valorBase.toString());
                    
                    if (window.updateAllDependentCells) {
                        window.updateAllDependentCells();
                    }
                    
                    setTimeout(() => {
                        const updatedCell = document.querySelector(`td[data-x="${xAumentar}"][data-y="${yAumentar}"]`);
                        if (updatedCell) {
                            const updatedSpan = updatedCell.querySelector('span');
                            const updatedInput = updatedCell.querySelector('input');
                            if (updatedSpan) updatedSpan.textContent = formattedValue;
                            if (updatedInput) updatedInput.value = formattedValue;
                        }
                    }, 50);
                }
                return;
            case 'DisminuirDecimales':
                const formatoAct = cell.dataset.format || 'Número';
                
                if (!cell.dataset.format) {
                    cell.dataset.format = 'Número';
                    cell.dataset.decimals = '2';
                }
                
                let decimalesAct = parseInt(cell.dataset.decimals || '0');
                if (decimalesAct > 0) {
                    decimalesAct--;
                    cell.dataset.decimals = decimalesAct;
                    
                    let valorOriginal;
                    if (cell.dataset.originalValue) {
                        valorOriginal = parseFloat(cell.dataset.originalValue);
                    } else {
                        valorOriginal = parseFloat(rawValue.replace(/[^0-9.-]/g, '')) || 0;
                        cell.dataset.originalValue = valorOriginal.toString();
                    }
                    
                    formattedValue = formatNumber(valorOriginal, decimalesAct);
                    
                    input.value = formattedValue;
                    span.textContent = formattedValue;
                    
                    const xDisminuir = parseInt(cell.dataset.x);
                    const yDisminuir = parseInt(cell.dataset.y);
                    if (window.updateCell) {
                        window.updateCell(xDisminuir, yDisminuir, formattedValue, valorOriginal.toString());
                        
                        if (window.updateAllDependentCells) {
                            window.updateAllDependentCells();
                        }
                        
                        setTimeout(() => {
                            const updatedCell = document.querySelector(`td[data-x="${xDisminuir}"][data-y="${yDisminuir}"]`);
                            if (updatedCell) {
                                const updatedSpan = updatedCell.querySelector('span');
                                const updatedInput = updatedCell.querySelector('input');
                                if (updatedSpan) updatedSpan.textContent = formattedValue;
                                if (updatedInput) updatedInput.value = formattedValue;
                            }
                        }, 50);
                    }
                }
                return;
            case 'Hora':
                try {
                    const valor24h = baseValue;
                    const totalSegundos = valor24h * 86400;
                    const horas = Math.floor(totalSegundos / 3600) % 24;
                    const minutos = Math.floor((totalSegundos % 3600) / 60);
                    const segundos = Math.floor(totalSegundos % 60);
                    
                    formattedValue = 
                        horas.toString().padStart(2, '0') + ':' + 
                        minutos.toString().padStart(2, '0') + ':' + 
                        segundos.toString().padStart(2, '0');
                } catch (e) {
                    formattedValue = '00:00:00';
                }
                break;
        }
        
        if (formattedValue === undefined || formattedValue === null || formattedValue === '') {
            switch (format) {
                case 'Hora':
                    formattedValue = '00:00:00';
                    break;
                case 'Porcentaje':
                    formattedValue = formatNumber(0, decimals) + '%';
                    break;
                default:
                    formattedValue = '0';
            }
        }
        
        if (format !== 'AumentarDecimales' && format !== 'DisminuirDecimales') {
            cell.dataset.format = format;
            cell.dataset.decimals = decimals;
            cell.dataset.originalValue = originalValue.toString();
        }
        
        input.value = formattedValue;
        span.textContent = formattedValue;
        
        if (format === 'Número' || format === 'Moneda' || format === 'Porcentaje') {
            cell.style.textAlign = 'right';
        } else if (format === 'General') {
            cell.style.textAlign = 'left';
        }
        
        const x = parseInt(cell.dataset.x);
        const y = parseInt(cell.dataset.y);
        if (window.updateCell) {
            window.updateCell(x, y, formattedValue, originalValue.toString());
        }
    }
    
    // ===== EVENT LISTENERS =====
    
    // Eventos de borde
    borderSelect.addEventListener('change', function(e) {
        const selectedCells = document.querySelectorAll('.selected-range, .celdailumida');
        
        if (selectedCells.length === 0) {
            alert('Seleccione una celda o rango de celdas primero');
            this.value = '';
            return;
        }

        const borderType = this.value;
        const color = borderColor.value;

        selectedCells.forEach(cell => {
            applyBorders(cell, borderType, color);
        });
        
        const currentValue = this.value;
        setTimeout(() => {
            this.value = '';
            setTimeout(() => {
                this.value = currentValue;
            }, 50);
        }, 50);
    });

    borderColor.addEventListener('change', function(e) {
        const selectedCells = document.querySelectorAll('.selected-range, .celdailumida');
        const borderType = borderSelect.value;
        
        if (selectedCells.length === 0) {
            alert('Seleccione una celda o rango de celdas primero');
            return;
        }

        const color = this.value;
        selectedCells.forEach(cell => {
            applyBorders(cell, borderType, color);
        });
    });

    // Combinar celdas
    if (btnCombinarCeldas) {
        btnCombinarCeldas.addEventListener('click', () => {
            const selectedCells = document.querySelectorAll('.selected-range');
            
            if (selectedCells.length <= 1) {
                alert('Seleccione al menos dos celdas para combinar');
                return;
            }
            
            let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
            
            selectedCells.forEach(cell => {
                const x = parseInt(cell.dataset.x);
                const y = parseInt(cell.dataset.y);
                minX = Math.min(minX, x);
                maxX = Math.max(maxX, x);
                minY = Math.min(minY, y);
                maxY = Math.max(maxY, y);
            });
            
            if ((maxX - minX + 1) * (maxY - minY + 1) !== selectedCells.length) {
                alert('Solo se pueden combinar celdas que formen un rectángulo continuo');
                return;
            }
            
            const firstCell = document.querySelector(`td[data-x="${minX}"][data-y="${minY}"]`);
            
            let firstCellValue = '';
            const input = firstCell.querySelector('input');
            const span = firstCell.querySelector('span');
            
            if (input && input.style.opacity !== '0' && input.value) {
                firstCellValue = String(input.value);
            } else if (span) {
                firstCellValue = String(span.textContent);
            }
            
            if (firstCell.dataset.originalValue && firstCell.dataset.originalFormula) {
                firstCellValue = String(firstCell.dataset.originalValue);
            }
            
            firstCell.setAttribute('rowspan', maxY - minY + 1);
            firstCell.setAttribute('colspan', maxX - minX + 1);
            firstCell.classList.add('combined-cell');
            
            firstCell.style.verticalAlign = 'middle';
            firstCell.style.textAlign = 'center';
            firstCell.style.height = `${(maxY - minY + 1) * 30}px`;
            firstCell.style.width = `${(maxX - minX + 1) * 128}px`;

            if (span) span.textContent = firstCellValue;
            if (input) input.value = firstCellValue;
            
            firstCell.dataset.originalValue = firstCellValue;
            
            selectedCells.forEach(cell => {
                if (cell !== firstCell) {
                    cell.style.display = 'none';
                }
            });
            
            document.querySelectorAll('.selected-range').forEach(cell => {
                cell.classList.remove('selected-range');
            });
            
            function preserveExactValue(x, y, value) {
                if (window.updateCell) {
                    window.updateCell(x, y, value, value, true);
                    
                    setTimeout(() => {
                        const updatedCell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
                        if (updatedCell) {
                            const updatedSpan = updatedCell.querySelector('span');
                            const updatedInput = updatedCell.querySelector('input');
                            if (updatedSpan) updatedSpan.textContent = value;
                            if (updatedInput) updatedInput.value = value;
                        }
                    }, 50);
                } else {
                    saveState();
                }
            }
            
            if (firstCellValue) {
                const x = parseInt(firstCell.dataset.x);
                const y = parseInt(firstCell.dataset.y);
                preserveExactValue(x, y, firstCellValue);
            } else {
                saveState();
            }
        });
    }

    // Eventos de formato numérico
    if (selectorNumero) {
        selectorNumero.addEventListener('change', (e) => {
            const format = e.target.value;
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, format);
            });
        });
        
        selectorNumero.addEventListener("input", (e) => {
            if (e.target.value === "Moneda") {
                const selectedCells = $$(".celdailumida");
                selectedCells.forEach(cell => {
                    const input = cell.querySelector("input");
                    if (input) {
                        const baseValue = getSimpleBaseNumber(input.value);
                        const newValue = `L${baseValue}`;
                        updateCellValue(cell, newValue, { textAlign: "right" });
                        
                        setTimeout(() => {
                            const updatedCell = $(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                            if (updatedCell && updatedCell.querySelector("input")) {
                                updatedCell.querySelector("input").value = newValue;
                                if (updatedCell.querySelector("span")) {
                                    updatedCell.querySelector("span").textContent = newValue;
                                }
                            }
                        }, 50);
                    }
                });
            }
        });
    }
    
    if (btnFormatoMoneda) {
        btnFormatoMoneda.addEventListener('click', () => {
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, 'Moneda');
            });
        });
    }
    
    if (btnPorcentaje) {
        btnPorcentaje.addEventListener('click', () => {
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, 'Porcentaje');
            });
        });
    }
    
    if (btnSeparadorMiles) {
        btnSeparadorMiles.addEventListener('click', () => {
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, 'SeparadorMiles');
            });
        });
    }
    
    if (btnAumentarDecimales) {
        btnAumentarDecimales.addEventListener('click', () => {
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, 'AumentarDecimales');
            });
        });
    }
    
    if (btnDisminuirDecimales) {
        btnDisminuirDecimales.addEventListener('click', () => {
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, 'DisminuirDecimales');
            });
        });
    }
    
    // Eventos de portapapeles
    if (btnCortar) {
        btnCortar.addEventListener("click", () => {
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
    
    if (btnCopiar) {
        btnCopiar.addEventListener("click", () => {
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
    
    if (btnPegar) {
        btnPegar.addEventListener("click", () => {
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
    
    if (btnCopiarFormato) {
        btnCopiarFormato.addEventListener("click", () => {
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
    
    // Eventos de formato de fuente
    if (selectorFuente) {
        selectorFuente.addEventListener("change", (e) => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { fontFamily: e.target.value });
            });
            focusActiveCell();
        });
    }
    
    if (selectorTamaño) {
        selectorTamaño.addEventListener("change", (e) => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { fontSize: `${e.target.value}px` });
            });
            focusActiveCell();
        });
    }
    
    if (btnNegrita) {
        btnNegrita.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                const isBold = cell.style.fontWeight === "bold";
                updateCellValue(cell, null, { fontWeight: isBold ? "normal" : "bold" });
            });
            focusActiveCell();
        });
    }
    
    if (btnCursiva) {
        btnCursiva.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                const isItalic = cell.style.fontStyle === "italic";
                updateCellValue(cell, null, { fontStyle: isItalic ? "normal" : "italic" });
            });
            focusActiveCell();
        });
    }
    
    if (btnSubrayado) {
        btnSubrayado.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                const isUnderlined = cell.style.textDecoration === "underline";
                updateCellValue(cell, null, { textDecoration: isUnderlined ? "none" : "underline" });
            });
            focusActiveCell();
        });
    }
    
    // Eventos de color
    if (btnColorRelleno) {
        btnColorRelleno.addEventListener("click", () => {
            const fillColorInput = $("#fill-color");
            if (fillColorInput) {
                fillColorInput.click();
            }
        });
    }
    
    if (btnColorFuente) {
        btnColorFuente.addEventListener("click", () => {
            const fontColorInput = $("#font-color");
            if (fontColorInput) {
                fontColorInput.click();
            }
        });
    }
    
    $("#fill-color")?.addEventListener("change", (e) => {
        const selectedCells = $$(".celdailumida");
        if (selectedCells.length === 0) return;
        
        selectedCells.forEach(cell => {
            updateCellValue(cell, null, { backgroundColor: e.target.value });
            setTimeout(() => {
                const updatedCell = document.querySelector(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                if (updatedCell) {
                    updatedCell.style.backgroundColor = e.target.value;
                }
            }, 50);
        });
        focusActiveCell();
    });
    
    $("#font-color")?.addEventListener("change", (e) => {
        const selectedCells = $$(".celdailumida");
        if (selectedCells.length === 0) return;
        
        selectedCells.forEach(cell => {
            updateCellValue(cell, null, { color: e.target.value });
            setTimeout(() => {
                const updatedCell = document.querySelector(`td[data-x="${cell.dataset.x}"][data-y="${cell.dataset.y}"]`);
                if (updatedCell) {
                    updatedCell.style.color = e.target.value;
                }
            }, 50);
        });
        focusActiveCell();
    });
    
    // Eventos de alineación
    if (btnAlinearIzq) {
        btnAlinearIzq.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { textAlign: "left" });
            });
            focusActiveCell();
        });
    }
    
    if (btnCentrar) {
        btnCentrar.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { textAlign: "center" });
            });
            focusActiveCell();
        });
    }
    
    if (btnAlinearDer) {
        btnAlinearDer.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { textAlign: "right" });
            });
            focusActiveCell();
        });
    }
    
    if (btnJustificar) {
        btnJustificar.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                updateCellValue(cell, null, { textAlign: "justify" });
            });
            focusActiveCell();
        });
    }
    
    if (btnOrientacion) {
        btnOrientacion.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                const isVertical = cell.style.writingMode === "vertical-rl";
                updateCellValue(cell, null, { writingMode: isVertical ? "horizontal-tb" : "vertical-rl" });
            });
            focusActiveCell();
        });
    }
    
    if (btnSangria) {
        btnSangria.addEventListener("click", () => {
            $$(".celdailumida").forEach(cell => {
                const hasIndent = cell.style.paddingLeft === "20px";
                updateCellValue(cell, null, { paddingLeft: hasIndent ? "0px" : "20px" });
            });
            focusActiveCell();
        });
    }
    
    // Funciones de búsqueda y reemplazo
    
    // Crear modal para búsqueda
    function createSearchModal() {
        // Eliminar modal existente si hay alguno
        const existingModal = document.getElementById('search-modal');
        if (existingModal) {
            existingModal.remove();
        }
        
        const modal = document.createElement('div');
        modal.id = 'search-modal';
        modal.className = 'modal';
        modal.style.display = 'block';
        modal.style.position = 'fixed';
        modal.style.zIndex = '1000';
        modal.style.left = '0';
        modal.style.top = '0';
        modal.style.width = '100%';
        modal.style.height = '100%';
        modal.style.overflow = 'auto';
        modal.style.backgroundColor = 'rgba(0,0,0,0.4)';
        
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        modalContent.style.backgroundColor = '#fefefe';
        modalContent.style.margin = '15% auto';
        modalContent.style.padding = '20px';
        modalContent.style.border = '1px solid #888';
        modalContent.style.width = '300px';
        modalContent.style.borderRadius = '5px';
        
        const closeBtn = document.createElement('span');
        closeBtn.className = 'close-btn';
        closeBtn.innerHTML = '&times;';
        closeBtn.style.color = '#aaa';
        closeBtn.style.float = 'right';
        closeBtn.style.fontSize = '28px';
        closeBtn.style.fontWeight = 'bold';
        closeBtn.style.cursor = 'pointer';
        
        closeBtn.addEventListener('click', () => {
            modal.remove();
        });
        
        const title = document.createElement('h3');
        title.textContent = 'Buscar';
        title.style.marginTop = '0';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Término a buscar';
        input.style.width = '100%';
        input.style.padding = '8px';
        input.style.marginBottom = '10px';
        input.style.boxSizing = 'border-box';
        
        const matchCase = document.createElement('div');
        matchCase.style.marginBottom = '10px';
        
        const matchCaseCheckbox = document.createElement('input');
        matchCaseCheckbox.type = 'checkbox';
        matchCaseCheckbox.id = 'match-case';
        
        const matchCaseLabel = document.createElement('label');
        matchCaseLabel.htmlFor = 'match-case';
        matchCaseLabel.textContent = 'Coincidir mayúsculas/minúsculas';
        matchCaseLabel.style.marginLeft = '5px';
        
        matchCase.appendChild(matchCaseCheckbox);
        matchCase.appendChild(matchCaseLabel);
        
        const buttons = document.createElement('div');
        buttons.style.textAlign = 'right';
        
        const searchBtn = document.createElement('button');
        searchBtn.textContent = 'Buscar';
        searchBtn.style.padding = '8px 16px';
        searchBtn.style.backgroundColor = '#217346';
        searchBtn.style.color = 'white';
        searchBtn.style.border = 'none';
        searchBtn.style.borderRadius = '4px';
        searchBtn.style.cursor = 'pointer';
        searchBtn.style.marginLeft = '10px';
        
        const findNextBtn = document.createElement('button');
        findNextBtn.textContent = 'Buscar siguiente';
        findNextBtn.style.padding = '8px 16px';
        findNextBtn.style.backgroundColor = '#f0f0f0';
        findNextBtn.style.border = '1px solid #ccc';
        findNextBtn.style.borderRadius = '4px';
        findNextBtn.style.cursor = 'pointer';
        findNextBtn.disabled = true;
        
        buttons.appendChild(findNextBtn);
        buttons.appendChild(searchBtn);
        
        modalContent.appendChild(closeBtn);
        modalContent.appendChild(title);
        modalContent.appendChild(input);
        modalContent.appendChild(matchCase);
        modalContent.appendChild(buttons);
        
        modal.appendChild(modalContent);
        document.body.appendChild(modal);
        
        input.focus();
        
        // Búsqueda en celdas
        searchBtn.addEventListener('click', () => {
            const searchTerm = input.value.trim();
            if (!searchTerm) return;
            
            const isCaseSensitive = matchCaseCheckbox.checked;
            
            searchResults = [];
            currentSearchIndex = -1;
            
            // Buscar en todas las celdas
            for (let x = 0; x < state.length; x++) {
                for (let y = 0; y < state[x].length; y++) {
                    const cellValue = state[x][y].value;
                    const cellComputedValue = state[x][y].computedValue;
                    
                    if (cellValue || cellComputedValue) {
                        let valueToSearch = cellValue || cellComputedValue;
                        if (typeof valueToSearch !== 'string') {
                            valueToSearch = String(valueToSearch);
                        }
                        
                        const termToFind = isCaseSensitive ? searchTerm : searchTerm.toLowerCase();
                        const valueToCompare = isCaseSensitive ? valueToSearch : valueToSearch.toLowerCase();
                        
                        if (valueToCompare.includes(termToFind)) {
                            searchResults.push({ x, y });
                        }
                    }
                }
            }
            
            if (searchResults.length > 0) {
                findNextBtn.disabled = false;
                highlightNextResult();
            } else {
                alert('No se encontraron resultados.');
                findNextBtn.disabled = true;
            }
        });
        
        findNextBtn.addEventListener('click', () => {
            highlightNextResult();
        });
        
        return modal;
    }
    
    // Crear modal para reemplazo
    function createReplaceModal() {
        // Eliminar modal existente si hay alguno
        const existingModal = document.getElementById('replace-modal');
        if (existingModal) {
            existingModal.remove();
        }
        
        const modal = document.createElement('div');
        modal.id = 'replace-modal';
        modal.className = 'modal';
        modal.style.display = 'block';
        modal.style.position = 'fixed';
        modal.style.zIndex = '1000';
        modal.style.left = '0';
        modal.style.top = '0';
        modal.style.width = '100%';
        modal.style.height = '100%';
        modal.style.overflow = 'auto';
        modal.style.backgroundColor = 'rgba(0,0,0,0.4)';
        
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        modalContent.style.backgroundColor = '#fefefe';
        modalContent.style.margin = '15% auto';
        modalContent.style.padding = '20px';
        modalContent.style.border = '1px solid #888';
        modalContent.style.width = '350px';
        modalContent.style.borderRadius = '5px';
        
        const closeBtn = document.createElement('span');
        closeBtn.className = 'close-btn';
        closeBtn.innerHTML = '&times;';
        closeBtn.style.color = '#aaa';
        closeBtn.style.float = 'right';
        closeBtn.style.fontSize = '28px';
        closeBtn.style.fontWeight = 'bold';
        closeBtn.style.cursor = 'pointer';
        
        closeBtn.addEventListener('click', () => {
            modal.remove();
        });
        
        const title = document.createElement('h3');
        title.textContent = 'Buscar y Reemplazar';
        title.style.marginTop = '0';
        
        const findLabel = document.createElement('label');
        findLabel.textContent = 'Buscar:';
        findLabel.style.display = 'block';
        findLabel.style.marginBottom = '5px';
        
        const findInput = document.createElement('input');
        findInput.type = 'text';
        findInput.placeholder = 'Término a buscar';
        findInput.style.width = '100%';
        findInput.style.padding = '8px';
        findInput.style.marginBottom = '10px';
        findInput.style.boxSizing = 'border-box';
        
        const replaceLabel = document.createElement('label');
        replaceLabel.textContent = 'Reemplazar con:';
        replaceLabel.style.display = 'block';
        replaceLabel.style.marginBottom = '5px';
        
        const replaceInput = document.createElement('input');
        replaceInput.type = 'text';
        replaceInput.placeholder = 'Término de reemplazo';
        replaceInput.style.width = '100%';
        replaceInput.style.padding = '8px';
        replaceInput.style.marginBottom = '10px';
        replaceInput.style.boxSizing = 'border-box';
        
        const matchCase = document.createElement('div');
        matchCase.style.marginBottom = '10px';
        
        const matchCaseCheckbox = document.createElement('input');
        matchCaseCheckbox.type = 'checkbox';
        matchCaseCheckbox.id = 'match-case-replace';
        
        const matchCaseLabel = document.createElement('label');
        matchCaseLabel.htmlFor = 'match-case-replace';
        matchCaseLabel.textContent = 'Coincidir mayúsculas/minúsculas';
        matchCaseLabel.style.marginLeft = '5px';
        
        matchCase.appendChild(matchCaseCheckbox);
        matchCase.appendChild(matchCaseLabel);
        
        const buttons = document.createElement('div');
        buttons.style.textAlign = 'right';
        buttons.style.marginTop = '15px';
        
        const findBtn = document.createElement('button');
        findBtn.textContent = 'Buscar';
        findBtn.style.padding = '8px 16px';
        findBtn.style.backgroundColor = '#f0f0f0';
        findBtn.style.border = '1px solid #ccc';
        findBtn.style.borderRadius = '4px';
        findBtn.style.cursor = 'pointer';
        
        const replaceBtn = document.createElement('button');
        replaceBtn.textContent = 'Reemplazar';
        replaceBtn.style.padding = '8px 16px';
        replaceBtn.style.backgroundColor = '#f0f0f0';
        replaceBtn.style.border = '1px solid #ccc';
        replaceBtn.style.borderRadius = '4px';
        replaceBtn.style.cursor = 'pointer';
        replaceBtn.style.marginLeft = '5px';
        
        const replaceAllBtn = document.createElement('button');
        replaceAllBtn.textContent = 'Reemplazar todo';
        replaceAllBtn.style.padding = '8px 16px';
        replaceAllBtn.style.backgroundColor = '#217346';
        replaceAllBtn.style.color = 'white';
        replaceAllBtn.style.border = 'none';
        replaceAllBtn.style.borderRadius = '4px';
        replaceAllBtn.style.cursor = 'pointer';
        replaceAllBtn.style.marginLeft = '5px';
        
        buttons.appendChild(findBtn);
        buttons.appendChild(replaceBtn);
        buttons.appendChild(replaceAllBtn);
        
        modalContent.appendChild(closeBtn);
        modalContent.appendChild(title);
        modalContent.appendChild(findLabel);
        modalContent.appendChild(findInput);
        modalContent.appendChild(replaceLabel);
        modalContent.appendChild(replaceInput);
        modalContent.appendChild(matchCase);
        modalContent.appendChild(buttons);
        
        modal.appendChild(modalContent);
        document.body.appendChild(modal);
        
        findInput.focus();
        
        // Buscar términos en celdas
        findBtn.addEventListener('click', () => {
            const searchTerm = findInput.value.trim();
            if (!searchTerm) return;
            
            const isCaseSensitive = matchCaseCheckbox.checked;
            
            searchResults = [];
            currentSearchIndex = -1;
            
            // Buscar en todas las celdas
            for (let x = 0; x < state.length; x++) {
                for (let y = 0; y < state[x].length; y++) {
                    const cellValue = state[x][y].value;
                    const cellComputedValue = state[x][y].computedValue;
                    
                    if (cellValue || cellComputedValue) {
                        let valueToSearch = cellValue || cellComputedValue;
                        if (typeof valueToSearch !== 'string') {
                            valueToSearch = String(valueToSearch);
                        }
                        
                        const termToFind = isCaseSensitive ? searchTerm : searchTerm.toLowerCase();
                        const valueToCompare = isCaseSensitive ? valueToSearch : valueToSearch.toLowerCase();
                        
                        if (valueToCompare.includes(termToFind)) {
                            searchResults.push({ x, y });
                        }
                    }
                }
            }
            
            if (searchResults.length > 0) {
                highlightNextResult();
            } else {
                alert('No se encontraron resultados.');
            }
        });
        
        // Reemplazar el término en la celda actual
        replaceBtn.addEventListener('click', () => {
            if (currentSearchIndex === -1 || searchResults.length === 0) {
                alert('Primero debe buscar un término.');
                return;
            }
            
            const searchTerm = findInput.value.trim();
            const replaceTerm = replaceInput.value;
            const isCaseSensitive = matchCaseCheckbox.checked;
            
            const { x, y } = searchResults[currentSearchIndex];
            let cellValue = state[x][y].value;
            
            if (cellValue.startsWith('=')) {
                // Si es una fórmula, no la reemplazamos
                alert('No se puede reemplazar texto en una fórmula.');
                return;
            }
            
            if (typeof cellValue !== 'string') {
                cellValue = String(cellValue);
            }
            
            let newValue;
            if (isCaseSensitive) {
                newValue = cellValue.replace(searchTerm, replaceTerm);
            } else {
                // Reemplazar sin importar mayúsculas/minúsculas
                const regex = new RegExp(searchTerm, 'gi');
                newValue = cellValue.replace(regex, replaceTerm);
            }
            
            // Actualizar la celda
            updateCell(x, y, newValue);
            
            // Mover al siguiente resultado si existe
            if (currentSearchIndex < searchResults.length - 1) {
                currentSearchIndex++;
                highlightNextResult();
            } else {
                // Volver al inicio si hemos llegado al final
                currentSearchIndex = -1;
                searchResults = [];
                alert('Reemplazo completado. No hay más resultados.');
            }
        });
        
        // Reemplazar todos los términos encontrados
        replaceAllBtn.addEventListener('click', () => {
            const searchTerm = findInput.value.trim();
            if (!searchTerm) return;
            
            const replaceTerm = replaceInput.value;
            const isCaseSensitive = matchCaseCheckbox.checked;
            
            let replacesCount = 0;
            
            // Buscar y reemplazar en todas las celdas
            for (let x = 0; x < state.length; x++) {
                for (let y = 0; y < state[x].length; y++) {
                    const cellValue = state[x][y].value;
                    
                    if (cellValue && !cellValue.startsWith('=')) {
                        if (typeof cellValue !== 'string') {
                            continue;
                        }
                        
                        const termToFind = isCaseSensitive ? searchTerm : searchTerm.toLowerCase();
                        const valueToCompare = isCaseSensitive ? cellValue : cellValue.toLowerCase();
                        
                        if (valueToCompare.includes(termToFind)) {
                            let newValue;
                            if (isCaseSensitive) {
                                newValue = cellValue.replace(new RegExp(searchTerm, 'g'), replaceTerm);
                            } else {
                                const regex = new RegExp(searchTerm, 'gi');
                                newValue = cellValue.replace(regex, replaceTerm);
                            }
                            
                            // Actualizar la celda
                            updateCell(x, y, newValue);
                            replacesCount++;
                        }
                    }
                }
            }
            
            alert(`Se realizaron ${replacesCount} reemplazos.`);
        });
        
        return modal;
    }
    
    // Resaltar siguiente resultado de búsqueda
    function highlightNextResult() {
        if (searchResults.length === 0) return;
        
        currentSearchIndex = (currentSearchIndex + 1) % searchResults.length;
        const { x, y } = searchResults[currentSearchIndex];
        
        // Limpiar selecciones anteriores
        clearSelections();
        
        // Seleccionar la celda con el resultado
        const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
        if (cell) {
            cell.classList.add('celdailumida');
            cell.scrollIntoView({ behavior: 'smooth', block: 'center' });
            updateActiveCellDisplay(x, y);
        }
    }
    
    // Eventos para botones de búsqueda y reemplazo
    if (btnBuscar) {
        btnBuscar.addEventListener('click', () => {
            createSearchModal();
        });
    }
    
    if (btnReemplazar) {
        btnReemplazar.addEventListener('click', () => {
            createReplaceModal();
        });
    }
    
    // Eventos adicionales de la tabla
    const table = $("table");
    if (table) {
        table.addEventListener("click", (e) => {
            const td = e.target.closest("td");
            if (td && td.querySelector("input")) {
                // click en celda
            }
        });
    }
});