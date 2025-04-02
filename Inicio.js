document.addEventListener('DOMContentLoaded', () => {
    const borderSelect = document.getElementById('border-select');
    const borderColor = document.getElementById('border-color');
    const defaultBorderColor = '#e2e2e2'; // Color por defecto
    const defaultBorderWidth = '1px'; // Grosor por defecto
    
    // Función auxiliar para aplicar bordes
    function applyBorders(cell, borderType, color) {
        // Primero limpiamos todos los estilos de borde anteriores
        cell.style.outline = '';
        cell.style.borderTop = '';
        cell.style.borderRight = '';
        cell.style.borderBottom = '';
        cell.style.borderLeft = '';
        
        // Remover la clase de selección que causa el resaltado verde
        cell.classList.remove('cell-with-border');
        
        switch(borderType) {
            case 'all':
                // Usar outline para "todos los bordes"
                cell.style.outline = `1px solid ${color}`;
                cell.style.position = 'relative';
                cell.style.zIndex = '100'; // Mayor z-index para asegurar visibilidad
                break;
            case 'top':
                // Restaurar posición y z-index primero
                cell.style.position = 'relative';
                cell.style.zIndex = '100';
                
                // Aplicar borde con !important para mayor prioridad
                cell.setAttribute('style', cell.getAttribute('style') + `;border-top: 2px solid ${color} !important;`);
                
                // Restaurar otros bordes
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
                // Restaurar todos los bordes al estilo por defecto
                cell.style.position = '';
                cell.style.zIndex = '';
                
                cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                break;
        }
    }
    
    borderSelect.addEventListener('change', function(e) {
        // Seleccionar tanto rangos como celdas individuales
        const selectedCells = document.querySelectorAll('.selected-range, .celdailumida');
        
        if (selectedCells.length === 0) {
            alert('Seleccione una celda o rango de celdas primero');
            this.value = '';
            return;
        }

        const borderType = this.value;
        const color = borderColor.value;

        // Aplicar bordes a cada celda seleccionada
        selectedCells.forEach(cell => {
            applyBorders(cell, borderType, color);
        });
        
        const currentValue = this.value;
        // Resetear el valor y luego restaurarlo después de un breve retraso
        setTimeout(() => {
            this.value = '';
            setTimeout(() => {
                this.value = currentValue;
            }, 50);
        }, 50);
    });

    // Evento para cambiar el color del borde
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

    // ===== COMBINAR CELDAS =====
    const btnCombinarCeldas = document.querySelector("button[title='Combinar celdas']");
    if (btnCombinarCeldas) {
        btnCombinarCeldas.addEventListener('click', () => {
            const selectedCells = document.querySelectorAll('.selected-range');
            
            if (selectedCells.length <= 1) {
                alert('Seleccione al menos dos celdas para combinar');
                return;
            }
            
            // Determinar el rectángulo que abarca el rango
            let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
            
            selectedCells.forEach(cell => {
                const x = parseInt(cell.dataset.x);
                const y = parseInt(cell.dataset.y);
                minX = Math.min(minX, x);
                maxX = Math.max(maxX, x);
                minY = Math.min(minY, y);
                maxY = Math.max(maxY, y);
            });
            
            // Verificar si las celdas forman un rectángulo continuo
            if ((maxX - minX + 1) * (maxY - minY + 1) !== selectedCells.length) {
                alert('Solo se pueden combinar celdas que formen un rectángulo continuo');
                return;
            }
            
            // Obtener el valor de la primera celda (esquina superior izquierda)
            const firstCell = document.querySelector(`td[data-x="${minX}"][data-y="${minY}"]`);
            
            // Obtener el valor exacto como texto sin procesar
            let firstCellValue = '';
            const input = firstCell.querySelector('input');
            const span = firstCell.querySelector('span');
            
            // Obtener el valor de la celda exactamente como aparece, sin conversión a número
            if (input && input.style.opacity !== '0' && input.value) {
                firstCellValue = String(input.value);
            } else if (span) {
                firstCellValue = String(span.textContent);
            }
            
            // Si el input tiene un valor de fórmula, obtener ese valor explícitamente
            if (firstCell.dataset.originalValue && firstCell.dataset.originalFormula) {
                firstCellValue = String(firstCell.dataset.originalValue);
            }
            
            // Crear la celda combinada en la posición de la primera celda
            firstCell.setAttribute('rowspan', maxY - minY + 1);
            firstCell.setAttribute('colspan', maxX - minX + 1);
            firstCell.classList.add('combined-cell');
            
            // Aplicar estilos para que el contenido ocupe todo el espacio de las celdas combinadas
            firstCell.style.verticalAlign = 'middle';
            firstCell.style.textAlign = 'center';
            firstCell.style.height = `${(maxY - minY + 1) * 30}px`; // Altura aproximada de cada celda
            firstCell.style.width = `${(maxX - minX + 1) * 128}px`; // Ancho aproximado de cada celda

            // Actualizar directamente el valor sin procesar o convertir
            if (span) span.textContent = firstCellValue;
            if (input) input.value = firstCellValue;
            
            // Almacena el valor original como string para evitar problemas de precisión
            firstCell.dataset.originalValue = firstCellValue;
            
            // Ocultar las demás celdas del rango
            selectedCells.forEach(cell => {
                if (cell !== firstCell) {
                    cell.style.display = 'none';
                }
            });
            
            // Limpieza de selección
            document.querySelectorAll('.selected-range').forEach(cell => {
                cell.classList.remove('selected-range');
            });
            
            // Función para preservar el valor exacto como string en el updateCell
            function preserveExactValue(x, y, value) {
                if (window.updateCell) {
                    // Pasar el valor exacto como string con flag para preservarlo
                    window.updateCell(x, y, value, value, true);
                    
                    // Asegurarse que el valor visible se mantiene exacto después de la actualización
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
            
            // Actualizar el estado con el valor exacto como string
            if (firstCellValue) {
                const x = parseInt(firstCell.dataset.x);
                const y = parseInt(firstCell.dataset.y);
                preserveExactValue(x, y, firstCellValue);
            } else {
                saveState();
            }
        });
    }

    // ===== FORMATO NUMÉRICO =====
    const selectorNumero = document.querySelector("#submenu-inicio .ribbon-group:nth-child(4) .fila.selectores select");
    const btnFormatoMoneda = document.querySelector("button[title='Formato de número']");
    const btnPorcentaje = document.querySelector("button[title='Porcentaje']");
    const btnSeparadorMiles = document.querySelector("button[title='Separador de miles']");
    const btnAumentarDecimales = document.querySelector("button[title='Aumentar decimales']");
    const btnDisminuirDecimales = document.querySelector("button[title='Disminuir decimales']");
    
    // Función para obtener el valor numérico base
    function getBaseNumber(value) {
        if (!value) return 0;
        
        // Manejar porcentajes existentes
        if (String(value).endsWith('%')) {
            return parseFloat(value.toString().replace(/[^0-9.-]/g, "")) / 100;
        }
        
        // Convertir directamente a número (manejar comas como decimales si es necesario)
        return parseFloat(value.toString().replace(/,/g, '')) || 0;
    }
    
    // Reescribir la función formatNumber para evitar problemas con separadores
    function formatNumber(number, decimals) {
        // Asegurarse de que trabajamos con un número limpio
        let num = parseFloat(number);
        if (isNaN(num)) num = 0;
        
        // Aplicar decimales exactos
        let formatted = num.toFixed(decimals);
        
        // Separar partes entera y decimal
        let parts = formatted.split('.');
        
        // Aplicar separador de miles solo a la parte entera
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        
        // Unir las partes
        return parts.join('.');
    }
    
    // Función para aplicar formato numérico
    function applyNumberFormat(cell, format, decimals = 2) {
        const input = cell.querySelector('input');
        const span = cell.querySelector('span');
        if (!input || !span) return;
        
        let value = input.value || span.textContent || '0';
        const baseValue = getBaseNumber(value);
        // Guardar el valor original antes de cualquier conversión
        let originalValue = parseFloat(value.toString().replace(/,/g, '')) || 0;
        
        let formattedValue = '';
        
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
            case 'Contabilidad':
                try {
                    if (baseValue < 0) {
                        formattedValue = '(L' + formatNumber(Math.abs(baseValue), decimals) + ')';
                    } else {
                        formattedValue = 'L' + formatNumber(baseValue, decimals);
                    }
                } catch (e) {
                    console.error("Error al formatear contabilidad:", e);
                    formattedValue = baseValue.toString();
                }
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
                    console.error("Error al formatear fecha:", e);
                    formattedValue = baseValue.toString();
                }
                break;
            case 'Porcentaje':
                // Siempre trabajar con el valor numérico puro
                // Si ya tiene formato de porcentaje, no volver a convertir
                if (cell.dataset.format === 'Porcentaje') {
                    // Ya está en formato porcentaje, mantener el valor base
                    formattedValue = formatNumber(originalValue * 100, decimals) + '%';
                } else {
                    // Es una nueva conversión a porcentaje, mostrar el mismo valor
                    formattedValue = formatNumber(originalValue, decimals) + '%';
                    // Guardar el valor base como decimal para cálculos (20% = 0.2)
                    originalValue = originalValue / 100;
                }
                break;
            case 'SeparadorMiles':
                formattedValue = formatNumber(baseValue, decimals);
                break;
            case 'AumentarDecimales':
                // Si no tiene formato, inicializar
                if (!cell.dataset.format) {
                    cell.dataset.format = 'Número';
                    cell.dataset.decimals = '0';
                }
                
                // Obtener e incrementar decimales
                let decActuales = parseInt(cell.dataset.decimals || '0');
                decActuales++;
                cell.dataset.decimals = decActuales;
                
                // Obtener valor limpio numérico
                let valorBase;
                if (cell.dataset.originalValue) {
                    valorBase = parseFloat(cell.dataset.originalValue);
                } else {
                    valorBase = parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
                    cell.dataset.originalValue = valorBase.toString();
                }
                
                // Aplicar formato con los decimales correctos
                formattedValue = formatNumber(valorBase, decActuales);
                
                // Actualizar visualización inmediatamente
                input.value = formattedValue;
                span.textContent = formattedValue;
                
                // Forzar actualización visual inmediata
                const xAumentar = parseInt(cell.dataset.x);
                const yAumentar = parseInt(cell.dataset.y);
                if (window.updateCell) {
                    // Actualizar con nuevo valor visible pero conservar valor original
                    window.updateCell(xAumentar, yAumentar, formattedValue, valorBase.toString());
                    
                    // Forzar actualización global
                    if (window.updateAllDependentCells) {
                        window.updateAllDependentCells();
                    }
                    
                    // Asegurar que el valor visible se mantiene después del refresco
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
                // Obtener el formato actual
                const formatoAct = cell.dataset.format || 'Número';
                
                // Si no tiene formato, aplicar uno primero
                if (!cell.dataset.format) {
                    cell.dataset.format = 'Número';
                    cell.dataset.decimals = '2';
                }
                
                // Decrementar decimales (mínimo 0)
                let decimalesAct = parseInt(cell.dataset.decimals || '0');
                if (decimalesAct > 0) {
                    decimalesAct--;
                    cell.dataset.decimals = decimalesAct;
                    
                    // Obtener valor original
                    let valorOriginal;
                    if (cell.dataset.originalValue) {
                        valorOriginal = parseFloat(cell.dataset.originalValue);
                    } else {
                        valorOriginal = parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
                        cell.dataset.originalValue = valorOriginal.toString();
                    }
                    
                    // Formatear según tipo actual
                    formattedValue = formatNumber(valorOriginal, decimalesAct);
                    
                    // Actualizar visualización inmediata
                    input.value = formattedValue;
                    span.textContent = formattedValue;
                    
                    // Forzar actualización del estado global
                    const xDisminuir = parseInt(cell.dataset.x);
                    const yDisminuir = parseInt(cell.dataset.y);
                    if (window.updateCell) {
                        window.updateCell(xDisminuir, yDisminuir, formattedValue, valorOriginal.toString());
                        
                        // Forzar actualización global
                        if (window.updateAllDependentCells) {
                            window.updateAllDependentCells();
                        }
                        
                        // Asegurar que el valor visible se mantiene después del refresco
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
                    // Verificar si el valor es un número
                    let horaValue = parseFloat(baseValue);
                    if (isNaN(horaValue)) {
                        formattedValue = baseValue.toString();
                    } else {
                        // Convertir a horas, minutos y segundos
                        const totalSeconds = horaValue * 86400; // 24h * 60m * 60s
                        const hours = Math.floor(totalSeconds / 3600) % 24;
                        const minutes = Math.floor((totalSeconds % 3600) / 60);
                        const seconds = Math.floor(totalSeconds % 60);
                        formattedValue = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
                    }
                } catch (e) {
                    console.error("Error al formatear hora:", e);
                    formattedValue = baseValue.toString();
                }
                break;
        }
        
        // Actualizar la celda con el nuevo valor
        if (format !== 'AumentarDecimales' && format !== 'DisminuirDecimales') {
            // Guardar el formato aplicado para futuras operaciones
            cell.dataset.format = format;
            cell.dataset.decimals = decimals;
            cell.dataset.originalValue = originalValue.toString(); // Guardar el valor original sin formato
        }
        
        // Verificar que formattedValue no es undefined
        if (formattedValue === undefined) {
            console.error("Valor formateado es undefined para formato:", format, "valor original:", baseValue);
            formattedValue = baseValue.toString();
        }
        
        // Actualizar la visualización
        input.value = formattedValue;
        span.textContent = formattedValue;
        
        // Alineación para valores numéricos
        if (format === 'Número' || format === 'Moneda' || format === 'Contabilidad' || format === 'Porcentaje') {
            cell.style.textAlign = 'right';
        } else if (format === 'General') {
            cell.style.textAlign = 'left';
        }
        
        // Actualizar el estado
        const x = parseInt(cell.dataset.x);
        const y = parseInt(cell.dataset.y);
        if (window.updateCell) {
            // Actualizar con el valor formateado para mostrar pero guardar el valor original
            window.updateCell(x, y, formattedValue, originalValue.toString());
        }
    }
    
    // Eventos para los elementos de formato numérico
    if (selectorNumero) {
        selectorNumero.addEventListener('change', (e) => {
            const format = e.target.value;
            document.querySelectorAll('.selected-range, .celdailumida').forEach(cell => {
                applyNumberFormat(cell, format);
            });
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
    
    // Función auxiliar para guardar el estado
    function saveState() {
        if (window.updateAllDependentCells) {
            window.updateAllDependentCells();
        }
    }
});