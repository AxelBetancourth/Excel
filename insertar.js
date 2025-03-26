// Funciones para manejo de las opciones de insertar

document.addEventListener('DOMContentLoaded', function() {
    // Referencia a elementos de la interfaz
    const btnImagenes = document.getElementById('btn-imagenes');
    const submenuImagenes = document.getElementById('submenu-imagenes');
    const btnFormas = document.getElementById('btn-formas');
    const submenuFormas = document.getElementById('submenu-formas');
    const uploadImageBtn = document.getElementById('upload-image-btn');
    const imageInput = document.getElementById('image-input');
    
    // Botones de gráficos
    const btnGraficoBarras = document.getElementById('btn-grafico-barras');
    const btnGraficoLineas = document.getElementById('btn-grafico-lineas');
    const btnGraficoCircular = document.getElementById('btn-grafico-circular');
    const btnGraficoDispersion = document.getElementById('btn-grafico-dispersion');
    
    // Modal de configuración de gráficos
    const chartModal = document.getElementById('chart-modal');
    const closeButton = document.querySelector('.close-button');
    const cancelChart = document.getElementById('cancel-chart');
    const createChart = document.getElementById('create-chart');
    
    // Variable para almacenar el tipo de gráfico seleccionado
    let selectedChartType = '';
    
    // Agregar manejador de eventos para botones con atributo data-menu
    const menuButtons = document.querySelectorAll('button[data-menu]');
    menuButtons.forEach(button => {
        button.addEventListener('click', function() {
            const menuId = this.getAttribute('data-menu');
            const menu = document.getElementById(menuId);
            if (menu) {
                // Cerrar otros menús abiertos
                document.querySelectorAll('.menu.open').forEach(openMenu => {
                    if (openMenu !== menu) {
                        openMenu.classList.remove('open');
                    }
                });
                // Alternar clase 'open' en este menú
                menu.classList.toggle('open');
            }
        });
    });
    
    // Cerrar menús desplegables al hacer clic fuera de ellos
    document.addEventListener('click', function(e) {
        if (!e.target.closest('button[data-menu]') && !e.target.closest('.menu')) {
            document.querySelectorAll('.menu.open').forEach(menu => {
                menu.classList.remove('open');
            });
        }
    });
    
    // Función para mostrar modal de gráfico con tipo específico
    function showChartModal(chartType) {
        selectedChartType = chartType;
        chartModal.style.display = 'flex';
    }
    
    // Manejador para botones de gráficos
    if (btnGraficoBarras) {
        btnGraficoBarras.addEventListener('click', () => showChartModal('bar'));
    }
    
    if (btnGraficoLineas) {
        btnGraficoLineas.addEventListener('click', () => showChartModal('line'));
    }
    
    if (btnGraficoCircular) {
        btnGraficoCircular.addEventListener('click', () => showChartModal('pie'));
    }
    
    if (btnGraficoDispersion) {
        btnGraficoDispersion.addEventListener('click', () => showChartModal('scatter'));
    }
    
    // Cerrar modal de gráficos
    if (closeButton) {
        closeButton.addEventListener('click', () => {
            chartModal.style.display = 'none';
        });
    }
    
    if (cancelChart) {
        cancelChart.addEventListener('click', () => {
            chartModal.style.display = 'none';
        });
    }
    
    // Manejar clic fuera del modal para cerrarlo
    window.addEventListener('click', (e) => {
        if (e.target === chartModal) {
            chartModal.style.display = 'none';
        }
    });
    
    // Manejo de subida de imágenes
    if (uploadImageBtn && imageInput) {
        uploadImageBtn.addEventListener('click', () => {
            imageInput.click();
        });
        
        imageInput.addEventListener('change', handleImageUpload);
    }
    
    // Función para insertar imágenes
    function handleImageUpload(event) {
        const file = event.target.files[0];
        if (file && file.type.match('image.*')) {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                insertImageAtCursor(e.target.result);
            };
            
            reader.readAsDataURL(file);
        }
    }
    
    // Insertar imagen en la celda seleccionada
    function insertImageAtCursor(imageSrc) {
        // Aquí se debe implementar la lógica para insertar la imagen en la celda activa
        console.log('Insertando imagen:', imageSrc);
        // Esta parte se integrará con la funcionalidad principal de la hoja de cálculo
    }
    
    // Función para insertar formas
    window.insertShape = function(shapeType) {
        // Aquí se debe implementar la lógica para insertar la forma seleccionada
        console.log('Insertando forma:', shapeType);
        // Esta parte se integrará con la funcionalidad principal de la hoja de cálculo
    };
    
    // Crear gráfico con los datos seleccionados
    if (createChart) {
        createChart.addEventListener('click', () => {
            const chartTitle = document.getElementById('chart-title').value;
            const dataRange = document.getElementById('data-range').value;
            
            // Aquí se debe implementar la lógica para crear el gráfico
            console.log('Creando gráfico:', {
                type: selectedChartType,
                title: chartTitle,
                dataRange: dataRange
            });
            
            // Cerrar modal después de crear
            chartModal.style.display = 'none';
        });
    }
});

// Funcionalidad para insertar tablas
document.addEventListener('DOMContentLoaded', function() {
    const insertTableBtn = document.getElementById('insert-table-btn');
    
    // Crear modal para tabla si no existe
    if (!document.getElementById('table-modal')) {
        const tableModal = document.createElement('div');
        tableModal.id = 'table-modal';
        tableModal.className = 'chart-modal'; // Reusamos estilos del modal de gráficos
        
        tableModal.innerHTML = `
            <div class="chart-modal-content">
                <div class="chart-modal-header">
                    <h2>Insertar Tabla</h2>
                    <button class="close-button" id="close-table-modal">&times;</button>
                </div>
                <div class="chart-modal-body">
                    <div class="form-group">
                        <label for="table-range">Rango de datos:</label>
                        <input type="text" id="table-range" placeholder="Ejemplo: A1:D5">
                    </div>
                    <div class="form-group">
                        <label for="table-has-headers">Mi tabla tiene encabezados</label>
                        <input type="checkbox" id="table-has-headers" checked>
                    </div>
                    <div class="form-group">
                        <label>Estilo de tabla:</label>
                        <div class="table-styles-container">
                            <div class="table-style-option selected" data-style="style1">
                                <div class="style-preview style1"></div>
                                <span>Estilo 1</span>
                            </div>
                            <div class="table-style-option" data-style="style2">
                                <div class="style-preview style2"></div>
                                <span>Estilo 2</span>
                            </div>
                            <div class="table-style-option" data-style="style3">
                                <div class="style-preview style3"></div>
                                <span>Estilo 3</span>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="chart-modal-footer">
                    <button id="create-table">Aceptar</button>
                    <button id="cancel-table">Cancelar</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(tableModal);
        
        // Agregar estilos para las previsualizaciones de tabla
        const styleElement = document.createElement('style');
        styleElement.textContent = `
            .table-styles-container {
                display: flex;
                gap: 15px;
                margin-top: 10px;
            }
            
            .table-style-option {
                cursor: pointer;
                padding: 5px;
                border: 2px solid transparent;
                text-align: center;
                border-radius: 4px;
            }
            
            .table-style-option.selected {
                border-color: #217346;
            }
            
            .style-preview {
                width: 80px;
                height: 60px;
                margin-bottom: 5px;
                border: 1px solid #ccc;
            }
            
            .style1 {
                background: linear-gradient(to bottom, #4472C4 20%, #E9EBF5 20%, #E9EBF5 40%, white 40%, white 60%, #E9EBF5 60%, #E9EBF5 80%, white 80%);
            }
            
            .style2 {
                background: linear-gradient(to bottom, #70AD47 20%, #E2EFDA 20%, #E2EFDA 40%, white 40%, white 60%, #E2EFDA 60%, #E2EFDA 80%, white 80%);
            }
            
            .style3 {
                background: linear-gradient(to bottom, #ED7D31 20%, #FBE5D6 20%, #FBE5D6 40%, white 40%, white 60%, #FBE5D6 60%, #FBE5D6 80%, white 80%);
            }
        `;
        document.head.appendChild(styleElement);
        
        // Event listeners para el modal
        const closeTableModal = document.getElementById('close-table-modal');
        const cancelTable = document.getElementById('cancel-table');
        const createTable = document.getElementById('create-table');
        const tableStyleOptions = document.querySelectorAll('.table-style-option');
        
        closeTableModal.addEventListener('click', () => {
            tableModal.style.display = 'none';
        });
        
        cancelTable.addEventListener('click', () => {
            tableModal.style.display = 'none';
        });
        
        // Selección de estilo
        tableStyleOptions.forEach(option => {
            option.addEventListener('click', () => {
                tableStyleOptions.forEach(opt => opt.classList.remove('selected'));
                option.classList.add('selected');
            });
        });
        
        // Crear tabla
        createTable.addEventListener('click', () => {
            const rangeInput = document.getElementById('table-range').value;
            const hasHeaders = document.getElementById('table-has-headers').checked;
            const selectedStyle = document.querySelector('.table-style-option.selected').dataset.style;
            
            if (!rangeInput) {
                alert('Por favor, ingrese un rango válido');
                return;
            }
            
            createTableFromRange(rangeInput, hasHeaders, selectedStyle);
            tableModal.style.display = 'none';
        });
        
        // Cerrar modal al hacer clic fuera
        window.addEventListener('click', (e) => {
            if (e.target === tableModal) {
                tableModal.style.display = 'none';
            }
        });
    }
    
    // Mostrar modal al hacer clic en el botón de tabla
    if (insertTableBtn) {
        insertTableBtn.addEventListener('click', () => {
            // Si hay un rango seleccionado, prellenamos el campo
            const activeRange = document.getElementById('active-cell').textContent;
            if (activeRange && activeRange.includes(':')) {
                document.getElementById('table-range').value = activeRange;
            } else {
                // Si solo hay una celda activa, seleccionamos un rango predeterminado alrededor de ella
                const cellRef = activeRange;
                if (cellRef.match(/^[A-Z]+\d+$/)) {
                    const col = cellRef.match(/[A-Z]+/)[0];
                    const row = parseInt(cellRef.match(/\d+/)[0]);
                    document.getElementById('table-range').value = `${col}${row}:${col}${row + 4}`;
                }
            }
            
            document.getElementById('table-modal').style.display = 'flex';
        });
    }
    
    // Función para crear tabla a partir de un rango
    function createTableFromRange(range, hasHeaders, style) {
        try {
            // Validar formato del rango (por ejemplo, A1:D5)
            if (!range.match(/^[A-Z]+\d+:[A-Z]+\d+$/)) {
                throw new Error('Formato de rango inválido');
            }
            
            // Extraer coordenadas
            const [startRef, endRef] = range.split(':');
            const startCoords = getCellCoords(startRef);
            const endCoords = getCellCoords(endRef);
            
            // Asegurar que las coordenadas son válidas
            if (!startCoords || !endCoords) {
                throw new Error('Coordenadas inválidas');
            }
            
            // Determinar el rango real (min a max)
            const startX = Math.min(startCoords[0], endCoords[0]);
            const endX = Math.max(startCoords[0], endCoords[0]);
            const startY = Math.min(startCoords[1], endCoords[1]);
            const endY = Math.max(startCoords[1], endCoords[1]);
            
            // Aplicar estilo de tabla a las celdas
            applyTableStyle(startX, endX, startY, endY, hasHeaders, style);
            
            console.log(`Tabla creada con estilo ${style} en el rango ${range}, con encabezados: ${hasHeaders}`);
        } catch (error) {
            console.error('Error al crear tabla:', error);
            alert('Error al crear la tabla: ' + error.message);
        }
    }
    
    // Función para aplicar estilo de tabla a un rango de celdas
    function applyTableStyle(startX, endX, startY, endY, hasHeaders, styleType) {
        // Clase base para todas las tablas
        const baseClass = 'excel-table';
        
        // Definir clases según el estilo seleccionado
        let styleClass;
        switch (styleType) {
            case 'style1':
                styleClass = 'table-style-blue';
                break;
            case 'style2':
                styleClass = 'table-style-green';
                break;
            case 'style3':
                styleClass = 'table-style-orange';
                break;
            default:
                styleClass = 'table-style-blue';
        }
        
        // Aplicar estilos a cada celda en el rango
        for (let y = startY; y <= endY; y++) {
            for (let x = startX; x <= endX; x++) {
                const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
                if (cell) {
                    // Clase base para todas las celdas de tabla
                    cell.classList.add(baseClass);
                    
                    // Clase de estilo específico
                    cell.classList.add(styleClass);
                    
                    // Clases específicas según la posición
                    if (hasHeaders && y === startY) {
                        cell.classList.add('table-header');
                    } else if (y === endY) {
                        cell.classList.add('table-footer');
                    }
                    
                    if (x === startX) {
                        cell.classList.add('table-left');
                    }
                    
                    if (x === endX) {
                        cell.classList.add('table-right');
                    }
                    
                    // Alternar filas para algunas filas
                    if ((y - startY) % 2 === 1 && (!hasHeaders || y > startY)) {
                        cell.classList.add('table-alternate-row');
                    }
                }
            }
        }
        
        // Agregar estilos si no existen
        if (!document.getElementById('table-styles')) {
            const styleElement = document.createElement('style');
            styleElement.id = 'table-styles';
            styleElement.textContent = `
                .excel-table {
                    border: 1px solid #ddd;
                }
                
                .table-header {
                    font-weight: bold;
                    text-align: center;
                }
                
                /* Estilo Azul */
                .table-style-blue.table-header {
                    background-color: #4472C4;
                    color: white;
                }
                
                .table-style-blue.table-alternate-row {
                    background-color: #E9EBF5;
                }
                
                /* Estilo Verde */
                .table-style-green.table-header {
                    background-color: #70AD47;
                    color: white;
                }
                
                .table-style-green.table-alternate-row {
                    background-color: #E2EFDA;
                }
                
                /* Estilo Naranja */
                .table-style-orange.table-header {
                    background-color: #ED7D31;
                    color: white;
                }
                
                .table-style-orange.table-alternate-row {
                    background-color: #FBE5D6;
                }
            `;
            document.head.appendChild(styleElement);
        }
    }

    // Función para borrar tablas cuando se selecciona el rango y se presiona Delete/Backspace
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
            // Verificar si hay celdas seleccionadas con la clase 'selected-range'
            const selectedCells = document.querySelectorAll('td.selected-range');
            
            if (selectedCells.length > 0) {
                // Verificar si alguna de las celdas seleccionadas es parte de una tabla
                let tableCellsFound = false;
                
                selectedCells.forEach(cell => {
                    if (cell.classList.contains('excel-table')) {
                        tableCellsFound = true;
                    }
                });
                
                if (tableCellsFound) {
                    // Si hay celdas de tabla en la selección, eliminar todos los estilos de tabla
                    selectedCells.forEach(cell => {
                        // Eliminar todas las clases relacionadas con tablas
                        cell.classList.remove('excel-table');
                        cell.classList.remove('table-style-blue');
                        cell.classList.remove('table-style-green');
                        cell.classList.remove('table-style-orange');
                        cell.classList.remove('table-header');
                        cell.classList.remove('table-footer');
                        cell.classList.remove('table-left');
                        cell.classList.remove('table-right');
                        cell.classList.remove('table-alternate-row');
                        
                        // También borrar el contenido de la celda
                        const x = parseInt(cell.dataset.x);
                        const y = parseInt(cell.dataset.y);
                        updateCell(x, y, "");
                    });
                    
                    console.log('Tabla eliminada');
                    e.preventDefault();
                }
            }
        }
    });
});

// Funcionalidad para Cuadros de Texto
document.addEventListener('DOMContentLoaded', function() {
    // Ahora usamos el ID específico para una selección más precisa
    const textBoxBtn = document.getElementById('text-box-btn');
    const spreadsheetContainer = document.querySelector('.spreadsheet-container');
    
    // Variable para seguir cuadros de texto activos
    let activeTextBox = null;
    let isCreatingTextBox = false;
    let textBoxDragStart = null;
    let textBoxes = [];
    let nextTextBoxId = 1;
    
    // Si existe el botón de cuadro de texto, añadir evento
    if (textBoxBtn) {
        console.log('Botón de cuadro de texto encontrado', textBoxBtn);
        textBoxBtn.addEventListener('click', () => {
            console.log('Clic en botón de cuadro de texto');
            // Cambiar cursor para indicar modo de inserción de cuadro de texto
            spreadsheetContainer.style.cursor = 'crosshair';
            isCreatingTextBox = true;
        });
    } else {
        console.error('Botón de cuadro de texto no encontrado');
    }
    
    // Evento para crear cuadro de texto al hacer clic en la hoja
    spreadsheetContainer.addEventListener('mousedown', (e) => {
        if (!isCreatingTextBox) return;
        
        // Obtener posición del clic
        const rect = spreadsheetContainer.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        
        // Marcar posición inicial
        textBoxDragStart = { x, y };
        
        // Prevenir selección de celdas durante creación de cuadro de texto
        e.preventDefault();
    });
    
    // Evento para finalizar creación del cuadro de texto
    spreadsheetContainer.addEventListener('mouseup', (e) => {
        if (!isCreatingTextBox || !textBoxDragStart) return;
        
        // Obtener posición final
        const rect = spreadsheetContainer.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        
        // Calcular dimensiones (asegurar tamaño mínimo)
        const width = Math.max(Math.abs(x - textBoxDragStart.x), 100);
        const height = Math.max(Math.abs(y - textBoxDragStart.y), 50);
        
        // Determinar posición superior izquierda
        const left = Math.min(textBoxDragStart.x, x);
        const top = Math.min(textBoxDragStart.y, y);
        
        // Crear el cuadro de texto
        createTextBox(left, top, width, height);
        
        // Restablecer estado
        isCreatingTextBox = false;
        textBoxDragStart = null;
        spreadsheetContainer.style.cursor = 'default';
    });
    
    // Crear un cuadro de texto
    function createTextBox(left, top, width, height) {
        // Crear elemento
        const textBox = document.createElement('div');
        const textBoxId = `text-box-${nextTextBoxId++}`;
        
        textBox.className = 'excel-text-box';
        textBox.id = textBoxId;
        textBox.style.position = 'absolute';
        textBox.style.left = `${left}px`;
        textBox.style.top = `${top}px`;
        // Hacemos el cuadro de texto más grande inicialmente
        textBox.style.width = `${Math.max(width, 250)}px`;
        textBox.style.height = `${Math.max(height, 120)}px`;
        textBox.setAttribute('contenteditable', 'true');
        textBox.dataset.type = 'textbox';
        
        // Prevenir comportamiento no deseado al editar el texto
        textBox.addEventListener('keydown', function(e) {
            // Detener la propagación para evitar que otros manejadores interfieran
            e.stopPropagation();
        });
        
        // Prevenir que el evento input cause comportamientos no deseados
        textBox.addEventListener('input', function(e) {
            e.stopPropagation();
        });
        
        // Manejar clics dentro del cuadro para evitar propagación no deseada
        textBox.addEventListener('click', function(e) {
            e.stopPropagation();
            // Asegurar que el cuadro esté activo cuando hacemos clic dentro
            if (activeTextBox !== textBox) {
                setActiveTextBox(textBox);
            }
        });
        
        // Añadir el cuadro de texto al contenedor
        spreadsheetContainer.appendChild(textBox);
        
        // Añadir a la lista de cuadros de texto
        textBoxes.push(textBox);
        
        // Establecer como activo y darle foco
        setActiveTextBox(textBox);
        textBox.focus();
    }
    
    // Eventos para la selección y movimiento de cuadros de texto
    document.addEventListener('mousedown', (e) => {
        // Verificar si el clic fue en un cuadro de texto o en un redimensionador
        const textBox = e.target.closest('.excel-text-box');
        const resizer = e.target.closest('.resizer');
        
        // Si se hizo clic en el redimensionador, no hacer nada aquí
        if (resizer) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        
        // Si se hizo clic en el cuadro de texto
        if (textBox) {
            // Verificar si el clic fue cerca del borde para selección
            const rect = textBox.getBoundingClientRect();
            const borderWidth = 10; // Ancho sensible del borde para selección
            
            const isNearBorder = 
                e.clientX - rect.left < borderWidth || 
                rect.right - e.clientX < borderWidth ||
                e.clientY - rect.top < borderWidth ||
                rect.bottom - e.clientY < borderWidth;
            
            if (isNearBorder) {
                e.preventDefault();
                setActiveTextBox(textBox);
                
                // Permitir mover el cuadro de texto
                const startPos = {
                    x: e.clientX,
                    y: e.clientY,
                    left: textBox.offsetLeft,
                    top: textBox.offsetTop
                };
                
                function onMouseMove(moveEvent) {
                    moveEvent.preventDefault();
                    textBox.style.left = `${startPos.left + (moveEvent.clientX - startPos.x)}px`;
                    textBox.style.top = `${startPos.top + (moveEvent.clientY - startPos.y)}px`;
                }
                
                function onMouseUp() {
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                }
                
                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            } else {
                // Clic en el interior del cuadro para editar
                setActiveTextBox(textBox);
                // No hacer nada más, dejar que el comportamiento de edición predeterminado ocurra
            }
            return;
        }
        
        // Si se hizo clic fuera de cualquier cuadro de texto
        if (!textBox && !resizer) {
            setActiveTextBox(null);
        }
    });
    
    // Función para establecer un cuadro de texto como activo
    function setActiveTextBox(textBox) {
        // Desactivar el cuadro activo anterior
        if (activeTextBox) {
            activeTextBox.classList.remove('active-text-box');
            
            // Quitar los manejadores de redimensión
            const resizers = activeTextBox.querySelectorAll('.resizer');
            resizers.forEach(resizer => {
                resizer.remove();
            });
        }
        
        // Establecer nuevo activo
        activeTextBox = textBox;
        
        if (textBox) {
            textBox.classList.add('active-text-box');
            
            // Añadir controles de redimensión
            setTimeout(() => {
                addResizers(textBox);
            }, 10);
        }
    }
    
    // Añadir manejadores de redimensión a un cuadro de texto
    function addResizers(textBox) {
        // Primero eliminamos cualquier redimensionador existente
        const existingResizers = textBox.querySelectorAll('.resizer');
        existingResizers.forEach(resizer => resizer.remove());
        
        // Posiciones de los redimensionadores
        const positions = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
        
        // Crear y añadir cada redimensionador
        positions.forEach(pos => {
            const resizer = document.createElement('div');
            resizer.className = `resizer resizer-${pos}`;
            resizer.dataset.position = pos;
            textBox.appendChild(resizer);
            
            // Añadir evento de mousedown específico para cada redimensionador
            resizer.addEventListener('mousedown', function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                // Asegurar que el evento no se propague a otros elementos
                const event = e || window.event;
                event.cancelBubble = true;
                
                // Capturar estado inicial de redimensionamiento
                const startResize = {
                    x: e.clientX,
                    y: e.clientY,
                    width: textBox.offsetWidth,
                    height: textBox.offsetHeight,
                    left: textBox.offsetLeft,
                    top: textBox.offsetTop
                };
                
                // Variables para el seguimiento del ratón
                let isDragging = true;
                
                function onMouseMove(moveEvent) {
                    if (!isDragging) return;
                    
                    moveEvent.preventDefault();
                    const dx = moveEvent.clientX - startResize.x;
                    const dy = moveEvent.clientY - startResize.y;
                    
                    // Aplicar cambios según la posición del redimensionador
                    if (pos.includes('e')) {
                        textBox.style.width = `${Math.max(startResize.width + dx, 120)}px`;
                    }
                    if (pos.includes('s')) {
                        textBox.style.height = `${Math.max(startResize.height + dy, 80)}px`;
                    }
                    if (pos.includes('w')) {
                        const newWidth = Math.max(startResize.width - dx, 120);
                        textBox.style.width = `${newWidth}px`;
                        textBox.style.left = `${startResize.left + (startResize.width - newWidth)}px`;
                    }
                    if (pos.includes('n')) {
                        const newHeight = Math.max(startResize.height - dy, 80);
                        textBox.style.height = `${newHeight}px`;
                        textBox.style.top = `${startResize.top + (startResize.height - newHeight)}px`;
                    }
                }
                
                function onMouseUp() {
                    isDragging = false;
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                    document.body.style.cursor = ''; // Restaurar cursor
                }
                
                // Cambiar cursor durante el redimensionamiento
                document.body.style.cursor = window.getComputedStyle(resizer).cursor;
                
                // Agregar oyentes a nivel de documento para capturar el movimiento incluso si el mouse se mueve fuera del resizador
                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            });
        });
    }
    
    // Eliminar cuadro de texto con Delete
    document.addEventListener('keydown', (e) => {
        // Verificar si hay un cuadro de texto activo y que no está en modo edición
        // (es decir, no tiene el foco de entrada de texto)
        if ((e.key === 'Delete' || e.key === 'Backspace') && 
            activeTextBox && 
            document.activeElement !== activeTextBox) {
            
            e.preventDefault();
            console.log('Eliminando cuadro de texto:', activeTextBox.id);
            
            // Eliminar el cuadro de texto
            activeTextBox.remove();
            
            // Actualizar referencia y lista
            const removedId = activeTextBox.id;
            activeTextBox = null;
            
            // Actualizar lista de cuadros de texto
            textBoxes = textBoxes.filter(box => box.id !== removedId);
        }
    });
});


//Desde aqui





