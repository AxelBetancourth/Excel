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
    const textBoxBtn = document.getElementById('text-box-btn');
    const spreadsheetContainer = document.querySelector('.spreadsheet-container');
    
    let activeTextBox = null;
    let nextTextBoxId = 1;
    let isMoving = false;
    let moveOffset = { x: 0, y: 0 };

    // Función para guardar el estado de un cuadro de texto
    function getTextBoxState(textBoxContainer) {
        return {
            id: textBoxContainer.querySelector('.excel-text-box').id,
            text: textBoxContainer.querySelector('.excel-text-box').value,
            left: textBoxContainer.style.left,
            top: textBoxContainer.style.top,
            width: textBoxContainer.querySelector('.excel-text-box').style.width || '250px',
            height: textBoxContainer.querySelector('.excel-text-box').style.height || '120px'
        };
    }

    // Función para crear un cuadro de texto desde un estado guardado
    function createTextBoxFromState(state) {
        const textBoxContainer = document.createElement('div');
        textBoxContainer.className = 'text-box-container';
        textBoxContainer.style.left = state.left;
        textBoxContainer.style.top = state.top;
        textBoxContainer.style.position = 'absolute';
        textBoxContainer.dataset.sheetIndex = currentSheetIndex;
        
        // Crear el contenedor del textarea y el botón X
        const textBoxWrapper = document.createElement('div');
        textBoxWrapper.className = 'text-box-wrapper';
        
        // Añadir botón X interno
        const closeButton = document.createElement('button');
        closeButton.className = 'text-box-close-btn';
        closeButton.innerHTML = 'x';
        closeButton.title = 'Cerrar cuadro de texto';
        
        closeButton.addEventListener('click', function(e) {
            e.stopPropagation();
            textBoxContainer.remove();
            if (activeTextBox === textBox) {
                activeTextBox = null;
            }
            // Actualizar los cuadros de texto en la hoja actual
            saveTextBoxesToSheet();
        });
        
        const textBox = document.createElement('textarea');
        textBox.className = 'excel-text-box';
        textBox.id = state.id;
        textBox.value = state.text;
        textBox.style.width = state.width;
        textBox.style.height = state.height;
        textBox.placeholder = 'Escriba su texto aquí...';

        // Agregar los mismos event listeners que en createNewTextBox
        addTextBoxEventListeners(textBox, textBoxContainer);

        const deleteButton = createDeleteButton(textBoxContainer, textBox);
        
        textBoxWrapper.appendChild(closeButton);
        textBoxWrapper.appendChild(textBox);
        textBoxContainer.appendChild(textBoxWrapper);
        textBoxContainer.appendChild(deleteButton);
        spreadsheetContainer.appendChild(textBoxContainer);
    }

    // Función para crear el botón de eliminar
    function createDeleteButton(container, textBox) {
        const deleteButton = document.createElement('button');
        deleteButton.className = 'delete-text-box';
        deleteButton.title = 'Eliminar cuadro de texto';
        
        deleteButton.addEventListener('click', function(e) {
            e.stopPropagation();
            container.remove();
            if (activeTextBox === textBox) {
                activeTextBox = null;
            }
            // Actualizar los cuadros de texto en la hoja actual
            saveTextBoxesToSheet();
        });

        return deleteButton;
    }

    // Función para agregar event listeners a un cuadro de texto
    function addTextBoxEventListeners(textBox, container) {
        let isDragging = false;
        let currentX;
        let currentY;
        let initialX;
        let initialY;

        textBox.addEventListener('copy', function(e) {
            e.stopPropagation();
        });
        
        textBox.addEventListener('paste', function(e) {
            e.stopPropagation();
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('text');
            const startPos = this.selectionStart;
            const endPos = this.selectionEnd;
            this.value = this.value.substring(0, startPos) + text + this.value.substring(endPos);
            this.selectionStart = this.selectionEnd = startPos + text.length;
            // Guardar el estado después de pegar
            saveTextBoxesToSheet();
        });

        textBox.addEventListener('keydown', function(e) {
            e.stopPropagation();
            if (e.key === 'Delete' || e.key === 'Backspace') {
                e.stopPropagation();
            }
        });

        textBox.addEventListener('input', function() {
            // Guardar el estado cuando se modifica el texto
            saveTextBoxesToSheet();
        });

        function handleDragStart(e) {
            if (e.target === textBox && e.target.selectionStart === 0) {
                isDragging = true;
                container.classList.add('moving');
                initialX = e.clientX - container.offsetLeft;
                initialY = e.clientY - container.offsetTop;
                setActiveTextBox(textBox);
            }
        }

        function handleDragEnd() {
            if (isDragging) {
                isDragging = false;
                container.classList.remove('moving');
                // Guardar el estado después de mover
                saveTextBoxesToSheet();
            }
        }

        function handleDrag(e) {
            if (isDragging) {
                e.preventDefault();
                currentX = e.clientX - initialX;
                currentY = e.clientY - initialY;

                const rect = spreadsheetContainer.getBoundingClientRect();
                const maxX = rect.width - container.offsetWidth;
                const maxY = rect.height - container.offsetHeight;

                currentX = Math.min(Math.max(0, currentX), maxX);
                currentY = Math.min(Math.max(0, currentY), maxY);

                container.style.left = `${currentX}px`;
                container.style.top = `${currentY}px`;
            }
        }

        textBox.addEventListener('mousedown', handleDragStart);
        document.addEventListener('mousemove', handleDrag);
        document.addEventListener('mouseup', handleDragEnd);
        
        textBox.addEventListener('click', function(e) {
            e.stopPropagation();
        });
    }
    
    // Función para guardar los cuadros de texto en la hoja actual
    function saveTextBoxesToSheet() {
        if (!sheets[currentSheetIndex].textBoxes) {
            sheets[currentSheetIndex].textBoxes = [];
        }
        
        const textBoxes = document.querySelectorAll(`.text-box-container[data-sheet-index="${currentSheetIndex}"]`);
        sheets[currentSheetIndex].textBoxes = Array.from(textBoxes).map(getTextBoxState);
    }

    // Función para cargar los cuadros de texto de una hoja
    function loadTextBoxesFromSheet() {
        // Ocultar todos los cuadros de texto existentes
        document.querySelectorAll('.text-box-container').forEach(container => {
            if (parseInt(container.dataset.sheetIndex) === currentSheetIndex) {
                container.style.display = 'block';
            } else {
                container.style.display = 'none';
            }
        });
        
        // Si la hoja actual no tiene cuadros de texto cargados pero tiene datos guardados, cargarlos
        if (sheets[currentSheetIndex].textBoxes && 
            !document.querySelector(`.text-box-container[data-sheet-index="${currentSheetIndex}"]`)) {
            sheets[currentSheetIndex].textBoxes.forEach(createTextBoxFromState);
        }
    }

    // Función para crear un nuevo cuadro de texto
    function createNewTextBox(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const rect = spreadsheetContainer.getBoundingClientRect();
        const centerX = rect.width / 2 - 125;
        const centerY = rect.height / 2 - 60;
        
        const textBoxContainer = document.createElement('div');
        textBoxContainer.className = 'text-box-container';
        textBoxContainer.style.left = `${centerX}px`;
        textBoxContainer.style.top = `${centerY}px`;
        textBoxContainer.style.position = 'absolute';
        textBoxContainer.dataset.sheetIndex = currentSheetIndex;
        
        // Crear el contenedor del textarea y el botón X
        const textBoxWrapper = document.createElement('div');
        textBoxWrapper.className = 'text-box-wrapper';
        
        // Añadir botón X interno
        const closeButton = document.createElement('button');
        closeButton.className = 'text-box-close-btn';
        closeButton.innerHTML = 'x';
        closeButton.title = 'Cerrar cuadro de texto';
        
        closeButton.addEventListener('click', function(e) {
            e.stopPropagation();
            textBoxContainer.remove();
            if (activeTextBox === textBox) {
                activeTextBox = null;
            }
            // Actualizar los cuadros de texto en la hoja actual
            saveTextBoxesToSheet();
        });
        
        const textBox = document.createElement('textarea');
        const textBoxId = `text-box-${nextTextBoxId++}`;
        
        textBox.className = 'excel-text-box';
        textBox.id = textBoxId;
        textBox.placeholder = 'Escriba su texto aquí...';

        addTextBoxEventListeners(textBox, textBoxContainer);
        
        const deleteButton = createDeleteButton(textBoxContainer, textBox);
        
        textBoxWrapper.appendChild(closeButton);
        textBoxWrapper.appendChild(textBox);
        textBoxContainer.appendChild(textBoxWrapper);
        textBoxContainer.appendChild(deleteButton);
        spreadsheetContainer.appendChild(textBoxContainer);
        
        setActiveTextBox(textBox);
        textBox.focus();
        
        // Guardar el nuevo cuadro de texto en la hoja actual
        saveTextBoxesToSheet();
        
        document.querySelectorAll('.menu.open').forEach(menu => {
            menu.classList.remove('open');
        });
    }

    // Configurar el botón de texto
    if (textBoxBtn) {
        textBoxBtn.addEventListener('click', createNewTextBox);
    }
    
    function setActiveTextBox(textBox) {
        if (activeTextBox) {
            activeTextBox.classList.remove('active-text-box');
        }
        
        activeTextBox = textBox;
        
        if (textBox) {
            textBox.classList.add('active-text-box');
        }
    }
    
    // Manejar la eliminación de cuadros de texto con teclas
    document.addEventListener('keydown', function(e) {
        if ((e.key === 'Delete' || e.key === 'Backspace') && 
            activeTextBox && 
            document.activeElement !== activeTextBox) {
            const container = activeTextBox.parentElement;
            container.remove();
            activeTextBox = null;
            // Actualizar los cuadros de texto en la hoja actual
            saveTextBoxesToSheet();
        }
    });
    
    // Deseleccionar cuadro de texto al hacer clic fuera
    document.addEventListener('mousedown', function(e) {
        if (!e.target.closest('.text-box-container')) {
            setActiveTextBox(null);
        }
    });

    // Escuchar el evento de cambio de hoja
    document.addEventListener('sheetChanged', function() {
        loadTextBoxesFromSheet();
    });

    // También cargar cuadros de texto al inicio
    // Verificamos si está definida la variable global sheets
    if (typeof sheets !== 'undefined' && sheets.length > 0) {
        // Asegurarse de que la hoja inicial tenga la propiedad textBoxes
        if (!sheets[currentSheetIndex].textBoxes) {
            sheets[currentSheetIndex].textBoxes = [];
        }
        loadTextBoxesFromSheet();
    }
});


//Desde aqui





