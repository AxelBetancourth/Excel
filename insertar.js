/* ================================
// Sección insertar (Dayana Espino)
// ================================*/

document.addEventListener("DOMContentLoaded", function () {

    // Lógica para los botones de insertar
    document.querySelectorAll(".insert-btn").forEach(button => {
        button.addEventListener("click", function (event) {
            event.stopPropagation(); // Prevenir que se cierre inmediatamente al hacer clic en el mismo botón
            
            let menuId = this.getAttribute("data-menu");
            let menu = document.getElementById(menuId);

            // Cerrar todos los menús abiertos antes de mostrar el actual
            const allMenus = document.querySelectorAll('.menu');
            allMenus.forEach(m => {
                if (m !== menu && !m.contains(event.target)) {
                    m.classList.remove("show"); // Cerrar otros menús
                }
            });

            // Mostrar u ocultar el menú correspondiente
            if (menu) {
                menu.classList.toggle("show");
            }
        });
    });
    
    // Función para alternar visibilidad de los submenús
    function toggleMenu(menuId) {
        const menu = document.getElementById(menuId);
        if (menu) {
            menu.classList.toggle("show");
        }
    }

    // Manejo de los botones de los submenús (Ilustraciones, Formas, Texto, etc.)
    const btnIlustraciones = document.querySelector('[data-menu="menu-ilustraciones"]');
    const btnFormas = document.getElementById('btn-formas');
    const btnTexto = document.getElementById('btn-texto');
    const btnSimbolos = document.getElementById('btn-simbolos');

    if (btnIlustraciones) {
        btnIlustraciones.addEventListener('click', function (event) {
            event.preventDefault();
            toggleMenu('menu-ilustraciones');
        });
    }
    if (btnFormas) {
        btnFormas.addEventListener('click', function (event) {
            event.preventDefault();
            toggleMenu('submenu-formas');
        });
    }
    if (btnTexto) {
        btnTexto.addEventListener('click', function (event) {
            event.preventDefault();
            toggleMenu('submenu-texto');
        });
    }
    if (btnSimbolos) {
        btnSimbolos.addEventListener('click', function (event) {
            event.preventDefault();
            toggleMenu('submenu-simbolos');
        });
    }

    // Lógica para cargar imagen
    const btnImagenes = document.getElementById('btn-imagenes');
    const imageInput = document.getElementById('image-input');

    if (btnImagenes && imageInput) {
        btnImagenes.addEventListener('click', function () {
            imageInput.click();
        });

        imageInput.addEventListener('change', function (event) {
            const file = event.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function (e) {
                    const imgElement = document.createElement('img');
                    imgElement.src = e.target.result;
                    imgElement.style.maxWidth = '100%';
                    imgElement.style.maxHeight = '100%';
                    imgElement.style.cursor = 'move';
                    imgElement.draggable = true;

                    const imgContainer = document.createElement('div');
                    imgContainer.classList.add('resizable');
                    imgContainer.style.position = 'absolute';
                    imgContainer.style.zIndex = '1000';
                    imgContainer.appendChild(imgElement);

                    const activeCell = document.querySelector('.table-cell.active');
                    if (activeCell) {
                        activeCell.style.position = 'relative';
                        activeCell.appendChild(imgContainer);
                    } else {
                        document.body.appendChild(imgContainer);
                    }

                    // Permitir mover la imagen
                    interact(imgContainer).draggable({
                        onmove: function (event) {
                            const target = event.target;
                            const x = (parseFloat(target.getAttribute('data-x')) || 0) + event.dx;
                            const y = (parseFloat(target.getAttribute('data-y')) || 0) + event.dy;

                            target.style.transform = `translate(${x}px, ${y}px)`;
                            target.setAttribute('data-x', x);
                            target.setAttribute('data-y', y);
                        }
                    });

                    // Permitir redimensionar la imagen
                    interact(imgContainer).resizable({
                        edges: { left: true, right: true, bottom: true, top: true }
                    }).on('resizemove', function (event) {
                        const target = event.target;
                        let x = (parseFloat(target.getAttribute('data-x')) || 0);
                        let y = (parseFloat(target.getAttribute('data-y')) || 0);

                        target.style.width = event.rect.width + 'px';
                        target.style.height = event.rect.height + 'px';

                        x += event.deltaRect.left;
                        y += event.deltaRect.top;

                        target.style.transform = `translate(${x}px, ${y}px)`;
                        target.setAttribute('data-x', x);
                        target.setAttribute('data-y', y);
                    });

                    // Agregar menú contextual
                    imgContainer.addEventListener('contextmenu', function (event) {
                        event.preventDefault();
                        const contextMenu = document.getElementById('context-menu');
                        contextMenu.style.top = `${event.clientY}px`;
                        contextMenu.style.left = `${event.clientX}px`;
                        contextMenu.classList.add('active');
                        contextMenu.currentTarget = imgContainer;
                    });
                }
                reader.readAsDataURL(file);
            }
        });
    }

    // Crear menú contextual dinámicamente
    const contextMenu = document.createElement('div');
    contextMenu.id = 'context-menu';
    contextMenu.classList.add('context-menu');
    contextMenu.innerHTML = `
        <ul>
            <li id="cut-option" class="cut-option">
                <i class="fas fa-cut" style="font-size: 16px; margin-right: 5px;"></i> Cortar
            </li>
        </ul>
    `;
    
    document.body.appendChild(contextMenu);

    // Funcionalidad de cortar imagen
    const cutOption = document.getElementById('cut-option');
    if (cutOption) {
        cutOption.addEventListener('click', function () {
            const contextMenu = document.getElementById('context-menu');
            const target = contextMenu.currentTarget;
            if (target) {
                target.remove();
                imageInput.value = "";
            }
            contextMenu.classList.remove('active');
        });
    }

    // Ocultar menús al hacer clic fuera de ellos
    document.addEventListener('click', function (event) {
        // Si el clic no es dentro de un menú, cerrar todos los menús
        const isClickInsideMenu = event.target.closest('.menu-container') || event.target.closest('.insert-btn');
        if (!isClickInsideMenu) {
            const allMenus = document.querySelectorAll('.menu');
            allMenus.forEach(menu => {
                menu.classList.remove("show");
            });
        }

        // Ocultar menú contextual
        const contextMenu = document.getElementById('context-menu');
        if (contextMenu) {
            contextMenu.classList.remove('active');
        }
    });
});

//FORMAS

function toggleSubmenu() {
    var submenu = document.getElementById("submenu-formas");
    submenu.classList.toggle("show");
}



// Función para insertar las formas en la celda seleccionada
function insertShape(shape) {
    // Obtener la celda seleccionada (suponiendo que el usuario hace clic en una celda editable)
    const selectedCell = document.querySelector('.editable-cell.selected');
    
    if (selectedCell) {
        let shapeElement;
        
        // Crear el elemento de forma según el tipo
        if (shape === 'rect') {
            shapeElement = document.createElement('div');
            shapeElement.classList.add('shape', 'rectangle');
        } else if (shape === 'circle') {
            shapeElement = document.createElement('div');
            shapeElement.classList.add('shape', 'circle');
        } else if (shape === 'triangle') {
            shapeElement = document.createElement('div');
            shapeElement.classList.add('shape', 'triangle');
        }

        // Insertar la forma en la celda
        selectedCell.appendChild(shapeElement);
    }
}

// Agregar la clase 'selected' a las celdas al hacer clic
document.querySelectorAll('.editable-cell').forEach(cell => {
    cell.addEventListener('click', function () {
        // Limpiar la selección anterior
        document.querySelectorAll('.editable-cell').forEach(c => c.classList.remove('selected'));
        // Marcar la celda como seleccionada
        this.classList.add('selected');
    });
});


// TABLAS 

/* ================================
// Funcionalidad para insertar tablas tipo Excel
// ================================*/

document.addEventListener("DOMContentLoaded", function () {
    // Crear el modal de tabla dinámicamente
    createTableModal();
    
    // Crear el menú contextual dinámicamente
    createContextMenu();
    
    const tableModal = document.getElementById('table-modal');
    const createTableBtn = document.getElementById('create-table');
    const cancelTableBtn = document.getElementById('cancel-table');
    const numRowsInput = document.getElementById('num-rows');
    const numColumnsInput = document.getElementById('num-columns');
    const tableStyleSelect = document.getElementById('table-style');
    
    // Abrir el modal de tabla al hacer clic en el botón
    document.getElementById('insert-table-btn').addEventListener('click', function () {
        tableModal.style.display = 'flex';
    });

    // Cerrar el modal al hacer clic en cancelar
    cancelTableBtn.addEventListener('click', function () {
        tableModal.style.display = 'none';
    });

    // Cerrar el modal haciendo clic fuera de él
    window.addEventListener('click', function (event) {
        if (event.target === tableModal) {
            tableModal.style.display = 'none';
        }
    });

    // Crear la tabla al hacer clic en el botón crear
    createTableBtn.addEventListener('click', function () {
        const numRows = parseInt(numRowsInput.value);
        const numColumns = parseInt(numColumnsInput.value);
        const tableStyle = tableStyleSelect.value;

        // Validar entradas
        if (isNaN(numRows) || isNaN(numColumns) || numRows < 1 || numColumns < 1) {
            alert("Por favor, ingresa un número válido de filas y columnas.");
            return;
        }

        // Obtener la celda activa (donde queremos insertar la tabla)
        const activeCell = document.querySelector('.table-cell.active') || document.querySelector('.editable-cell.selected');
        
        if (activeCell) {
            // Crear el contenedor para la tabla anidada
            const nestedTableContainer = document.createElement('div');
            nestedTableContainer.className = 'nested-table-container';
            
            // Crear la tabla
            const table = createExcelTable(numRows, numColumns, tableStyle);
            
            // Agregar la tabla al contenedor
            nestedTableContainer.appendChild(table);
            
            // Limpiar el contenido existente de la celda y agregar la tabla
            activeCell.innerHTML = '';
            activeCell.appendChild(nestedTableContainer);
            
            // Inicializar eventos para la tabla
            initializeTableEvents(table);
        } else {
            // Si no hay celda activa, insertar al final del documento
            const tableContainer = document.createElement('div');
            tableContainer.className = 'table-container';
            
            const table = createExcelTable(numRows, numColumns, tableStyle);
            tableContainer.appendChild(table);
            
            document.body.appendChild(tableContainer);
            
            // Inicializar eventos para la tabla
            initializeTableEvents(table);
        }

        // Cerrar el modal
        tableModal.style.display = 'none';
    });
    
    // Función para crear el modal de tabla dinámicamente
    function createTableModal() {
        const modalHTML = `
            <div id="table-modal" class="table-modal" style="display: none;">
                <div class="table-modal-content">
                    <h3>Insertar tabla</h3>
                    <div class="modal-row">
                        <label for="num-rows">Número de filas:</label>
                        <input type="number" id="num-rows" min="1" max="20" value="3">
                    </div>
                    <div class="modal-row">
                        <label for="num-columns">Número de columnas:</label>
                        <input type="number" id="num-columns" min="1" max="10" value="3">
                    </div>
                    <div class="modal-row">
                        <label for="table-style">Estilo de tabla:</label>
                        <select id="table-style">
                            <option value="table-default">Tabla predeterminada</option>
                            <option value="table-grid">Cuadrícula</option>
                            <option value="table-striped">Filas alternas</option>
                            <option value="table-borderless">Sin bordes</option>
                        </select>
                    </div>
                    <div class="modal-buttons">
                        <button id="cancel-table">Cancelar</button>
                        <button id="create-table">Crear tabla</button>
                    </div>
                </div>
            </div>
        `;
        
        // Crear un elemento temporal para convertir la cadena HTML en elementos DOM
        const temp = document.createElement('div');
        temp.innerHTML = modalHTML;
        
        // Agregar el modal al cuerpo del documento
        document.body.appendChild(temp.firstElementChild);
    }
    
    // Función para crear el menú contextual
    function createContextMenu() {
        const contextMenu = document.createElement('div');
        contextMenu.id = 'table-context-menu';
        contextMenu.className = 'context-menu';
        contextMenu.style.display = 'none';
        contextMenu.style.position = 'absolute';
        contextMenu.style.zIndex = '1000';
        contextMenu.style.backgroundColor = 'white';
        contextMenu.style.border = '1px solid #ccc';
        contextMenu.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
        contextMenu.style.padding = '5px 0';
        
        // Crear la opción para eliminar tabla
        const deleteOption = document.createElement('div');
        deleteOption.className = 'context-menu-item';
        deleteOption.style.padding = '8px 15px';
        deleteOption.style.cursor = 'pointer';
        deleteOption.style.fontSize = '14px';
        deleteOption.innerHTML = '<i class="fas fa-trash" style="margin-right: 8px;"></i>Eliminar tabla';
        deleteOption.addEventListener('click', function() {
            // Eliminar la tabla seleccionada
            if (contextMenu.targetTable) {
                const container = contextMenu.targetTable.closest('.nested-table-container') || 
                                 contextMenu.targetTable.closest('.table-container');
                if (container) {
                    container.remove();
                } else {
                    contextMenu.targetTable.remove();
                }
            }
            // Ocultar menú contextual
            contextMenu.style.display = 'none';
        });
        
        // Agregar hover effect
        deleteOption.addEventListener('mouseover', function() {
            this.style.backgroundColor = '#f2f2f2';
        });
        deleteOption.addEventListener('mouseout', function() {
            this.style.backgroundColor = '';
        });
        
        // Agregar la opción al menú
        contextMenu.appendChild(deleteOption);
        
        // Agregar el menú al documento
        document.body.appendChild(contextMenu);
        
        // Cerrar el menú al hacer clic en cualquier parte
        document.addEventListener('click', function() {
            contextMenu.style.display = 'none';
        });
    }
    
    // Función para crear una tabla estilo Excel
    function createExcelTable(rows, columns, style) {
        const table = document.createElement('table');
        table.className = `excel-table ${style}`;
        
        // Crear cuerpo de la tabla
        const tbody = document.createElement('tbody');
        
        // Crear filas y columnas
        for (let i = 0; i < rows; i++) {
            const row = document.createElement('tr');
            
            for (let j = 0; j < columns; j++) {
                const cell = document.createElement('td');
                cell.contentEditable = 'true';
                cell.className = 'excel-cell';
                cell.setAttribute('data-row', i);
                cell.setAttribute('data-col', j);
                
                // Agregar manijas para redimensionamiento
                const resizeHandle = document.createElement('div');
                resizeHandle.className = 'resize-handle';
                cell.appendChild(resizeHandle);
                
                const resizeHandleRow = document.createElement('div');
                resizeHandleRow.className = 'resize-handle-row';
                cell.appendChild(resizeHandleRow);
                
                row.appendChild(cell);
            }
            
            tbody.appendChild(row);
        }
        
        table.appendChild(tbody);
        
        // Añadir evento de click derecho para mostrar menú contextual
        table.addEventListener('contextmenu', function(e) {
            e.preventDefault();
            
            const contextMenu = document.getElementById('table-context-menu');
            contextMenu.style.display = 'block';
            contextMenu.style.left = e.pageX + 'px';
            contextMenu.style.top = e.pageY + 'px';
            contextMenu.targetTable = this;
            
            // Evitar que el menú se cierre inmediatamente
            e.stopPropagation();
        });
        
        return table;
    }
    
    // Inicializar eventos para las celdas de la tabla
    function initializeTableEvents(table) {
        // Selección de celda
        table.querySelectorAll('td').forEach(cell => {
            // Evento de clic para seleccionar celda
            cell.addEventListener('click', function(e) {
                e.stopPropagation();
                
                // Quitar selección previa
                document.querySelectorAll('.excel-cell').forEach(c => {
                    c.classList.remove('cell-selected');
                });
                
                // Aplicar selección a la celda actual
                this.classList.add('cell-selected');
            });
        });
        
        // Implementar redimensionamiento de columnas
        table.querySelectorAll('.resize-handle').forEach(handle => {
            handle.addEventListener('mousedown', function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                const cell = this.parentElement;
                const initialWidth = cell.offsetWidth;
                const initialX = e.clientX;
                
                function handleMouseMove(moveEvent) {
                    const deltaX = moveEvent.clientX - initialX;
                    cell.style.width = (initialWidth + deltaX) + 'px';
                }
                
                function handleMouseUp() {
                    document.removeEventListener('mousemove', handleMouseMove);
                    document.removeEventListener('mouseup', handleMouseUp);
                }
                
                document.addEventListener('mousemove', handleMouseMove);
                document.addEventListener('mouseup', handleMouseUp);
            });
        });
        
        // Implementar redimensionamiento de filas
        table.querySelectorAll('.resize-handle-row').forEach(handle => {
            handle.addEventListener('mousedown', function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                const cell = this.parentElement;
                const initialHeight = cell.offsetHeight;
                const initialY = e.clientY;
                
                function handleMouseMove(moveEvent) {
                    const deltaY = moveEvent.clientY - initialY;
                    cell.style.height = (initialHeight + deltaY) + 'px';
                }
                
                function handleMouseUp() {
                    document.removeEventListener('mousemove', handleMouseMove);
                    document.removeEventListener('mouseup', handleMouseUp);
                }
                
                document.addEventListener('mousemove', handleMouseMove);
                document.addEventListener('mouseup', handleMouseUp);
            });
        });
    }
    
    // Evento para establecer celda activa
    document.addEventListener('click', function(event) {
        // Verificar si se hizo clic en una celda con clase 'table-cell'
        const cellClicked = event.target.closest('.table-cell');
        
        // Eliminar clase 'active' de todas las celdas
        document.querySelectorAll('.table-cell').forEach(cell => {
            cell.classList.remove('active');
        });
        
        // Si se hizo clic en una celda, marcarla como activa
        if (cellClicked) {
            cellClicked.classList.add('active');
        }
    });
    
    // Configuración para celdas editables
    document.querySelectorAll('.editable-cell').forEach(cell => {
        cell.addEventListener('click', function() {
            // Quitar selección previa
            document.querySelectorAll('.editable-cell').forEach(c => {
                c.classList.remove('selected');
            });
            
            // Marcar celda como seleccionada
            this.classList.add('selected');
        });
    });
});