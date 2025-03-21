/* ================================
// Sección insertar (Dayana Espino)
// ================================*/


document.addEventListener("DOMContentLoaded", function () {

    
    // ================================================================
    //  FUNCIONES PARA ILUSTRACIONES
    // ================================================================

    // Lógica para los botones de insertar
     document.querySelectorAll(".insert-btn").forEach(button => {
        button.addEventListener("click", function (event) {
            event.stopPropagation(); // Prevenir que se cierre inmediatamente al hacer clic en el mismo botón
            
            let menuId = this.getAttribute("data-menu");
            let menu = document.getElementById(menuId);

            // Cerrar todos los menús abiertos antes de mostrar el actual
            const allMenus = document.querySelectorAll('.menu');
            allMenus.forEach(m => {
                if (m !== menu) {
                    m.classList.remove("show"); // Cerrar otros menús
                }
            });

            // Mostrar u ocultar el menú correspondiente
            if (menu) {
                menu.classList.toggle("show");
            }
        });
    });

    // Cerrar el menú desplegable al hacer clic fuera de él
    document.addEventListener("click", function (event) {
        const allMenus = document.querySelectorAll('.menu');
        allMenus.forEach(menu => {
            if (!menu.contains(event.target)) {
                menu.classList.remove("show"); // Cerrar el menú si el clic fue fuera de él
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

    $body.addEventListener('mouseup', () => {
        isSelecting = false;
        startCell = null;
      });
      
      // Función para limpiar las selecciones
      function clearSelections() {
        document.querySelectorAll('.selected-cell').forEach(cell => cell.classList.remove('selected-cell'));
      }
      
      // Función para actualizar la celda activa
      function updateActiveCellDisplay(x, y) {
        activeCell = { x, y };
        const columnLetter = letras[x];
        const rowNumber = y + 1;
        activeCellDisplay.textContent = `${columnLetter}${rowNumber}`;
        formulaInput.value = state[x][y].value;
      }
      
      //Insertar y eliminar imagenes y forma 
      
      document.getElementById('btn-imagenes').addEventListener('click', function() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.style.display = 'none';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (file) {
                handleImageUpload(file);
            }
        };
        
        document.body.appendChild(input);
        input.click();
        document.body.removeChild(input);
    });
    
  
      // Evento para manejar la carga de la imagen
      function handleImageUpload(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = document.createElement('img');
            img.src = e.target.result;
            img.style.position = 'absolute';
            img.style.maxWidth = '200px';
            img.style.height = 'auto';
            img.style.top = '50px';
            img.style.left = '50px';
            img.style.cursor = 'move';
            img.classList.add('draggable');
    
            const spreadsheetContainer = document.querySelector('.spreadsheet-container');
            if (spreadsheetContainer) {
                spreadsheetContainer.appendChild(img);
                makeImageInteractive(img);
            } else {
                alert('No se encontró el contenedor de la hoja de cálculo.');
            }
        };
        reader.readAsDataURL(file);
    }
      
      function toggleDropdown(id) {
        const dropdown = document.getElementById(id);
        dropdown.classList.toggle('show');
      }
      
      function insertShape(shape) {
        const shapeElement = document.createElement('div');
        shapeElement.style.position = 'absolute';
        shapeElement.style.top = '50px';
        shapeElement.style.left = '50px';
        shapeElement.style.width = '100px';
        shapeElement.style.height = '100px';
        shapeElement.style.zIndex = '10'; // Asegura que la forma esté encima de las celdas
      
        switch (shape) {
          case 'rect':
            shapeElement.style.backgroundColor = 'blue';
            break;
          case 'circle':
            shapeElement.style.backgroundColor = 'green';
            shapeElement.style.borderRadius = '50%';
            break;
          case 'triangle':
            shapeElement.style.width = '0';
            shapeElement.style.height = '0';
            shapeElement.style.borderLeft = '50px solid transparent';
            shapeElement.style.borderRight = '50px solid transparent';
            shapeElement.style.borderBottom = '100px solid red';
            break;
        }
      
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        if (spreadsheetContainer) {
          spreadsheetContainer.appendChild(shapeElement);
          makeShapeInteractive(shapeElement);
        } else {
          alert('No se encontró el contenedor de la hoja de cálculo.');
        }
      }
      
      function makeImageInteractive(img) {
        interact(img)
          .draggable({
            onmove: dragMoveListener
          })
          .resizable({
            edges: { left: true, right: true, bottom: true, top: true }
          })
          .on('resizemove', function (event) {
            let { x, y } = event.target.dataset;
      
            x = (parseFloat(x) || 0) + event.deltaRect.left;
            y = (parseFloat(y) || 0) + event.deltaRect.top;
      
            Object.assign(event.target.style, {
              width: `${event.rect.width}px`,
              height: `${event.rect.height}px`,
              transform: `translate(${x}px, ${y}px)`
            });
      
            Object.assign(event.target.dataset, { x, y });
          });
      
        img.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          showContextMenu(event, img);
        });
      }
      
      function makeShapeInteractive(shape) {
        interact(shape)
          .draggable({
            onmove: dragMoveListener
          })
          .resizable({
            edges: { left: true, right: true, bottom: true, top: true }
          })
          .on('resizemove', function (event) {
            let { x, y } = event.target.dataset;
      
            x = (parseFloat(x) || 0) + event.deltaRect.left;
            y = (parseFloat(y) || 0) + event.deltaRect.top;
      
            Object.assign(event.target.style, {
              width: `${event.rect.width}px`,
              height: `${event.rect.height}px`,
              transform: `translate(${x}px, ${y}px)`
            });
      
            Object.assign(event.target.dataset, { x, y });
          });
      
        shape.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          showContextMenu(event, shape);
        });
      }
      
      function dragMoveListener(event) {
        const target = event.target;
        const x = (parseFloat(target.dataset.x) || 0) + event.dx;
        const y = (parseFloat(target.dataset.y) || 0) + event.dy;
      
        target.style.transform = `translate(${x}px, ${y}px)`;
        target.dataset.x = x;
        target.dataset.y = y;
      }
      
      let copiedElement = null;
      
      function showContextMenu(event, element = null) {
        // Eliminar cualquier menú contextual existente
        const existingMenu = document.querySelector('.context-menu');
        if (existingMenu) {
          existingMenu.remove();
        }
      
        // Crear el menú contextual
        const contextMenu = document.createElement('div');
        contextMenu.className = 'context-menu';
      
        // Obtener el contenedor principal
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        const containerRect = spreadsheetContainer.getBoundingClientRect();
      
        // Calcular la posición del menú contextual
        const top = event.clientY - containerRect.top;
        const left = event.clientX - containerRect.left;
      
        contextMenu.style.top = `${top}px`;
        contextMenu.style.left = `${left}px`;
        contextMenu.style.position = 'absolute';
        contextMenu.style.display = 'block';
        contextMenu.style.zIndex = '1000'; // Asegura que el menú esté encima de otros elementos
        contextMenu.style.backgroundColor = 'white'; // Fondo blanco
        contextMenu.style.border = '1px solid #ccc'; // Borde gris claro
        contextMenu.style.padding = '10px'; // Espaciado interno
        contextMenu.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)'; // Sombra para darle un efecto de elevación
        contextMenu.style.minWidth = '150px'; // Asegura un ancho mínimo para el menú
      
        let optionsHTML = `
          <ul style="list-style: none; margin: 0; padding: 0;">
        `;
      
        if (element) {
          optionsHTML += `
            <li id="cut-element-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              ✂️ Cortar
            </li>
            <li id="copy-element-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              📋 Copiar
            </li>
            <li id="format-element-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              🖌️ Formato de imagen
            </li>
          `;
        } else if (copiedElement) {
          optionsHTML += `
            <li id="paste-element-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              📋 Pegar
            </li>
          `;
        }
      
        optionsHTML += `</ul>`;
        contextMenu.innerHTML = optionsHTML;
      
        spreadsheetContainer.appendChild(contextMenu);
      
        if (element) {
          document.getElementById('cut-element-btn').onclick = () => {
            // Implementar la funcionalidad de cortar
            element.remove();
            contextMenu.remove();
          };
      
          document.getElementById('copy-element-btn').onclick = () => {
            // Implementar la funcionalidad de copiar
            copiedElement = element.cloneNode(true);
            contextMenu.remove();
          };
      
          document.getElementById('format-element-btn').onclick = () => {
            // Implementar la funcionalidad de formato de imagen
            contextMenu.remove();
          };
        } else if (copiedElement) {
          document.getElementById('paste-element-btn').onclick = (event) => {
            pasteElement(event);
            contextMenu.remove();
          };
        }
      
        document.addEventListener('click', () => {
          contextMenu.remove();
        }, { once: true });
      }
      
      // Función para pegar el elemento copiado
      function pasteElement(event) {
        if (copiedElement) {
          const spreadsheetContainer = document.querySelector('.spreadsheet-container');
          const containerRect = spreadsheetContainer.getBoundingClientRect();
      
          const top = event.clientY - containerRect.top;
          const left = event.clientX - containerRect.left;
      
          const newElement = copiedElement.cloneNode(true);
          newElement.style.top = `${top}px`;
          newElement.style.left = `${left}px`;
      
          spreadsheetContainer.appendChild(newElement);
      
          if (newElement.tagName.toLowerCase() === 'img') {
            makeImageInteractive(newElement);
          } else {
            makeShapeInteractive(newElement);
          }
      
          copiedElement = null;
        }
      }
      
      // Evento para mostrar el menú contextual con la opción de pegar
      document.addEventListener('contextmenu', (event) => {
        if (!event.target.closest('.draggable')) {
          event.preventDefault();
          showContextMenu(event);
        }
      });
      
      // Evento para pegar con Ctrl + V
      document.addEventListener('keydown', (event) => {
        if (event.ctrlKey && event.key === 'v') {
          const fakeEvent = {
            clientY: window.innerHeight / 2,
            clientX: window.innerWidth / 2
          };
          pasteElement(fakeEvent);
        }
      });
      
      // Función para hacer las imágenes interactivas
      function makeImageInteractive(img) {
        interact(img)
          .draggable({
            onmove: dragMoveListener
          })
          .resizable({
            edges: { left: true, right: true, bottom: true, top: true }
          })
          .on('resizemove', function (event) {
            let { x, y } = event.target.dataset;
      
            x = (parseFloat(x) || 0) + event.deltaRect.left;
            y = (parseFloat(y) || 0) + event.deltaRect.top;
      
            Object.assign(event.target.style, {
              width: `${event.rect.width}px`,
              height: `${event.rect.height}px`,
              transform: `translate(${x}px, ${y}px)`
            });
      
            Object.assign(event.target.dataset, { x, y });
          });
      
        img.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          showImageContextMenu(event, img);
        });
      }
      
      function showImageContextMenu(event, img) {
        // Eliminar cualquier menú contextual existente
        const existingMenu = document.querySelector('.context-menu');
        if (existingMenu) {
          existingMenu.remove();
        }
      
        // Crear el menú contextual
        const contextMenu = document.createElement('div');
        contextMenu.className = 'context-menu';
      
        // Obtener el contenedor principal
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        const containerRect = spreadsheetContainer.getBoundingClientRect();
      
        // Calcular la posición del menú contextual
        const top = event.clientY - containerRect.top;
        const left = event.clientX - containerRect.left;
      
        contextMenu.style.top = `${top}px`;
        contextMenu.style.left = `${left}px`;
        contextMenu.style.position = 'absolute';
        contextMenu.style.display = 'block';
        contextMenu.style.zIndex = '1000'; // Asegura que el menú esté encima de otros elementos
        contextMenu.style.backgroundColor = 'white'; // Fondo blanco
        contextMenu.style.border = '1px solid #ccc'; // Borde gris claro
        contextMenu.style.padding = '10px'; // Espaciado interno
        contextMenu.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)'; // Sombra para darle un efecto de elevación
        contextMenu.style.minWidth = '150px'; // Asegura un ancho mínimo para el menú
      
        let optionsHTML = `
          <ul style="list-style: none; margin: 0; padding: 0;">
            <li id="cut-image-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              ✂️ Cortar
            </li>
            <li id="copy-image-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              📋 Copiar
            </li>
            <li id="format-image-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              🖌️ Formato de imagen
            </li>
          </ul>
        `;
        contextMenu.innerHTML = optionsHTML;
      
        spreadsheetContainer.appendChild(contextMenu);
      
        document.getElementById('cut-image-btn').onclick = () => {
          img.remove();
          contextMenu.remove();
        };
      
        document.getElementById('copy-image-btn').onclick = () => {
          copiedElement = img.cloneNode(true);
          contextMenu.remove();
        };
      
        document.getElementById('format-image-btn').onclick = () => {
          // Implementar la funcionalidad de formato de imagen
          contextMenu.remove();
        };
      
        document.addEventListener('click', () => {
          contextMenu.remove();
        }, { once: true });
      }
      
      // Codigo para formas
      
      function makeShapeInteractive(shape) {
        interact(shape)
          .draggable({
            onmove: dragMoveListener
          })
          .resizable({
            edges: { left: true, right: true, bottom: true, top: true }
          })
          .on('resizemove', function (event) {
            let { x, y } = event.target.dataset;
      
            x = (parseFloat(x) || 0) + event.deltaRect.left;
            y = (parseFloat(y) || 0) + event.deltaRect.top;
      
            Object.assign(event.target.style, {
              width: `${event.rect.width}px`,
              height: `${event.rect.height}px`,
              transform: `translate(${x}px, ${y}px)`
            });
      
            Object.assign(event.target.dataset, { x, y });
          });
      
        shape.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          showShapeContextMenu(event, shape);
        });
      }
      
      function showShapeContextMenu(event, shape) {
        // Eliminar cualquier menú contextual existente
        const existingMenu = document.querySelector('.context-menu');
        if (existingMenu) {
          existingMenu.remove();
        }
      
        // Crear el menú contextual
        const contextMenu = document.createElement('div');
        contextMenu.className = 'context-menu';
      
        // Obtener el contenedor principal
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        const containerRect = spreadsheetContainer.getBoundingClientRect();
      
        // Calcular la posición del menú contextual
        const top = event.clientY - containerRect.top;
        const left = event.clientX - containerRect.left;
      
        contextMenu.style.top = `${top}px`;
        contextMenu.style.left = `${left}px`;
        contextMenu.style.position = 'absolute';
        contextMenu.style.display = 'block';
        contextMenu.style.zIndex = '1000'; // Asegura que el menú esté encima de otros elementos
        contextMenu.style.backgroundColor = 'white'; // Fondo blanco
        contextMenu.style.border = '1px solid #ccc'; // Borde gris claro
        contextMenu.style.padding = '10px'; // Espaciado interno
        contextMenu.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)'; // Sombra para darle un efecto de elevación
        contextMenu.style.minWidth = '150px'; // Asegura un ancho mínimo para el menú
      
        let optionsHTML = `
          <ul style="list-style: none; margin: 0; padding: 0;">
            <li id="cut-shape-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              ✂️ Cortar
            </li>
            <li id="copy-shape-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              📋 Copiar
            </li>
            <li id="format-shape-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              🖌️ Formato de imagen
            </li>
          </ul>
        `;
        contextMenu.innerHTML = optionsHTML;
      
        spreadsheetContainer.appendChild(contextMenu);
      
        document.getElementById('cut-shape-btn').onclick = () => {
          shape.remove();
          contextMenu.remove();
        };
      
        document.getElementById('copy-shape-btn').onclick = () => {
          copiedElement = shape.cloneNode(true);
          contextMenu.remove();
        };
      
        document.getElementById('format-shape-btn').onclick = () => {
          // Implementar la funcionalidad de formato de imagen
          contextMenu.remove();
        };
      
        document.addEventListener('click', () => {
          contextMenu.remove();
        }, { once: true });
      }
      
      // Función para manejar el movimiento de arrastre
      function dragMoveListener(event) {
        const target = event.target;
        const x = (parseFloat(target.dataset.x) || 0) + event.dx;
        const y = (parseFloat(target.dataset.y) || 0) + event.dy;
      
        target.style.transform = `translate(${x}px, ${y}px)`;
        target.dataset.x = x;
        target.dataset.y = y;
      }
      
      // Función para pegar el elemento copiado
      function pasteElement(event) {
        if (copiedElement) {
          const spreadsheetContainer = document.querySelector('.spreadsheet-container');
          const containerRect = spreadsheetContainer.getBoundingClientRect();
      
          const top = event.clientY - containerRect.top;
          const left = event.clientX - containerRect.left;
      
          const newElement = copiedElement.cloneNode(true);
          newElement.style.top = `${top}px`;
          newElement.style.left = `${left}px`;
      
          spreadsheetContainer.appendChild(newElement);
      
          if (newElement.tagName.toLowerCase() === 'img') {
            makeImageInteractive(newElement);
          } else {
            makeShapeInteractive(newElement);
          }
      
          copiedElement = null;
        }
      }
      
      // Evento para mostrar el menú contextual con la opción de pegar
      document.addEventListener('contextmenu', (event) => {
        if (!event.target.closest('.draggable')) {
          event.preventDefault();
          showContextMenu(event);
        }
      });
      
      // Evento para pegar con Ctrl + V
      document.addEventListener('keydown', (event) => {
        if (event.ctrlKey && event.key === 'v') {
          const fakeEvent = {
            clientY: window.innerHeight / 2,
            clientX: window.innerWidth / 2
          };
          pasteElement(fakeEvent);
        }
      });
      
      function showContextMenu(event) {
        // Eliminar cualquier menú contextual existente
        const existingMenu = document.querySelector('.context-menu');
        if (existingMenu) {
          existingMenu.remove();
        }
      
        // Crear el menú contextual
        const contextMenu = document.createElement('div');
        contextMenu.className = 'context-menu';
      
        // Obtener el contenedor principal
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        const containerRect = spreadsheetContainer.getBoundingClientRect();
      
        // Calcular la posición del menú contextual
        const top = event.clientY - containerRect.top;
        const left = event.clientX - containerRect.left;
      
        contextMenu.style.top = `${top}px`;
        contextMenu.style.left = `${left}px`;
        contextMenu.style.position = 'absolute';
        contextMenu.style.display = 'block';
        contextMenu.style.zIndex = '1000'; // Asegura que el menú esté encima de otros elementos
        contextMenu.style.backgroundColor = 'white'; // Fondo blanco
        contextMenu.style.border = '1px solid #ccc'; // Borde gris claro
        contextMenu.style.padding = '10px'; // Espaciado interno
        contextMenu.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)'; // Sombra para darle un efecto de elevación
        contextMenu.style.minWidth = '150px'; // Asegura un ancho mínimo para el menú
      
        let optionsHTML = `
          <ul style="list-style: none; margin: 0; padding: 0;">
            <li id="paste-element-btn" class="context-option" style="color: black; cursor: pointer; padding: 5px 10px;">
              📋 Pegar
            </li>
          </ul>
        `;
        contextMenu.innerHTML = optionsHTML;
      
        spreadsheetContainer.appendChild(contextMenu);
      
        document.getElementById('paste-element-btn').onclick = (event) => {
          pasteElement(event);
          contextMenu.remove();
        };
      
        document.addEventListener('click', () => {
          contextMenu.remove();
        }, { once: true });
      }
      



    // ================================================================
    // FIN DEFUNCIONES PARA ILUSTRACIONES
    // ================================================================



    // ================================================================
    // FUNCIONES PARA TABLAS DINAMICAS
    // ================================================================



    // ================================================================
    // FIN DE FUNCIONES PARA TABLAS DINAMICAS
    // ================================================================




    // ================================================================
    // FUNCIONES PARA TABLAS DINAMICAS RECOMENDADAS
    // ================================================================



    // ================================================================
    // FIN DE FUNCIONES PARA TABLAS DINAMICAS RECOMENDADAS
    // ================================================================



    
    // ================================================================
    // FUNCIONES PARA TABLAS
    // ================================================================



// Obtener referencia al botón de tabla
const insertTableBtn = document.getElementById('insert-table-btn');

// Agregar evento de clic al botón de inserción de tabla
insertTableBtn.addEventListener('click', createTable);

// Variable para almacenar el panel de opciones de tabla
let tableOptionsPanel = null;

// Función para crear una tabla en las celdas seleccionadas
function createTable() {
  // Comprobar si hay una selección de celdas
  if (!activeCell) {
    alert("Por favor, selecciona una celda para insertar la tabla");
    return;
  }
  
  // Crear un modal para configurar la tabla
  const modal = document.createElement('div');
  modal.className = 'table-modal';
  modal.innerHTML = `
    <div class="table-modal-content">
      <h3>Insertar Tabla</h3>
      <div class="modal-row">
        <label for="table-rows">Filas:</label>
        <input type="number" id="table-rows" min="1" max="20" value="3">
      </div>
      <div class="modal-row">
        <label for="table-cols">Columnas:</label>
        <input type="number" id="table-cols" min="1" max="10" value="3">
      </div>
      <div class="modal-row">
        <label for="table-style">Estilo:</label>
        <select id="table-style">
          <option value="simple">Simple</option>
          <option value="striped">Con rayas</option>
          <option value="bordered">Bordes completos</option>
          <option value="excel-blue">Excel Azul</option>
        </select>
      </div>
      <div class="modal-buttons">
        <button id="cancel-table">Cancelar</button>
        <button id="create-table">Crear</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  // Manejar eventos de botones del modal
  document.getElementById('cancel-table').addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  document.getElementById('create-table').addEventListener('click', () => {
    const rows = parseInt(document.getElementById('table-rows').value);
    const cols = parseInt(document.getElementById('table-cols').value);
    const style = document.getElementById('table-style').value;
    
    if (rows > 0 && cols > 0) {
      generateTable(activeCell.x, activeCell.y, rows, cols, style);
      document.body.removeChild(modal);
    }
  });
}

// Función para generar la tabla en las celdas
function generateTable(startX, startY, rows, cols, style = 'simple') {
  // Verificar que hay suficientes celdas disponibles
  if (startX + cols > state.length || startY + rows > state[0].length) {
    alert("No hay suficientes celdas disponibles para esta tabla");
    return;
  }
  
  const tableId = `table_${Date.now()}`;
  
  // Crear borde alrededor de todas las celdas de la tabla
  for (let y = 0; y < rows; y++) {
    for (let x = 0; x < cols; x++) {
      const cellX = startX + x;
      const cellY = startY + y;
      
      // Aplicar estilo de borde a la celda
      const selector = `td[data-x="${cellX}"][data-y="${cellY}"]`;
      const cell = document.querySelector(selector);
      
      if (cell) {
        cell.classList.add('table-cell');
        cell.dataset.tableId = tableId;
        cell.dataset.tableRow = y;
        cell.dataset.tableCol = x;
        
        // Aplicar estilos según la posición y el estilo seleccionado
        if (y === 0) cell.classList.add('table-top');
        if (y === rows - 1) cell.classList.add('table-bottom');
        if (x === 0) cell.classList.add('table-left');
        if (x === cols - 1) cell.classList.add('table-right');
        
        // Aplicar estilo seleccionado
        if (style) {
          cell.classList.add(`table-style-${style}`);
          
          // Si es estilo con rayas, alternar colores
          if (style === 'striped' && y % 2 === 1) {
            cell.classList.add('table-row-alt');
          }
          
          // Si es el estilo Excel, hacer la primera fila en negrita con fondo azul
          if (style === 'excel-blue') {
            if (y === 0) {
              cell.classList.add('table-header-excel');
            }
          }
        }
        
        // Guardar el valor actual
        const currentValue = state[cellX][cellY].value;
        
        // Si está vacía y es la primera fila, añadir texto de columna
        if (y === 0 && currentValue === "") {
          updateCell(cellX, cellY, `Columna ${x + 1}`);
        }
        
        // Si es la primera columna y está vacía, añadir el número de fila (excepto en el encabezado)
        if (x === 0 && y > 0 && currentValue === "") {
          updateCell(cellX, cellY, `Fila ${y}`);
        }
      }
    }
  }
  
  // Registrar la tabla creada para referencia futura
  const tableInfo = {
    startX,
    startY,
    rows,
    cols,
    id: tableId,
    style
  };
  
  // Si no existe, crear un array para almacenar las tablas
  if (!window.excelTables) window.excelTables = [];
  window.excelTables.push(tableInfo);
  
  // Guardar datos en localStorage
  saveTableData();
}

// Función para guardar los datos de tablas
function saveTableData() {
  if (window.excelTables) {
    localStorage.setItem('excelTables', JSON.stringify(window.excelTables));
  }
}

// Función para cargar los datos de tablas guardados
function loadTableData() {
  const savedTables = localStorage.getItem('excelTables');
  if (savedTables) {
    window.excelTables = JSON.parse(savedTables);
    
    // Restaurar el estilo de las tablas guardadas
    window.excelTables.forEach(table => {
      for (let y = 0; y < table.rows; y++) {
        for (let x = 0; x < table.cols; x++) {
          const cellX = table.startX + x;
          const cellY = table.startY + y;
          
          const selector = `td[data-x="${cellX}"][data-y="${cellY}"]`;
          const cell = document.querySelector(selector);
          
          if (cell) {
            cell.classList.add('table-cell');
            cell.dataset.tableId = table.id;
            cell.dataset.tableRow = y;
            cell.dataset.tableCol = x;
            
            if (y === 0) cell.classList.add('table-top');
            if (y === table.rows - 1) cell.classList.add('table-bottom');
            if (x === 0) cell.classList.add('table-left');
            if (x === table.cols - 1) cell.classList.add('table-right');
            
            // Aplicar estilo guardado
            if (table.style) {
              cell.classList.add(`table-style-${table.style}`);
              
              if (table.style === 'striped' && y % 2 === 1) {
                cell.classList.add('table-row-alt');
              }
              
              if (table.style === 'excel-blue') {
                if (y === 0) {
                  cell.classList.add('table-header-excel');
                }
              }
            }
          }
        }
      }
    });
  }
}

// Función para mostrar panel de opciones de tabla
function showTableOptions(x, y) {
  // Primero eliminamos cualquier panel existente
  hideTableOptions();
  
  // Buscar si la celda pertenece a una tabla
  const tableId = findTableIdForCell(x, y);
  if (!tableId) return;
  
  // Obtener información de la tabla
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Crear el panel de opciones
  tableOptionsPanel = document.createElement('div');
  tableOptionsPanel.className = 'table-options-panel';
  
  // Obtener coordenadas de celda dentro de la tabla
  const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
  const tableRow = parseInt(cell.dataset.tableRow);
  const tableCol = parseInt(cell.dataset.tableCol);
  
  // Preparar opciones según la posición de la celda
  let optionsHTML = `
    <div class="table-options-title">Opciones de tabla</div>
    <button id="format-table-btn">Formato de tabla</button>
    <button id="delete-table-btn">Eliminar tabla</button>
  `;
  
  // Añadir opciones específicas para filas/columnas
  optionsHTML += `
    <div class="table-options-divider"></div>
    <button id="insert-row-above-btn">Insertar fila arriba</button>
    <button id="insert-row-below-btn">Insertar fila abajo</button>
    <button id="insert-col-left-btn">Insertar columna izquierda</button>
    <button id="insert-col-right-btn">Insertar columna derecha</button>
    <div class="table-options-divider"></div>
    <button id="delete-row-btn">Eliminar fila</button>
    <button id="delete-col-btn">Eliminar columna</button>
  `;
  
  tableOptionsPanel.innerHTML = optionsHTML;
  
  // Posicionar el panel cerca de la celda activa
  if (cell) {
    const rect = cell.getBoundingClientRect();
    tableOptionsPanel.style.top = rect.bottom + 'px';
    tableOptionsPanel.style.left = rect.left + 'px';
    document.body.appendChild(tableOptionsPanel);
    
    // Añadir eventos
    document.getElementById('delete-table-btn').addEventListener('click', () => {
      if (confirm('¿Estás seguro de que deseas eliminar esta tabla?')) {
        deleteTable(tableId);
        hideTableOptions();
      }
    });
    
    document.getElementById('format-table-btn').addEventListener('click', () => {
      showTableFormatOptions(tableId);
      hideTableOptions();
    });
    
    // Eventos para insertar/eliminar filas y columnas
    document.getElementById('insert-row-above-btn').addEventListener('click', () => {
      insertTableRow(tableId, tableRow, 'above');
      hideTableOptions();
    });
    
    document.getElementById('insert-row-below-btn').addEventListener('click', () => {
      insertTableRow(tableId, tableRow, 'below');
      hideTableOptions();
    });
    
    document.getElementById('insert-col-left-btn').addEventListener('click', () => {
      insertTableColumn(tableId, tableCol, 'left');
      hideTableOptions();
    });
    
    document.getElementById('insert-col-right-btn').addEventListener('click', () => {
      insertTableColumn(tableId, tableCol, 'right');
      hideTableOptions();
    });
    
    document.getElementById('delete-row-btn').addEventListener('click', () => {
      if (tableInfo.rows > 1) {
        deleteTableRow(tableId, tableRow);
        hideTableOptions();
      } else {
        alert('No se puede eliminar la última fila de la tabla.');
      }
    });
    
    document.getElementById('delete-col-btn').addEventListener('click', () => {
      if (tableInfo.cols > 1) {
        deleteTableColumn(tableId, tableCol);
        hideTableOptions();
      } else {
        alert('No se puede eliminar la última columna de la tabla.');
      }
    });
  }
}

// Función para mostrar opciones de formato de tabla
function showTableFormatOptions(tableId) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Crear modal para opciones de formato
  const modal = document.createElement('div');
  modal.className = 'table-modal';
  modal.innerHTML = `
    <div class="table-modal-content">
      <h3>Formato de Tabla</h3>
      <div class="modal-row">
        <label for="table-style-select">Estilo:</label>
        <select id="table-style-select">
          <option value="simple" ${tableInfo.style === 'simple' ? 'selected' : ''}>Simple</option>
          <option value="striped" ${tableInfo.style === 'striped' ? 'selected' : ''}>Con rayas</option>
          <option value="bordered" ${tableInfo.style === 'bordered' ? 'selected' : ''}>Bordes completos</option>
          <option value="excel-blue" ${tableInfo.style === 'excel-blue' ? 'selected' : ''}>Excel Azul</option>
        </select>
      </div>
      <div class="modal-row">
        <label for="header-row-checkbox">Primera fila como encabezado:</label>
        <input type="checkbox" id="header-row-checkbox" ${tableInfo.hasHeaderRow ? 'checked' : ''}>
      </div>
      <div class="modal-buttons">
        <button id="cancel-format">Cancelar</button>
        <button id="apply-format">Aplicar</button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
  
  document.getElementById('cancel-format').addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  document.getElementById('apply-format').addEventListener('click', () => {
    const newStyle = document.getElementById('table-style-select').value;
    const hasHeaderRow = document.getElementById('header-row-checkbox').checked;
    
    // Actualizar estilo de la tabla
    reformatTable(tableId, newStyle, hasHeaderRow);
    document.body.removeChild(modal);
  });
}

// Función para reformatear una tabla existente
function reformatTable(tableId, newStyle, hasHeaderRow) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Actualizar información de la tabla
  tableInfo.style = newStyle;
  tableInfo.hasHeaderRow = hasHeaderRow;
  
  // Eliminar estilos anteriores y aplicar nuevos
  for (let y = 0; y < tableInfo.rows; y++) {
    for (let x = 0; x < tableInfo.cols; x++) {
      const cellX = tableInfo.startX + x;
      const cellY = tableInfo.startY + y;
      
      const selector = `td[data-x="${cellX}"][data-y="${cellY}"]`;
      const cell = document.querySelector(selector);
      
      if (cell) {
        // Eliminar clases de estilo antiguas
        cell.classList.remove('table-style-simple', 'table-style-striped', 'table-style-bordered', 'table-style-excel-blue');
        cell.classList.remove('table-row-alt', 'table-header-excel');
        
        // Aplicar nuevo estilo
        cell.classList.add(`table-style-${newStyle}`);
        
        if (newStyle === 'striped' && y % 2 === 1) {
          cell.classList.add('table-row-alt');
        }
        
        if (newStyle === 'excel-blue' && y === 0 && hasHeaderRow) {
          cell.classList.add('table-header-excel');
        }
      }
    }
  }
  
  // Guardar cambios
  saveTableData();
}

// Función para insertar una fila en la tabla
function insertTableRow(tableId, rowIndex, position) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  const insertAt = position === 'above' ? rowIndex : rowIndex + 1;
  
  // Mover todas las celdas por debajo hacia abajo
  for (let y = tableInfo.rows - 1; y >= insertAt; y--) {
    for (let x = 0; x < tableInfo.cols; x++) {
      const fromX = tableInfo.startX + x;
      const fromY = tableInfo.startY + y;
      const toY = tableInfo.startY + y + 1;
      
      // Comprobar si existe celda de destino
      if (toY < state[0].length) {
        // Mover valor y estilo
        const fromCell = document.querySelector(`td[data-x="${fromX}"][data-y="${fromY}"]`);
        const toCell = document.querySelector(`td[data-x="${fromX}"][data-y="${toY}"]`);
        
        if (fromCell && toCell) {
          // Copiar valor
          updateCell(fromX, toY, state[fromX][fromY].value);
          
          // Actualizar atributos
          toCell.dataset.tableId = tableId;
          toCell.dataset.tableRow = (parseInt(fromCell.dataset.tableRow) + 1).toString();
          toCell.dataset.tableCol = fromCell.dataset.tableCol;
          
          // Copiar clases
          toCell.className = fromCell.className;
        }
      }
    }
  }
  
  // Añadir la nueva fila
  for (let x = 0; x < tableInfo.cols; x++) {
    const cellX = tableInfo.startX + x;
    const cellY = tableInfo.startY + insertAt;
    
    const cell = document.querySelector(`td[data-x="${cellX}"][data-y="${cellY}"]`);
    
    if (cell) {
      // Limpiar celda
      updateCell(cellX, cellY, "");
      
      // Aplicar estilo
      cell.className = 'table-cell';
      cell.classList.add(`table-style-${tableInfo.style || 'simple'}`);
      
      // Si es estilo con rayas y es fila impar
      if (tableInfo.style === 'striped' && insertAt % 2 === 1) {
        cell.classList.add('table-row-alt');
      }
      
      // Actualizar atributos
      cell.dataset.tableId = tableId;
      cell.dataset.tableRow = insertAt.toString();
      cell.dataset.tableCol = x.toString();
      
      // Aplicar bordes laterales si es necesario
      if (x === 0) cell.classList.add('table-left');
      if (x === tableInfo.cols - 1) cell.classList.add('table-right');
    }
  }
  
  // Actualizar tamaño de la tabla
  tableInfo.rows++;
  
  // Actualizar estilos de bordes
  updateTableBorders(tableId);
  
  // Guardar cambios
  saveTableData();
}

// Función para insertar una columna en la tabla
function insertTableColumn(tableId, colIndex, position) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  const insertAt = position === 'left' ? colIndex : colIndex + 1;
  
  // Mover todas las celdas a la derecha
  for (let x = tableInfo.cols - 1; x >= insertAt; x--) {
    for (let y = 0; y < tableInfo.rows; y++) {
      const fromX = tableInfo.startX + x;
      const fromY = tableInfo.startY + y;
      const toX = tableInfo.startX + x + 1;
      
      // Comprobar si existe celda de destino
      if (toX < state.length) {
        // Mover valor y estilo
        const fromCell = document.querySelector(`td[data-x="${fromX}"][data-y="${fromY}"]`);
        const toCell = document.querySelector(`td[data-x="${toX}"][data-y="${fromY}"]`);
        
        if (fromCell && toCell) {
          // Copiar valor
          updateCell(toX, fromY, state[fromX][fromY].value);
          
          // Actualizar atributos
          toCell.dataset.tableId = tableId;
          toCell.dataset.tableRow = fromCell.dataset.tableRow;
          toCell.dataset.tableCol = (parseInt(fromCell.dataset.tableCol) + 1).toString();
          
          // Copiar clases
          toCell.className = fromCell.className;
        }
      }
    }
  }
  
  // Añadir la nueva columna
  for (let y = 0; y < tableInfo.rows; y++) {
    const cellX = tableInfo.startX + insertAt;
    const cellY = tableInfo.startY + y;
    
    const cell = document.querySelector(`td[data-x="${cellX}"][data-y="${cellY}"]`);
    
    if (cell) {
      // Limpiar celda
      updateCell(cellX, cellY, "");
      
      // Aplicar estilo
      cell.className = 'table-cell';
      cell.classList.add(`table-style-${tableInfo.style || 'simple'}`);
      
      // Si es estilo con rayas y es fila impar
      if (tableInfo.style === 'striped' && y % 2 === 1) {
        cell.classList.add('table-row-alt');
      }
      
      // Si es la primera fila y tiene estilo Excel
      if (tableInfo.style === 'excel-blue' && y === 0 && tableInfo.hasHeaderRow) {
        cell.classList.add('table-header-excel');
      }
      
      // Actualizar atributos
      cell.dataset.tableId = tableId;
      cell.dataset.tableRow = y.toString();
      cell.dataset.tableCol = insertAt.toString();
      
      // Aplicar bordes superior e inferior si es necesario
      if (y === 0) cell.classList.add('table-top');
      if (y === tableInfo.rows - 1) cell.classList.add('table-bottom');
    }
  }
  
  // Actualizar tamaño de la tabla
  tableInfo.cols++;
  
  // Actualizar estilos de bordes
  updateTableBorders(tableId);
  
  // Guardar cambios
  saveTableData();
}

// Función para eliminar una fila de la tabla
function deleteTableRow(tableId, rowIndex) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Eliminar fila y mover todas las filas de abajo hacia arriba
  for (let y = rowIndex; y < tableInfo.rows - 1; y++) {
    for (let x = 0; x < tableInfo.cols; x++) {
      const toX = tableInfo.startX + x;
      const toY = tableInfo.startY + y;
      const fromY = tableInfo.startY + y + 1;
      
      const toCell = document.querySelector(`td[data-x="${toX}"][data-y="${toY}"]`);
      const fromCell = document.querySelector(`td[data-x="${toX}"][data-y="${fromY}"]`);
      
      if (toCell && fromCell) {
        // Copiar valor
        updateCell(toX, toY, state[toX][fromY].value);
        
        // Actualizar atributos
        toCell.dataset.tableRow = (parseInt(fromCell.dataset.tableRow) - 1).toString();
        
        // Copiar clases excepto bordes que se actualizarán después
        toCell.className = fromCell.className.replace(/table-(top|bottom|left|right)/g, '');
      }
    }
  }
  
  // Limpiar la última fila
  const lastRowY = tableInfo.startY + tableInfo.rows - 1;
  for (let x = 0; x < tableInfo.cols; x++) {
    const cellX = tableInfo.startX + x;
    const cell = document.querySelector(`td[data-x="${cellX}"][data-y="${lastRowY}"]`);
    
    if (cell) {
      // Limpiar celda
      updateCell(cellX, lastRowY, "");
      
      // Quitar estilos de tabla
      cell.className = '';
      
      // Quitar atributos de tabla
      delete cell.dataset.tableId;
      delete cell.dataset.tableRow;
      delete cell.dataset.tableCol;
    }
  }
  
  // Actualizar tamaño de tabla
  tableInfo.rows--;
  
  // Actualizar estilos de bordes
  updateTableBorders(tableId);
  
  // Guardar cambios
  saveTableData();
}

// Función para eliminar una columna de la tabla
function deleteTableColumn(tableId, colIndex) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Eliminar columna y mover todas las columnas de la derecha hacia la izquierda
  for (let x = colIndex; x < tableInfo.cols - 1; x++) {
    for (let y = 0; y < tableInfo.rows; y++) {
      const toX = tableInfo.startX + x;
      const toY = tableInfo.startY + y;
      const fromX = tableInfo.startX + x + 1;
      
      const toCell = document.querySelector(`td[data-x="${toX}"][data-y="${toY}"]`);
      const fromCell = document.querySelector(`td[data-x="${fromX}"][data-y="${toY}"]`);
      
      if (toCell && fromCell) {
        // Copiar valor
        updateCell(toX, toY, state[fromX][toY].value);
        
        // Actualizar atributos
        toCell.dataset.tableCol = (parseInt(fromCell.dataset.tableCol) - 1).toString();
        
        // Copiar clases excepto bordes que se actualizarán después
        toCell.className = fromCell.className.replace(/table-(top|bottom|left|right)/g, '');
      }
    }
  }
  
  // Limpiar la última columna
  const lastColX = tableInfo.startX + tableInfo.cols - 1;
  for (let y = 0; y < tableInfo.rows; y++) {
    const cellY = tableInfo.startY + y;
    const cell = document.querySelector(`td[data-x="${lastColX}"][data-y="${cellY}"]`);
    
    if (cell) {
      // Limpiar celda
      updateCell(lastColX, cellY, "");
      
      // Quitar estilos de tabla
      cell.className = '';
      
      // Quitar atributos de tabla
      delete cell.dataset.tableId;
      delete cell.dataset.tableRow;
      delete cell.dataset.tableCol;
    }
  }
  
  // Actualizar tamaño de tabla
  tableInfo.cols--;
  
  // Actualizar estilos de bordes
  updateTableBorders(tableId);
  
  // Guardar cambios
  saveTableData();
}

// Función para actualizar los bordes después de modificar la tabla
function updateTableBorders(tableId) {
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  // Quitar todos los bordes primero
  for (let y = 0; y < tableInfo.rows; y++) {
    for (let x = 0; x < tableInfo.cols; x++) {
      const cellX = tableInfo.startX + x;
      const cellY = tableInfo.startY + y;
      
      const cell = document.querySelector(`td[data-x="${cellX}"][data-y="${cellY}"]`);
      
      if (cell) {
        cell.classList.remove('table-top', 'table-bottom', 'table-left', 'table-right');
      }
    }
  }
  
  // Aplicar bordes actualizados
  for (let y = 0; y < tableInfo.rows; y++) {
    for (let x = 0; x < tableInfo.cols; x++) {
      const cellX = tableInfo.startX + x;
      const cellY = tableInfo.startY + y;
      
      const cell = document.querySelector(`td[data-x="${cellX}"][data-y="${cellY}"]`);
      
      if (cell) {
        // Añadir bordes según posición
        if (y === 0) cell.classList.add('table-top');
        if (y === tableInfo.rows - 1) cell.classList.add('table-bottom');
        if (x === 0) cell.classList.add('table-left');
        if (x === tableInfo.cols - 1) cell.classList.add('table-right');
      }
    }
  }
}

// Función para ocultar el panel de opciones
function hideTableOptions() {
  if (tableOptionsPanel && tableOptionsPanel.parentNode) {
    tableOptionsPanel.parentNode.removeChild(tableOptionsPanel);
    tableOptionsPanel = null;
  }
}

// Función para encontrar el ID de tabla para una celda dada
function findTableIdForCell(x, y) {
  const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
  return cell ? cell.dataset.tableId : null;
}

// Función para encontrar la información de una tabla por su ID
function findTableInfoById(tableId) {
  if (!window.excelTables) return null;
  return window.excelTables.find(table => table.id === tableId);
}

// Función para eliminar una tabla
function deleteTable(tableId) {
  if (!window.excelTables) return;
  
  const tableInfo = findTableInfoById(tableId);
  if (!tableInfo) return;
  
  const { startX, startY, rows, cols } = tableInfo;
  
  // Eliminar estilos y datos de las celdas
  for (let y = 0; y < rows; y++) {
    for (let x = 0; x < cols; x++) {
      const cellX = startX + x;
      const cellY = startY + y;
      
      const selector = `td[data-x="${cellX}"][data-y="${cellY}"]`;
      const cell = document.querySelector(selector);
      
      if (cell) {
        // Eliminar todas las clases asociadas a la celda
        cell.classList.remove('table-cell', 'table-top', 'table-bottom', 'table-left', 'table-right');
        cell.classList.remove('table-style-simple', 'table-style-striped', 'table-style-bordered', 'table-style-excel-blue');
        cell.classList.remove('table-row-alt', 'table-header-excel');
        
        // Eliminar el atributo data-tableId si existe
        delete cell.dataset.tableId;
        delete cell.dataset.tableRow;
        delete cell.dataset.tableCol;

        // Opcional: Limpiar cualquier contenido adicional dentro de la celda (si es necesario)
        cell.innerHTML = ''; // Esto vacía cualquier contenido en la celda, si es necesario

        // También podrías restaurar otros atributos si es necesario, por ejemplo:
        // cell.style = '';  // Para borrar cualquier estilo en línea
      }
    }
  }

  // Eliminar la tabla del array de tablas
  window.excelTables = window.excelTables.filter(table => table.id !== tableId);

  // Guardar cambios
  saveTableData();
}

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

$body.addEventListener('contextmenu', event => {
  event.preventDefault();
  const td = event.target.closest('td');
  if (td) {
    clearSelections();
    $$('.celdailumida').forEach(celda => celda.classList.remove('celdailumida'));
    td.classList.add('celdailumida');
    const x = parseInt(td.dataset.x);
    const y = parseInt(td.dataset.y);
    updateActiveCellDisplay(x, y);
    showTableOptions(x, y); // Mostrar opciones de tabla al hacer clic derecho
  }
});

let isSelecting = false;
let startCell = null;

$body.addEventListener('mousedown', event => {
  const td = event.target.closest('td');
  if (td) {
    isSelecting = true;
    startCell = td;
    clearSelections();
    td.classList.add('selected-cell');
  }
});

$body.addEventListener('mousemove', event => {
  if (isSelecting) {
    const td = event.target.closest('td');
    if (td && startCell) {
      clearSelections();
      const startX = parseInt(startCell.dataset.x);
      const startY = parseInt(startCell.dataset.y);
      const endX = parseInt(td.dataset.x);
      const endY = parseInt(td.dataset.y);

      const minX = Math.min(startX, endX);
      const maxX = Math.max(startX, endX);
      const minY = Math.min(startY, endY);
      const maxY = Math.max(startY, endY);

      for (let x = minX; x <= maxX; x++) {
        for (let y = minY; y <= maxY; y++) {
          const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
          if (cell) {
            cell.classList.add('selected-cell');
          }
        }
      }
    }
  }
});

$body.addEventListener('mouseup', () => {
  isSelecting = false;
  startCell = null;
});

// Función para limpiar las selecciones
function clearSelections() {
  document.querySelectorAll('.selected-cell').forEach(cell => cell.classList.remove('selected-cell'));
}

// Función para actualizar la celda activa
function updateActiveCellDisplay(x, y) {
  activeCell = { x, y };
  const columnLetter = letras[x];
  const rowNumber = y + 1;
  activeCellDisplay.textContent = `${columnLetter}${rowNumber}`;
  formulaInput.value = state[x][y].value;
}

let activeCell = null; // Variable para almacenar la celda activa

// Evento para manejar clic en las celdas
document.querySelectorAll('td').forEach(cell => {
    cell.addEventListener('click', function () {
        // Limpiar la selección previa
        document.querySelectorAll('.active-cell').forEach(c => c.classList.remove('active-cell'));

        // Marcar la celda actual como activa
        this.classList.add('active-cell');
        activeCell = this;

        // Mostrar la referencia de la celda activa (por ejemplo, "A1")
        const col = this.dataset.x; // Columna
        const row = this.dataset.y; // Fila
        console.log(`Celda activa: ${col}${row}`);
    });
});



    // ================================================================
    // FIN DEFUNCIONES PARA TABLAS
    // ================================================================




    // ================================================================
    // FUNCIONES PARA GRÁFICOS RECOMENDADOS
    // ================================================================




    // ================================================================
    // FIN DEFUNCIONES PARA GRÁFICOS RECOMENDADOS
    // ================================================================



// ================================================================
// FUNCIONES PARA GRÁFICOS
// ================================================================


   // Referencias a elementos DOM
   const chartModal = document.getElementById('chart-modal');
   const createChartBtn = document.getElementById('create-chart');
   const closeModalBtn = document.querySelector('.chart-modal .close-button');
   const cancelChartBtn = document.getElementById('cancel-chart');
   const previewCanvas = document.getElementById('preview-canvas');

   // Función para mostrar el modal
   function showChartModal(chartType) {
    if (!chartModal) return;

    chartModal.style.display = 'block';
    chartModal.dataset.chartType = chartType;

    // Limpiar campos
    document.getElementById('chart-title').value = '';
    document.getElementById('data-range').value = '';

    // Actualizar título del modal
    const modalTitle = chartModal.querySelector('.chart-modal-header h2');
    if (modalTitle) {
        modalTitle.textContent = `Configurar Gráfico de ${getChartTypeLabel(chartType)}`;
    }

    // Generar vista previa después de un pequeño retraso para asegurar que el modal esté visible
    setTimeout(() => {
        generatePreview(chartType);
    }, 100);
}

   // Función para generar vista previa
   function generatePreview(chartType) {
    const previewContainer = document.querySelector('.preview-container');
    
    // Limpiar el contenedor
    previewContainer.innerHTML = '<label>Vista previa:</label>';
    
    // Crear nuevo canvas
    const canvas = document.createElement('canvas');
    canvas.style.width = '100%';
    canvas.style.height = '200px';
    previewContainer.appendChild(canvas);

    // Datos de ejemplo para la vista previa
    const previewData = {
        labels: ['A', 'B', 'C', 'D'],
        datasets: [{
            label: 'Vista previa',
            data: [12, 19, 3, 5],
            backgroundColor: [
                'rgba(255, 99, 132, 0.2)',
                'rgba(54, 162, 235, 0.2)',
                'rgba(255, 206, 86, 0.2)',
                'rgba(75, 192, 192, 0.2)'
            ],
            borderColor: [
                'rgba(255, 99, 132, 1)',
                'rgba(54, 162, 235, 1)',
                'rgba(255, 206, 86, 1)',
                'rgba(75, 192, 192, 1)'
            ],
            borderWidth: 1
        }]
    };

    // Crear el gráfico de vista previa
    new Chart(canvas.getContext('2d'), {
        type: chartType,
        data: previewData,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 500
            },
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}

// Función para obtener datos reales de las celdas
function obtenerDatosDeCeldas(rango) {
    try {
        const [inicio, fin] = rango.split(':');
        const inicioCol = inicio.match(/[A-Z]+/)[0];
        const inicioFila = parseInt(inicio.match(/\d+/)[0]);
        const finCol = fin.match(/[A-Z]+/)[0];
        const finFila = parseInt(fin.match(/\d+/)[0]);

        const datos = [];
        const etiquetas = [];
        const tabla = document.querySelector('#spreadsheet tbody');

        // Convertir letras a índices (A=0, B=1, etc.)
        const colInicio = inicioCol.charCodeAt(0) - 65;
        const colFin = finCol.charCodeAt(0) - 65;

        // Generar etiquetas para las columnas (primera fila del rango)
        const filaEtiquetas = tabla.rows[inicioFila - 1];
        for (let col = colInicio; col <= colFin; col++) {
            const celda = filaEtiquetas?.cells[col];
            etiquetas.push(celda?.textContent.trim() || `Columna ${col - colInicio + 1}`);
        }

        // Recolectar datos de las filas posteriores a la primera fila
        for (let fila = inicioFila + 1; fila <= finFila; fila++) {
            const filaActual = tabla.rows[fila - 1];
            if (filaActual) {
                const filaDatos = [];
                for (let col = colInicio; col <= colFin; col++) {
                    const celda = filaActual.cells[col];
                    const valor = parseFloat(celda?.textContent.trim());
                    filaDatos.push(!isNaN(valor) ? valor : null); // Usar `null` para celdas vacías
                }
                datos.push(filaDatos);
            }
        }

        // Validar si se obtuvieron etiquetas y datos
        if (etiquetas.length === 0 || datos.length === 0) {
            throw new Error('No se encontraron datos o etiquetas válidas en el rango especificado.');
        }

        return { datos, etiquetas };
    } catch (error) {
        console.error('Error al obtener datos:', error);
        return { datos: [], etiquetas: [] };
    }
}


function validarRango(rango) {
    const rangoRegex = /^[A-Z]+\d+:[A-Z]+\d+$/;
    return rangoRegex.test(rango);
}

// Evento para crear el gráfico
createChartBtn?.addEventListener('click', function () {
    const chartType = chartModal.dataset.chartType;
    const titulo = document.getElementById('chart-title').value || 'Gráfico';
    const rangoDatos = document.getElementById('data-range').value;

    if (!validarRango(rangoDatos)) {
        alert('Por favor, introduce un rango de datos válido en el formato A1:B2.');
        return;
    }

    const { datos, etiquetas } = obtenerDatosDeCeldas(rangoDatos);

    if (datos.length === 0 || etiquetas.length === 0) {
        alert('No se encontraron datos válidos en el rango especificado.');
        return;
    }

    insertarGraficoEncima(chartType, rangoDatos, titulo);
    chartModal.style.display = 'none';
});

// Función para insertar el gráfico dentro del rango seleccionado
function insertarGraficoEncima(chartType, rangoDatos, titulo) {
    const { datos, etiquetas } = obtenerDatosDeCeldas(rangoDatos);

    // Validar que se obtuvieron datos válidos
    if (datos.length === 0 || etiquetas.length === 0) {
        alert('No se encontraron datos válidos en el rango especificado.');
        return;
    }

    // Crear el contenedor flotante del gráfico
    const chartContainer = document.createElement('div');
    chartContainer.className = 'chart-container-flotante';
    chartContainer.style.position = 'absolute';
    chartContainer.style.top = '100px'; // Posición inicial
    chartContainer.style.left = '100px'; // Posición inicial
    chartContainer.style.width = '500px'; // Ancho inicial
    chartContainer.style.height = '300px'; // Alto inicial
    chartContainer.style.zIndex = '1000'; // Asegura que esté encima de otros elementos
    chartContainer.style.backgroundColor = 'white'; // Fondo blanco
    chartContainer.style.border = '1px solid #ccc'; // Borde gris claro
    chartContainer.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)'; // Sombra para darle un efecto de elevación
    chartContainer.style.resize = 'both'; // Permitir redimensionar
    chartContainer.style.overflow = 'hidden'; // Ocultar contenido desbordado

    // Crear el canvas para el gráfico
    const canvas = document.createElement('canvas');
    chartContainer.appendChild(canvas);

    // Insertar el contenedor en el cuerpo del documento
    document.body.appendChild(chartContainer);

    // Generar colores para cada serie (columna)
    const colores = [
        'rgba(255, 99, 132, 0.2)',
        'rgba(54, 162, 235, 0.2)',
        'rgba(255, 206, 86, 0.2)',
        'rgba(75, 192, 192, 0.2)',
        'rgba(153, 102, 255, 0.2)',
        'rgba(255, 159, 64, 0.2)'
    ];
    const bordes = [
        'rgba(255, 99, 132, 1)',
        'rgba(54, 162, 235, 1)',
        'rgba(255, 206, 86, 1)',
        'rgba(75, 192, 192, 1)',
        'rgba(153, 102, 255, 1)',
        'rgba(255, 159, 64, 1)'
    ];

    // Crear datasets para cada columna (serie)
    const datasets = datos.map((serie, index) => ({
        label: `Columna ${index + 1}`,
        data: serie,
        backgroundColor: colores[index % colores.length],
        borderColor: bordes[index % bordes.length],
        borderWidth: 1
    }));

    // Crear el gráfico usando Chart.js
    new Chart(canvas, {
        type: chartType,
        data: {
            labels: etiquetas, // Etiquetas para las filas
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: titulo
                }
            }
        }
    });

    // Hacer el contenedor del gráfico interactivo (arrastrable y redimensionable)
    interact(chartContainer)
        .draggable({
            onmove: dragMoveListener
        })
        .resizable({
            edges: { left: true, right: true, bottom: true, top: true }
        })
        .on('resizemove', function (event) {
            const { width, height } = event.rect;
            event.target.style.width = `${width}px`;
            event.target.style.height = `${height}px`;

            // Redimensionar el canvas dentro del contenedor
            const canvas = event.target.querySelector('canvas');
            if (canvas) {
                canvas.width = width;
                canvas.height = height;
            }
        });

    // Agregar evento para mostrar el menú contextual
    chartContainer.addEventListener('contextmenu', (event) => {
        showChartContextMenu(event, chartContainer);
    });
}



// Función para manejar el movimiento de arrastre
function dragMoveListener(event) {
    const target = event.target;
    const x = (parseFloat(target.dataset.x) || 0) + event.dx;
    const y = (parseFloat(target.dataset.y) || 0) + event.dy;

    target.style.transform = `translate(${x}px, ${y}px)`;
    target.dataset.x = x;
    target.dataset.y = y;
}


   // Función para crear el gráfico en la hoja
   function createChart(chartType, rangoDatos, titulo) {
       const tbody = document.querySelector('#spreadsheet tbody');
       if (!tbody) return;

       // Crear elementos
       const nuevaFila = document.createElement('tr');
       const celda = document.createElement('td');
       const chartContainer = document.createElement('div');
       const canvas = document.createElement('canvas');

       // Configurar contenedor
       chartContainer.className = 'chart-container';
       celda.colSpan = 26;
       
       // Ensamblar elementos
       chartContainer.appendChild(canvas);
       celda.appendChild(chartContainer);
       nuevaFila.appendChild(celda);
       tbody.appendChild(nuevaFila);

       // Obtener datos reales
       const { datos, etiquetas } = obtenerDatosDeCeldas(rangoDatos);

       // Crear gráfico
       new Chart(canvas, {
           type: chartType,
           data: {
               labels: etiquetas,
               datasets: [{
                   label: titulo,
                   data: datos,
                   backgroundColor: 'rgba(54, 162, 235, 0.2)',
                   borderColor: 'rgba(54, 162, 235, 1)',
                   borderWidth: 1
               }]
           },
           options: {
               responsive: true,
               maintainAspectRatio: false,
               plugins: {
                   title: {
                       display: true,
                       text: titulo,
                       font: {
                           size: 16
                       }
                   }
               }
           }
       });
   }

   // Event Listeners
   document.getElementById('btn-grafico-barras')?.addEventListener('click', () => showChartModal('bar'));
   document.getElementById('btn-grafico-lineas')?.addEventListener('click', () => showChartModal('line'));
   document.getElementById('btn-grafico-circular')?.addEventListener('click', () => showChartModal('pie'));
   document.getElementById('btn-grafico-dispersion')?.addEventListener('click', () => showChartModal('scatter'));


   closeModalBtn?.addEventListener('click', () => chartModal.style.display = 'none');
   cancelChartBtn?.addEventListener('click', () => chartModal.style.display = 'none');


   function showChartContextMenu(event, chartContainer) {
    // Eliminar cualquier menú contextual existente
    const existingMenu = document.querySelector('.chart-context-menu');
    if (existingMenu) {
        existingMenu.remove();
    }

    // Crear el menú contextual
    const contextMenu = document.createElement('div');
    contextMenu.className = 'chart-context-menu';
    contextMenu.style.position = 'absolute';
    contextMenu.style.top = `${event.clientY}px`;
    contextMenu.style.left = `${event.clientX}px`;
    contextMenu.style.backgroundColor = 'white';
    contextMenu.style.border = '1px solid #ccc';
    contextMenu.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)';
    contextMenu.style.padding = '10px';
    contextMenu.style.zIndex = '1000';
    contextMenu.style.minWidth = '150px';

    // Opciones del menú
    contextMenu.innerHTML = `
        <ul style="list-style: none; margin: 0; padding: 0;">
            <li id="cut-chart-btn" style="padding: 5px; cursor: pointer;">✂️ Cortar</li>
        </ul>
    `;

    // Agregar el menú al documento
    document.body.appendChild(contextMenu);

    // Manejar la opción de cortar
    document.getElementById('cut-chart-btn').addEventListener('click', () => {
        chartContainer.remove(); // Eliminar el gráfico
        contextMenu.remove(); // Eliminar el menú contextual
    });

    // Eliminar el menú contextual al hacer clic fuera de él
    document.addEventListener('click', () => {
        contextMenu.remove();
    }, { once: true });

    // Prevenir el menú contextual predeterminado
    event.preventDefault();
}


// ================================================================
// FIN DE FUNCIONES PARA GRÁFICOS
// ================================================================


    //=================================================================
    // CELDAS ACTIVAS
    //=================================================================

    

   

    // FIN CELDAS
    
 


    // ================================================================
    // FUNCIONES PARA MAPAS
    // ================================================================




    // ================================================================
    // FIN DEFUNCIONES PARA MAPAS
    // ================================================================




    // ================================================================
    // FUNCIONES PARA SEGMENTACION DE DATOS
    // ================================================================




    // ================================================================
    // FIN DEFUNCIONES PARA SEGMENTACION DE DATOS
    // ================================================================



    // ================================================================
    // FUNCIONES PARA TEXTO
    // ================================================================




    // ================================================================
    // FIN DEFUNCIONES PARA TEXTO
    // ================================================================






    // ================================================================
    // FUNCIONES PARA SIMBOLOS
    // ================================================================




    // ================================================================
    // FIN DEFUNCIONES PARA SIMBOLOS
    // ================================================================


});
