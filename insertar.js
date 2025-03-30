// Funciones para manejo de las opciones de insertar

document.addEventListener("DOMContentLoaded", function () {
  // Referencia a elementos de la interfaz
  const btnImagenes = document.getElementById("btn-imagenes");
  const submenuImagenes = document.getElementById("submenu-imagenes");
  const btnFormas = document.getElementById("btn-formas");
  const submenuFormas = document.getElementById("submenu-formas");
  const uploadImageBtn = document.getElementById("upload-image-btn");
  const imageInput = document.getElementById("image-input");

  // Botones de gráficos
  const btnGraficoBarras = document.getElementById("btn-grafico-barras");
  const btnGraficoLineas = document.getElementById("btn-grafico-lineas");
  const btnGraficoCircular = document.getElementById("btn-grafico-circular");
  const btnGraficoDispersion = document.getElementById(
    "btn-grafico-dispersion"
  );

  // Modal de configuración de gráficos
  const chartModal = document.getElementById("chart-modal");
  const closeButton = document.querySelector(".close-button");
  const cancelChart = document.getElementById("cancel-chart");
  const createChart = document.getElementById("create-chart");

  // Variable para almacenar el tipo de gráfico seleccionado
  let selectedChartType = "";

  // Agregar manejador de eventos para botones con atributo data-menu
  const menuButtons = document.querySelectorAll("button[data-menu]");
  menuButtons.forEach((button) => {
    button.addEventListener("click", function () {
      const menuId = this.getAttribute("data-menu");
      const menu = document.getElementById(menuId);
      if (menu) {
        // Cerrar otros menús abiertos
        document.querySelectorAll(".menu.open").forEach((openMenu) => {
          if (openMenu !== menu) {
            openMenu.classList.remove("open");
          }
        });
        // Alternar clase 'open' en este menú
        menu.classList.toggle("open");
      }
    });
  });

  // Cerrar menús desplegables al hacer clic fuera de ellos
  document.addEventListener("click", function (e) {
    if (!e.target.closest("button[data-menu]") && !e.target.closest(".menu")) {
      document.querySelectorAll(".menu.open").forEach((menu) => {
        menu.classList.remove("open");
      });
    }
  });

  // Función para mostrar modal de gráfico con tipo específico
  function showChartModal(chartType) {
    selectedChartType = chartType;
    chartModal.style.display = "flex";
  }

  // Manejador para botones de gráficos
  if (btnGraficoBarras) {
    btnGraficoBarras.addEventListener("click", () => showChartModal("bar"));
  }

  if (btnGraficoLineas) {
    btnGraficoLineas.addEventListener("click", () => showChartModal("line"));
  }

  if (btnGraficoCircular) {
    btnGraficoCircular.addEventListener("click", () => showChartModal("pie"));
  }

  if (btnGraficoDispersion) {
    btnGraficoDispersion.addEventListener("click", () =>
      showChartModal("scatter")
    );
  }

  // Cerrar modal de gráficos
  if (closeButton) {
    closeButton.addEventListener("click", () => {
      chartModal.style.display = "none";
    });
  }

  if (cancelChart) {
    cancelChart.addEventListener("click", () => {
      chartModal.style.display = "none";
    });
  }

  // Manejar clic fuera del modal para cerrarlo
  window.addEventListener("click", (e) => {
    if (e.target === chartModal) {
      chartModal.style.display = "none";
    }
  });

  //Manejo de insertar imagenes

  const spreadsheetContainer = document.querySelector(".spreadsheet-container");
  const imageInputInsertarImg = document.createElement("input");

  imageInputInsertarImg.type = "file";
  imageInputInsertarImg.accept = "image/*";
  imageInputInsertarImg.style.display = "none";
  document.body.appendChild(imageInputInsertarImg);

  // Evento para abrir el selector de imágenes
  if (uploadImageBtn && imageInputInsertarImg) {
    uploadImageBtn.addEventListener("click", () => {
      imageInputInsertarImg.click();
    });

    imageInputInsertarImg.addEventListener(
      "change",
      handleImageUploadInsertarImg
    );
  }

  // Función para insertar la imagen dentro de la hoja de cálculo
  function insertImageInsideSpreadsheet(imageSrc) {
    if (!spreadsheetContainer) {
      alert("No se encontró la hoja de cálculo.");
      return;
    }

    const imgContainer = document.createElement("div");
    imgContainer.style.position = "absolute";
    imgContainer.style.width = "150px";
    imgContainer.style.height = "150px";
    imgContainer.style.top = "50px";
    imgContainer.style.left = "50px";
    imgContainer.style.border = "2px solid black";
    imgContainer.style.borderRadius = "5px";
    imgContainer.style.cursor = "move";
    imgContainer.style.zIndex = "1000";
    imgContainer.style.overflow = "hidden";

    const img = document.createElement("img");
    img.src = imageSrc;
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = "cover";

    imgContainer.appendChild(img);
    spreadsheetContainer.appendChild(imgContainer);

    makeImageInteractiveInsertarImg(imgContainer);
  }

  // Función para procesar la imagen cargada
  function handleImageUploadInsertarImg(event) {
    const file = event.target.files[0];
    if (file && file.type.match("image.*")) {
      const reader = new FileReader();
      reader.onload = function (e) {
        insertImageInsideSpreadsheet(e.target.result);
      };
      reader.readAsDataURL(file);
    }
  }

  // Función para hacer la imagen interactiva dentro de la hoja de cálculo
  function makeImageInteractiveInsertarImg(imgContainer) {
    let offsetX,
      offsetY,
      isDragging = false;

    // Permitir mover la imagen
    imgContainer.addEventListener("mousedown", (e) => {
      isDragging = true;
      offsetX = e.clientX - imgContainer.getBoundingClientRect().left;
      offsetY = e.clientY - imgContainer.getBoundingClientRect().top;
    });

    document.addEventListener("mousemove", (e) => {
      if (isDragging) {
        const containerRect = spreadsheetContainer.getBoundingClientRect();
        let newX = e.clientX - containerRect.left - offsetX;
        let newY = e.clientY - containerRect.top - offsetY;

        // Restringir el movimiento dentro de la hoja de cálculo
        newX = Math.max(
          0,
          Math.min(containerRect.width - imgContainer.clientWidth, newX)
        );
        newY = Math.max(
          0,
          Math.min(containerRect.height - imgContainer.clientHeight, newY)
        );

        imgContainer.style.left = `${newX}px`;
        imgContainer.style.top = `${newY}px`;
      }
    });

    document.addEventListener("mouseup", () => {
      isDragging = false;
    });

    // Agregar manejadores de redimensionamiento en las esquinas
    ["bottom-right", "bottom-left", "top-right", "top-left"].forEach(
      (corner) => {
        const resizeHandle = document.createElement("div");
        resizeHandle.className = `resize-handle ${corner}`;
        resizeHandle.style.position = "absolute";
        resizeHandle.style.width = "10px";
        resizeHandle.style.height = "10px";
        resizeHandle.style.background = "blue";
        resizeHandle.style.cursor = `${
          corner.includes("bottom") ? "ns" : "ew"
        }-resize`;

        if (corner === "bottom-right") {
          resizeHandle.style.right = "0";
          resizeHandle.style.bottom = "0";
          resizeHandle.style.cursor = "nwse-resize";
        } else if (corner === "bottom-left") {
          resizeHandle.style.left = "0";
          resizeHandle.style.bottom = "0";
          resizeHandle.style.cursor = "nesw-resize";
        } else if (corner === "top-right") {
          resizeHandle.style.right = "0";
          resizeHandle.style.top = "0";
          resizeHandle.style.cursor = "nesw-resize";
        } else if (corner === "top-left") {
          resizeHandle.style.left = "0";
          resizeHandle.style.top = "0";
          resizeHandle.style.cursor = "nwse-resize";
        }

        imgContainer.appendChild(resizeHandle);

        // Hacer que funcione el redimensionamiento
        resizeHandle.addEventListener("mousedown", (e) => {
          e.stopPropagation();
          let isResizing = true;

          document.addEventListener("mousemove", (e) => {
            if (isResizing) {
              let newWidth =
                e.clientX - imgContainer.getBoundingClientRect().left;
              let newHeight =
                e.clientY - imgContainer.getBoundingClientRect().top;

              // Evitar tamaños negativos
              newWidth = Math.max(50, newWidth);
              newHeight = Math.max(50, newHeight);

              imgContainer.style.width = `${newWidth}px`;
              imgContainer.style.height = `${newHeight}px`;
            }
          });

          document.addEventListener(
            "mouseup",
            () => {
              isResizing = false;
            },
            { once: true }
          );
        });
      }
    );

    // Agregar menú contextual con opción de cortar
    imgContainer.addEventListener("contextmenu", (e) => {
      e.preventDefault();

      // Crear cuadro contextual
      const contextMenu = document.createElement("div");
      contextMenu.style.position = "absolute";
      contextMenu.style.top = `${e.clientY}px`;
      contextMenu.style.left = `${e.clientX}px`;
      contextMenu.style.background = "#fff";
      contextMenu.style.border = "1px solid #ccc";
      contextMenu.style.padding = "5px 10px";
      contextMenu.style.cursor = "pointer";
      contextMenu.style.boxShadow = "0px 0px 5px rgba(0,0,0,0.2)";
      contextMenu.style.zIndex = "2000";

      // Crear icono de tijeras
      const scissorsIcon = document.createElement("i");
      scissorsIcon.classList.add("fas", "fa-scissors"); // Font Awesome scissors icon
      scissorsIcon.style.marginRight = "10px"; // Espacio entre el icono y el texto

      // Crear texto para el contexto
      const contextText = document.createElement("span");
      contextText.textContent = "Cortar";

      // Agregar el ícono y el texto al menú
      contextMenu.appendChild(scissorsIcon);
      contextMenu.appendChild(contextText);

      document.body.appendChild(contextMenu);

      // Eliminar imagen al hacer clic en "Cortar"
      contextMenu.addEventListener("click", () => {
        imgContainer.remove();
        contextMenu.remove();
      });

      // Cerrar el menú si se hace clic fuera de él
      document.addEventListener(
        "click",
        () => {
          contextMenu.remove();
        },
        { once: true }
      );
    });
  }

  //fin imagenes

  //Funcion de formas
  window.insertShape = function (shapeType) {
    console.log("Insertando forma:", shapeType);

    const insertarForma = document.createElement("div");
    insertarForma.classList.add("shape");
    insertarForma.style.position = "absolute"; // Asegura que las formas se posicionen dentro del contenedor
    insertarForma.style.zIndex = "9999"; // Asegura que las formas estén por encima de las celdas de la tabla

    // Definir el tipo de forma
    if (shapeType === "rect") {
      insertarForma.style.width = "100px";
      insertarForma.style.height = "100px";
      insertarForma.style.backgroundColor = "lightblue";
    } else if (shapeType === "circle") {
      insertarForma.style.width = "100px";
      insertarForma.style.height = "100px";
      insertarForma.style.borderRadius = "50%";
      insertarForma.style.backgroundColor = "lightgreen";
    } else if (shapeType === "triangle") {
      insertarForma.style.width = "0";
      insertarForma.style.height = "0";
      insertarForma.style.borderLeft = "50px solid transparent";
      insertarForma.style.borderRight = "50px solid transparent";
      insertarForma.style.borderBottom = "100px solid lightcoral";
    }

    // Seleccionar el contenedor principal
    const spreadsheetContainer = document.querySelector(
      ".spreadsheet-container"
    );

    // Verificar si el contenedor existe
    if (!spreadsheetContainer) {
      console.log("No se encontró el contenedor principal.");
      return;
    }

    // Insertar la forma en el contenedor
    spreadsheetContainer.appendChild(insertarForma);

    // Configurar la posición inicial dentro del contenedor
    insertarForma.style.top = "50px";
    insertarForma.style.left = "50px";

    // Hacer la forma interactiva
    hacerFormaInteractiva(insertarForma, spreadsheetContainer);
  };

  // Función para hacer la forma interactiva
  function hacerFormaInteractiva(insertarForma, contenedor) {
    let isDragging = false;
    let offsetX, offsetY;

    insertarForma.addEventListener("mousedown", (e) => {
      isDragging = true;
      offsetX = e.clientX - insertarForma.getBoundingClientRect().left;
      offsetY = e.clientY - insertarForma.getBoundingClientRect().top;

      // Evitar que el menú contextual interfiera
      e.stopPropagation();
    });

    document.addEventListener("mousemove", (e) => {
      if (isDragging) {
        let newX =
          e.clientX - offsetX - contenedor.getBoundingClientRect().left;
        let newY = e.clientY - offsetY - contenedor.getBoundingClientRect().top;

        // Restringir el movimiento dentro del contenedor
        newX = Math.max(
          0,
          Math.min(contenedor.offsetWidth - insertarForma.offsetWidth, newX)
        );
        newY = Math.max(
          0,
          Math.min(contenedor.offsetHeight - insertarForma.offsetHeight, newY)
        );

        insertarForma.style.left = `${newX}px`;
        insertarForma.style.top = `${newY}px`;
      }
    });

    document.addEventListener("mouseup", () => {
      isDragging = false;
    });

    // Crear el manejador de redimensionado
    const resizeHandle = document.createElement("div");
    resizeHandle.classList.add("resize-handle");
    resizeHandle.style.width = "10px";
    resizeHandle.style.height = "10px";
    resizeHandle.style.position = "absolute";
    resizeHandle.style.bottom = "0";
    resizeHandle.style.right = "0";
    resizeHandle.style.backgroundColor = "black";
    resizeHandle.style.cursor = "nwse-resize";
    insertarForma.appendChild(resizeHandle);

    let isResizing = false;

    resizeHandle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      isResizing = true;
      const initialWidth = insertarForma.offsetWidth;
      const initialHeight = insertarForma.offsetHeight;
      const initialX = e.clientX;
      const initialY = e.clientY;

      document.addEventListener("mousemove", (e) => {
        if (isResizing) {
          const widthChange = e.clientX - initialX;
          const heightChange = e.clientY - initialY;

          insertarForma.style.width = `${Math.max(
            50,
            initialWidth + widthChange
          )}px`;
          insertarForma.style.height = `${Math.max(
            50,
            initialHeight + heightChange
          )}px`;
        }
      });

      document.addEventListener("mouseup", () => {
        isResizing = false;
      });
    });

    // Menú contextual para cortar con ícono de tijeras
    insertarForma.addEventListener("contextmenu", (e) => {
      e.preventDefault();

      // Crear cuadro contextual para eliminar la forma
      const contextMenu = document.createElement("div");
      contextMenu.style.position = "absolute";

      // Usar las coordenadas del evento de clic para posicionar el cuadro contextual
      const topPosition = e.clientY; // Ubicación del clic en el eje Y
      const leftPosition = e.clientX; // Ubicación del clic en el eje X

      contextMenu.style.top = `${topPosition}px`;
      contextMenu.style.left = `${leftPosition}px`;

      contextMenu.style.background = "#fff";
      contextMenu.style.border = "1px solid #ccc";
      contextMenu.style.padding = "5px 10px";
      contextMenu.style.cursor = "pointer";
      contextMenu.style.boxShadow = "0px 0px 5px rgba(0,0,0,0.2)";
      contextMenu.style.zIndex = "10000"; // Asegura que el menú esté encima de la forma

      // Crear ícono de tijeras
      const scissorsIcon = document.createElement("i");
      scissorsIcon.classList.add("fas", "fa-scissors"); // Font Awesome scissors icon
      scissorsIcon.style.marginRight = "10px"; // Espacio entre el icono y el texto

      const contextText = document.createElement("span");
      contextText.textContent = "Cortar";

      // Agregar el ícono y el texto al menú
      contextMenu.appendChild(scissorsIcon);
      contextMenu.appendChild(contextText);

      document.body.appendChild(contextMenu);

      // Eliminar la forma al hacer clic en "Cortar"
      contextMenu.addEventListener("click", () => {
        insertarForma.remove();
        contextMenu.remove();
      });

      // Cerrar el menú si se hace clic fuera de él
      document.addEventListener(
        "click",
        () => {
          contextMenu.remove();
        },
        { once: true }
      );
    });
  }

  //fin forma

  // Crear gráfico con los datos seleccionados
  if (createChart) {
    createChart.addEventListener("click", () => {
      const chartTitle = document.getElementById("chart-title").value;
      const dataRange = document.getElementById("data-range").value;

      // Aquí se debe implementar la lógica para crear el gráfico
      console.log("Creando gráfico:", {
        type: selectedChartType,
        title: chartTitle,
        dataRange: dataRange,
      });

      // Cerrar modal después de crear
      chartModal.style.display = "none";
    });
  }
});

// Funcionalidad para insertar tablas
document.addEventListener("DOMContentLoaded", function () {
  // Utilizamos const y let para definir variables (estilo moderno)
  const insertTableBtn = document.getElementById("insert-table-btn");
  let activeTableCells = []; // Para rastrear celdas que pertenecen a la tabla activa

  // Crear modal para tabla si no existe
  if (!document.getElementById("table-modal")) {
    createTableModal();
  }

  // Mostrar modal al hacer clic en el botón de tabla
  if (insertTableBtn) {
    insertTableBtn.addEventListener("click", mostrarModalTablaInsertarTabla);
  }

  // Agregar evento para manejo del click derecho en celdas
  const spreadsheet =
    document.querySelector(".spreadsheet") ||
    document.getElementById("spreadsheet");
  if (spreadsheet) {
    spreadsheet.addEventListener(
      "contextmenu",
      mostrarMenuContextualInsertarTabla
    );
  }

  // Función para crear el modal de la tabla
  function createTableModal() {
    const tableModal = document.createElement("div");
    tableModal.id = "table-modal";
    tableModal.className = "chart-modal"; // Reusamos estilos del modal de gráficos

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
    agregarEstilosTablaInsertarTabla();

    // Event listeners para el modal
    const closeTableModal = document.getElementById("close-table-modal");
    const cancelTable = document.getElementById("cancel-table");
    const createTable = document.getElementById("create-table");
    const tableStyleOptions = document.querySelectorAll(".table-style-option");

    closeTableModal.addEventListener("click", () => {
      tableModal.style.display = "none";
    });

    cancelTable.addEventListener("click", () => {
      tableModal.style.display = "none";
    });

    // Selección de estilo
    tableStyleOptions.forEach((option) => {
      option.addEventListener("click", () => {
        tableStyleOptions.forEach((opt) => opt.classList.remove("selected"));
        option.classList.add("selected");
      });
    });

    // Crear tabla
    createTable.addEventListener("click", () => {
      const rangeInput = document.getElementById("table-range").value;
      const hasHeaders = document.getElementById("table-has-headers").checked;
      const selectedStyle = document.querySelector(
        ".table-style-option.selected"
      ).dataset.style;

      if (!rangeInput) {
        alert("Por favor, ingrese un rango válido");
        return;
      }

      createTableFromRange(rangeInput, hasHeaders, selectedStyle);
      tableModal.style.display = "none";
    });

    // Cerrar modal al hacer clic fuera
    window.addEventListener("click", (e) => {
      if (e.target === tableModal) {
        tableModal.style.display = "none";
      }
    });
  }

  // Función para agregar estilos de tabla
  function agregarEstilosTablaInsertarTabla() {
    if (!document.getElementById("table-preview-styles")) {
      const styleElement = document.createElement("style");
      styleElement.id = "table-preview-styles";
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
            transition: all 0.2s ease;
          }
          .table-style-option:hover {
            background-color: #f5f5f5;
          }
          .table-style-option.selected {
            border-color: #217346;
          }
          .style-preview {
            width: 80px;
            height: 60px;
            margin-bottom: 5px;
            border: 1px solid #ccc;
            border-radius: 2px;
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
          
          /* Estilos para menú contextual */
          .context-menu-insertarTabla {
            position: absolute;
            background-color: white;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            padding: 5px 0;
            z-index: 1000;
          }
          .context-menu-insertarTabla ul {
            list-style: none;
            margin: 0;
            padding: 0;
          }
          .context-menu-insertarTabla li {
            padding: 8px 15px;
            cursor: pointer;
            font-size: 14px;
          }
          .context-menu-insertarTabla li:hover {
            background-color: #f0f0f0;
          }
          .context-menu-insertarTabla .separator {
            height: 1px;
            background-color: #ccc;
            margin: 5px 0;
          }
        `;
      document.head.appendChild(styleElement);
    }

    // Estilos para las tablas
    if (!document.getElementById("table-styles")) {
      const styleElement = document.createElement("style");
      styleElement.id = "table-styles";
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
          /* Destacar tabla activa */
          .table-active {
            outline: 2px solid #0078d7;
          }
        `;
      document.head.appendChild(styleElement);
    }
  }

  // Función para mostrar el modal de tabla
  function mostrarModalTablaInsertarTabla() {
    // Si hay un rango seleccionado, prellenamos el campo
    const activeRange = document.getElementById("active-cell")?.textContent;
    if (activeRange && activeRange.includes(":")) {
      document.getElementById("table-range").value = activeRange;
    } else {
      // Si solo hay una celda activa, seleccionamos un rango predeterminado alrededor de ella
      const cellRef = activeRange;
      if (cellRef && cellRef.match(/^[A-Z]+\d+$/)) {
        const col = cellRef.match(/[A-Z]+/)[0];
        const row = parseInt(cellRef.match(/\d+/)[0]);
        document.getElementById("table-range").value = `${col}${row}:${col}${
          row + 4
        }`;
      }
    }
    document.getElementById("table-modal").style.display = "flex";
  }

  // Función para crear tabla a partir de un rango
  function createTableFromRange(range, hasHeaders, style) {
    try {
      // Validar formato del rango (por ejemplo, A1:D5)
      if (!range.match(/^[A-Z]+\d+:[A-Z]+\d+$/)) {
        throw new Error("Formato de rango inválido");
      }

      // Extraer coordenadas
      const [startRef, endRef] = range.split(":");
      const startCoords = getCellCoords(startRef);
      const endCoords = getCellCoords(endRef);

      // Asegurar que las coordenadas son válidas
      if (!startCoords || !endCoords) {
        throw new Error("Coordenadas inválidas");
      }

      // Determinar el rango real (min a max)
      const startX = Math.min(startCoords[0], endCoords[0]);
      const endX = Math.max(startCoords[0], endCoords[0]);
      const startY = Math.min(startCoords[1], endCoords[1]);
      const endY = Math.max(startCoords[1], endCoords[1]);

      // Aplicar estilo de tabla a las celdas
      const tableId = `table-${Date.now()}`; // ID único para la tabla
      applyTableStyle(startX, endX, startY, endY, hasHeaders, style, tableId);

      // Guardar información de la tabla para futuras referencias
      activeTableCells = [];
      for (let y = startY; y <= endY; y++) {
        for (let x = startX; x <= endX; x++) {
          activeTableCells.push({ x, y, tableId });
        }
      }

      console.log(
        `Tabla creada con estilo ${style} en el rango ${range}, con encabezados: ${hasHeaders}`
      );
    } catch (error) {
      console.error("Error al crear tabla:", error);
      alert("Error al crear la tabla: " + error.message);
    }
  }

  // Función para aplicar estilo de tabla a un rango de celdas
  function applyTableStyle(
    startX,
    endX,
    startY,
    endY,
    hasHeaders,
    styleType,
    tableId
  ) {
    // Clase base para todas las tablas
    const baseClass = "excel-table";

    // Definir clases según el estilo seleccionado
    let styleClass;
    switch (styleType) {
      case "style1":
        styleClass = "table-style-blue";
        break;
      case "style2":
        styleClass = "table-style-green";
        break;
      case "style3":
        styleClass = "table-style-orange";
        break;
      default:
        styleClass = "table-style-blue";
    }

    // Aplicar estilos a cada celda en el rango
    for (let y = startY; y <= endY; y++) {
      for (let x = startX; x <= endX; x++) {
        const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
        if (cell) {
          // Almacenar las coordenadas de la tabla
          cell.dataset.tableId = tableId;
          cell.dataset.tableStartX = startX;
          cell.dataset.tableStartY = startY;
          cell.dataset.tableEndX = endX;
          cell.dataset.tableEndY = endY;
          cell.dataset.tableHasHeaders = hasHeaders;
          cell.dataset.tableStyle = styleType;

          // Clase base para todas las celdas de tabla
          cell.classList.add(baseClass);

          // Clase de estilo específico
          cell.classList.add(styleClass);

          // Clases específicas según la posición
          if (hasHeaders && y === startY) {
            cell.classList.add("table-header");
          } else if (y === endY) {
            cell.classList.add("table-footer");
          }

          if (x === startX) {
            cell.classList.add("table-left");
          }

          if (x === endX) {
            cell.classList.add("table-right");
          }

          // Alternar filas para algunas filas
          if ((y - startY) % 2 === 1 && (!hasHeaders || y > startY)) {
            cell.classList.add("table-alternate-row");
          }
        }
      }
    }
  }

  // Función para mostrar el menú contextual
  function mostrarMenuContextualInsertarTabla(e) {
    // Verificar si el click fue en una celda
    const cell = e.target.closest("td");
    if (!cell) return; // Si no es una celda, salimos

    // Verificar si la celda es parte de una tabla
    if (cell.classList.contains("excel-table")) {
      e.preventDefault(); // Prevenir el menú contextual por defecto

      // Obtener información de la tabla
      const tableId = cell.dataset.tableId;
      const startX = parseInt(cell.dataset.tableStartX);
      const startY = parseInt(cell.dataset.tableStartY);
      const endX = parseInt(cell.dataset.tableEndX);
      const endY = parseInt(cell.dataset.tableEndY);
      const hasHeaders = cell.dataset.tableHasHeaders === "true";
      const style = cell.dataset.tableStyle;

      // Resaltar la tabla actual
      highlightTableCells(startX, endX, startY, endY);

      // Crear menú contextual
      crearMenuContextual(e.clientX, e.clientY, {
        cell,
        x: parseInt(cell.dataset.x),
        y: parseInt(cell.dataset.y),
        tableId,
        startX,
        startY,
        endX,
        endY,
        hasHeaders,
        style,
      });
    }
  }

  // Función para destacar visualmente todas las celdas de la tabla
  function highlightTableCells(startX, endX, startY, endY) {
    // Primero eliminamos el resaltado de todas las celdas
    document.querySelectorAll(".table-active").forEach((cell) => {
      cell.classList.remove("table-active");
    });

    // Luego resaltamos las celdas de la tabla actual
    for (let y = startY; y <= endY; y++) {
      for (let x = startX; x <= endX; x++) {
        const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
        if (cell) {
          cell.classList.add("table-active");
        }
      }
    }
  }

  // Función para crear el menú contextual
  function crearMenuContextual(x, y, tableInfo) {
    // Eliminar cualquier menú contextual existente
    const existingMenu = document.querySelector(".context-menu-insertarTabla");
    if (existingMenu) {
      existingMenu.remove();
    }

    // Crear el nuevo menú
    const menu = document.createElement("div");
    menu.className = "context-menu-insertarTabla";
    menu.innerHTML = `
        <ul>
          <li data-action="insert-row-above">Insertar fila arriba</li>
          <li data-action="insert-row-below">Insertar fila abajo</li>
          <li class="separator"></li>
          <li data-action="insert-column-left">Insertar columna a la izquierda</li>
          <li data-action="insert-column-right">Insertar columna a la derecha</li>
          <li class="separator"></li>
          <li data-action="delete-row">Eliminar fila</li>
          <li data-action="delete-column">Eliminar columna</li>
          <li class="separator"></li>
          <li data-action="delete-table">Eliminar tabla</li>
        </ul>
      `;

    // Posicionar el menú
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;

    // Agregar el menú al documento
    document.body.appendChild(menu);

    // Asignar eventos a las opciones del menú
    menu.querySelectorAll("li[data-action]").forEach((item) => {
      item.addEventListener("click", () => {
        const action = item.dataset.action;
        handleTableAction(action, tableInfo);
        menu.remove();
      });
    });

    // Cerrar el menú al hacer clic fuera
    document.addEventListener("click", function closeMenu(e) {
      if (!menu.contains(e.target)) {
        menu.remove();
        document.removeEventListener("click", closeMenu);
      }
    });
  }

  // Función para manejar las acciones de la tabla
  function handleTableAction(action, tableInfo) {
    const { x, y, startX, startY, endX, endY, hasHeaders, style, tableId } =
      tableInfo;

    switch (action) {
      case "insert-row-above":
        insertarFilaEnTabla(
          startX,
          endX,
          y,
          "above",
          tableId,
          hasHeaders,
          style
        );
        break;
      case "insert-row-below":
        insertarFilaEnTabla(
          startX,
          endX,
          y,
          "below",
          tableId,
          hasHeaders,
          style
        );
        break;
      case "insert-column-left":
        insertarColumnaEnTabla(
          startY,
          endY,
          x,
          "left",
          tableId,
          hasHeaders,
          style
        );
        break;
      case "insert-column-right":
        insertarColumnaEnTabla(
          startY,
          endY,
          x,
          "right",
          tableId,
          hasHeaders,
          style
        );
        break;
      case "delete-row":
        eliminarFilaDeTabla(
          startX,
          endX,
          y,
          startY,
          endY,
          tableId,
          hasHeaders,
          style
        );
        break;
      case "delete-column":
        eliminarColumnaDeTabla(
          startY,
          endY,
          x,
          startX,
          endX,
          tableId,
          hasHeaders,
          style
        );
        break;
      case "delete-table":
        eliminarTabla(startX, endX, startY, endY);
        break;
    }
  }

  // Función para insertar una fila en la tabla
  function insertarFilaEnTabla(
    startX,
    endX,
    atY,
    position,
    tableId,
    hasHeaders,
    style
  ) {
    // Determinar la posición para insertar
    const insertPosition = position === "above" ? atY : atY + 1;

    // Función para mover celdas hacia abajo desde una fila específica
    function moverCeldasHaciaAbajo(desdeY) {
      // Obtener el número máximo de filas y columnas en la hoja de cálculo
      const maxRow = getMaxRow();
      const maxCol = getMaxCol();

      // Empezar desde la última fila y mover hacia arriba
      for (let y = maxRow; y >= desdeY; y--) {
        for (let x = 0; x <= maxCol; x++) {
          const sourceCell = document.querySelector(
            `td[data-x="${x}"][data-y="${y}"]`
          );
          const targetCell = document.querySelector(
            `td[data-x="${x}"][data-y="${y + 1}"]`
          );

          if (sourceCell && targetCell) {
            // Transferir contenido
            targetCell.innerHTML = sourceCell.innerHTML;
            // Transferir clases relacionadas con tablas
            sourceCell.classList.forEach((className) => {
              if (
                className.startsWith("table-") ||
                className === "excel-table"
              ) {
                targetCell.classList.add(className);
                sourceCell.classList.remove(className);
              }
            });
            // Limpiar celda de origen si es la fila que se está insertando
            if (y === desdeY) {
              sourceCell.innerHTML = "";
              // Eliminar clases de tabla
              sourceCell.classList.forEach((className) => {
                if (
                  className.startsWith("table-") ||
                  className === "excel-table"
                ) {
                  sourceCell.classList.remove(className);
                }
              });
            }
          }
        }
      }
    }

    // Primero, debemos mover todas las celdas hacia abajo para hacer espacio
    moverCeldasHaciaAbajo(insertPosition);

    // Luego, aplicar estilo de tabla a la nueva fila
    for (let x = startX; x <= endX; x++) {
      const cell = document.querySelector(
        `td[data-x="${x}"][data-y="${insertPosition}"]`
      );
      if (cell) {
        // Actualizar las clases y datos de la celda
        cell.classList.add("excel-table");

        // Determinar el estilo según el tipo de tabla
        let styleClass;
        switch (style) {
          case "style1":
            styleClass = "table-style-blue";
            break;
          case "style2":
            styleClass = "table-style-green";
            break;
          case "style3":
            styleClass = "table-style-orange";
            break;
          default:
            styleClass = "table-style-blue";
        }

        cell.classList.add(styleClass);

        // Agregar clases para bordes
        if (x === startX) cell.classList.add("table-left");
        if (x === endX) cell.classList.add("table-right");

        // Alternar filas
        const rowOffset = insertPosition - startY;
        if (rowOffset % 2 === 1 && (!hasHeaders || insertPosition > startY)) {
          cell.classList.add("table-alternate-row");
        }

        // Establecer atributos de tabla
        cell.dataset.tableId = tableId;
      }
    }

    // Actualizar información de la tabla en todas las celdas
    actualizarInfoTabla(
      startX,
      endX,
      startY,
      endY + 1,
      hasHeaders,
      style,
      tableId
    );
  }

  // Función para insertar una columna en la tabla
  function insertarColumnaEnTabla(
    startY,
    endY,
    atX,
    position,
    tableId,
    hasHeaders,
    style
  ) {
    // Determinar la posición para insertar
    const insertPosition = position === "left" ? atX : atX + 1;

    // Función para mover celdas hacia la derecha desde una columna específica
    function moverCeldasHaciaDerecha(desdeX) {
      // Obtener el número máximo de filas y columnas en la hoja de cálculo
      const maxRow = getMaxRow();
      const maxCol = getMaxCol();

      // Empezar desde la última columna y mover hacia la izquierda
      for (let x = maxCol; x >= desdeX; x--) {
        for (let y = 0; y <= maxRow; y++) {
          const sourceCell = document.querySelector(
            `td[data-x="${x}"][data-y="${y}"]`
          );
          const targetCell = document.querySelector(
            `td[data-x="${x + 1}"][data-y="${y}"]`
          );

          if (sourceCell && targetCell) {
            // Transferir contenido
            targetCell.innerHTML = sourceCell.innerHTML;
            // Transferir clases relacionadas con tablas
            sourceCell.classList.forEach((className) => {
              if (
                className.startsWith("table-") ||
                className === "excel-table"
              ) {
                targetCell.classList.add(className);
                sourceCell.classList.remove(className);
              }
            });
            // Limpiar celda de origen si es la columna que se está insertando
            if (x === desdeX) {
              sourceCell.innerHTML = "";
              // Eliminar clases de tabla
              sourceCell.classList.forEach((className) => {
                if (
                  className.startsWith("table-") ||
                  className === "excel-table"
                ) {
                  sourceCell.classList.remove(className);
                }
              });
            }
          }
        }
      }
    }

    // Funciones auxiliares para obtener el máximo número de filas y columnas
    function getMaxRow() {
      // Buscar la fila con el valor más alto de data-y
      const rows = document.querySelectorAll("td[data-y]");
      let maxRow = 0;

      rows.forEach((cell) => {
        const y = parseInt(cell.dataset.y);
        if (y > maxRow) maxRow = y;
      });

      return maxRow;
    }

    function getMaxCol() {
      // Buscar la columna con el valor más alto de data-x
      const cols = document.querySelectorAll("td[data-x]");
      let maxCol = 0;

      cols.forEach((cell) => {
        const x = parseInt(cell.dataset.x);
        if (x > maxCol) maxCol = x;
      });

      return maxCol;
    }

    // Primero, debemos mover todas las celdas hacia la derecha para hacer espacio
    moverCeldasHaciaDerecha(insertPosition);

    // Luego, aplicar estilo de tabla a la nueva columna
    for (let y = startY; y <= endY; y++) {
      const cell = document.querySelector(
        `td[data-x="${insertPosition}"][data-y="${y}"]`
      );
      if (cell) {
        // Actualizar las clases y datos de la celda
        cell.classList.add("excel-table");

        // Determinar el estilo según el tipo de tabla
        let styleClass;
        switch (style) {
          case "style1":
            styleClass = "table-style-blue";
            break;
          case "style2":
            styleClass = "table-style-green";
            break;
          case "style3":
            styleClass = "table-style-orange";
            break;
          default:
            styleClass = "table-style-blue";
        }

        cell.classList.add(styleClass);

        // Agregar clases especiales
        if (hasHeaders && y === startY) cell.classList.add("table-header");
        if (y === endY) cell.classList.add("table-footer");

        // Alternar filas
        const rowOffset = y - startY;
        if (rowOffset % 2 === 1 && (!hasHeaders || y > startY)) {
          cell.classList.add("table-alternate-row");
        }

        // Establecer atributos de tabla
        cell.dataset.tableId = tableId;
      }
    }

    // Actualizar información de la tabla en todas las celdas
    actualizarInfoTabla(
      startX,
      endX + 1,
      startY,
      endY,
      hasHeaders,
      style,
      tableId
    );
  }

  // Función para eliminar una fila de la tabla
  function eliminarFilaDeTabla(
    startX,
    endX,
    atY,
    startY,
    endY,
    tableId,
    hasHeaders,
    style
  ) {
    // Si solo queda una fila, simplemente eliminamos toda la tabla
    if (endY - startY === 0) {
      eliminarTabla(startX, endX, startY, endY);
      return;
    }

    // Eliminar las celdas de la fila
    for (let x = startX; x <= endX; x++) {
      const cell = document.querySelector(`td[data-x="${x}"][data-y="${atY}"]`);
      if (cell) {
        // Eliminar todas las clases de tabla
        cell.classList.remove("excel-table");
        cell.classList.remove("table-style-blue");
        cell.classList.remove("table-style-green");
        cell.classList.remove("table-style-orange");
        cell.classList.remove("table-header");
        cell.classList.remove("table-footer");
        cell.classList.remove("table-left");
        cell.classList.remove("table-right");
        cell.classList.remove("table-alternate-row");
        cell.classList.remove("table-active");

        // Borrar el contenido y atributos de tabla
        updateCell(x, atY, "");
        delete cell.dataset.tableId;
        delete cell.dataset.tableStartX;
        delete cell.dataset.tableStartY;
        delete cell.dataset.tableEndX;
        delete cell.dataset.tableEndY;
        delete cell.dataset.tableHasHeaders;
        delete cell.dataset.tableStyle;
      }
    }

    // Mover las celdas hacia arriba
    for (let y = atY + 1; y <= endY; y++) {
      for (let x = startX; x <= endX; x++) {
        const sourceCell = document.querySelector(
          `td[data-x="${x}"][data-y="${y}"]`
        );
        const targetCell = document.querySelector(
          `td[data-x="${x}"][data-y="${y - 1}"]`
        );

        if (sourceCell && targetCell) {
          // Mover el contenido
          targetCell.innerHTML = sourceCell.innerHTML;

          // Mover las clases
          sourceCell.classList.forEach((className) => {
            if (className.startsWith("table-") || className === "excel-table") {
              targetCell.classList.add(className);
              sourceCell.classList.remove(className);
            }
          });

          function updateCell(x, y, value) {
            const cell = document.querySelector(
              `td[data-x="${x}"][data-y="${y}"]`
            );
            if (cell) {
              cell.innerHTML = value;
            }
          }

          // Actualizar datos de la celda
          updateCell(x, y, "");
        }
      }
    }

    // Función para actualizar la información de la tabla en todas las celdas
    function actualizarInfoTabla(
      startX,
      endX,
      startY,
      endY,
      hasHeaders,
      style,
      tableId
    ) {
      for (let y = startY; y <= endY; y++) {
        for (let x = startX; x <= endX; x++) {
          const cell = document.querySelector(
            `td[data-x="${x}"][data-y="${y}"]`
          );
          if (cell) {
            // Actualizar atributos de la tabla
            cell.dataset.tableId = tableId;
            cell.dataset.tableStartX = startX;
            cell.dataset.tableStartY = startY;
            cell.dataset.tableEndX = endX;
            cell.dataset.tableEndY = endY;
            cell.dataset.tableHasHeaders = hasHeaders;
            cell.dataset.tableStyle = style;
          }
        }
      }
    }

    // Actualizar información de la tabla en todas las celdas
    actualizarInfoTabla(
      startX,
      endX,
      startY,
      endY - 1,
      hasHeaders,
      style,
      tableId
    );
  }

  // Función para convertir referencia de celda a coordenadas
  function getCellCoords(cellRef) {
    // Extraer letra y número
    const match = cellRef.match(/([A-Z]+)(\d+)/);
    if (!match) return null;

    const colStr = match[1];
    const row = parseInt(match[2]) - 1; // Restar 1 porque las filas empiezan en 0 internamente

    // Convertir letra a número (A=0, B=1, ...)
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - "A".charCodeAt(0));
    }

    return [col, row];
  }

  // Función para eliminar una columna de la tabla
  function eliminarColumnaDeTabla(
    startY,
    endY,
    atX,
    startX,
    endX,
    tableId,
    hasHeaders,
    style
  ) {
    // Si solo queda una columna, simplemente eliminamos toda la tabla
    if (endX - startX === 0) {
      eliminarTabla(startX, endX, startY, endY);
      return;
    }

    // Eliminar las celdas de la columna
    for (let y = startY; y <= endY; y++) {
      const cell = document.querySelector(`td[data-x="${atX}"][data-y="${y}"]`);
      if (cell) {
        // Eliminar todas las clases de tabla
        cell.classList.remove("excel-table");
        cell.classList.remove("table-style-blue");
        cell.classList.remove("table-style-green");
        cell.classList.remove("table-style-orange");
        cell.classList.remove("table-header");
        cell.classList.remove("table-footer");
        cell.classList.remove("table-left");
        cell.classList.remove("table-right");
        cell.classList.remove("table-alternate-row");
        cell.classList.remove("table-active");

        // Borrar el contenido y atributos de tabla
        updateCell(atX, y, "");
        delete cell.dataset.tableId;
        delete cell.dataset.tableStartX;
        delete cell.dataset.tableStartY;
        delete cell.dataset.tableEndX;
        delete cell.dataset.tableEndY;
        delete cell.dataset.tableHasHeaders;
        delete cell.dataset.tableStyle;
      }
    }

    // Mover las celdas hacia la izquierda
    for (let x = atX + 1; x <= endX; x++) {
      for (let y = startY; y <= endY; y++) {
        const sourceCell = document.querySelector(
          `td[data-x="${x}"][data-y="${y}"]`
        );
        const targetCell = document.querySelector(
          `td[data-x="${x - 1}"][data-y="${y}"]`
        );

        if (sourceCell && targetCell) {
          // Mover el contenido
          targetCell.innerHTML = sourceCell.innerHTML;

          // Mover las clases
          sourceCell.classList.forEach((className) => {
            if (className.startsWith("table-") || className === "excel-table") {
              targetCell.classList.add(className);
              sourceCell.classList.remove(className);
            }
          });

          // Actualizar datos de la celda
          updateCell(x, y, "");
        }
      }
    }

    // Actualizar información de la tabla en todas las celdas
    actualizarInfoTabla(
      startX,
      endX - 1,
      startY,
      endY,
      hasHeaders,
      style,
      tableId
    );
  }

  // Función para eliminar toda la tabla
  function eliminarTabla(startX, endX, startY, endY) {
    for (let y = startY; y <= endY; y++) {
      for (let x = startX; x <= endX; x++) {
        const cell = document.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
        if (cell) {
          // Eliminar todas las clases relacionadas con tablas
          cell.classList.remove("excel-table");
          cell.classList.remove("table-style-blue");
          cell.classList.remove("table-style-green");
          cell.classList.remove("table-style-orange");
          cell.classList.remove("table-header");
          cell.classList.remove("table-footer");
          cell.classList.remove("table-border-top");
          cell.classList.remove("table-border-bottom");
          cell.classList.remove("table-border-left");
          cell.classList.remove("table-border-right");
          cell.classList.remove("table-cell");

          // Restaurar el estilo original
          cell.style.backgroundColor = "";
          cell.style.border = "";
          cell.style.fontWeight = "";
          cell.style.color = "";
        }
      }
    }

    // Actualizamos el estado de la tabla
    tablaActiva = false;
    tablaBordes = {
      startX: null,
      startY: null,
      endX: null,
      endY: null,
    };
  }
});

// Fin de tablas

//Cuadro de texto

// Funcionalidad para Cuadros de Texto
document.addEventListener("DOMContentLoaded", function () {
  // Ahora usamos el ID específico para una selección más precisa
  const textBoxBtn = document.getElementById("text-box-btn");
  const spreadsheetContainer = document.querySelector(".spreadsheet-container");
  // Variable para seguir cuadros de texto activos
  let activeTextBox = null;
  let isCreatingTextBox = false;
  let textBoxDragStart = null;
  let textBoxes = [];
  let nextTextBoxId = 1;

  // Crear elemento de menú contextual una sola vez
  const contextMenu = document.createElement("div");
  contextMenu.className = "textbox-context-menu";
  contextMenu.style.position = "absolute";
  contextMenu.style.display = "none";
  contextMenu.style.zIndex = "1000";
  contextMenu.style.background = "white";
  contextMenu.style.border = "1px solid #ccc";
  contextMenu.style.boxShadow = "2px 2px 5px rgba(0,0,0,0.2)";
  contextMenu.style.padding = "5px 0";
  contextMenu.style.borderRadius = "4px";
  document.body.appendChild(contextMenu);

  // Agregar opción de eliminación
  const deleteOption = document.createElement("div");
  deleteOption.innerText = "Eliminar cuadro de texto";
  deleteOption.style.padding = "6px 15px";
  deleteOption.style.cursor = "pointer";

  // Agregar evento de hover
  deleteOption.addEventListener("mouseover", function () {
    this.style.backgroundColor = "#f0f0f0";
  });

  deleteOption.addEventListener("mouseout", function () {
    this.style.backgroundColor = "transparent";
  });

  contextMenu.appendChild(deleteOption);

  // Si existe el botón de cuadro de texto, añadir evento
  if (textBoxBtn) {
    console.log("Botón de cuadro de texto encontrado", textBoxBtn);
    textBoxBtn.addEventListener("click", () => {
      console.log("Clic en botón de cuadro de texto");
      // Cambiar cursor para indicar modo de inserción de cuadro de texto
      spreadsheetContainer.style.cursor = "crosshair";
      isCreatingTextBox = true;
    });
  } else {
    console.error("Botón de cuadro de texto no encontrado");
  }

  // Evento para crear cuadro de texto al hacer clic en la hoja
  spreadsheetContainer.addEventListener("mousedown", (e) => {
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
  spreadsheetContainer.addEventListener("mouseup", (e) => {
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
    spreadsheetContainer.style.cursor = "default";
  });

  // Crear un cuadro de texto
  function createTextBox(left, top, width, height) {
    // Crear elemento
    const textBox = document.createElement("div");
    const textBoxId = `text-box-${nextTextBoxId++}`;
    textBox.className = "excel-text-box";
    textBox.id = textBoxId;
    textBox.style.position = "absolute";
    textBox.style.left = `${left}px`;
    textBox.style.top = `${top}px`;
    // Hacemos el cuadro de texto más grande inicialmente
    textBox.style.width = `${Math.max(width, 250)}px`;
    textBox.style.height = `${Math.max(height, 120)}px`;
    textBox.setAttribute("contenteditable", "true");
    textBox.dataset.type = "textbox";

    // Añadir estas líneas después de configurar otras propiedades de estilo:
    textBox.style.direction = "ltr"; // Asegura dirección de izquierda a derecha
    textBox.style.textAlign = "left"; // Alinea el texto a la izquierda
    textBox.style.unicodeBidi = "normal"; // Comportamiento normal de texto

    // Prevenir comportamiento no deseado al editar el texto
    textBox.addEventListener("keydown", function (e) {
      // Solo detenemos la propagación para teclas de navegación
      // pero permitimos la escritura normal
      if (
        ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Tab"].includes(
          e.key
        )
      ) {
        e.stopPropagation();
      }
    });

    // Permite la edición normal del texto mientras mantiene la propagación para otros eventos
    textBox.addEventListener("input", function (e) {
      e.stopPropagation();
    });

    // Manejar clics dentro del cuadro para evitar propagación no deseada
    textBox.addEventListener("click", function (e) {
      e.stopPropagation();
      // Asegurar que el cuadro esté activo cuando hacemos clic dentro
      if (activeTextBox !== textBox) {
        setActiveTextBox(textBox);
      }
    });

    // Agregar evento contextmenu (clic derecho) MEJORADO
    textBox.addEventListener("contextmenu", handleContextMenu);

    function handleContextMenu(e) {
      // Prevenir el menú contextual predeterminado
      e.preventDefault();

      // Asegurar que este textbox esté activo
      setActiveTextBox(textBox);

      // Obtener la posición precisa para el menú
      const rect = textBox.getBoundingClientRect();

      // Posicionar el menú relativo al cuadro de texto,
      // no solo donde se hizo clic
      contextMenu.style.left = `${e.pageX}px`;
      contextMenu.style.top = `${e.pageY}px`;
      contextMenu.style.display = "block";

      // Configurar la acción de eliminación para este textBox específico
      deleteOption.onclick = function () {
        console.log("Eliminando cuadro de texto:", textBox.id);
        textBox.remove();
        contextMenu.style.display = "none";

        // Actualizar la lista de cuadros de texto
        textBoxes = textBoxes.filter((box) => box !== textBox);

        // Limpiar la referencia si este era el activo
        if (activeTextBox === textBox) {
          activeTextBox = null;
        }
      };

      // Asegurar que el evento no burbujee para evitar cierre inmediato
      e.stopPropagation();
    }

    // Añadir el cuadro de texto al contenedor
    spreadsheetContainer.appendChild(textBox);
    // Añadir a la lista de cuadros de texto
    textBoxes.push(textBox);
    // Establecer como activo y darle foco
    setActiveTextBox(textBox);
    textBox.focus();
  }

  // Eventos para la selección y movimiento de cuadros de texto
  document.addEventListener("mousedown", (e) => {
    // No ocultar el menú contextual inmediatamente, solo si se hace clic fuera de él
    if (!e.target.closest(".textbox-context-menu")) {
      contextMenu.style.display = "none";
    }

    // Verificar si el clic fue en un cuadro de texto o en un redimensionador
    const textBox = e.target.closest(".excel-text-box");
    const resizer = e.target.closest(".resizer");
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
          top: textBox.offsetTop,
        };
        function onMouseMove(moveEvent) {
          moveEvent.preventDefault();
          textBox.style.left = `${
            startPos.left + (moveEvent.clientX - startPos.x)
          }px`;
          textBox.style.top = `${
            startPos.top + (moveEvent.clientY - startPos.y)
          }px`;
        }
        function onMouseUp() {
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);
        }
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      } else {
        // Clic en el interior del cuadro para editar
        setActiveTextBox(textBox);
        // No prevenir el comportamiento predeterminado aquí para permitir la edición
      }
      return;
    }
    // Si se hizo clic fuera de cualquier cuadro de texto
    if (!textBox && !resizer && !e.target.closest(".textbox-context-menu")) {
      setActiveTextBox(null);
    }
  });

  // Función para establecer un cuadro de texto como activo
  function setActiveTextBox(textBox) {
    // Desactivar el cuadro activo anterior
    if (activeTextBox) {
      activeTextBox.classList.remove("active-text-box");
      // Quitar los manejadores de redimensión
      const resizers = activeTextBox.querySelectorAll(".resizer");
      resizers.forEach((resizer) => {
        resizer.remove();
      });
    }
    // Establecer nuevo activo
    activeTextBox = textBox;
    if (textBox) {
      textBox.classList.add("active-text-box");
      // Añadir controles de redimensión
      setTimeout(() => {
        addResizers(textBox);
      }, 10);
    }
  }

  // Añadir manejadores de redimensión a un cuadro de texto
  function addResizers(textBox) {
    // Primero eliminamos cualquier redimensionador existente
    const existingResizers = textBox.querySelectorAll(".resizer");
    existingResizers.forEach((resizer) => resizer.remove());
    // Posiciones de los redimensionadores
    const positions = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
    // Crear y añadir cada redimensionador
    positions.forEach((pos) => {
      const resizer = document.createElement("div");
      resizer.className = `resizer resizer-${pos}`;
      resizer.dataset.position = pos;

      textBox.appendChild(resizer);
      // Añadir evento de mousedown específico para cada redimensionador
      resizer.addEventListener("mousedown", function (e) {
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
          top: textBox.offsetTop,
        };
        // Variables para el seguimiento del ratón
        let isDragging = true;
        function onMouseMove(moveEvent) {
          if (!isDragging) return;
          moveEvent.preventDefault();
          const dx = moveEvent.clientX - startResize.x;
          const dy = moveEvent.clientY - startResize.y;
          // Aplicar cambios según la posición del redimensionador
          if (pos.includes("e")) {
            textBox.style.width = `${Math.max(startResize.width + dx, 120)}px`;
          }
          if (pos.includes("s")) {
            textBox.style.height = `${Math.max(startResize.height + dy, 80)}px`;
          }
          if (pos.includes("w")) {
            const newWidth = Math.max(startResize.width - dx, 120);
            textBox.style.width = `${newWidth}px`;
            textBox.style.left = `${
              startResize.left + (startResize.width - newWidth)
            }px`;
          }

          if (pos.includes("n")) {
            const newHeight = Math.max(startResize.height - dy, 80);
            textBox.style.height = `${newHeight}px`;
            textBox.style.top = `${
              startResize.top + (startResize.height - newHeight)
            }px`;
          }
        }
        function onMouseUp() {
          isDragging = false;
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);
          document.body.style.cursor = ""; // Restaurar cursor
        }
        // Cambiar cursor durante el redimensionamiento
        document.body.style.cursor = window.getComputedStyle(resizer).cursor;
        // Agregar oyentes a nivel de documento para capturar el movimiento
        // incluso si el mouse se mueve fuera del resizador
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      });
    });
  }

  // Cerrar el menú contextual al hacer clic en cualquier parte fuera del menú
  document.addEventListener("click", function (e) {
    if (!e.target.closest(".textbox-context-menu")) {
      contextMenu.style.display = "none";
    }
  });

  // Mantener la función de eliminación con tecla Delete
  document.addEventListener("keydown", (e) => {
    // Verificar si hay un cuadro de texto activo y que no está en modo edición
    if (
      (e.key === "Delete" || e.key === "Backspace") &&
      activeTextBox &&
      document.activeElement !== activeTextBox
    ) {
      e.preventDefault();
      console.log("Eliminando cuadro de texto:", activeTextBox.id);
      // Eliminar el cuadro de texto
      activeTextBox.remove();
      // Actualizar referencia y lista
      const removedId = activeTextBox.id;
      activeTextBox = null;

      // Actualizar lista de cuadros de texto
      textBoxes = textBoxes.filter((box) => box.id !== removedId);
    }
  });
});

//Desde aqui
