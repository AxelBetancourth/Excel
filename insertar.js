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
      insertarForma.style.backgroundColor = "lightblue";
    } else if (shapeType === "triangle") {
      insertarForma.style.width = "0";
      insertarForma.style.height = "0";
      insertarForma.style.borderLeft = "50px solid transparent";
      insertarForma.style.borderRight = "50px solid transparent";
      insertarForma.style.borderBottom = "100px solid lightblue";
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

// Funcionalidad para gráficos
function initChartFunctionality() {
    const chartModal = document.getElementById('chart-modal');
    const btnGraficoBarras = document.getElementById('btn-grafico-barras');
    const btnGraficoLineas = document.getElementById('btn-grafico-lineas');
    const btnGraficoCircular = document.getElementById('btn-grafico-circular');
    const createChartBtn = document.getElementById('create-chart');
    const previewCanvas = document.getElementById('preview-canvas');
    let currentChartType = '';
    let previewChart = null;

    // Función para mostrar el modal y limpiar la vista previa
    function showChartModal(type) {
        currentChartType = type;
        chartModal.style.display = 'flex';

        // Limpiar campos y vista previa
        document.getElementById('chart-title').value = '';
        document.getElementById('data-range').value = '';
        clearPreview();
    }

    // Función para limpiar la vista previa
    function clearPreview() {
        if (previewChart) {
            previewChart.destroy();
            previewChart = null;
        }
        const ctx = previewCanvas.getContext('2d');
        ctx.clearRect(0, 0, previewCanvas.width, previewCanvas.height);
    }

    // Función para actualizar la vista previa
    function updatePreview() {
        const range = document.getElementById('data-range').value;
        const title = document.getElementById('chart-title').value;

        if (!isValidRange(range)) {
            console.error('Rango inválido. Asegúrate de ingresar un rango válido como A1:B5.');
            return;
        }

        try {
            const data = getDataFromRange(range);
            const ctx = previewCanvas.getContext('2d');

            // Limpiar el gráfico previo
            clearPreview();

            // Crear el nuevo gráfico
            previewChart = new Chart(ctx, {
                type: currentChartType,
                data: {
                    labels: data.labels,
                    datasets: [{
                        label: title || 'Datos',
                        data: data.values,
                        backgroundColor: [
                            'rgba(255, 99, 132, 0.5)',
                            'rgba(54, 162, 235, 0.5)',
                            'rgba(255, 206, 86, 0.5)',
                            'rgba(75, 192, 192, 0.5)',
                            'rgba(153, 102, 255, 0.5)'
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)',
                            'rgba(75, 192, 192, 1)',
                            'rgba(153, 102, 255, 1)'
                        ],
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    scales: currentChartType === 'pie' ? {} : {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: data.yAxisLabel
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: data.xAxisLabel
                            }
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: title || 'Vista previa del gráfico'
                        },
                        legend: {
                            display: currentChartType === 'pie',
                            position: 'bottom'
                        }
                    }
                }
            });
        } catch (error) {
            console.error('Error al actualizar la vista previa:', error);
        }
    }

    // Función para crear el gráfico en la hoja
    function createChart() {
        const range = document.getElementById('data-range').value;
        const title = document.getElementById('chart-title').value;

        if (!isValidRange(range)) {
            alert('Por favor, ingrese un rango válido como A1:B5.');
            return;
        }

        try {
            const data = getDataFromRange(range);
            const chartContainer = document.createElement('div');
            chartContainer.className = 'chart-container';
            chartContainer.dataset.sheetIndex = currentSheetIndex; // Asociar con la hoja actual
            chartContainer.style.position = 'absolute';
            chartContainer.style.width = '400px';
            chartContainer.style.height = '300px';
            chartContainer.style.backgroundColor = 'white';
            chartContainer.style.padding = '10px';
            chartContainer.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
            chartContainer.style.cursor = 'move';
            chartContainer.style.left = '50px';
            chartContainer.style.top = '50px';
            chartContainer.style.zIndex = '1000';

            // Contenedor para el canvas
            const canvasContainer = document.createElement('div');
            canvasContainer.style.width = '100%';
            canvasContainer.style.height = '100%';
            chartContainer.appendChild(canvasContainer);

            const canvas = document.createElement('canvas');
            canvas.style.width = '100%';
            canvas.style.height = '100%';
            canvasContainer.appendChild(canvas);

            // Botón de cerrar
            const closeBtn = document.createElement('button');
            closeBtn.innerHTML = '×';
            closeBtn.className = 'chart-close-btn';
            closeBtn.style.position = 'absolute';
            closeBtn.style.right = '5px';
            closeBtn.style.top = '5px';
            closeBtn.style.zIndex = '1001';
            closeBtn.onclick = () => chartContainer.remove();
            chartContainer.appendChild(closeBtn);

            // Controladores de redimensionamiento
            const resizers = document.createElement('div');
            resizers.className = 'resizers';
            ['nw', 'ne', 'sw', 'se'].forEach(pos => {
                const handle = document.createElement('div');
                handle.className = `resizer ${pos}`;
                resizers.appendChild(handle);
            });
            chartContainer.appendChild(resizers);

            document.querySelector('.spreadsheet-container').appendChild(chartContainer);

            // Crear el gráfico
            new Chart(canvas, {
                type: currentChartType,
                data: {
                    labels: data.labels,
                    datasets: [{
                        label: title || 'Datos',
                        data: data.values,
                        backgroundColor: [
                            'rgba(255, 99, 132, 0.5)',
                            'rgba(54, 162, 235, 0.5)',
                            'rgba(255, 206, 86, 0.5)',
                            'rgba(75, 192, 192, 0.5)',
                            'rgba(153, 102, 255, 0.5)'
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)',
                            'rgba(75, 192, 192, 1)',
                            'rgba(153, 102, 255, 1)'
                        ],
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: currentChartType === 'pie' ? {} : {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: data.yAxisLabel
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: data.xAxisLabel
                            }
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: title || 'Gráfico'
                        },
                        legend: {
                            display: currentChartType === 'pie',
                            position: 'bottom'
                        }
                    }
                }
            });

            // Hacer el gráfico arrastrable
            makeChartDraggable(chartContainer);
            makeChartResizable(chartContainer);

            // Cerrar el modal después de crear el gráfico
            chartModal.style.display = 'none';
        } catch (error) {
            console.error('Error al crear el gráfico:', error);
            alert('Error al crear el gráfico. Verifique el rango de datos.');
        }
    }

    function makeChartDraggable(element) {
        let isDragging = false;
        let currentX;
        let currentY;
        let initialX;
        let initialY;

        element.onmousedown = dragStart;

        function dragStart(e) {
            if (e.target.classList.contains('resizer') || e.target.classList.contains('chart-close-btn')) {
                return;
            }

            isDragging = true;
            initialX = e.clientX - element.offsetLeft;
            initialY = e.clientY - element.offsetTop;

            document.onmousemove = drag;
            document.onmouseup = dragEnd;
        }

        function drag(e) {
            if (isDragging) {
                e.preventDefault();
                currentX = e.clientX - initialX;
                currentY = e.clientY - initialY;

                // Obtener la celda A1 y Z100 para definir los límites
                const cellA1 = document.querySelector('td[data-x="0"][data-y="0"]');
                const cellZ100 = document.querySelector(`td[data-x="${cols-1}"][data-y="${rows-1}"]`);

                if (!cellA1 || !cellZ100) return;

                // Obtener los límites exactos del rango de celdas
                const tableLimits = {
                    left: cellA1.getBoundingClientRect().left,
                    top: cellA1.getBoundingClientRect().top,
                    right: cellZ100.getBoundingClientRect().right,
                    bottom: cellZ100.getBoundingClientRect().bottom
                };

                // Obtener el contenedor de la hoja de cálculo
                const spreadsheetContainer = document.querySelector('.spreadsheet-container');
                const containerRect = spreadsheetContainer.getBoundingClientRect();

                // Calcular los límites relativos al contenedor
                const minX = tableLimits.left - containerRect.left + spreadsheetContainer.scrollLeft;
                const minY = tableLimits.top - containerRect.top + spreadsheetContainer.scrollTop;
                const maxX = tableLimits.right - containerRect.left + spreadsheetContainer.scrollLeft - element.offsetWidth;
                const maxY = tableLimits.bottom - containerRect.top + spreadsheetContainer.scrollTop - element.offsetHeight;

                // Restringir el movimiento dentro del rango A1-Z100
                currentX = Math.max(minX, Math.min(currentX, maxX));
                currentY = Math.max(minY, Math.min(currentY, maxY));

                // Aplicar las nuevas posiciones
                element.style.left = `${currentX}px`;
                element.style.top = `${currentY}px`;
            }
        }

        function dragEnd() {
            isDragging = false;
            document.onmousemove = null;
            document.onmouseup = null;
        }
    }

    function makeChartResizable(element) {
        const resizers = element.querySelectorAll('.resizer');
        let currentResizer;
        let isResizing = false;
        let originalWidth;
        let originalHeight;
        let originalX;
        let originalY;
        let originalMouseX;
        let originalMouseY;

        resizers.forEach(resizer => {
            resizer.addEventListener('mousedown', initResize);
        });

        function initResize(e) {
            isResizing = true;
            currentResizer = e.target;

            originalWidth = parseFloat(element.style.width);
            originalHeight = parseFloat(element.style.height);
            originalX = element.offsetLeft;
            originalY = element.offsetTop;
            originalMouseX = e.pageX;
            originalMouseY = e.pageY;

            document.addEventListener('mousemove', resize);
            document.addEventListener('mouseup', stopResize);
        }

        function resize(e) {
            if (!isResizing) return;

            const minSize = 200; // Tamaño mínimo en píxeles
            let newWidth = originalWidth;
            let newHeight = originalHeight;
            let newX = originalX;
            let newY = originalY;

            const dx = e.pageX - originalMouseX;
            const dy = e.pageY - originalMouseY;

            if (currentResizer.classList.contains('se')) {
                newWidth = originalWidth + dx;
                newHeight = originalHeight + dy;
            } else if (currentResizer.classList.contains('sw')) {
                newWidth = originalWidth - dx;
                newHeight = originalHeight + dy;
                newX = originalX + dx;
            } else if (currentResizer.classList.contains('ne')) {
                newWidth = originalWidth + dx;
                newHeight = originalHeight - dy;
                newY = originalY + dy;
            } else if (currentResizer.classList.contains('nw')) {
                newWidth = originalWidth - dx;
                newHeight = originalHeight - dy;
                newX = originalX + dx;
                newY = originalY + dy;
            }

            // Aplicar límites mínimos
            if (newWidth >= minSize) {
                element.style.width = newWidth + 'px';
                if (newX !== originalX) element.style.left = newX + 'px';
            }
            if (newHeight >= minSize) {
                element.style.height = newHeight + 'px';
                if (newY !== originalY) element.style.top = newY + 'px';
            }
        }

        function stopResize() {
            isResizing = false;
            document.removeEventListener('mousemove', resize);
            document.removeEventListener('mouseup', stopResize);
        }
    }

    // Event listeners para los botones de gráficos
    btnGraficoBarras?.addEventListener('click', () => showChartModal('bar'));
    btnGraficoLineas?.addEventListener('click', () => showChartModal('line'));
    btnGraficoCircular?.addEventListener('click', () => showChartModal('pie'));

    // Event listener para el rango de datos
    document.getElementById('data-range')?.addEventListener('input', updatePreview);

    // Event listener para crear el gráfico
    createChartBtn?.addEventListener('click', createChart);
}

// Inicializar la funcionalidad de gráficos
document.addEventListener('DOMContentLoaded', initChartFunctionality);

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

                // Obtener la celda A1 y Z100 para definir los límites
                const cellA1 = document.querySelector('td[data-x="0"][data-y="0"]');
                const cellZ100 = document.querySelector(`td[data-x="${cols-1}"][data-y="${rows-1}"]`);

                if (!cellA1 || !cellZ100) return;

                // Obtener los límites exactos del rango de celdas
                const tableLimits = {
                    left: cellA1.getBoundingClientRect().left,
                    top: cellA1.getBoundingClientRect().top,
                    right: cellZ100.getBoundingClientRect().right,
                    bottom: cellZ100.getBoundingClientRect().bottom
                };

                // Obtener el contenedor de la hoja de cálculo
                const spreadsheetContainer = document.querySelector('.spreadsheet-container');
                const containerRect = spreadsheetContainer.getBoundingClientRect();

                // Calcular los límites relativos al contenedor
                const minX = tableLimits.left - containerRect.left + spreadsheetContainer.scrollLeft;
                const minY = tableLimits.top - containerRect.top + spreadsheetContainer.scrollTop;
                const maxX = tableLimits.right - containerRect.left + spreadsheetContainer.scrollLeft - container.offsetWidth;
                const maxY = tableLimits.bottom - containerRect.top + spreadsheetContainer.scrollTop - container.offsetHeight;

                // Restringir el movimiento dentro del rango A1-Z100
                currentX = Math.max(minX, Math.min(currentX, maxX));
                currentY = Math.max(minY, Math.min(currentY, maxY));

                // Aplicar las nuevas posiciones
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

function getDataFromRange(range) {
    const [start, end] = range.split(':').map(cell => getCellCoords(cell));
    const labels = [];
    const values = [];
    let xAxisLabel = '';
    let yAxisLabel = '';

    // Validar que las coordenadas sean válidas
    if (!start || !end) {
        throw new Error('Rango inválido. Asegúrate de ingresar un rango válido como A1:B5.');
    }

    // Determinar si el rango incluye encabezados
    const hasHeaders = start[1] === 0; // Si la primera fila es un encabezado

    if (hasHeaders) {
        // Usar la primera fila como etiquetas de los ejes
        xAxisLabel = state[start[0]][start[1]].computedValue; // Encabezado del eje X
        yAxisLabel = state[end[0]][start[1]].computedValue; // Encabezado del eje Y

        // Extraer datos desde la segunda fila
        for (let i = start[1] + 1; i <= end[1]; i++) {
            const label = state[start[0]][i].computedValue; // Etiqueta del eje X
            const value = state[end[0]][i].computedValue; // Valor del eje Y
            if (!isNaN(value)) {
                labels.push(label);
                values.push(parseFloat(value));
            }
        }
    } else {
        // Si no hay encabezados, usar índices como etiquetas
        for (let i = start[1]; i <= end[1]; i++) {
            const label = state[start[0]][i].computedValue; // Etiqueta del eje X
            const value = state[end[0]][i].computedValue; // Valor del eje Y
            if (!isNaN(value)) {
                labels.push(label);
                values.push(parseFloat(value));
            }
        }
        xAxisLabel = 'Eje X';
        yAxisLabel = 'Eje Y';
    }

    return { labels, values, xAxisLabel, yAxisLabel };
}

function isValidRange(range) {
    return /^[A-Z]+\d+:[A-Z]+\d+$/.test(range);
}

// Agregar evento para mostrar/ocultar gráficos al cambiar de hoja
document.addEventListener('sheetChanged', function() {
    document.querySelectorAll('.chart-container').forEach(container => {
        if (parseInt(container.dataset.sheetIndex) === currentSheetIndex) {
            container.style.display = 'block';
        } else {
            container.style.display = 'none';
        }
    });
});








