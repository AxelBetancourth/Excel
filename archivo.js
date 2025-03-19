//anadido por CLAUDIO para --> archivo
	
document.addEventListener("DOMContentLoaded", function () {
  const buttons = document.querySelectorAll(".nav-btn");
  const sections = document.querySelectorAll(".contenido-seccion");

  function activarSeccion(target) {
    // Ocultar todas las secciones
    sections.forEach(section => section.classList.remove("activo"));

    // Quitar la clase 'activo' de todos los botones
    buttons.forEach(button => button.classList.remove("activo"));

    // Mostrar la sección correspondiente
    document.getElementById(target).classList.add("activo");

    // Activar el botón correspondiente
    document.querySelector(`.nav-btn[data-target="${target}"]`).classList.add("activo");
  }

  buttons.forEach(button => {
    button.addEventListener("click", function () {
      const target = this.getAttribute("data-target");
      activarSeccion(target);
    });
  });

  // Activar la sección de inicio por defecto
  activarSeccion("inicio");
});

//Ventana modal

document.getElementById('openSaveModal').addEventListener('click', function() {
  document.getElementById('saveModal').style.display = 'flex';  // Cambié 'block' a 'flex' para mejor posicionamiento
});

// Función para cerrar la ventana modal
document.getElementById('cancelSave').addEventListener('click', function() {
  document.getElementById('saveModal').style.display = 'none';
});

// Función para guardar el archivo cuando el usuario haga clic en "Guardar"
document.getElementById('saveFile').addEventListener('click', function() {
  const nombreArchivo = document.getElementById('filename').value.trim() || "datos_guardados";
  const tabla = document.getElementById('spreadsheet');
// Crear array de datos con formato correcto para Excel
const datos = [];
const formulas = {}; // Objeto para almacenar las fórmulas con sus coordenadas en Excel

// Extraer los datos de la tabla
for (let y = 0; y < rows; y++) {
  const filaDatos = [];
  for (let x = 0; x < cols; x++) {
      const celda = state[x][y];

      if (celda.value.startsWith('=')) { // Si es una fórmula
          filaDatos.push(null); // Insertar una celda vacía
          const cellRef = XLSX.utils.encode_cell({ c: x, r: y });
          formulas[cellRef] = celda.value.substring(1); // Quitar '=' y guardar la fórmula
      } else {
          filaDatos.push(celda.computedValue);
      }
  }
  datos.push(filaDatos);
}
// Crear el libro de trabajo de Excel
const wb = XLSX.utils.book_new();

// Extraer correctamente los nombres de las hojas desde el DOM
document.querySelectorAll(".sheet").forEach((sheetElement, index) => {
  sheets[index].name = sheetElement.textContent.trim(); // Guardamos el nombre correcto en sheets
});

sheets.forEach((sheet, sheetIndex) => {
  const datos = [];
  const sheetName = sheet.name || `Hoja${sheetIndex + 1}`;
  const formulas = {}; // Objeto para almacenar las fórmulas

  // Limpiar caracteres inválidos en nombres de hojas de Excel
  const nombreValido = sheetName.replace(/[\[\]:\*\?\/\\]/g, "").trim();
  if (nombreValido.length > 31) nombreValido = nombreValido.substring(0, 31);

  // Asegurar que la hoja tiene datos válidos
  if (!sheet.data || !Array.isArray(sheet.data)) return;

  for (let y = 0; y < rows; y++) {
      const filaDatos = [];
      for (let x = 0; x < cols; x++) {
          const celda = sheet.data[x]?.[y] || { value: "", computedValue: "" };

          if (typeof celda.value === "string" && celda.value.startsWith('=')) {
              filaDatos.push(null); // Dejar la celda vacía
              const cellRef = XLSX.utils.encode_cell({ c: x, r: y }); 
              formulas[cellRef] = celda.value.substring(1); // Guardar la fórmula sin '='
          } else {
              filaDatos.push(celda.computedValue || "");
          }
      }
      datos.push(filaDatos);
  }

  // Crear hoja de Excel con los datos
  const ws = XLSX.utils.aoa_to_sheet(datos);

  // Aplicar manualmente las fórmulas en las celdas correctas
  Object.keys(formulas).forEach(cellRef => {
      if (!ws[cellRef]) ws[cellRef] = {}; // Asegurar que la celda existe
      ws[cellRef].f = formulas[cellRef];  // Asignar la fórmula correctamente
  });

  // Añadir la hoja al libro de trabajo con su nombre correcto
  XLSX.utils.book_append_sheet(wb, ws, nombreValido);
});

// Convertir el libro en binario
const wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'binary' });

// Convertir el binario a Blob
const blob = new Blob([stringToArrayBuffer(wbout)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

// Crear un enlace invisible para la descarga
const enlace = document.createElement('a');
enlace.href = URL.createObjectURL(blob);
enlace.download = nombreArchivo + ".xlsx";

// Simular un clic en el enlace para descargar el archivo
document.body.appendChild(enlace);
enlace.click();
document.body.removeChild(enlace);

  
    // Guardar en archivos recientes
    agregarArchivoReciente(nombreArchivo + ".xlsx");
  
    // Cerrar la ventana modal
    document.getElementById('saveModal').style.display = 'none';
});

// Función para convertir string binario a ArrayBuffer
function stringToArrayBuffer(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

// Función para agregar archivos recientes
function agregarArchivoReciente(nombreArchivo) {
  let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
  
  // Evitar duplicados
  const index = archivos.indexOf(nombreArchivo);
  if (index !== -1) {
    // Si ya existe, quítalo primero para ponerlo al inicio
    archivos.splice(index, 1);
  }
  
  archivos.unshift(nombreArchivo); // Agregar al inicio
  if (archivos.length > 5) archivos.pop(); // Máximo 5 recientes
  localStorage.setItem('file_' + nombreArchivo + ".xlsx", JSON.stringify(state));
  
  mostrarArchivosRecientes();
}

// Función para mostrar archivos recientes con mejor formato
function mostrarArchivosRecientes() {
  let archivos = JSON.parse(localStorage.getItem("archivosRecientes")) || [];
  let lista = document.getElementById("lista-recientes");
  lista.innerHTML = ""; // Limpiar antes de actualizar

  if (archivos.length === 0) {
    let li = document.createElement("li");
    li.className = "no-recent-files";
    li.textContent = "No hay archivos recientes";
    lista.appendChild(li);
    return;
  }

  archivos.forEach(archivo => {
    let li = document.createElement("li");
    li.className = "recent-file-item";
    
    // Crear icono
    let icon = document.createElement("img");
    icon.src = "images/logoexcel.png"; // Asegúrate de tener este icono
    icon.alt = "Excel";
    icon.className = "file-icon";
    
    // Crear nombre de archivo
    let nameSpan = document.createElement("span");
    nameSpan.textContent = archivo;
    nameSpan.className = "file-name";
    
    // Añadir elementos al li
    li.appendChild(icon);
    li.appendChild(nameSpan);
    
    // Añadir evento de clic para abrir el archivo
    li.addEventListener("click", function() {
      // Guarda el estado actual antes de abrir el archivo
      localStorage.setItem("tempState", JSON.stringify(state));
      
      // Abre el archivo seleccionado
      window.location.href = `main.html?archivo=${encodeURIComponent(archivo)}`;
    });
    
    lista.appendChild(li);
  });
}

// Llamar a la función al cargar la página
document.addEventListener("DOMContentLoaded", mostrarArchivosRecientes);

//script para poder abrir una nueva pestana

document.getElementById("btnNuevo").addEventListener("click", function () {
  // Crear un nombre de archivo único
  let nombreArchivo = "Archivo_" + new Date().toISOString().replace(/[:.-]/g, "_") + ".xlsx";

  // Almacenar en archivos recientes
  agregarArchivoReciente(nombreArchivo);

  // Abrir nueva pestaña con main.html pasando el nombre del archivo como parámetro
  window.open(`main.html?archivo=${encodeURIComponent(nombreArchivo)}`, "_blank");
});

//anadido por CLAUDIO para --> archivo

//AÑADIR MAS HOJAS DE CALCULO, aqui debe de ir el codigo que te eh pedido


// Eventos en los botones de la cinta
tabs.forEach(btn => {
  btn.addEventListener('click', () => {
    activateTab(btn.dataset.tab);
  });
});

// Evento para cerrar la barra lateral "Archivo"
sidebarClose.addEventListener('click', () => {
  sidebar.classList.remove('active');
  submenuContainer.style.display = 'block';
});

// Al cargar la página, se activa la pestaña guardada (o "inicio" por defecto)
document.addEventListener('DOMContentLoaded', () => {
  const activeTab = localStorage.getItem('activeTab') || 'inicio';
  activateTab(activeTab);
});

//mostrar mensaje de bienvenida inicio

document.addEventListener("DOMContentLoaded", function () {
  const mensajeBienvenida = document.getElementById("mensaje-bienvenida");
  const hora = new Date().getHours();

  if (hora >= 6 && hora < 12) {
      mensajeBienvenida.textContent = "Buenos días";
  } else if (hora >= 12 && hora < 18) {
      mensajeBienvenida.textContent = "Buenas tardes";
  } else {
      mensajeBienvenida.textContent = "Buenas noches";
  }
});

//Funcion para el boton abrir

function abrirExcel() {
  let fileInput = document.getElementById('fileInput');
  let file = fileInput.files[0]; // Obtener el archivo seleccionado

  if (!file) {
      alert("Por favor, selecciona un archivo de Excel.");
      return;
  }

  let reader = new FileReader();

  reader.onload = function (e) {
      let data = new Uint8Array(e.target.result);
      let workbook = XLSX.read(data, { type: 'array' });

      let sheetName = workbook.SheetNames[0]; // Tomar la primera hoja
      let sheet = workbook.Sheets[sheetName];

      let htmlTable = XLSX.utils.sheet_to_html(sheet); // Convertir la hoja a HTML
      document.getElementById('excelTable').innerHTML = htmlTable;
  };

  reader.readAsArrayBuffer(file); // Leer el archivo como ArrayBuffer
}


