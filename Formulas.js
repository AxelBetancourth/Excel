// Función para calcular la suma de celdas seleccionadas
function sumarCeldas() {
    // Obtener todas las celdas seleccionadas
    const selectedCells = document.querySelectorAll('.selected');
    
    let suma = 0;
    
    selectedCells.forEach(celda => {
        // Obtener el valor numérico de la celda
        const valor = parseFloat(celda.textContent) || 0;
        suma += valor;
    });
    
    // Mostrar el resultado en la barra de fórmulas
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = `=SUMA(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
}


// Función para calcular el promedio de celdas seleccionadas
function calcularPromedio() {
    const selectedCells = document.querySelectorAll('.selected');
    let suma = 0;
    let conteo = 0;

    selectedCells.forEach(celda => {
        const valor = parseFloat(celda.textContent) || 0;
        suma += valor;
        conteo++;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = `=PROMEDIO(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
}

// Función para contar números de celdas seleccionadas
function contarNumeros() {
    const selectedCells = document.querySelectorAll('.selected');
    let conteo = 0;

    selectedCells.forEach(celda => {
        const valor = parseFloat(celda.textContent);
        if (!isNaN(valor)) {
            conteo++;
        }
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = `=CONTAR(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
}

// Función para encontrar el máximo de celdas seleccionadas
function encontrarMax() {
    const selectedCells = document.querySelectorAll('.selected');
    let max = -Infinity;

    selectedCells.forEach(celda => {
        const valor = parseFloat(celda.textContent) || 0;
        if (valor > max) {
            max = valor;
        }
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = `=MAX(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
}

// Función para encontrar el mínimo de celdas seleccionadas
function encontrarMin() {
    const selectedCells = document.querySelectorAll('.selected');
    let min = Infinity;

    selectedCells.forEach(celda => {
        const valor = parseFloat(celda.textContent) || 0;
        if (valor < min) {
            min = valor;
        }
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = `=MIN(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
}

// Evento para el dropdown de operaciones
document.querySelector('.dropdown').addEventListener('click', function(e) {
    const texto = e.target.textContent;
    
    switch(texto) {
        case 'Suma':
            sumarCeldas();
            break;
        case 'Promedio':
            calcularPromedio();
            break;
        case 'Contar números':
            contarNumeros();
            break;
        case 'Máx':
            encontrarMax();
            break;
        case 'Mín':
            encontrarMin();
            break;
    }
});

// Funciones para fecha y hora


// Función para insertar la fecha y hora actual (AHORA)
function insertarAhora() {
    const selectedCells = document.querySelectorAll('.selected');
   
    const fechaHora = new Date().toLocaleString(); // Obtener la fecha y hora actual
    
    selectedCells.forEach(celda => {
        celda.textContent = fechaHora; // Insertar el valor visible en la celda
    });

    // Mostrar la función en la barra de fórmulas
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=AHORA()';
}

function insertarAño() {
    const selectedCells = document.querySelectorAll('.selected');
   
    const añoActual = new Date().getFullYear(); // Obtener solo el año actual
    
    selectedCells.forEach(celda => {
        celda.textContent = añoActual; // Insertar el valor visible en la celda
    });

    // Mostrar la función en la barra de fórmulas
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=AÑO(AHORA())';
}

function insertarDia() {
    const selectedCells = document.querySelectorAll('.selected');
    const diaActual = new Date().getDate();
    
    selectedCells.forEach(celda => {
        celda.textContent = diaActual;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=DIA(AHORA())';
}

function insertarFecha() {
    const selectedCells = document.querySelectorAll('.selected');
    const fecha = new Date().toLocaleDateString();
    
    selectedCells.forEach(celda => {
        celda.textContent = fecha;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=FECHA(AHORA())';
}

function insertarHora() {
    const selectedCells = document.querySelectorAll('.selected');
    const hora = new Date().getHours();
    
    selectedCells.forEach(celda => {
        celda.textContent = hora;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=HORA(AHORA())';
}

function insertarHoy() {
    const selectedCells = document.querySelectorAll('.selected');
    const hoy = new Date().toLocaleDateString();
    
    selectedCells.forEach(celda => {
        celda.textContent = hoy;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=HOY()';
}

function insertarMes() {
    const selectedCells = document.querySelectorAll('.selected');
    const mesActual = new Date().getMonth() + 1; // getMonth() devuelve 0-11
    
    selectedCells.forEach(celda => {
        celda.textContent = mesActual;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=MES(AHORA())';
}

function insertarMinuto() {
    const selectedCells = document.querySelectorAll('.selected');
    const minutoActual = new Date().getMinutes();
    
    selectedCells.forEach(celda => {
        celda.textContent = minutoActual;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=MINUTO(AHORA())';
}

function insertarSegundo() {
    const selectedCells = document.querySelectorAll('.selected');
    const segundoActual = new Date().getSeconds();
    
    selectedCells.forEach(celda => {
        celda.textContent = segundoActual;
    });

    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=SEGUNDO(AHORA())';
}

function insertarTiempoTranscurrido() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 2) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=TIEMPO_TRANSCURRIDO(${selectedCells[0].id},${selectedCells[1].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=TIEMPO_TRANSCURRIDO(Tiempo_inicial,Tiempo_final)';
    }
}

// Evento para el dropdown de fecha y hora
document.querySelector('.dropdown-content-fecha').addEventListener('click', function(e) {
    e.preventDefault(); // Prevenir el comportamiento predeterminado del enlace
    
    const texto = e.target.textContent;
    
    switch(texto) {
        case 'AHORA':
            insertarAhora();
            break;
        case 'AÑO':
            insertarAño();
            break;
        case 'DÍA':
            insertarDia();
            break;
        case 'FECHA':
            insertarFecha();
            break;
        case 'HORA':
            insertarHora();
            break;
        case 'HOY':
            insertarHoy();
            break;
        case 'MES':
            insertarMes();
            break;
        case 'MINUTO':
            insertarMinuto();
            break;
        case 'SEGUNDO':
            insertarSegundo();
            break;
        case 'TIEMPO TRANSCURRIDO':
            insertarTiempoTranscurrido();
            break;
    }
});

// Funciones para operaciones lógicas

// Función para insertar la función CAMBIAR
function insertarCambiar() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=CAMBIAR(${selectedCells[0].id},"valor_si_verdadero","valor_si_falso")`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=CAMBIAR(expresión_lógica,"valor_si_verdadero","valor_si_falso")';
    }
}

// Función para insertar el valor lógico FALSO
function insertarFalso() {
    const selectedCells = document.querySelectorAll('.selected');
    selectedCells.forEach(celda => {
        celda.textContent = "FALSO";
    });
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=FALSO()';
}

// Función para insertar la función NO (negación lógica)
function insertarNo() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=NO(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=NO(valor_lógico)';
    }
}

// Función para insertar la función O (OR lógico)
function insertarO() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 2) {
        const celdasIds = Array.from(selectedCells)
            .map(celda => celda.id)
            .join(',');
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=O(${celdasIds})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=O(valor_lógico1,valor_lógico2,...)';
    }
}

// Función para insertar la función SI (condicional)
function insertarSi() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=SI(${selectedCells[0].id},"valor_si_verdadero","valor_si_falso")`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=SI(prueba_lógica,"valor_si_verdadero","valor_si_falso")';
    }
}

// Función para insertar el valor lógico VERDADERO
function insertarVerdadero() {
    const selectedCells = document.querySelectorAll('.selected');
    selectedCells.forEach(celda => {
        celda.textContent = "VERDADERO";
    });
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=VERDADERO()';
}

// Evento para el dropdown de funciones lógicas
document.querySelector('.dropdown-content').addEventListener('click', function(e) {
    e.preventDefault(); // Prevenir el comportamiento predeterminado del enlace
    
    const texto = e.target.textContent;
    
    switch(texto) {
        case 'CAMBIAR':
            insertarCambiar();
            break;
        case 'FALSO':
            insertarFalso();
            break;
        case 'NO':
            insertarNo();
            break;
        case 'O':
            insertarO();
            break;
        case 'SI':
            insertarSi();
            break;
        case 'VERDADERO':
            insertarVerdadero();
            break;
    }
});

// Funciones para operaciones de texto
function insertarConcatenar() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 2) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=CONCATENAR(${Array.from(selectedCells).map(celda => celda.id).join(',')})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=CONCATENAR("texto1","texto2")';
    }
}

function insertarMayusculas() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=MAYUSCULAS(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=MAYUSCULAS(texto)';
    }
}

function insertarMinusculas() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=MINUSCULAS(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=MINUSCULAS(texto)';
    }
}

function insertarLargo() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=LARGO(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=LARGO(texto)';
    }
}

function insertarReemplazar() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=REEMPLAZAR(${selectedCells[0].id}Texto,posición_inicial,número_caracteres,"texto_nuevo")`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=REEMPLAZAR(texto,posición_inicial,número_caracteres,"texto_nuevo")';
    }
}

function insertarExtrae() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=EXTRAE(${selectedCells[0].id}Texto;posición_inicial;posición_final)`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=EXTRAE(Texto;posición_inicial;posición_final)';
    }
}

// Evento para el dropdown de texto
document.querySelector('.dropdown-content-texto').addEventListener('click', function(e) {
    e.preventDefault();
    
    const texto = e.target.textContent;
    
    switch(texto) {
        case 'CONCATENAR':
            insertarConcatenar();
            break;
        case 'MAYUSCULAS':
            insertarMayusculas();
            break;
        case 'MINUSCULAS':
            insertarMinusculas();
            break;
        case 'LARGO':
            insertarLargo();
            break;
        case 'REEMPLAZAR':
            insertarReemplazar();
            break;
        case 'EXTRAE':
            insertarExtrae();
            break;
    }
});
