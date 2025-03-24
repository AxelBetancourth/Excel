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



// Función para insertar el año actual o de una fecha
function insertarAño() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        // Si hay una celda seleccionada, intentamos extraer una fecha de ella
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=AÑO(${selectedCells[0].id})`;
    } else {
        // Si no hay celdas seleccionadas o hay más de una
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=AÑO(AHORA())';
    }
}

// Función para insertar el día del mes
function insertarDia() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=DÍA(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=DÍA(AHORA())';
    }
}

// Función para insertar una fecha específica
function insertarFecha() {
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=FECHA(AÑO,MES,DÍA)';
}

// Función para insertar la hora actual
function insertarHora() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=HORA(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=HORA(AHORA())';
    }
}

// Función para insertar la fecha actual
function insertarHoy() {
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=HOY()';
}

// Función para insertar el mes
function insertarMes() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=MES(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=MES(AHORA())';
    }
}

// Función para insertar el minuto
function insertarMinuto() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=MINUTO(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=MINUTO(AHORA())';
    }
}

// Función para insertar el segundo
function insertarSegundo() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=SEGUNDO(${selectedCells[0].id})`;
    } else {
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=SEGUNDO(AHORA())';
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
    }
});

// Funciones para operaciones lógicas

// Función para insertar la función CAMBIAR
function insertarCambiar() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        // Si hay al menos una celda seleccionada
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=CAMBIAR(${selectedCells[0].id},valor_si_verdadero,valor_si_falso)`;
    } else {
        // Si no hay celdas seleccionadas
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=CAMBIAR(expresión_lógica,valor_si_verdadero,valor_si_falso)';
    }
}

// Función para insertar el valor lógico FALSO
function insertarFalso() {
    const formulaInput = document.getElementById('formula-input');
    formulaInput.value = '=FALSO()';
}

// Función para insertar la función NO (negación lógica)
function insertarNo() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length === 1) {
        // Si hay una celda seleccionada
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=NO(${selectedCells[0].id})`;
    } else {
        // Si no hay celdas seleccionadas o hay más de una
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=NO(valor_lógico)';
    }
}

// Función para insertar la función O (OR lógico)
function insertarO() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 2) {
        // Si hay al menos dos celdas seleccionadas
        const celdasIds = Array.from(selectedCells)
            .map(celda => celda.id)
            .join(',');
            
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=O(${celdasIds})`;
    } else {
        // Si no hay suficientes celdas seleccionadas
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=O(valor_lógico1,valor_lógico2,...)';
    }
}

// Función para insertar la función SI (condicional)
function insertarSi() {
    const selectedCells = document.querySelectorAll('.selected');
    
    if (selectedCells.length >= 1) {
        // Si hay al menos una celda seleccionada
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = `=SI(${selectedCells[0].id}=condición,valor_si_verdadero,valor_si_falso)`;
    } else {
        // Si no hay celdas seleccionadas
        const formulaInput = document.getElementById('formula-input');
        formulaInput.value = '=SI(prueba_lógica,valor_si_verdadero,valor_si_falso)';
    }
}

// Función para insertar el valor lógico VERDADERO
function insertarVerdadero() {
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

