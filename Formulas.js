// Función para calcular la suma de celdas seleccionadas
function sumarCeldas() {
    // Obtener todas las celdas seleccionadas
    const selectedCells = document.getSelectedRange('.selected');
    
    let suma = 0;
    
    selectedCells.forEach(celda => {
        // Obtener el valor numérico de la celda
        const valor = parseFloat(celda.textContent) || 0;
        suma += valor;
    });
    
    // Mostrar el resultado en la barra de fórmulas
    const formulaInput = document.getSelectedRangell('formula-input');
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

