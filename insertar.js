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
