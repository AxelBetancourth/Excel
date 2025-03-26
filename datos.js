// Script para manejar la funcionalidad del menú de Datos

document.addEventListener('DOMContentLoaded', function() {
    // Obtener todos los botones del primer nivel
    const botonesNivel1 = document.querySelectorAll('#submenu-datos .menu-btn-nivel1');
    
    // Agregar event listeners a los botones de nivel 1
    botonesNivel1.forEach(boton => {
        boton.addEventListener('click', function(e) {
            e.stopPropagation(); // Evitar que el clic se propague
            
            // Obtener el submenú asociado a este botón (siguiente elemento ul)
            const submenu = this.nextElementSibling;
            
            // Verificar si el submenú existe
            if (submenu && submenu.tagName === 'UL') {
                // Si ya está visible, ocultarlo
                if (submenu.classList.contains('show')) {
                    submenu.classList.remove('show');
                } else {
                    // Primero, ocultar todos los submenús abiertos
                    document.querySelectorAll('#submenu-datos ul.show').forEach(menu => {
                        menu.classList.remove('show');
                    });
                    
                    // Mostrar este submenú
                    submenu.classList.add('show');
                }
            }
        });
    });
    
    // Agregar event listeners a los botones de nivel 2
    const botonesNivel2 = document.querySelectorAll('#submenu-datos .menu-btn-nivel2');
    
    botonesNivel2.forEach(boton => {
        boton.addEventListener('click', function(e) {
            e.stopPropagation(); // Evitar que el clic se propague
            
            // Obtener el submenú asociado a este botón
            const submenu = this.nextElementSibling;
            
            // Verificar si el submenú existe
            if (submenu && submenu.tagName === 'UL') {
                // Si ya está visible, ocultarlo
                if (submenu.classList.contains('show')) {
                    submenu.classList.remove('show');
                } else {
                    // Primero, ocultar todos los submenús de nivel 3
                    const submenusNivel3 = document.querySelectorAll('#submenu-datos .menu-btn-nivel2 + ul');
                    submenusNivel3.forEach(menu => {
                        menu.classList.remove('show');
                    });
                    
                    // Mostrar este submenú
                    submenu.classList.add('show');
                }
            }
        });
    });
    
    // Manejar los botones PDF y Excel
    const btnPDF = document.getElementById('btn-pdf');
    const btnExcel = document.getElementById('btn-excel');
    
    if (btnPDF) {
        btnPDF.addEventListener('click', function(e) {
            e.stopPropagation();
            document.getElementById('archivoPDF').click();
        });
    }
    
    if (btnExcel) {
        btnExcel.addEventListener('click', function(e) {
            e.stopPropagation();
            document.getElementById('archivoExcel').click();
        });
    }
    
    // Cerrar submenús al hacer clic fuera de ellos
    document.addEventListener('click', function(e) {
        if (!e.target.closest('#submenu-datos .container')) {
            const submenus = document.querySelectorAll('#submenu-datos ul.show');
            submenus.forEach(menu => {
                menu.classList.remove('show');
            });
        }
    });
});
