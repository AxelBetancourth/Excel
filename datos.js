// Script para manejar la funcionalidad del menú de Datos

document.addEventListener('DOMContentLoaded', function() {
    // Función para posicionar todos los elementos a la izquierda
    function moverElementosIzquierda() {
        // Mover el menú principal a la izquierda
        const submenuDatos = document.getElementById('submenu-datos');
        if (submenuDatos) {
            submenuDatos.style.flexDirection = 'row';
            submenuDatos.style.justifyContent = 'flex-start';
            submenuDatos.style.flexWrap = 'nowrap';
            submenuDatos.style.gap = '0';
            
            // Configurar el wrapper
            const menuWrapper = submenuDatos.querySelector('.menu-wrapper');
            if (menuWrapper) {
                menuWrapper.style.display = 'flex';
                menuWrapper.style.flexDirection = 'row';
                menuWrapper.style.marginLeft = '10px';
            }
            
            // Mover el contenedor a la izquierda de la página
            const container = submenuDatos.querySelector('.container');
            if (container) {
                container.style.marginRight = '0';
                container.style.borderTopRightRadius = '0';
                container.style.borderBottomRightRadius = '0';
                container.style.display = 'flex';
                container.style.alignItems = 'center';
                
                // Centrar verticalmente los elementos dentro del contenedor
                const ul = container.querySelector('ul');
                if (ul) {
                    ul.style.display = 'flex';
                    ul.style.flexDirection = 'column';
                    ul.style.justifyContent = 'center';
                }
            }
            
            // Mover los botones de ordenar junto al contenedor principal
            const botonesOrdenar = submenuDatos.querySelector('.BotonesOrdenar');
            if (botonesOrdenar) {
                botonesOrdenar.style.marginLeft = '0';
                botonesOrdenar.style.borderLeft = 'none';
                botonesOrdenar.style.borderTopLeftRadius = '0';
                botonesOrdenar.style.borderBottomLeftRadius = '0';
            }
        }
    }
    
    // Llamar a la función al cargar la página
    moverElementosIzquierda();
    
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
                    
                    // Posicionar a la derecha del botón
                    submenu.style.left = '100%';
                    submenu.style.right = 'auto';
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
                    
                    // Posicionar a la derecha del botón
                    submenu.style.left = '100%';
                    submenu.style.right = 'auto';
                }
            }
        });
    });
    
    // Función para ajustar la posición del submenú si se sale de la pantalla
    function ajustarPosicionSubmenu(submenu) {
        // Esta función ya no se usa, pero la mantenemos por si se necesita en el futuro
    }
    
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
