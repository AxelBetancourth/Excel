// Aqui ira el codigo javascrip del main

const tabs = document.querySelectorAll('.ribbon-tabs button');
const submenus = document.querySelectorAll('.submenu');
const submenuContainer = document.getElementById('submenu-container');
const sidebar = document.getElementById('archivo-sidebar');
const sidebarClose = document.getElementById('sidebar-close');

/**
* Activa la pestaña solicitada. Si es "archivo", muestra la barra lateral.
* Para las demás pestañas, muestra el submenú correspondiente.
*/

function activateTab(tabName){
    if(tabName === 'archivo'){
        sidebar.classList.add('active');
        // Oculta el submenu cuando se abre el archivo
        submenuContainer.style.display = 'none';
    } else {
        // Asegurarse de ocultar la barra lateral si estuviera abierta
        sidebar.classList.remove('active');
        submenuContainer.style.display = 'block';
        // Actualizar la clase active en los botones
        tabs.forEach(btn => {
            if(btn.dataset.tab === tabName){
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        // Mostrar el submenu correspondiente y ocultar los demas
        submenus.forEach(menu => {
            if(menu.id === 'submenu-' + tabName){
                menu.classList.add('active');
            } else {
                menu.classList.remove('active');
            }
        });

        // Guardar la pestaña seleccionada para mantenerla al recargar
        localStorage.setItem('activeTab', tabName);
    }
}

// Asignamos el evento a cada botón de la cinta

tabs.forEach(btn => {
    btn.addEventListener('click', () => {
        const tab = btn.dataset.tab;
        activateTab(tab);
    });
}
);


// Botón para cerrar la barra lateral y regresar al contenido principal

sidebarClose.addEventListener('click', () => {
    sidebar.classList.remove('active');
    submenuContainer.style.display = 'block';
}
);

// Al cargar la página, se lee la pestaña activa almacenada o se muestra "inicio" por defecto

document.addEventListener('DOMContentLoaded', () => {
    const activeTab = localStorage.getItem('activeTab') || 'inicio';
    activateTab(activeTab);
}
);