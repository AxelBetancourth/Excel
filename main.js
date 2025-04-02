	// Elementos de la cinta y barra lateral
	const tabs = document.querySelectorAll('.ribbon-tabs button');
	const submenus = document.querySelectorAll('.submenu');
	const submenuContainer = document.getElementById('submenu-container');
	const sidebar = document.getElementById('archivo-sidebar');
	const sidebarClose = document.getElementById('sidebar-close');
	
	// Elementos de la barra de fórmulas y hojas
	const formulaInput = document.getElementById('formula-input');
	const activeCellDisplay = document.getElementById('active-cell');
	const sheetsContainer = document.querySelector('.sheets-container');
	const addSheetButton = document.getElementById('add-sheet');
	
	/**
	 * Activa la pestaña indicada. Si es "archivo" muestra la barra lateral.
	 */
	function activateTab(tabName) {
	  if (tabName === 'archivo') {
	    sidebar.classList.add('active');
	    submenuContainer.style.display = 'none';
	  } else {
	    sidebar.classList.remove('active');
	    submenuContainer.style.display = 'block';
	    tabs.forEach(btn => {
	      btn.classList.toggle('active', btn.dataset.tab === tabName);
	    });
	    submenus.forEach(menu => {
	      menu.classList.toggle('active', menu.id === 'submenu-' + tabName);
	    });
	    localStorage.setItem('activeTab', tabName);
	  }
	}
	
	
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
  
  // Evento para confirmar antes de recargar o cerrar la página
  window.addEventListener('beforeunload', function(e) {
    // Mensaje para mostrar al usuario
    const mensajeConfirmacion = '¿Estás seguro que deseas salir? Los cambios no guardados se perderán.';
    
    // Configurar el mensaje de retorno (requerido para la mayoría de navegadores)
    e.returnValue = mensajeConfirmacion;
    
    // Para compatibilidad con navegadores antiguos
    return mensajeConfirmacion;
  });
  
  // Al cargar la página, se activa la pestaña guardada (o "inicio" por defecto)
  document.addEventListener('DOMContentLoaded', () => {
	const activeTab = localStorage.getItem('activeTab') || 'inicio';
	activateTab(activeTab);
	
	// Inicializar la hoja de cálculo
	renderSpreadsheet();
	
	// Cargar archivo si viene por parámetro URL
	const params = new URLSearchParams(window.location.search);
	const archivo = params.get("archivo");
	if (archivo) {
	  const savedState = localStorage.getItem('file_' + archivo);
	  if (savedState) {
		state = JSON.parse(savedState);
		sheets[currentSheetIndex].data = state;
		renderSpreadsheet();
	  }
	}
  });

//Funcion para la pantalla de carga
window.addEventListener('load', function() {
	var progressBar = document.getElementById('progress-bar');
	var progressPercentage = document.getElementById('progress-percentage');
	var loader = document.getElementById('loader');
  
	var progress = 0;
	var totalDuration = 1000;
	var intervalTime = 30;
	var increment = (intervalTime / totalDuration) * 100;
  
	var interval = setInterval(function() {
	  progress += increment;
	  if (progress >= 100) {
		progress = 100;
		clearInterval(interval);
		setTimeout(function(){
		  loader.style.opacity = '0';
		  setTimeout(function(){
			loader.style.display = 'none';
		  }, 500);
		}, 500);
	  }
	  progressBar.style.width = progress + '%';
	  progressPercentage.textContent = Math.floor(progress) + '%';
	}, intervalTime);
});
  
  // Exportar funciones y variables que necesiten ser accesibles desde otros módulos
  window.activateTab = activateTab;
	
	