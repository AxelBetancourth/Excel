document.addEventListener('DOMContentLoaded', () => {
    const borderSelect = document.getElementById('border-select');
    const borderColor = document.getElementById('border-color');
    const defaultBorderColor = '#e2e2e2'; // Color por defecto
    const defaultBorderWidth = '1px'; // Grosor por defecto
    
    borderSelect.addEventListener('change', function(e) {
        const selectedCells = document.querySelectorAll('.selected-range, .celdailumida');
        
        if (selectedCells.length === 0) {
            alert('Seleccione una celda o rango de celdas primero');
            this.value = '';
            return;
        }

        const borderType = this.value;
        const color = borderColor.value;

        selectedCells.forEach(cell => {
            // Remover la clase de selección que causa el resaltado verde
            cell.classList.remove('cell-with-border');
            
            switch(borderType) {
                case 'all':
                    // Aplicar todos los bordes individualmente
                    cell.style.borderTop = `2px solid ${color}`;
                    cell.style.borderRight = `2px solid ${color}`;
                    cell.style.borderBottom = `2px solid ${color}`;
                    cell.style.borderLeft = `2px solid ${color}`;
                    break;
                case 'top':
                    cell.style.borderTop = `2px solid ${color}`;
                    // Mantener los otros bordes con el estilo por defecto
                    cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    break;
                case 'right':
                    cell.style.borderRight = `2px solid ${color}`;
                    cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    break;
                case 'bottom':
                    cell.style.borderBottom = `2px solid ${color}`;
                    cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    break;
                case 'left':
                    cell.style.borderLeft = `2px solid ${color}`;
                    cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    break;
                case 'none':
                    // Restaurar todos los bordes al estilo por defecto
                    cell.style.borderTop = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderRight = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderBottom = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    cell.style.borderLeft = `${defaultBorderWidth} solid ${defaultBorderColor}`;
                    break;
            }
        });
    });

    // Evento para cambiar el color del borde
    borderColor.addEventListener('change', function(e) {
        const selectedCells = document.querySelectorAll('.selected-range, .celdailumida');
        const borderType = borderSelect.value;
        
        if (selectedCells.length === 0) {
            alert('Seleccione una celda o rango de celdas primero');
            return;
        }

        const color = this.value;
        selectedCells.forEach(cell => {
            cell.classList.remove('cell-with-border');
            
            if (borderType === 'all' || !borderType) {
                cell.style.borderTop = `2px solid ${color}`;
                cell.style.borderRight = `2px solid ${color}`;
                cell.style.borderBottom = `2px solid ${color}`;
                cell.style.borderLeft = `2px solid ${color}`;
            } else if (borderType !== 'none') {
                const borderSide = `border${borderType.charAt(0).toUpperCase() + borderType.slice(1)}`;
                cell.style[borderSide] = `2px solid ${color}`;
            }
        });
    });
});