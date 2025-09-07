/**
 * Funciones utilitarias generales para la aplicación
 */
const Utils = {
    
    /**
     * Cambia entre pestañas de la interfaz
     * @param {string} tabName - Nombre de la pestaña a activar
     * @param {Event} event - Evento del click
     */
    switchTab(tabName, event) {
        // Remover clases active de todas las pestañas y contenidos
        document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        // Activar la pestaña clickeada
        if (event && event.target) {
            event.target.classList.add('active');
        }
        
        // Activar el contenido correspondiente
        const tabContent = document.getElementById(tabName);
        if (tabContent) {
            tabContent.classList.add('active');
        }
    },

    /**
     * Divide texto por comas respetando paréntesis anidados
     * @param {string} text - Texto a dividir
     * @returns {Array<string>} - Array de elementos divididos
     */
    splitByComma(text) {
        const result = [];
        let current = '';
        let parenCount = 0;
        
        for (let char of text) {
            if (char === '(') parenCount++;
            else if (char === ')') parenCount--;
            else if (char === ',' && parenCount === 0) {
                if (current.trim()) result.push(current.trim());
                current = '';
                continue;
            }
            current += char;
        }
        
        if (current.trim()) result.push(current.trim());
        return result;
    },

    /**
     * Parsea reglas de renombrado del textarea
     * @param {string} rulesText - Texto con reglas de renombrado
     * @returns {Object} - Objeto con reglas de mapeo
     */
    parseRenameRules(rulesText) {
        const rules = {};
        rulesText.split('\n').forEach(line => {
            const [original, renamed] = line.split(':').map(s => s.trim());
            if (original && renamed) {
                rules[original] = renamed;
            }
        });
        return rules;
    },

    /**
     * Determina la función de agregación basada en el tipo de dato
     * @param {string} dataType - Tipo de dato de la columna
     * @returns {string} - 'sum' o 'count'
     */
    getAggregateFunction(dataType) {
        return RegexUtils.isNumericDataType(dataType) ? 'sum' : 'count';
    },

    /**
     * Limpia y formatea períodos para uso en queries
     * @param {string} periodsText - Texto con períodos separados por coma
     * @returns {string} - Períodos formateados sin espacios
     */
    formatPeriods(periodsText) {
        return periodsText.replace(/\s/g, '');
    },

    /**
     * Crea un elemento HTML con clase y contenido
     * @param {string} tag - Tag del elemento
     * @param {string} className - Clase CSS
     * @param {string} content - Contenido HTML
     * @returns {HTMLElement} - Elemento creado
     */
    createElement(tag, className = '', content = '') {
        const element = document.createElement(tag);
        if (className) element.className = className;
        if (content) element.innerHTML = content;
        return element;
    },

    /**
     * Copia texto al portapapeles
     * @param {string} text - Texto a copiar
     * @param {string} successMessage - Mensaje de éxito (opcional)
     */
    async copyToClipboard(text, successMessage = 'Texto copiado al portapapeles') {
        try {
            await navigator.clipboard.writeText(text);
            alert(successMessage);
        } catch (err) {
            console.error('Error al copiar al portapapeles:', err);
            // Fallback para navegadores que no soportan clipboard API
            const textArea = document.createElement('textarea');
            textArea.value = text;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            alert(successMessage);
        }
    },

    /**
     * Debounce para optimizar búsquedas
     * @param {Function} func - Función a ejecutar
     * @param {number} wait - Tiempo de espera en ms
     * @returns {Function} - Función debounced
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    },

    /**
     * Valida si un string es un CREATE TABLE válido
     * @param {string} text - Texto a validar
     * @returns {boolean} - true si es válido
     */
    isValidCreateTable(text) {
        return text.trim().length > 0 && 
               /CREATE\s+TABLE/i.test(text) && 
               text.includes('(') && 
               RegexUtils.extractTableName(text) !== null;
    },

    /**
     * Obtiene fecha actual en formato ISO
     * @returns {string} - Fecha en formato ISO
     */
    getCurrentISODate() {
        return new Date().toISOString();
    },

    /**
     * Formatea fecha para mostrar en UI
     * @param {string} isoDate - Fecha en formato ISO
     * @returns {string} - Fecha formateada
     */
    formatDisplayDate(isoDate) {
        const date = new Date(isoDate);
        return date.toLocaleDateString('es-ES', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    },

    /**
     * Genera sugerencias de esquema EDV basado en esquema DDV
     * @param {string} ddvSchema - Esquema DDV
     * @returns {string} - Esquema EDV sugerido
     */
    generateEDVSchema(ddvSchema) {
        return ddvSchema
            .replace(/ddv/gi, 'edv')
            .replace(/matrizvariables/gi, 'trdata_012');
    },

    /**
     * Genera sugerencias de tabla EDV basado en tabla DDV
     * @param {string} ddvTable - Tabla DDV
     * @returns {string} - Tabla EDV sugerida
     */
    generateEDVTable(ddvTable) {
        // Agregar sufijo común si no lo tiene
        if (!ddvTable.includes('_ruben') && !ddvTable.includes('_dev')) {
            return ddvTable + '_ruben';
        }
        return ddvTable;
    },

    /**
     * Valida configuración de parámetros
     * @param {Object} params - Parámetros a validar
     * @returns {Object} - {isValid: boolean, errors: Array<string>}
     */
    validateParameters(params) {
        const errors = [];
        
        if (!params.esquemaDDV) errors.push('Esquema DDV es requerido');
        if (!params.tablaDDV) errors.push('Tabla DDV es requerida');
        if (!params.esquemaEDV) errors.push('Esquema EDV es requerido');
        if (!params.tablaEDV) errors.push('Tabla EDV es requerida');
        if (!params.periodos) errors.push('Períodos son requeridos');
        
        // Validar formato de períodos
        if (params.periodos && !/^\d{6}(\s*,\s*\d{6})*$/.test(params.periodos)) {
            errors.push('Formato de períodos inválido (usar YYYYMM)');
        }
        
        return {
            isValid: errors.length === 0,
            errors
        };
    },

    /**
     * Filtra elementos de una lista basado en múltiples criterios
     * @param {Array} items - Items a filtrar
     * @param {Object} filters - Filtros a aplicar
     * @returns {Array} - Items filtrados
     */
    filterItems(items, filters) {
        return items.filter(item => {
            return Object.entries(filters).every(([key, value]) => {
                if (!value) return true; // Sin filtro
                
                if (typeof value === 'string') {
                    const itemValue = item[key] || '';
                    return itemValue.toLowerCase().includes(value.toLowerCase());
                }
                
                if (typeof value === 'function') {
                    return value(item[key], item);
                }
                
                return item[key] === value;
            });
        });
    }
};

// Hacer función switchTab disponible globalmente para los event handlers del HTML
window.switchTab = function(tabName) {
    const tabs = document.querySelectorAll('.tab');
    const activeTab = Array.from(tabs).find(tab => 
        tab.textContent.toLowerCase().includes(tabName.toLowerCase())
    );
    
    if (activeTab) {
        Utils.switchTab(tabName, { target: activeTab });
    } else {
        Utils.switchTab(tabName);
    }
};