/**
 * Módulo para gestión de parámetros de la aplicación
 */
const ParametersModule = {
    
    // Variables del módulo
    parameters: {},
    
    /**
     * Carga datos de ejemplo al inicializar
     */
    loadExampleData() {
        document.getElementById('esquemaDDV').value = 'catalog_lhcl_prod_bcp.bcp_ddv_matrizvariables';
        document.getElementById('tablaDDV').value = 'hm_matrizdemografico';
        document.getElementById('esquemaEDV').value = 'catalog_lhcl_prod_bcp_expl.bcp_edv_trdata_012';
        document.getElementById('tablaEDV').value = 'hm_matrizdemografico_ruben';
        document.getElementById('periodos').value = '202505, 202506, 202507';
        document.getElementById('renameRules').value = 'codclaveunicocli:cod_uni_cocli';
    },

    /**
     * Guarda los parámetros actuales
     */
    saveParameters() {
        // Recopilar datos del formulario
        const formData = {
            esquemaDDV: document.getElementById('esquemaDDV').value.trim(),
            tablaDDV: document.getElementById('tablaDDV').value.trim(),
            esquemaEDV: document.getElementById('esquemaEDV').value.trim(),
            tablaEDV: document.getElementById('tablaEDV').value.trim(),
            periodos: document.getElementById('periodos').value.trim(),
            renameRules: Utils.parseRenameRules(document.getElementById('renameRules').value)
        };
        
        // Validar parámetros
        const validation = Utils.validateParameters(formData);
        
        if (!validation.isValid) {
            alert('Errores en los parámetros:\n• ' + validation.errors.join('\n• '));
            return false;
        }
        
        // Guardar parámetros
        this.parameters = formData;
        
        // Guardar en localStorage para persistencia
        localStorage.setItem('cuadreParameters', JSON.stringify(this.parameters));
        
        alert('Parámetros guardados correctamente');
        
        // Auto-navegar a la siguiente pestaña
        window.switchTab('describe');
        
        return true;
    },

    /**
     * Obtiene los parámetros actuales
     * @returns {Object} - Parámetros actuales
     */
    getCurrentParameters() {
        return this.parameters;
    },

    /**
     * Carga parámetros desde localStorage
     */
    loadParametersFromStorage() {
        try {
            const stored = localStorage.getItem('cuadreParameters');
            if (stored) {
                this.parameters = JSON.parse(stored);
                this.populateForm();
            }
        } catch (error) {
            console.error('Error cargando parámetros:', error);
        }
    },

    /**
     * Llena el formulario con los parámetros cargados
     */
    populateForm() {
        if (!this.parameters) return;
        
        document.getElementById('esquemaDDV').value = this.parameters.esquemaDDV || '';
        document.getElementById('tablaDDV').value = this.parameters.tablaDDV || '';
        document.getElementById('esquemaEDV').value = this.parameters.esquemaEDV || '';
        document.getElementById('tablaEDV').value = this.parameters.tablaEDV || '';
        document.getElementById('periodos').value = this.parameters.periodos || '';
        
        // Reconstruir reglas de renombrado
        if (this.parameters.renameRules) {
            const rulesText = Object.entries(this.parameters.renameRules)
                .map(([original, renamed]) => `${original}:${renamed}`)
                .join('\n');
            document.getElementById('renameRules').value = rulesText;
        }
    },

    /**
     * Auto-completa esquemas basado en detección automática
     * @param {string} tableName - Nombre de tabla detectado
     * @param {string} schema - Esquema detectado
     */
    autoFillSchemas(tableName, schema) {
        if (!schema || !tableName) return;
        
        const schemaType = RegexUtils.getSchemaType(schema);
        
        if (schemaType === 'ddv') {
            // Llenar campos DDV
            document.getElementById('esquemaDDV').value = schema;
            document.getElementById('tablaDDV').value = tableName;
            
            // Sugerir campos EDV
            const edvSchema = Utils.generateEDVSchema(schema);
            const edvTable = Utils.generateEDVTable(tableName);
            
            document.getElementById('esquemaEDV').value = edvSchema;
            document.getElementById('tablaEDV').value = edvTable;
            
            // Mostrar mensaje informativo
            this.showAutoFillMessage(schema, edvSchema, tableName, edvTable);
            
        } else if (schemaType === 'edv') {
            // Llenar campos EDV
            document.getElementById('esquemaEDV').value = schema;
            document.getElementById('tablaEDV').value = tableName;
            
            alert(`Esquema EDV detectado automáticamente: ${schema}`);
        }
    },

    /**
     * Muestra mensaje de auto-completado
     * @param {string} ddvSchema - Esquema DDV
     * @param {string} edvSchema - Esquema EDV sugerido
     * @param {string} ddvTable - Tabla DDV
     * @param {string} edvTable - Tabla EDV sugerida
     */
    showAutoFillMessage(ddvSchema, edvSchema, ddvTable, edvTable) {
        const message = `Esquemas auto-completados:\n\n` +
                       `DDV: ${ddvSchema}.${ddvTable}\n` +
                       `EDV: ${edvSchema}.${edvTable}\n\n` +
                       `¿Los esquemas EDV sugeridos son correctos?`;
        
        if (confirm(message)) {
            // Guardar automáticamente si el usuario confirma
            this.saveParameters();
        }
    },

    /**
     * Resetea todos los parámetros
     */
    resetParameters() {
        if (confirm('¿Resetear todos los parámetros? Se perderán los datos actuales.')) {
            this.parameters = {};
            localStorage.removeItem('cuadreParameters');
            
            // Limpiar formulario
            document.getElementById('esquemaDDV').value = '';
            document.getElementById('tablaDDV').value = '';
            document.getElementById('esquemaEDV').value = '';
            document.getElementById('tablaEDV').value = '';
            document.getElementById('periodos').value = '';
            document.getElementById('renameRules').value = '';
            
            // Cargar datos de ejemplo
            this.loadExampleData();
            
            alert('Parámetros reseteados');
        }
    },

    /**
     * Exporta parámetros actuales
     * @returns {string} - Parámetros en formato JSON
     */
    exportParameters() {
        if (!Object.keys(this.parameters).length) {
            alert('No hay parámetros para exportar');
            return '';
        }
        
        return JSON.stringify(this.parameters, null, 2);
    },

    /**
     * Importa parámetros desde JSON
     * @param {string} jsonString - Parámetros en formato JSON
     */
    importParameters(jsonString) {
        try {
            const imported = JSON.parse(jsonString);
            
            // Validar estructura
            const validation = Utils.validateParameters(imported);
            if (!validation.isValid) {
                throw new Error('Parámetros inválidos: ' + validation.errors.join(', '));
            }
            
            this.parameters = imported;
            this.populateForm();
            
            alert('Parámetros importados correctamente');
            
        } catch (error) {
            alert('Error importando parámetros: ' + error.message);
        }
    },

    /**
     * Obtiene sugerencias de configuración basadas en el historial
     * @returns {Array<Object>} - Array de sugerencias
     */
    getSuggestions() {
        const suggestions = [];
        
        // Sugerencias basadas en repositorio
        const repo = RepositoryModule.getRepository();
        const schemas = [...new Set(Object.values(repo).map(t => t.schema))];
        
        schemas.forEach(schema => {
            const schemaType = RegexUtils.getSchemaType(schema);
            if (schemaType === 'ddv') {
                suggestions.push({
                    type: 'schema_pair',
                    ddvSchema: schema,
                    edvSchema: Utils.generateEDVSchema(schema),
                    description: `Par de esquemas: ${schema} → ${Utils.generateEDVSchema(schema)}`
                });
            }
        });
        
        return suggestions;
    },

    /**
     * Aplica una sugerencia de configuración
     * @param {Object} suggestion - Sugerencia a aplicar
     */
    applySuggestion(suggestion) {
        if (suggestion.type === 'schema_pair') {
            document.getElementById('esquemaDDV').value = suggestion.ddvSchema;
            document.getElementById('esquemaEDV').value = suggestion.edvSchema;
            
            alert(`Esquemas aplicados:\nDDV: ${suggestion.ddvSchema}\nEDV: ${suggestion.edvSchema}`);
        }
    },

    /**
     * Valida que los parámetros estén listos para generar queries
     * @returns {boolean} - true si están listos
     */
    areParametersReady() {
        const validation = Utils.validateParameters(this.parameters);
        return validation.isValid;
    }
};