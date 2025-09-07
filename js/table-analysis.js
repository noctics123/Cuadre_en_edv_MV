/**
 * Módulo para análisis y parseo de CREATE TABLE statements
 */
const TableAnalysisModule = {
    
    // Variables del módulo
    tableStructure: [],
    
    /**
     * Detecta CREATE TABLE automáticamente en el texto pegado
     */
    detectCreateTable() {
        const text = document.getElementById('createTableInput').value.trim();
        
        if (!text) {
            alert('Por favor, pega algún contenido en el área de texto');
            return;
        }
        
        // Buscar CREATE TABLE en el texto usando los nuevos patrones regex
        const createTableMatches = RegexUtils.extractMultipleCreateTables(text);
        
        if (createTableMatches.length === 0) {
            alert('No se encontró ningún CREATE TABLE en el texto pegado');
            return;
        }
        
        if (createTableMatches.length === 1) {
            // Si hay solo uno, usarlo directamente
            this.processSingleCreateTable(createTableMatches[0]);
        } else {
            // Si hay múltiples, mostrar opciones
            this.showMultipleCreateTableOptions(createTableMatches);
        }
    },

    /**
     * Procesa un único CREATE TABLE detectado
     * @param {string} createStatement - CREATE TABLE statement
     */
    processSingleCreateTable(createStatement) {
        document.getElementById('createTableInput').value = createStatement;
        this.parseCreateTable();
        
        // Auto-llenar esquemas si es posible
        const tableName = RegexUtils.extractTableName(createStatement);
        const schema = RegexUtils.extractSchemaName(createStatement);
        this.autoFillSchemas(tableName, schema);
        
        alert(`CREATE TABLE detectado: ${tableName}`);
    },

    /**
     * Muestra opciones cuando hay múltiples CREATE TABLE
     * @param {Array<string>} createTables - Array de CREATE TABLE statements
     */
    showMultipleCreateTableOptions(createTables) {
        let options = 'Se encontraron múltiples CREATE TABLE:\n\n';
        createTables.forEach((create, index) => {
            const tableName = RegexUtils.extractTableName(create);
            const schema = RegexUtils.extractSchemaName(create);
            options += `${index + 1}. ${tableName} (${schema})\n`;
        });
        options += '\n¿Cuál quieres usar? (Introduce el número)';
        
        const choice = prompt(options);
        const index = parseInt(choice) - 1;
        
        if (index >= 0 && index < createTables.length) {
            this.processSingleCreateTable(createTables[index]);
        }
    },

    /**
     * Auto-llena esquemas basado en la tabla detectada
     * @param {string} tableName - Nombre de la tabla
     * @param {string} schema - Esquema detectado
     */
    autoFillSchemas(tableName, schema) {
        if (!schema || !tableName) return;
        
        const schemaType = RegexUtils.getSchemaType(schema);
        
        if (schemaType === 'ddv') {
            // Auto-completar DDV
            document.getElementById('esquemaDDV').value = schema;
            document.getElementById('tablaDDV').value = tableName;
            
            // Sugerir esquema EDV
            const edvSchema = Utils.generateEDVSchema(schema);
            const edvTable = Utils.generateEDVTable(tableName);
            document.getElementById('esquemaEDV').value = edvSchema;
            document.getElementById('tablaEDV').value = edvTable;
            
            alert(`Esquemas auto-completados:\nDDV: ${schema}\nEDV: ${edvSchema}`);
        } else if (schemaType === 'edv') {
            // Auto-completar EDV
            document.getElementById('esquemaEDV').value = schema;
            document.getElementById('tablaEDV').value = tableName;
            
            alert(`Esquema EDV detectado: ${schema}`);
        }
    },

    /**
     * Parsea CREATE TABLE y extrae la estructura
     */
    parseCreateTable() {
        const createTableText = document.getElementById('createTableInput').value;
        
        if (!createTableText.trim()) {
            alert('Por favor, pega un CREATE TABLE statement');
            return;
        }
        
        try {
            this.tableStructure = this.extractTableStructure(createTableText);
            this.displayFieldMapping();
            document.getElementById('fieldMapping').style.display = 'block';
            
            // Actualizar columnas en repositorio si existe
            const tableName = RegexUtils.extractTableName(createTableText);
            if (tableName && RepositoryModule.tablesRepository[tableName]) {
                RepositoryModule.tablesRepository[tableName].columns = this.tableStructure.length;
                RepositoryModule.saveRepositoryToStorage();
            }
            
        } catch (error) {
            alert('Error al analizar la tabla: ' + error.message);
            console.error('Error completo:', error);
        }
    },

    /**
     * Extrae la estructura de columnas del CREATE TABLE
     * @param {string} sql - CREATE TABLE statement
     * @returns {Array<Object>} - Array de objetos con información de columnas
     */
    extractTableStructure(sql) {
        const columnsText = RegexUtils.extractColumnsContent(sql);
        const columns = [];
        const columnDefinitions = Utils.splitByComma(columnsText);
        
        // Obtener reglas de renombrado actuales
        const currentParams = ParametersModule.getCurrentParameters();
        const renameRules = currentParams.renameRules || {};
        
        for (let colDef of columnDefinitions) {
            colDef = colDef.trim();
            if (!colDef) continue;
            
            // Saltar definiciones que no son columnas (constraints, etc.)
            if (this.isConstraintDefinition(colDef)) continue;
            
            const columnInfo = this.parseColumnDefinition(colDef, renameRules);
            if (columnInfo) {
                columns.push(columnInfo);
            }
        }
        
        return columns;
    },

    /**
     * Verifica si una definición es un constraint y no una columna
     * @param {string} definition - Definición a verificar
     * @returns {boolean} - true si es un constraint
     */
    isConstraintDefinition(definition) {
        const constraintKeywords = [
            'PRIMARY KEY', 'FOREIGN KEY', 'UNIQUE', 'CHECK', 
            'CONSTRAINT', 'INDEX', 'KEY'
        ];
        
        const upperDef = definition.toUpperCase();
        return constraintKeywords.some(keyword => upperDef.includes(keyword));
    },

    /**
     * Parsea una definición individual de columna
     * @param {string} colDef - Definición de columna
     * @param {Object} renameRules - Reglas de renombrado
     * @returns {Object|null} - Información de la columna o null si no es válida
     */
    parseColumnDefinition(colDef, renameRules) {
        const parts = colDef.split(/\s+/);
        if (parts.length < 2) return null;
        
        const columnName = parts[0].replace(/[`"]/g, ''); // Remover comillas
        let dataType = parts[1];
        
        // Incluir paréntesis si existen (para tipos como DECIMAL(10,2))
        if (parts.length > 2 && parts[2].startsWith('(')) {
            let parenContent = '';
            for (let i = 2; i < parts.length; i++) {
                parenContent += parts[i];
                if (parts[i].includes(')')) break;
            }
            dataType += parenContent;
        }
        
        // Determinar función de agregación
        const aggregateFunction = Utils.getAggregateFunction(dataType);
        
        return {
            columnName,
            dataType,
            aggregateFunction,
            edvName: renameRules[columnName] || columnName,
            isNullable: !colDef.toUpperCase().includes('NOT NULL'),
            hasDefault: colDef.toUpperCase().includes('DEFAULT')
        };
    },

    /**
     * Muestra el mapeo de campos en la interfaz
     */
    displayFieldMapping() {
        const content = document.getElementById('fieldMappingContent');
        content.innerHTML = '';
        
        let countFields = 0, sumFields = 0;
        
        this.tableStructure.forEach((field, index) => {
            if (field.aggregateFunction === 'count') countFields++;
            else sumFields++;
            
            const row = Utils.createElement('div', 'field-row');
            row.innerHTML = `
                <div title="${field.columnName}">${field.columnName}</div>
                <div title="${field.dataType}">${field.dataType}</div>
                <div title="Función de agregación">${field.aggregateFunction}()</div>
                <input type="text" class="rename-input" value="${field.edvName}" 
                       onchange="TableAnalysisModule.updateFieldName(${index}, this.value)"
                       title="Nombre del campo en EDV">
            `;
            content.appendChild(row);
        });
        
        // Mostrar estadísticas
        this.displayFieldStatistics(countFields, sumFields);
    },

    /**
     * Muestra estadísticas de los campos
     * @param {number} countFields - Número de campos COUNT
     * @param {number} sumFields - Número de campos SUM
     */
    displayFieldStatistics(countFields, sumFields) {
        const statsSection = document.getElementById('statsSection');
        statsSection.innerHTML = `
            <div class="stat-card">
                <div class="stat-number">${this.tableStructure.length}</div>
                <div class="stat-label">Total Campos</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${countFields}</div>
                <div class="stat-label">COUNT() Campos</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${sumFields}</div>
                <div class="stat-label">SUM() Campos</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${Math.round((sumFields / this.tableStructure.length) * 100) || 0}%</div>
                <div class="stat-label">Campos Numéricos</div>
            </div>
        `;
    },

    /**
     * Actualiza el nombre de campo EDV
     * @param {number} index - Índice del campo
     * @param {string} newName - Nuevo nombre
     */
    updateFieldName(index, newName) {
        if (this.tableStructure[index]) {
            this.tableStructure[index].edvName = newName.trim();
        }
    },

    /**
     * Obtiene la estructura actual de la tabla
     * @returns {Array<Object>} - Estructura de la tabla
     */
    getTableStructure() {
        return this.tableStructure;
    },

    /**
     * Valida que la estructura de tabla esté completa
     * @returns {Object} - {isValid: boolean, errors: Array<string>}
     */
    validateTableStructure() {
        const errors = [];
        
        if (!this.tableStructure.length) {
            errors.push('No hay estructura de tabla definida');
            return { isValid: false, errors };
        }
        
        // Verificar nombres de campos EDV únicos
        const edvNames = this.tableStructure.map(f => f.edvName);
        const duplicates = edvNames.filter((name, index) => edvNames.indexOf(name) !== index);
        
        if (duplicates.length > 0) {
            errors.push(`Nombres EDV duplicados: ${[...new Set(duplicates)].join(', ')}`);
        }
        
        // Verificar campos vacíos
        const emptyFields = this.tableStructure.filter(f => !f.columnName || !f.edvName);
        if (emptyFields.length > 0) {
            errors.push('Hay campos con nombres vacíos');
        }
        
        return {
            isValid: errors.length === 0,
            errors
        };
    },

    /**
     * Resetea la estructura de tabla
     */
    resetTableStructure() {
        this.tableStructure = [];
        document.getElementById('fieldMapping').style.display = 'none';
        document.getElementById('createTableInput').value = '';
    }
};