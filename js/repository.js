/**
 * Módulo para gestión del repositorio de CREATE TABLEs
 */
const RepositoryModule = {
    
    // Variables del módulo
    tablesRepository: {},
    
    /**
     * Inicializa el módulo cargando datos del localStorage
     */
    init() {
        this.loadRepositoryFromStorage();
        this.updateUI();
    },

    /**
     * Carga repositorio desde localStorage
     */
    loadRepositoryFromStorage() {
        try {
            const stored = localStorage.getItem('tablesRepository');
            this.tablesRepository = stored ? JSON.parse(stored) : {};
        } catch (error) {
            console.error('Error cargando repositorio:', error);
            this.tablesRepository = {};
        }
    },

    /**
     * Guarda repositorio en localStorage
     */
    saveRepositoryToStorage() {
        try {
            localStorage.setItem('tablesRepository', JSON.stringify(this.tablesRepository));
        } catch (error) {
            console.error('Error guardando repositorio:', error);
        }
    },

    /**
     * Actualiza toda la UI del repositorio
     */
    updateUI() {
        this.updateTableSelector();
        this.updateRepositoryStats();
        this.loadRepositoryList();
        this.updateFilters();
    },

    /**
     * Actualiza selector de tablas en la pestaña Describe
     */
    updateTableSelector() {
        const selector = document.getElementById('tableSelector');
        if (!selector) return;
        
        selector.innerHTML = '<option value="">-- Seleccionar tabla existente --</option>';
        
        // Agrupar por esquema para mejor organización
        const groupedTables = this.groupTablesBySchema();
        
        Object.entries(groupedTables).forEach(([schema, tables]) => {
            const optgroup = document.createElement('optgroup');
            optgroup.label = schema;
            
            tables.forEach(tableName => {
                const option = document.createElement('option');
                option.value = tableName;
                option.textContent = `${tableName} (${this.tablesRepository[tableName].columns || 0} cols)`;
                optgroup.appendChild(option);
            });
            
            selector.appendChild(optgroup);
        });
    },

    /**
     * Agrupa tablas por esquema para organización
     * @returns {Object} - Objeto agrupado por esquema
     */
    groupTablesBySchema() {
        const grouped = {};
        
        Object.entries(this.tablesRepository).forEach(([tableName, tableData]) => {
            const schema = tableData.schema || 'Sin esquema';
            if (!grouped[schema]) {
                grouped[schema] = [];
            }
            grouped[schema].push(tableName);
        });
        
        // Ordenar esquemas y tablas
        Object.keys(grouped).forEach(schema => {
            grouped[schema].sort();
        });
        
        return grouped;
    },

    /**
     * Actualiza filtros en la pestaña repositorio (incluyendo los nuevos dropdowns)
     */
    updateFilters() {
        this.updateSchemaFilter();
        this.updateTableFilter();
    },

    /**
     * Actualiza filtro de esquemas
     */
    updateSchemaFilter() {
        const filter = document.getElementById('schemaFilter');
        if (!filter) return;
        
        const schemas = [...new Set(Object.values(this.tablesRepository).map(t => t.schema))];
        filter.innerHTML = '<option value="">Todos los esquemas</option>';
        
        schemas.sort().forEach(schema => {
            const option = document.createElement('option');
            option.value = schema;
            option.textContent = schema;
            filter.appendChild(option);
        });
    },

    /**
     * Actualiza filtro de tablas (nuevo dropdown agregado)
     */
    updateTableFilter() {
        const filter = document.getElementById('tableFilter');
        if (!filter) return;
        
        const tables = Object.keys(this.tablesRepository).sort();
        filter.innerHTML = '<option value="">Todas las tablas</option>';
        
        tables.forEach(tableName => {
            const option = document.createElement('option');
            option.value = tableName;
            option.textContent = tableName;
            filter.appendChild(option);
        });
    },

    /**
     * Filtra tablas en el selector de la pestaña Describe
     */
    filterTables() {
        const searchTerm = document.getElementById('tableSearch').value.toLowerCase();
        const selector = document.getElementById('tableSelector');
        
        if (!selector) return;
        
        // Filtrar opciones en optgroups
        Array.from(selector.querySelectorAll('optgroup')).forEach(optgroup => {
            let hasVisibleOptions = false;
            
            Array.from(optgroup.querySelectorAll('option')).forEach(option => {
                const tableName = option.value.toLowerCase();
                const isVisible = tableName.includes(searchTerm);
                option.style.display = isVisible ? 'block' : 'none';
                
                if (isVisible) hasVisibleOptions = true;
            });
            
            optgroup.style.display = hasVisibleOptions ? 'block' : 'none';
        });
    },

    /**
     * Carga tabla desde repositorio (llamado desde selector)
     */
    loadTableFromRepo() {
        const selectedTable = document.getElementById('tableSelector').value;
        if (!selectedTable || !this.tablesRepository[selectedTable]) return;
        
        const tableData = this.tablesRepository[selectedTable];
        document.getElementById('createTableInput').value = tableData.createStatement;
        
        // Auto-detectar esquemas si coinciden
        this.autoDetectSchemas(selectedTable);
        
        // Auto-parsear la tabla
        TableAnalysisModule.parseCreateTable();
        
        alert(`Tabla "${selectedTable}" cargada desde el repositorio`);
    },

    /**
     * Auto-detecta esquemas basado en el nombre de la tabla
     * @param {string} tableName - Nombre de la tabla
     */
    autoDetectSchemas(tableName) {
        const tableData = this.tablesRepository[tableName];
        if (!tableData) return;
        
        const schemaType = RegexUtils.getSchemaType(tableData.schema);
        
        if (schemaType === 'ddv') {
            document.getElementById('esquemaDDV').value = tableData.schema;
            document.getElementById('tablaDDV').value = tableName;
            
            // Sugerir esquema EDV
            const edvSchema = Utils.generateEDVSchema(tableData.schema);
            const edvTable = Utils.generateEDVTable(tableName);
            document.getElementById('esquemaEDV').value = edvSchema;
            document.getElementById('tablaEDV').value = edvTable;
        } else if (schemaType === 'edv') {
            document.getElementById('esquemaEDV').value = tableData.schema;
            document.getElementById('tablaEDV').value = tableName;
        }
    },

    /**
     * Guarda tabla en repositorio
     */
    saveToRepository() {
        const createStatement = document.getElementById('createTableInput').value.trim();
        
        if (!createStatement) {
            alert('No hay CREATE TABLE para guardar');
            return;
        }
        
        if (!Utils.isValidCreateTable(createStatement)) {
            alert('El CREATE TABLE no parece ser válido');
            return;
        }
        
        try {
            const tableName = RegexUtils.extractTableName(createStatement);
            const schema = RegexUtils.extractSchemaName(createStatement);
            
            if (!tableName) {
                throw new Error('No se pudo extraer el nombre de la tabla del CREATE TABLE');
            }
            
            const existsInRepo = this.tablesRepository[tableName];
            
            if (existsInRepo && !confirm(`La tabla "${tableName}" ya existe en el repositorio. ¿Sobrescribir?`)) {
                return;
            }
            
            // Contar columnas si hay estructura parseada
            const columns = TableAnalysisModule.getTableStructure().length || 0;
            
            this.tablesRepository[tableName] = {
                tableName: tableName,
                schema: schema,
                createStatement: createStatement,
                dateAdded: Utils.getCurrentISODate(),
                columns: columns,
                lastUsed: Utils.getCurrentISODate()
            };
            
            this.saveRepositoryToStorage();
            this.updateUI();
            
            alert(`Tabla "${tableName}" guardada en el repositorio\nEsquema: ${schema}\nColumnas: ${columns}`);
            
        } catch (error) {
            alert('Error al guardar en repositorio: ' + error.message);
            console.error('Error completo:', error);
        }
    },

    /**
     * Muestra gestor de repositorio
     */
    showRepositoryManager() {
        window.switchTab('repositorio');
    },

    /**
     * Actualiza estadísticas del repositorio
     */
    updateRepositoryStats() {
        const totalTables = Object.keys(this.tablesRepository).length;
        const schemas = [...new Set(Object.values(this.tablesRepository).map(t => t.schema))];
        const totalColumns = Object.values(this.tablesRepository).reduce((sum, t) => sum + (t.columns || 0), 0);
        const avgColumns = totalTables > 0 ? Math.round(totalColumns / totalTables) : 0;
        
        const statsHtml = `
            <div class="stat-card">
                <div class="stat-number">${totalTables}</div>
                <div class="stat-label">Tablas Guardadas</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${schemas.length}</div>
                <div class="stat-label">Esquemas</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${totalColumns}</div>
                <div class="stat-label">Total Columnas</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">${avgColumns}</div>
                <div class="stat-label">Promedio Columnas</div>
            </div>
        `;
        
        const repoStats = document.getElementById('repoStats');
        if (repoStats) {
            repoStats.innerHTML = statsHtml;
        }
    },

    /**
     * Carga lista del repositorio
     */
    loadRepositoryList() {
        const container = document.getElementById('repositoryList');
        if (!container) return;
        
        container.innerHTML = '';
        
        const tables = Object.entries(this.tablesRepository);
        
        if (tables.length === 0) {
            container.innerHTML = `
                <div style="text-align: center; padding: 40px; color: #6c757d;">
                    <h4>Repositorio vacío</h4>
                    <p>Agrega tu primer CREATE TABLE desde la pestaña "Describe Tabla"</p>
                </div>
            `;
            return;
        }
        
        tables.forEach(([tableKey, tableData]) => {
            const item = this.createRepositoryItem(tableKey, tableData);
            container.appendChild(item);
        });
    },

    /**
     * Crea item del repositorio
     * @param {string} tableKey - Clave de la tabla
     * @param {Object} tableData - Datos de la tabla
     * @returns {HTMLElement} - Elemento HTML del item
     */
    createRepositoryItem(tableKey, tableData) {
        const item = Utils.createElement('div', 'repository-item');
        item.id = `repo-item-${tableKey}`;
        
        const preview = (tableData.createStatement || '').substring(0, 150) + '...';
        const dateAdded = tableData.dateAdded ? Utils.formatDisplayDate(tableData.dateAdded) : 'Fecha desconocida';
        const schemaType = RegexUtils.getSchemaType(tableData.schema);
        const schemaBadge = schemaType === 'ddv' ? 'DDV' : schemaType === 'edv' ? 'EDV' : 'OTRO';
        
        item.innerHTML = `
            <div class="repo-item-header">
                <div>
                    <div class="repo-item-title">
                        ${tableData.tableName}
                        <span class="schema-badge schema-badge-${schemaType}">${schemaBadge}</span>
                    </div>
                    <div class="repo-item-schema">${tableData.schema} • ${tableData.columns || 0} columnas • ${dateAdded}</div>
                </div>
                <div class="repo-item-actions">
                    <button class="btn btn-small" onclick="RepositoryModule.loadTableFromRepoItem('${tableKey}')">Usar</button>
                    <button class="btn btn-small btn-secondary" onclick="RepositoryModule.viewTableFull('${tableKey}')">Ver</button>
                    <button class="btn btn-small" style="background: #dc3545;" onclick="RepositoryModule.deleteFromRepo('${tableKey}')">Eliminar</button>
                </div>
            </div>
            <div class="repo-item-preview">${preview}</div>
        `;
        
        return item;
    },

    /**
     * Filtra repositorio basado en múltiples criterios
     */
    filterRepository() {
        const schemaFilter = document.getElementById('schemaFilter')?.value || '';
        const tableFilter = document.getElementById('tableFilter')?.value || '';
        const searchTerm = document.getElementById('repoSearch')?.value.toLowerCase() || '';
        const showFilter = document.getElementById('showFilter')?.value || 'all';
        
        Object.keys(this.tablesRepository).forEach(tableKey => {
            const item = document.getElementById(`repo-item-${tableKey}`);
            const tableData = this.tablesRepository[tableKey];
            
            if (!item || !tableData) return;
            
            let isVisible = true;
            
            // Filtro por esquema
            if (schemaFilter && tableData.schema !== schemaFilter) {
                isVisible = false;
            }
            
            // Filtro por tabla específica
            if (tableFilter && tableKey !== tableFilter) {
                isVisible = false;
            }
            
            // Filtro por búsqueda de texto
            if (searchTerm) {
                const searchableText = `${tableData.tableName} ${tableData.schema}`.toLowerCase();
                if (!searchableText.includes(searchTerm)) {
                    isVisible = false;
                }
            }
            
            // Filtro por tipo de esquema
            if (showFilter !== 'all') {
                const schemaType = RegexUtils.getSchemaType(tableData.schema);
                if (showFilter === 'ddv' && schemaType !== 'ddv') isVisible = false;
                if (showFilter === 'edv' && schemaType !== 'edv') isVisible = false;
                if (showFilter === 'recent') {
                    const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
                    const tableDate = new Date(tableData.dateAdded);
                    if (tableDate < weekAgo) isVisible = false;
                }
            }
            
            item.style.display = isVisible ? 'block' : 'none';
        });
    },

    /**
     * Carga tabla desde item del repositorio
     * @param {string} tableKey - Clave de la tabla
     */
    loadTableFromRepoItem(tableKey) {
        const tableData = this.tablesRepository[tableKey];
        if (!tableData) return;
        
        document.getElementById('createTableInput').value = tableData.createStatement;
        this.autoDetectSchemas(tableKey);
        TableAnalysisModule.parseCreateTable();
        
        // Actualizar último uso
        tableData.lastUsed = Utils.getCurrentISODate();
        this.saveRepositoryToStorage();
        
        window.switchTab('describe');
        alert(`Tabla "${tableKey}" cargada desde el repositorio`);
    },

    /**
     * Ver tabla completa en modal
     * @param {string} tableKey - Clave de la tabla
     */
    viewTableFull(tableKey) {
        const tableData = this.tablesRepository[tableKey];
        if (!tableData) return;
        
        document.getElementById('modalTitle').textContent = `${tableData.tableName} (${tableData.schema})`;
        document.getElementById('modalContent').textContent = tableData.createStatement;
        document.getElementById('tableModal').style.display = 'block';
    },

    /**
     * Elimina tabla del repositorio
     * @param {string} tableKey - Clave de la tabla
     */
    deleteFromRepo(tableKey) {
        if (confirm(`¿Eliminar "${tableKey}" del repositorio? Esta acción no se puede deshacer.`)) {
            delete this.tablesRepository[tableKey];
            this.saveRepositoryToStorage();
            this.updateUI();
            alert(`Tabla "${tableKey}" eliminada del repositorio`);
        }
    },

    /**
     * Importa desde archivo
     */
    importFromFile() {
        document.getElementById('fileImport').click();
    },

    /**
     * Maneja importación de archivo
     * @param {Event} event - Evento del input file
     */
    handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result;
            const tables = RegexUtils.extractMultipleCreateTables(content);
            
            if (tables.length === 0) {
                alert('No se encontraron CREATE TABLE statements en el archivo');
                return;
            }
            
            this.processBatchImport(tables);
        };
        
        reader.readAsText(file);
    },

    /**
     * Procesa importación masiva de tablas
     * @param {Array<string>} tables - Array de CREATE TABLE statements
     */
    processBatchImport(tables) {
        let imported = 0;
        let errors = 0;
        const importResults = [];
        
        tables.forEach(createStatement => {
            try {
                const tableName = RegexUtils.extractTableName(createStatement);
                const schema = RegexUtils.extractSchemaName(createStatement);
                
                if (tableName) {
                    const alreadyExists = this.tablesRepository[tableName];
                    
                    this.tablesRepository[tableName] = {
                        tableName,
                        schema,
                        createStatement,
                        dateAdded: Utils.getCurrentISODate(),
                        columns: 0, // Se calculará al parsear
                        lastUsed: null,
                        imported: true
                    };
                    
                    importResults.push({
                        table: tableName,
                        schema: schema,
                        status: alreadyExists ? 'sobrescrito' : 'nuevo'
                    });
                    
                    imported++;
                } else {
                    errors++;
                }
            } catch (error) {
                console.log('Error procesando tabla:', error);
                errors++;
            }
        });
        
        this.saveRepositoryToStorage();
        this.updateUI();
        
        // Mostrar resultados detallados
        this.showImportResults(imported, errors, importResults);
    },

    /**
     * Muestra resultados de importación
     * @param {number} imported - Número de tablas importadas
     * @param {number} errors - Número de errores
     * @param {Array} results - Detalles de importación
     */
    showImportResults(imported, errors, results) {
        let message = `Importación completada:\n✅ ${imported} tablas importadas\n❌ ${errors} errores\n\n`;
        
        if (results.length > 0) {
            message += 'Detalles:\n';
            results.slice(0, 10).forEach(result => {
                const status = result.status === 'nuevo' ? '🆕' : '🔄';
                message += `${status} ${result.table} (${result.schema})\n`;
            });
            
            if (results.length > 10) {
                message += `... y ${results.length - 10} tablas más`;
            }
        }
        
        alert(message);
    },

    /**
     * Exporta repositorio
     */
    exportRepository() {
        const tables = Object.values(this.tablesRepository);
        
        if (tables.length === 0) {
            alert('No hay tablas en el repositorio para exportar');
            return;
        }
        
        const data = tables.map(t => {
            return `-- Tabla: ${t.tableName}\n-- Esquema: ${t.schema}\n-- Fecha: ${t.dateAdded}\n\n${t.createStatement}`;
        }).join('\n\n-- =====================================\n\n');
        
        const blob = new Blob([data], { type: 'text/sql' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `repositorio_create_tables_${new Date().toISOString().split('T')[0]}.sql`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        alert(`Repositorio exportado: ${tables.length} tablas`);
    },

    /**
     * Limpia repositorio
     */
    clearRepository() {
        const totalTables = Object.keys(this.tablesRepository).length;
        
        if (totalTables === 0) {
            alert('El repositorio ya está vacío');
            return;
        }
        
        if (confirm(`¿Eliminar TODAS las ${totalTables} tablas del repositorio? Esta acción no se puede deshacer.`)) {
            this.tablesRepository = {};
            this.saveRepositoryToStorage();
            this.updateUI();
            alert('Repositorio limpiado');
        }
    },

    /**
     * Obtiene referencia al repositorio
     * @returns {Object} - Repositorio de tablas
     */
    getRepository() {
        return this.tablesRepository;
    }
};