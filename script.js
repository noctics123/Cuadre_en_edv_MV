// Variables globales
let tableStructure = [];
let parameters = {};
let generatedQueries = {};
let tablesRepository = JSON.parse(localStorage.getItem('tablesRepository')) || {};

// Inicializar al cargar la página
window.onload = function() {
    loadExampleData();
    updateTableSelector();
    updateRepositoryStats();
    loadRepositoryList();
};

// Cargar datos de ejemplo
function loadExampleData() {
    document.getElementById('esquemaDDV').value = 'catalog_lhcl_prod_bcp.bcp_ddv_matrizvariables_v';
    document.getElementById('tablaDDV').value = 'hm_matrizdemografico';
    document.getElementById('esquemaEDV').value = 'catalog_lhcl_prod_bcp_expl.bcp_edv_trdata_012';
    document.getElementById('tablaEDV').value = 'hm_matrizdemografico_ruben';
}

// FUNCIONES BÁSICAS

// Cambio de tabs
function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tabName).classList.add('active');
}

// Guardar parámetros
function saveParameters() {
    parameters = {
        esquemaDDV: document.getElementById('esquemaDDV').value,
        tablaDDV: document.getElementById('tablaDDV').value,
        esquemaEDV: document.getElementById('esquemaEDV').value,
        tablaEDV: document.getElementById('tablaEDV').value,
        periodos: document.getElementById('periodos').value,
        renameRules: parseRenameRules(document.getElementById('renameRules').value)
    };
    
    alert('Parámetros guardados correctamente');
    switchTab('describe');
}

// Parsear reglas de renombrado
function parseRenameRules(rulesText) {
    const rules = {};
    rulesText.split('\n').forEach(line => {
        const [original, renamed] = line.split(':').map(s => s.trim());
        if (original && renamed) {
            rules[original] = renamed;
        }
    });
    return rules;
}

// FUNCIONES DEL REPOSITORIO

// Actualizar selector de tablas
function updateTableSelector() {
    const selector = document.getElementById('tableSelector');
    selector.innerHTML = '<option value="">-- Seleccionar tabla existente --</option>';
    
    Object.keys(tablesRepository).forEach(tableName => {
        const option = document.createElement('option');
        option.value = tableName;
        option.textContent = `${tableName} (${tablesRepository[tableName].schema})`;
        selector.appendChild(option);
    });
}

// Cargar tabla desde repositorio
function loadTableFromRepo() {
    const selectedTable = document.getElementById('tableSelector').value;
    if (selectedTable && tablesRepository[selectedTable]) {
        document.getElementById('createTableInput').value = tablesRepository[selectedTable].createStatement;
        
        // Auto-detectar esquemas si coinciden
        autoDetectSchemas(selectedTable);
        
        // Auto-parsear la tabla
        parseCreateTable();
        
        alert(`Tabla "${selectedTable}" cargada desde el repositorio`);
    }
}

// Auto-detectar esquemas basado en el nombre de la tabla
function autoDetectSchemas(tableName) {
    const tableData = tablesRepository[tableName];
    
    // Si el esquema de la tabla coincide con DDV, autocompletar
    if (tableData.schema.includes('ddv') || tableData.schema.includes('matrizvariables')) {
        document.getElementById('esquemaDDV').value = tableData.schema;
        document.getElementById('tablaDDV').value = tableName;
    }
    
    // Sugerir esquema EDV basado en patrones comunes
    if (tableData.schema.includes('ddv')) {
        const edvSchema = tableData.schema.replace('ddv', 'edv').replace('matrizvariables', 'trdata_012');
        document.getElementById('esquemaEDV').value = edvSchema;
        
        // Sugerir nombre de tabla EDV (agregar sufijo común)
        document.getElementById('tablaEDV').value = tableName + '_ruben';
    }
}

// Filtrar tablas en selector
function filterTables() {
    const searchTerm = document.getElementById('tableSearch').value.toLowerCase();
    const selector = document.getElementById('tableSelector');
    
    Array.from(selector.options).forEach(option => {
        if (option.value === '') return; // Skip placeholder
        const tableName = option.value.toLowerCase();
        option.style.display = tableName.includes(searchTerm) ? 'block' : 'none';
    });
}

// Guardar tabla en repositorio
function saveToRepository() {
    const createStatement = document.getElementById('createTableInput').value.trim();
    
    if (!createStatement) {
        alert('No hay CREATE TABLE para guardar');
        return;
    }
    
    try {
        const tableName = extractTableName(createStatement);
        const schema = extractSchemaName(createStatement);
        
        if (!tableName) {
            throw new Error('No se pudo extraer el nombre de la tabla');
        }
        
        const tableKey = tableName;
        const existsInRepo = tablesRepository[tableKey];
        
        if (existsInRepo && !confirm(`La tabla "${tableName}" ya existe en el repositorio. ¿Sobrescribir?`)) {
            return;
        }
        
        tablesRepository[tableKey] = {
            tableName: tableName,
            schema: schema,
            createStatement: createStatement,
            dateAdded: new Date().toISOString(),
            columns: tableStructure.length || 0
        };
        
        // Guardar en localStorage
        localStorage.setItem('tablesRepository', JSON.stringify(tablesRepository));
        
        // Actualizar UI
        updateTableSelector();
        updateRepositoryStats();
        loadRepositoryList();
        
        alert(`Tabla "${tableName}" guardada en el repositorio`);
        
    } catch (error) {
        alert('Error al guardar: ' + error.message);
    }
}

// Extraer nombre de tabla del CREATE TABLE
function extractTableName(createStatement) {
    const match = createStatement.match(/CREATE\s+TABLE\s+(?:\w+\.)*(\w+)/i);
    return match ? match[1] : null;
}

// Extraer esquema del CREATE TABLE
function extractSchemaName(createStatement) {
    const match = createStatement.match(/CREATE\s+TABLE\s+([\w.]+)\.\w+/i);
    return match ? match[1] : 'unknown_schema';
}

// Mostrar gestor de repositorio
function showRepositoryManager() {
    switchTab('repositorio');
}

// Actualizar estadísticas del repositorio
function updateRepositoryStats() {
    const totalTables = Object.keys(tablesRepository).length;
    const schemas = [...new Set(Object.values(tablesRepository).map(t => t.schema))];
    const totalColumns = Object.values(tablesRepository).reduce((sum, t) => sum + (t.columns || 0), 0);
    
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
    `;
    
    const repoStats = document.getElementById('repoStats');
    if (repoStats) {
        repoStats.innerHTML = statsHtml;
    }
}

// Cargar lista del repositorio
function loadRepositoryList() {
    const container = document.getElementById('repositoryList');
    if (!container) return;
    
    container.innerHTML = '';
    
    // Actualizar filtro de esquemas
    updateSchemaFilter();
    
    Object.entries(tablesRepository).forEach(([tableKey, tableData]) => {
        const item = createRepositoryItem(tableKey, tableData);
        container.appendChild(item);
    });
}

// Crear item del repositorio
function createRepositoryItem(tableKey, tableData) {
    const item = document.createElement('div');
    item.className = 'repository-item';
    item.id = `repo-item-${tableKey}`;
    
    const preview = tableData.createStatement.substring(0, 150) + '...';
    
    item.innerHTML = `
        <div class="repo-item-header">
            <div>
                <div class="repo-item-title">${tableData.tableName}</div>
                <div class="repo-item-schema">${tableData.schema} • ${tableData.columns} columnas</div>
            </div>
            <div class="repo-item-actions">
                <button class="btn btn-small" onclick="loadTableFromRepoItem('${tableKey}')">Usar</button>
                <button class="btn btn-small btn-secondary" onclick="viewTableFull('${tableKey}')">Ver</button>
                <button class="btn btn-small" style="background: #dc3545;" onclick="deleteFromRepo('${tableKey}')">Eliminar</button>
            </div>
        </div>
        <div class="repo-item-preview">${preview}</div>
    `;
    
    return item;
}

// Actualizar filtro de esquemas
function updateSchemaFilter() {
    const filter = document.getElementById('schemaFilter');
    if (!filter) return;
    
    const schemas = [...new Set(Object.values(tablesRepository).map(t => t.schema))];
    filter.innerHTML = '<option value="">Todos los esquemas</option>';
    
    schemas.forEach(schema => {
        const option = document.createElement('option');
        option.value = schema;
        option.textContent = schema;
        filter.appendChild(option);
    });
}

// Filtrar repositorio
function filterRepository() {
    const schemaFilter = document.getElementById('schemaFilter').value;
    const searchTerm = document.getElementById('repoSearch').value.toLowerCase();
    
    Object.keys(tablesRepository).forEach(tableKey => {
        const item = document.getElementById(`repo-item-${tableKey}`);
        const tableData = tablesRepository[tableKey];
        
        const matchesSchema = !schemaFilter || tableData.schema === schemaFilter;
        const matchesSearch = !searchTerm || tableData.tableName.toLowerCase().includes(searchTerm);
        
        if (item) {
            item.style.display = (matchesSchema && matchesSearch) ? 'block' : 'none';
        }
    });
}

// Cargar tabla desde item del repositorio
function loadTableFromRepoItem(tableKey) {
    document.getElementById('createTableInput').value = tablesRepository[tableKey].createStatement;
    autoDetectSchemas(tableKey);
    parseCreateTable();
    switchTab('describe');
    alert(`Tabla "${tableKey}" cargada`);
}

// Ver tabla completa en modal
function viewTableFull(tableKey) {
    const tableData = tablesRepository[tableKey];
    document.getElementById('modalTitle').textContent = `${tableData.tableName} (${tableData.schema})`;
    document.getElementById('modalContent').textContent = tableData.createStatement;
    document.getElementById('tableModal').style.display = 'block';
}

// Cerrar modal
function closeModal() {
    document.getElementById('tableModal').style.display = 'none';
}

// Eliminar del repositorio
function deleteFromRepo(tableKey) {
    if (confirm(`¿Eliminar "${tableKey}" del repositorio?`)) {
        delete tablesRepository[tableKey];
        localStorage.setItem('tablesRepository', JSON.stringify(tablesRepository));
        updateTableSelector();
        updateRepositoryStats();
        loadRepositoryList();
    }
}

// Importar desde archivo
function importFromFile() {
    document.getElementById('fileImport').click();
}

// Manejar importación de archivo
function handleFileImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const content = e.target.result;
        const tables = extractMultipleCreateTables(content);
        
        if (tables.length === 0) {
            alert('No se encontraron CREATE TABLE statements en el archivo');
            return;
        }
        
        let imported = 0;
        tables.forEach(createStatement => {
            try {
                const tableName = extractTableName(createStatement);
                const schema = extractSchemaName(createStatement);
                
                if (tableName) {
                    tablesRepository[tableName] = {
                        tableName,
                        schema,
                        createStatement,
                        dateAdded: new Date().toISOString(),
                        columns: 0 // Se calculará al parsear
                    };
                    imported++;
                }
            } catch (error) {
                console.log('Error procesando tabla:', error);
            }
        });
        
        localStorage.setItem('tablesRepository', JSON.stringify(tablesRepository));
        updateTableSelector();
        updateRepositoryStats();
        loadRepositoryList();
        
        alert(`${imported} tablas importadas al repositorio`);
    };
    
    reader.readAsText(file);
}

// Extraer múltiples CREATE TABLE de un archivo
function extractMultipleCreateTables(content) {
    const createTableRegex = /CREATE\s+TABLE\s+[^;]+;/gi;
    return content.match(createTableRegex) || [];
}

// Exportar repositorio
function exportRepository() {
    const data = Object.values(tablesRepository).map(t => t.createStatement).join('\n\n-- =====================================\n\n');
    
    const blob = new Blob([data], { type: 'text/sql' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'repositorio_create_tables.sql';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

// Limpiar repositorio
function clearRepository() {
    if (confirm('¿Eliminar TODAS las tablas del repositorio? Esta acción no se puede deshacer.')) {
        tablesRepository = {};
        localStorage.setItem('tablesRepository', JSON.stringify(tablesRepository));
        updateTableSelector();
        updateRepositoryStats();
        loadRepositoryList();
        alert('Repositorio limpiado');
    }
}

// FUNCIONES DE ANÁLISIS DE TABLAS

// Parsear CREATE TABLE
function parseCreateTable() {
    const createTableText = document.getElementById('createTableInput').value;
    
    if (!createTableText.trim()) {
        alert('Por favor, pega un CREATE TABLE statement');
        return;
    }
    
    try {
        tableStructure = extractTableStructure(createTableText);
        displayFieldMapping();
        document.getElementById('fieldMapping').style.display = 'block';
        
        // Actualizar columnas en repositorio si existe
        const tableName = extractTableName(createTableText);
        if (tableName && tablesRepository[tableName]) {
            tablesRepository[tableName].columns = tableStructure.length;
            localStorage.setItem('tablesRepository', JSON.stringify(tablesRepository));
        }
        
    } catch (error) {
        alert('Error al analizar la tabla: ' + error.message);
    }
}

// Extraer estructura de tabla
function extractTableStructure(sql) {
    const cleanSql = sql.replace(/\s+/g, ' ').trim();
    const tableMatch = cleanSql.match(/CREATE\s+TABLE\s+[^\(]+\(\s*(.*?)\s*\)\s*(?:USING|PARTITIONED|LOCATION|TBLPROPERTIES|$)/i);
    
    if (!tableMatch) {
        throw new Error('No se pudo encontrar la definición de columnas');
    }
    
    const columnsText = tableMatch[1];
    const columns = [];
    const columnDefinitions = splitByComma(columnsText);
    
    for (let colDef of columnDefinitions) {
        colDef = colDef.trim();
        if (!colDef) continue;
        
        const parts = colDef.split(/\s+/);
        if (parts.length >= 2) {
            const columnName = parts[0];
            let dataType = parts[1];
            
            // Incluir paréntesis si existen
            if (parts.length > 2 && parts[2].startsWith('(')) {
                let parenContent = '';
                for (let i = 2; i < parts.length; i++) {
                    parenContent += parts[i];
                    if (parts[i].includes(')')) break;
                }
                dataType += parenContent;
            }
            
            // Determinar función (count o sum)
            const aggregateFunction = getAggregateFunction(dataType);
            
            columns.push({
                columnName,
                dataType,
                aggregateFunction,
                edvName: parameters.renameRules?.[columnName] || columnName
            });
        }
    }
    
    return columns;
}

// Dividir por comas respetando paréntesis
function splitByComma(text) {
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
}

// Determinar función de agregación
function getAggregateFunction(dataType) {
    const upperType = dataType.toUpperCase();
    if (upperType === 'DOUBLE' || upperType.includes('DECIMAL')) {
        return 'sum';
    }
    return 'count';
}

// Mostrar mapeo de campos
function displayFieldMapping() {
    const content = document.getElementById('fieldMappingContent');
    content.innerHTML = '';
    
    let countFields = 0, sumFields = 0;
    
    tableStructure.forEach((field, index) => {
        if (field.aggregateFunction === 'count') countFields++;
        else sumFields++;
        
        const row = document.createElement('div');
        row.className = 'field-row';
        row.innerHTML = `
            <div>${field.columnName}</div>
            <div>${field.dataType}</div>
            <div>${field.aggregateFunction}()</div>
            <input type="text" class="rename-input" value="${field.edvName}" 
                   onchange="updateFieldName(${index}, this.value)">
        `;
        content.appendChild(row);
    });
    
    // Mostrar estadísticas
    const statsSection = document.getElementById('statsSection');
    statsSection.innerHTML = `
        <div class="stat-card">
            <div class="stat-number">${tableStructure.length}</div>
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
    `;
}

// Actualizar nombre de campo
function updateFieldName(index, newName) {
    tableStructure[index].edvName = newName;
}

// FUNCIONES DE GENERACIÓN DE QUERIES

// Generar todos los queries
function generateAllQueries() {
    if (!tableStructure.length || !parameters.esquemaDDV) {
        alert('Primero completa los parámetros y la estructura de tabla');
        return;
    }
    
    generatedQueries = {
        universos: generateUniversosQuery(),
        agrupados: generateAgrupadosQuery(),
        minus1: generateMinusQuery(true),
        minus2: generateMinusQuery(false)
    };
    
    displayQueries();
    switchTab('queries');
}

// Generar query de universos
function generateUniversosQuery() {
    const periodos = parameters.periodos.replace(/\s/g, '');
    
    return `--UNIVERSOS--
select edv.codmes, numreg_ddv, numreg_edv, numreg_ddv - numreg_edv as diff_numreg 
from (
    select codmes, count(*) numreg_edv 
    from ${parameters.esquemaEDV}.${parameters.tablaEDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
) edv 
inner join (
    select codmes, count(*) numreg_ddv 
    from ${parameters.esquemaDDV}.${parameters.tablaDDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
) ddv on edv.codmes = ddv.codmes 
order by edv.codmes;`;
}

// Generar query de agrupados
function generateAgrupadosQuery() {
    const periodos = parameters.periodos.replace(/\s/g, '');
    const selectFields = tableStructure.map(field => 
        `${field.aggregateFunction}(${field.edvName})`
    ).join(', ');
    
    const selectFieldsDDV = tableStructure.map(field => 
        `${field.aggregateFunction}(${field.columnName})`
    ).join(', ');
    
    return `--AGRUPADOS--
select * from (
select 
'EDV' capa, codmes, ${selectFields}
from ${parameters.esquemaEDV}.${parameters.tablaEDV} 
where codmes in ( ${periodos} ) 
group by codmes
union all
select 
'DDV' capa, codmes, ${selectFieldsDDV}
from ${parameters.esquemaDDV}.${parameters.tablaDDV} 
where codmes in ( ${periodos} ) 
group by codmes
) order by codmes, capa`;
}

// Generar query MINUS
function generateMinusQuery(edvFirst) {
    const periodos = parameters.periodos.replace(/\s/g, '');
    const fieldsEDV = tableStructure.map(f => f.edvName).join(',');
    const fieldsDDV = tableStructure.map(f => f.columnName).join(',');
    
    if (edvFirst) {
        return `--MINUS EDV - DDV--
select ${fieldsEDV}
from ${parameters.esquemaEDV}.${parameters.tablaEDV}
where codmes in ( ${periodos} )
minus all
select ${fieldsDDV}
from ${parameters.esquemaDDV}.${parameters.tablaDDV}
where codmes in ( ${periodos} )`;
    } else {
        return `--MINUS DDV - EDV--
select ${fieldsDDV}
from ${parameters.esquemaDDV}.${parameters.tablaDDV}
where codmes in ( ${periodos} )
minus all
select ${fieldsEDV}
from ${parameters.esquemaEDV}.${parameters.tablaEDV}
where codmes in ( ${periodos} )`;
    }
}

// Mostrar queries generados
function displayQueries() {
    const outputDiv = document.getElementById('queryOutputs');
    outputDiv.innerHTML = '';
    
    const queryTypes = [
        { key: 'universos', title: 'Query de Universos' },
        { key: 'agrupados', title: 'Query de Agrupados' },
        { key: 'minus1', title: 'Query MINUS (EDV - DDV)' },
        { key: 'minus2', title: 'Query MINUS (DDV - EDV)' }
    ];
    
    queryTypes.forEach(({ key, title }) => {
        const section = document.createElement('div');
        section.className = 'output-section';
        section.innerHTML = `
            <h4>${title}</h4>
            <div class="query-output">${generatedQueries[key]}</div>
            <button class="btn btn-secondary" onclick="copyQuery('${key}')">Copiar Query</button>
        `;
        outputDiv.appendChild(section);
    });
}

// Copiar query al portapapeles
function copyQuery(queryKey) {
    navigator.clipboard.writeText(generatedQueries[queryKey]);
    alert('Query copiado al portapapeles');
}

// FUNCIONES DE EXPORTACIÓN

// Exportar a Excel
function exportToExcel() {
    if (!Object.keys(generatedQueries).length) {
        alert('Primero genera los queries');
        return;
    }
    
    const wb = XLSX.utils.book_new();
    
    // Hoja 1: Parámetros
    const wsParams = XLSX.utils.aoa_to_sheet([
        ['ESQUEMA', 'TABLA', 'FINAL (SE USARA EN LAS QUERYS DE RATIFICACION)'],
        ['TABLA EDV', `${parameters.esquemaEDV}`, `${parameters.esquemaEDV}.${parameters.tablaEDV}`],
        ['TABLA DDV', `${parameters.esquemaDDV}`, `${parameters.esquemaDDV}.${parameters.tablaDDV}`],
        [],
        ['PERIODOS', parameters.periodos.split(',')[0].trim(), `in ( ${parameters.periodos} )`]
    ]);
    XLSX.utils.book_append_sheet(wb, wsParams, 'PARAMETROS');
    
    // Hoja 2: Describe
    const describeData = [
        ['col_name', 'data_type', 'comment', 'col_name_EDV', 'METRICA', 'METRICA EDV', 'QUERY'],
        ...tableStructure.map(field => [
            field.columnName,
            field.dataType,
            'null',
            field.edvName,
            `${field.aggregateFunction}(${field.columnName})`,
            `${field.aggregateFunction}(${field.edvName})`,
            `${field.aggregateFunction === 'count' ? 'DDV' : 'EDV'} ${field.aggregateFunction}(${field.columnName})`
        ])
    ];
    const wsDescribe = XLSX.utils.aoa_to_sheet(describeData);
    XLSX.utils.book_append_sheet(wb, wsDescribe, 'TABLA DESCRIBE');
    
    // Hoja 3: Queries
    const queryData = [
        ['TIPO', 'QUERY'],
        ['UNIVERSOS', generatedQueries.universos],
        ['AGRUPADOS', generatedQueries.agrupados],
        ['MINUS_EDV_DDV', generatedQueries.minus1],
        ['MINUS_DDV_EDV', generatedQueries.minus2]
    ];
    const wsQueries = XLSX.utils.aoa_to_sheet(queryData);
    XLSX.utils.book_append_sheet(wb, wsQueries, 'QUERYS DE RATIFICACION');
    
    // Descargar archivo
    XLSX.writeFile(wb, 'queries_ratificacion.xlsx');
}