/**
 * Módulo para generación de queries de ratificación
 */
const QueryModule = {
    
    // Variables del módulo
    generatedQueries: {},
    
    /**
     * Genera todos los queries de ratificación
     */
    generateAllQueries() {
        // Validar requisitos previos
        const validation = this.validateRequirements();
        if (!validation.isValid) {
            alert('No se pueden generar queries:\n• ' + validation.errors.join('\n• '));
            return;
        }
        
        try {
            const params = ParametersModule.getCurrentParameters();
            const tableStructure = TableAnalysisModule.getTableStructure();
            
            // Generar cada tipo de query
            this.generatedQueries = {
                universos: this.generateUniversosQuery(params),
                agrupados: this.generateAgrupadosQuery(params, tableStructure),
                minus1: this.generateMinusQuery(params, tableStructure, 'edv_minus_ddv'),
                minus2: this.generateMinusQuery(params, tableStructure, 'ddv_minus_edv')
            };
            
            // Mostrar queries en la interfaz
            this.displayQueries();
            
            // Cambiar a la pestaña de queries
            window.switchTab('queries');
            
            alert('Todos los queries han sido generados correctamente');
            
        } catch (error) {
            alert('Error generando queries: ' + error.message);
            console.error('Error completo:', error);
        }
    },

    /**
     * Valida los requisitos para generar queries
     * @returns {Object} - {isValid: boolean, errors: Array<string>}
     */
    validateRequirements() {
        const errors = [];
        
        // Validar parámetros
        const params = ParametersModule.getCurrentParameters();
        const paramValidation = Utils.validateParameters(params);
        if (!paramValidation.isValid) {
            errors.push(...paramValidation.errors);
        }
        
        // Validar estructura de tabla
        const tableValidation = TableAnalysisModule.validateTableStructure();
        if (!tableValidation.isValid) {
            errors.push(...tableValidation.errors);
        }
        
        return {
            isValid: errors.length === 0,
            errors
        };
    },

    /**
     * Genera query de universos (conteo de registros)
     * @param {Object} params - Parámetros de configuración
     * @returns {string} - Query de universos
     */
    generateUniversosQuery(params) {
        const periodos = Utils.formatPeriods(params.periodos);
        
        return `-- QUERY DE UNIVERSOS
-- Compara el número total de registros entre DDV y EDV
select 
    edv.codmes, 
    numreg_ddv, 
    numreg_edv, 
    numreg_ddv - numreg_edv as diff_numreg,
    CASE 
        WHEN numreg_ddv = numreg_edv THEN '✅ IGUALES'
        WHEN numreg_ddv > numreg_edv THEN '⚠️ FALTAN EN EDV'
        ELSE '⚠️ SOBRAN EN EDV'
    END as status
from (
    select codmes, count(*) numreg_edv 
    from ${params.esquemaEDV}.${params.tablaEDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
) edv 
inner join (
    select codmes, count(*) numreg_ddv 
    from ${params.esquemaDDV}.${params.tablaDDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
) ddv on edv.codmes = ddv.codmes 
order by edv.codmes;`;
    },

    /**
     * Genera query de agrupados (métricas por campo)
     * @param {Object} params - Parámetros de configuración
     * @param {Array} tableStructure - Estructura de la tabla
     * @returns {string} - Query de agrupados
     */
    generateAgrupadosQuery(params, tableStructure) {
        const periodos = Utils.formatPeriods(params.periodos);
        
        // Generar campos para EDV
        const selectFieldsEDV = tableStructure.map(field => 
            `${field.aggregateFunction}(${field.edvName}) as ${field.columnName}_${field.aggregateFunction}`
        ).join(',\n    ');
        
        // Generar campos para DDV
        const selectFieldsDDV = tableStructure.map(field => 
            `${field.aggregateFunction}(${field.columnName}) as ${field.columnName}_${field.aggregateFunction}`
        ).join(',\n    ');
        
        return `-- QUERY DE AGRUPADOS
-- Compara métricas agregadas por campo entre DDV y EDV
select * from (
    select 
        'EDV' as capa, 
        codmes,
        ${selectFieldsEDV}
    from ${params.esquemaEDV}.${params.tablaEDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
    
    union all
    
    select 
        'DDV' as capa, 
        codmes,
        ${selectFieldsDDV}
    from ${params.esquemaDDV}.${params.tablaDDV} 
    where codmes in ( ${periodos} ) 
    group by codmes
) 
order by codmes, capa;`;
    },

    /**
     * Genera query MINUS para detectar diferencias
     * @param {Object} params - Parámetros de configuración
     * @param {Array} tableStructure - Estructura de la tabla
     * @param {string} type - Tipo de MINUS ('edv_minus_ddv' o 'ddv_minus_edv')
     * @returns {string} - Query MINUS
     */
    generateMinusQuery(params, tableStructure, type) {
        const periodos = Utils.formatPeriods(params.periodos);
        
        // Generar lista de campos
        const fieldsEDV = tableStructure.map(f => f.edvName).join(',\n    ');
        const fieldsDDV = tableStructure.map(f => f.columnName).join(',\n    ');
        
        if (type === 'edv_minus_ddv') {
            return `-- QUERY MINUS (EDV - DDV)
-- Registros que están en EDV pero NO en DDV
select 
    ${fieldsEDV}
from ${params.esquemaEDV}.${params.tablaEDV}
where codmes in ( ${periodos} )

minus all

select 
    ${fieldsDDV}
from ${params.esquemaDDV}.${params.tablaDDV}
where codmes in ( ${periodos} );`;
        } else {
            return `-- QUERY MINUS (DDV - EDV)
-- Registros que están en DDV pero NO en EDV
select 
    ${fieldsDDV}
from ${params.esquemaDDV}.${params.tablaDDV}
where codmes in ( ${periodos} )

minus all

select 
    ${fieldsEDV}
from ${params.esquemaEDV}.${params.tablaEDV}
where codmes in ( ${periodos} );`;
        }
    },

    /**
     * Genera query personalizado
     * @param {string} queryType - Tipo de query personalizado
     * @param {Object} options - Opciones adicionales
     * @returns {string} - Query personalizado
     */
    generateCustomQuery(queryType, options = {}) {
        const params = ParametersModule.getCurrentParameters();
        const tableStructure = TableAnalysisModule.getTableStructure();
        const periodos = Utils.formatPeriods(params.periodos);
        
        switch (queryType) {
            case 'sample_data':
                return this.generateSampleDataQuery(params, tableStructure, periodos, options);
            
            case 'data_quality':
                return this.generateDataQualityQuery(params, tableStructure, periodos, options);
            
            case 'field_comparison':
                return this.generateFieldComparisonQuery(params, tableStructure, periodos, options);
            
            default:
                throw new Error(`Tipo de query no soportado: ${queryType}`);
        }
    },

    /**
     * Genera query para muestra de datos
     * @param {Object} params - Parámetros
     * @param {Array} tableStructure - Estructura
     * @param {string} periodos - Períodos
     * @param {Object} options - Opciones
     * @returns {string} - Query de muestra
     */
    generateSampleDataQuery(params, tableStructure, periodos, options) {
        const limit = options.limit || 100;
        const fields = tableStructure.slice(0, 10).map(f => f.columnName).join(', ');
        
        return `-- MUESTRA DE DATOS
-- Primeros ${limit} registros para validación manual
select top ${limit}
    ${fields}
from ${params.esquemaDDV}.${params.tablaDDV}
where codmes in ( ${periodos} )
order by codmes, codclaveunicocli;`;
    },

    /**
     * Genera query para validación de calidad de datos
     * @param {Object} params - Parámetros
     * @param {Array} tableStructure - Estructura
     * @param {string} periodos - Períodos
     * @param {Object} options - Opciones
     * @returns {string} - Query de calidad
     */
    generateDataQualityQuery(params, tableStructure, periodos, options) {
        const numericFields = tableStructure.filter(f => f.aggregateFunction === 'sum');
        const nullChecks = numericFields.map(f => 
            `sum(case when ${f.columnName} is null then 1 else 0 end) as ${f.columnName}_nulls`
        ).join(',\n    ');
        
        return `-- VALIDACIÓN DE CALIDAD DE DATOS
select 
    'DDV' as capa,
    codmes,
    count(*) as total_registros,
    ${nullChecks}
from ${params.esquemaDDV}.${params.tablaDDV}
where codmes in ( ${periodos} )
group by codmes
order by codmes;`;
    },

    /**
     * Genera query para comparación específica de campos
     * @param {Object} params - Parámetros
     * @param {Array} tableStructure - Estructura
     * @param {string} periodos - Períodos
     * @param {Object} options - Opciones
     * @returns {string} - Query de comparación
     */
    generateFieldComparisonQuery(params, tableStructure, periodos, options) {
        const targetField = options.field || tableStructure[0]?.columnName || 'codclaveunicocli';
        
        return `-- COMPARACIÓN ESPECÍFICA DE CAMPO: ${targetField}
select 
    coalesce(ddv.${targetField}, edv.${targetField}) as ${targetField},
    ddv.codmes as ddv_codmes,
    edv.codmes as edv_codmes,
    case 
        when ddv.${targetField} is null then 'SOLO_EN_EDV'
        when edv.${targetField} is null then 'SOLO_EN_DDV'
        else 'EN_AMBOS'
    end as status
from (
    select distinct ${targetField}, codmes
    from ${params.esquemaDDV}.${params.tablaDDV}
    where codmes in ( ${periodos} )
) ddv
full outer join (
    select distinct ${targetField}, codmes
    from ${params.esquemaEDV}.${params.tablaEDV}
    where codmes in ( ${periodos} )
) edv on ddv.${targetField} = edv.${targetField} and ddv.codmes = edv.codmes
where ddv.${targetField} is null or edv.${targetField} is null
order by ${targetField}, ddv_codmes, edv_codmes;`;
    },

    /**
     * Muestra queries generados en la interfaz
     */
    displayQueries() {
        const outputDiv = document.getElementById('queryOutputs');
        if (!outputDiv) return;
        
        outputDiv.innerHTML = '';
        
        const queryTypes = [
            { 
                key: 'universos', 
                title: '📊 Query de Universos',
                description: 'Compara el número total de registros entre DDV y EDV'
            },
            { 
                key: 'agrupados', 
                title: '📈 Query de Agrupados',
                description: 'Compara métricas agregadas por cada campo'
            },
            { 
                key: 'minus1', 
                title: '🔍 Query MINUS (EDV - DDV)',
                description: 'Registros que están en EDV pero NO en DDV'
            },
            { 
                key: 'minus2', 
                title: '🔍 Query MINUS (DDV - EDV)',
                description: 'Registros que están en DDV pero NO en EDV'
            }
        ];
        
        queryTypes.forEach(({ key, title, description }) => {
            if (this.generatedQueries[key]) {
                const section = this.createQuerySection(key, title, description, this.generatedQueries[key]);
                outputDiv.appendChild(section);
            }
        });
        
        // Agregar sección de queries personalizados
        this.addCustomQuerySection(outputDiv);
    },

    /**
     * Crea una sección de query
     * @param {string} key - Clave del query
     * @param {string} title - Título
     * @param {string} description - Descripción
     * @param {string} query - Query SQL
     * @returns {HTMLElement} - Elemento de la sección
     */
    createQuerySection(key, title, description, query) {
        const section = Utils.createElement('div', 'output-section');
        
        section.innerHTML = `
            <div class="query-header">
                <h4>${title}</h4>
                <p class="query-description">${description}</p>
                <div class="query-stats">
                    Líneas: ${query.split('\n').length} | 
                    Caracteres: ${query.length} |
                    Palabras: ${query.split(/\s+/).length}
                </div>
            </div>
            <div class="query-output">${this.highlightSQL(query)}</div>
            <div class="query-actions">
                <button class="btn btn-secondary" onclick="QueryModule.copyQuery('${key}')">📋 Copiar Query</button>
                <button class="btn btn-secondary" onclick="QueryModule.downloadQuery('${key}')">💾 Descargar</button>
                <button class="btn btn-secondary" onclick="QueryModule.validateQuery('${key}')">✅ Validar</button>
            </div>
        `;
        
        return section;
    },

    /**
     * Resalta sintaxis SQL básica
     * @param {string} sql - Query SQL
     * @returns {string} - SQL con sintaxis resaltada
     */
    highlightSQL(sql) {
        return sql
            .replace(/\b(SELECT|FROM|WHERE|GROUP BY|ORDER BY|UNION|JOIN|INNER|LEFT|RIGHT|OUTER|ON|AS|CASE|WHEN|THEN|ELSE|END|COUNT|SUM|AVG|MIN|MAX|DISTINCT|TOP|LIMIT)\b/gi, '<span class="sql-keyword">$1</span>')
            .replace(/\b(AND|OR|NOT|IN|EXISTS|BETWEEN|LIKE|IS|NULL)\b/gi, '<span class="sql-operator">$1</span>')
            .replace(/--([^\n]*)/g, '<span class="sql-comment">--$1</span>')
            .replace(/('[^']*')/g, '<span class="sql-string">$1</span>');
    },

    /**
     * Agrega sección de queries personalizados
     * @param {HTMLElement} container - Contenedor
     */
    addCustomQuerySection(container) {
        const customSection = Utils.createElement('div', 'output-section custom-queries');
        
        customSection.innerHTML = `
            <h4>🛠️ Queries Personalizados</h4>
            <div class="custom-query-controls">
                <select id="customQueryType">
                    <option value="">-- Seleccionar tipo --</option>
                    <option value="sample_data">Muestra de Datos</option>
                    <option value="data_quality">Validación de Calidad</option>
                    <option value="field_comparison">Comparación de Campo</option>
                </select>
                <button class="btn" onclick="QueryModule.generateAndShowCustomQuery()">Generar Query Personalizado</button>
            </div>
            <div id="customQueryOutput" style="display: none;"></div>
        `;
        
        container.appendChild(customSection);
    },

    /**
     * Genera y muestra query personalizado
     */
    generateAndShowCustomQuery() {
        const queryType = document.getElementById('customQueryType').value;
        if (!queryType) {
            alert('Selecciona un tipo de query personalizado');
            return;
        }
        
        try {
            const customQuery = this.generateCustomQuery(queryType);
            const outputDiv = document.getElementById('customQueryOutput');
            
            outputDiv.innerHTML = `
                <div class="query-output">${this.highlightSQL(customQuery)}</div>
                <button class="btn btn-secondary" onclick="Utils.copyToClipboard(\`${customQuery.replace(/`/g, '\\`')}\`, 'Query personalizado copiado')">📋 Copiar Query</button>
            `;
            outputDiv.style.display = 'block';
            
        } catch (error) {
            alert('Error generando query personalizado: ' + error.message);
        }
    },

    /**
     * Copia query al portapapeles
     * @param {string} queryKey - Clave del query
     */
    async copyQuery(queryKey) {
        if (!this.generatedQueries[queryKey]) {
            alert('Query no encontrado');
            return;
        }
        
        await Utils.copyToClipboard(this.generatedQueries[queryKey], 'Query copiado al portapapeles');
    },

    /**
     * Descarga query como archivo
     * @param {string} queryKey - Clave del query
     */
    downloadQuery(queryKey) {
        if (!this.generatedQueries[queryKey]) {
            alert('Query no encontrado');
            return;
        }
        
        const query = this.generatedQueries[queryKey];
        const filename = `query_${queryKey}_${new Date().toISOString().split('T')[0]}.sql`;
        
        const blob = new Blob([query], { type: 'text/sql' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    },

    /**
     * Valida sintaxis básica del query
     * @param {string} queryKey - Clave del query
     */
    validateQuery(queryKey) {
        if (!this.generatedQueries[queryKey]) {
            alert('Query no encontrado');
            return;
        }
        
        const query = this.generatedQueries[queryKey];
        const issues = [];
        
        // Validaciones básicas
        if (!query.includes('SELECT')) issues.push('No contiene SELECT');
        if (!query.includes('FROM')) issues.push('No contiene FROM');
        if (query.split('(').length !== query.split(')').length) {
            issues.push('Paréntesis desbalanceados');
        }
        
        // Verificar nombres de esquemas y tablas
        const params = ParametersModule.getCurrentParameters();
        if (!query.includes(params.esquemaDDV)) issues.push('Falta esquema DDV');
        if (!query.includes(params.esquemaEDV)) issues.push('Falta esquema EDV');
        
        if (issues.length === 0) {
            alert('✅ Query validado correctamente');
        } else {
            alert('⚠️ Problemas encontrados:\n• ' + issues.join('\n• '));
        }
    },

    /**
     * Obtiene queries generados
     * @returns {Object} - Queries generados
     */
    getGeneratedQueries() {
        return this.generatedQueries;
    },

    /**
     * Resetea queries generados
     */
    resetQueries() {
        this.generatedQueries = {};
        const outputDiv = document.getElementById('queryOutputs');
        if (outputDiv) {
            outputDiv.innerHTML = '<p style="text-align: center; color: #6c757d;">No hay queries generados</p>';
        }
        
        // Ocultar botones de export rápido
        this.hideQuickExportButtons();
        
        // Actualizar botones de export individual
        if (typeof ExportModule !== 'undefined' && ExportModule.updateIndividualExportButtons) {
            ExportModule.updateIndividualExportButtons();
        }
    }
};