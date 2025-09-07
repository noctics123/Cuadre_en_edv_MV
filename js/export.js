/**
 * Módulo para exportación a Excel y otros formatos
 */
const ExportModule = {
    
    /**
     * Exporta toda la información a Excel
     */
    exportToExcel() {
        const validation = this.validateExportRequirements();
        if (!validation.isValid) {
            alert('No se puede exportar:\n• ' + validation.errors.join('\n• '));
            return;
        }
        
        try {
            const format = document.getElementById('excelFormat').value;
            const workbook = this.createWorkbook(format);
            
            const filename = this.generateFilename();
            XLSX.writeFile(workbook, filename);
            
            this.showExportPreview(workbook);
            alert(`Excel exportado correctamente: ${filename}`);
            
        } catch (error) {
            alert('Error al exportar Excel: ' + error.message);
            console.error('Error completo:', error);
        }
    },

    /**
     * Valida requisitos para exportación
     * @returns {Object} - {isValid: boolean, errors: Array<string>}
     */
    validateExportRequirements() {
        const errors = [];
        
        const params = ParametersModule.getCurrentParameters();
        if (!params || !params.esquemaDDV) {
            errors.push('Faltan parámetros de configuración');
        }
        
        const tableStructure = TableAnalysisModule.getTableStructure();
        if (!tableStructure || tableStructure.length === 0) {
            errors.push('No hay estructura de tabla definida');
        }
        
        const queries = QueryModule.getGeneratedQueries();
        if (!queries || Object.keys(queries).length === 0) {
            errors.push('No hay queries generados');
        }
        
        return {
            isValid: errors.length === 0,
            errors
        };
    },

    /**
     * Crea el workbook de Excel
     * @param {string} format - Formato de export ('standard' o 'cuadre')
     * @returns {Object} - Workbook de Excel
     */
    createWorkbook(format) {
        const wb = XLSX.utils.book_new();
        
        if (format === 'cuadre') {
            // Formato específico para cuadre (como el template original)
            this.addParametersSheet(wb);
            this.addDescribeSheet(wb);
            this.addQueriesSheet(wb);
            this.addMetadataSheet(wb);
        } else {
            // Formato estándar más completo
            this.addSummarySheet(wb);
            this.addParametersSheet(wb);
            this.addTableStructureSheet(wb);
            this.addQueriesSheet(wb);
            this.addValidationSheet(wb);
            this.addMetadataSheet(wb);
        }
        
        return wb;
    },

    /**
     * Agrega hoja de resumen
     * @param {Object} wb - Workbook
     */
    addSummarySheet(wb) {
        const params = ParametersModule.getCurrentParameters();
        const tableStructure = TableAnalysisModule.getTableStructure();
        const queries = QueryModule.getGeneratedQueries();
        
        const summaryData = [
            ['🏗️ RESUMEN DE CUADRE DDV vs EDV', '', ''],
            ['', '', ''],
            ['📊 INFORMACIÓN GENERAL', '', ''],
            ['Tabla DDV', `${params.esquemaDDV}.${params.tablaDDV}`, ''],
            ['Tabla EDV', `${params.esquemaEDV}.${params.tablaEDV}`, ''],
            ['Períodos', params.periodos, ''],
            ['Total Campos', tableStructure.length, ''],
            ['Campos COUNT', tableStructure.filter(f => f.aggregateFunction === 'count').length, ''],
            ['Campos SUM', tableStructure.filter(f => f.aggregateFunction === 'sum').length, ''],
            ['', '', ''],
            ['📝 QUERIES GENERADOS', '', ''],
            ['Query Universos', queries.universos ? '✅ Generado' : '❌ No generado', ''],
            ['Query Agrupados', queries.agrupados ? '✅ Generado' : '❌ No generado', ''],
            ['Query MINUS (EDV-DDV)', queries.minus1 ? '✅ Generado' : '❌ No generado', ''],
            ['Query MINUS (DDV-EDV)', queries.minus2 ? '✅ Generado' : '❌ No generado', ''],
            ['', '', ''],
            ['📅 METADATOS', '', ''],
            ['Fecha de Generación', new Date().toLocaleString('es-ES'), ''],
            ['Herramienta', 'Generador de Queries de Ratificación v2', ''],
            ['Usuario', 'Sistema', '']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(summaryData);
        
        // Formatear la hoja
        this.formatSummarySheet(ws);
        
        XLSX.utils.book_append_sheet(wb, ws, 'RESUMEN');
    },

    /**
     * Formatea la hoja de resumen
     * @param {Object} ws - Worksheet
     */
    formatSummarySheet(ws) {
        // Configurar anchos de columna
        ws['!cols'] = [
            { width: 25 },
            { width: 50 },
            { width: 15 }
        ];
        
        // Configurar rangos
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } } // Título principal
        ];
    },

    /**
     * Agrega hoja de parámetros
     * @param {Object} wb - Workbook
     */
    addParametersSheet(wb) {
        const params = ParametersModule.getCurrentParameters();
        
        const parametersData = [
            ['PARÁMETRO', 'VALOR', 'DESCRIPCIÓN'],
            ['ESQUEMA DDV', params.esquemaDDV, 'Esquema de producción (DDV)'],
            ['TABLA DDV', params.tablaDDV, 'Tabla de producción'],
            ['ESQUEMA EDV', params.esquemaEDV, 'Esquema de desarrollo (EDV)'],
            ['TABLA EDV', params.tablaEDV, 'Tabla de desarrollo'],
            ['PERÍODOS', params.periodos, 'Períodos a evaluar (formato YYYYMM)'],
            ['', '', ''],
            ['REGLAS DE RENOMBRADO', '', ''],
            ['CAMPO ORIGINAL', 'CAMPO EDV', 'APLICADO']
        ];
        
        // Agregar reglas de renombrado
        if (params.renameRules) {
            Object.entries(params.renameRules).forEach(([original, renamed]) => {
                parametersData.push([original, renamed, '✅']);
            });
        }
        
        // Agregar información de tablas finales
        parametersData.push(['', '', '']);
        parametersData.push(['TABLAS FINALES PARA QUERIES', '', '']);
        parametersData.push(['DDV', `${params.esquemaDDV}.${params.tablaDDV}`, 'Tabla fuente']);
        parametersData.push(['EDV', `${params.esquemaEDV}.${params.tablaEDV}`, 'Tabla destino']);
        
        const ws = XLSX.utils.aoa_to_sheet(parametersData);
        
        // Formatear
        ws['!cols'] = [{ width: 20 }, { width: 40 }, { width: 25 }];
        
        XLSX.utils.book_append_sheet(wb, ws, 'PARAMETROS');
    },

    /**
     * Agrega hoja de estructura de tabla (describe)
     * @param {Object} wb - Workbook
     */
    addDescribeSheet(wb) {
        const tableStructure = TableAnalysisModule.getTableStructure();
        
        const describeData = [
            ['CAMPO_ORIGINAL', 'TIPO_DATO', 'COMENTARIO', 'CAMPO_EDV', 'FUNCION', 'METRICA_DDV', 'METRICA_EDV']
        ];
        
        tableStructure.forEach(field => {
            describeData.push([
                field.columnName,
                field.dataType,
                'null', // Por ahora sin comentarios
                field.edvName,
                field.aggregateFunction.toUpperCase(),
                `${field.aggregateFunction}(${field.columnName})`,
                `${field.aggregateFunction}(${field.edvName})`
            ]);
        });
        
        const ws = XLSX.utils.aoa_to_sheet(describeData);
        
        // Formatear
        ws['!cols'] = [
            { width: 20 }, // CAMPO_ORIGINAL
            { width: 15 }, // TIPO_DATO
            { width: 15 }, // COMENTARIO
            { width: 20 }, // CAMPO_EDV
            { width: 12 }, // FUNCION
            { width: 25 }, // METRICA_DDV
            { width: 25 }  // METRICA_EDV
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'TABLA_DESCRIBE');
    },

    /**
     * Agrega hoja de estructura detallada de tabla
     * @param {Object} wb - Workbook
     */
    addTableStructureSheet(wb) {
        const tableStructure = TableAnalysisModule.getTableStructure();
        
        const structureData = [
            ['#', 'CAMPO_ORIGINAL', 'TIPO_DATO', 'FUNCION_AGREGACION', 'CAMPO_EDV', 'ES_NUMERICO', 'PERMITE_NULOS']
        ];
        
        tableStructure.forEach((field, index) => {
            structureData.push([
                index + 1,
                field.columnName,
                field.dataType,
                field.aggregateFunction,
                field.edvName,
                field.aggregateFunction === 'sum' ? 'SÍ' : 'NO',
                field.isNullable ? 'SÍ' : 'NO'
            ]);
        });
        
        // Agregar estadísticas
        const countFields = tableStructure.filter(f => f.aggregateFunction === 'count').length;
        const sumFields = tableStructure.filter(f => f.aggregateFunction === 'sum').length;
        
        structureData.push(['', '', '', '', '', '', '']);
        structureData.push(['ESTADÍSTICAS', '', '', '', '', '', '']);
        structureData.push(['Total Campos', tableStructure.length, '', '', '', '', '']);
        structureData.push(['Campos COUNT', countFields, '', '', '', '', '']);
        structureData.push(['Campos SUM', sumFields, '', '', '', '', '']);
        structureData.push(['% Numéricos', `${Math.round((sumFields / tableStructure.length) * 100)}%`, '', '', '', '', '']);
        
        const ws = XLSX.utils.aoa_to_sheet(structureData);
        
        // Formatear
        ws['!cols'] = [
            { width: 5 },  // #
            { width: 20 }, // CAMPO_ORIGINAL
            { width: 15 }, // TIPO_DATO
            { width: 15 }, // FUNCION_AGREGACION
            { width: 20 }, // CAMPO_EDV
            { width: 12 }, // ES_NUMERICO
            { width: 15 }  // PERMITE_NULOS
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'ESTRUCTURA_TABLA');
    },

    /**
     * Agrega hoja de queries
     * @param {Object} wb - Workbook
     */
    addQueriesSheet(wb) {
        const queries = QueryModule.getGeneratedQueries();
        
        const queryData = [
            ['TIPO_QUERY', 'DESCRIPCION', 'QUERY_SQL', 'LINEAS', 'CARACTERES'],
            [
                'UNIVERSOS',
                'Compara número total de registros entre DDV y EDV',
                queries.universos || '',
                queries.universos ? queries.universos.split('\n').length : 0,
                queries.universos ? queries.universos.length : 0
            ],
            [
                'AGRUPADOS',
                'Compara métricas agregadas por cada campo',
                queries.agrupados || '',
                queries.agrupados ? queries.agrupados.split('\n').length : 0,
                queries.agrupados ? queries.agrupados.length : 0
            ],
            [
                'MINUS_EDV_DDV',
                'Registros que están en EDV pero NO en DDV',
                queries.minus1 || '',
                queries.minus1 ? queries.minus1.split('\n').length : 0,
                queries.minus1 ? queries.minus1.length : 0
            ],
            [
                'MINUS_DDV_EDV',
                'Registros que están en DDV pero NO en EDV',
                queries.minus2 || '',
                queries.minus2 ? queries.minus2.split('\n').length : 0,
                queries.minus2 ? queries.minus2.length : 0
            ]
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(queryData);
        
        // Formatear
        ws['!cols'] = [
            { width: 20 }, // TIPO_QUERY
            { width: 40 }, // DESCRIPCION
            { width: 80 }, // QUERY_SQL
            { width: 10 }, // LINEAS
            { width: 12 }  // CARACTERES
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'QUERIES_RATIFICACION');
    },

    /**
     * Agrega hoja de validación
     * @param {Object} wb - Workbook
     */
    addValidationSheet(wb) {
        const params = ParametersModule.getCurrentParameters();
        const tableStructure = TableAnalysisModule.getTableStructure();
        const queries = QueryModule.getGeneratedQueries();
        
        const validationData = [
            ['ASPECTO', 'ESTADO', 'DETALLES', 'RECOMENDACIÓN'],
            ['Parámetros', '', '', ''],
            ['- Esquema DDV', params.esquemaDDV ? '✅' : '❌', params.esquemaDDV || 'No definido', params.esquemaDDV ? '' : 'Definir esquema DDV'],
            ['- Esquema EDV', params.esquemaEDV ? '✅' : '❌', params.esquemaEDV || 'No definido', params.esquemaEDV ? '' : 'Definir esquema EDV'],
            ['- Períodos', params.periodos ? '✅' : '❌', params.periodos || 'No definidos', params.periodos ? '' : 'Definir períodos'],
            ['', '', '', ''],
            ['Estructura de Tabla', '', '', ''],
            ['- Campos definidos', tableStructure.length > 0 ? '✅' : '❌', `${tableStructure.length} campos`, tableStructure.length > 0 ? '' : 'Analizar CREATE TABLE'],
            ['- Campos numéricos', '', `${tableStructure.filter(f => f.aggregateFunction === 'sum').length} campos SUM`, ''],
            ['- Mapeo EDV', '', `${tableStructure.filter(f => f.edvName !== f.columnName).length} campos renombrados`, ''],
            ['', '', '', ''],
            ['Queries Generados', '', '', ''],
            ['- Query Universos', queries.universos ? '✅' : '❌', queries.universos ? 'Generado' : 'No generado', queries.universos ? '' : 'Generar queries'],
            ['- Query Agrupados', queries.agrupados ? '✅' : '❌', queries.agrupados ? 'Generado' : 'No generado', queries.agrupados ? '' : 'Generar queries'],
            ['- Queries MINUS', (queries.minus1 && queries.minus2) ? '✅' : '❌', (queries.minus1 && queries.minus2) ? 'Ambos generados' : 'Faltantes', (queries.minus1 && queries.minus2) ? '' : 'Generar queries']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(validationData);
        
        // Formatear
        ws['!cols'] = [
            { width: 25 }, // ASPECTO
            { width: 10 }, // ESTADO
            { width: 40 }, // DETALLES
            { width: 30 }  // RECOMENDACIÓN
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'VALIDACION');
    },

    /**
     * Agrega hoja de metadatos
     * @param {Object} wb - Workbook
     */
    addMetadataSheet(wb) {
        const tableStructure = TableAnalysisModule.getTableStructure();
        const repository = RepositoryModule.getRepository();
        
        const metadataData = [
            ['METADATO', 'VALOR', 'DESCRIPCIÓN'],
            ['Fecha de Generación', new Date().toISOString(), 'Timestamp de creación del archivo'],
            ['Herramienta', 'Generador de Queries de Ratificación v2', 'Versión de la herramienta utilizada'],
            ['Versión Regex', '2.0', 'Versión de patrones regex utilizados'],
            ['', '', ''],
            ['ESTADÍSTICAS DE SESIÓN', '', ''],
            ['Tablas en Repositorio', Object.keys(repository).length, 'Total de tablas guardadas'],
            ['Campos Analizados', tableStructure.length, 'Total de campos en la tabla actual'],
            ['Reglas de Renombrado', Object.keys(ParametersModule.getCurrentParameters().renameRules || {}).length, 'Número de reglas aplicadas'],
            ['', '', ''],
            ['INFORMACIÓN TÉCNICA', '', ''],
            ['Navegador', navigator.userAgent, 'User Agent del navegador'],
            ['Resolución', `${screen.width}x${screen.height}`, 'Resolución de pantalla'],
            ['Zona Horaria', Intl.DateTimeFormat().resolvedOptions().timeZone, 'Zona horaria del usuario'],
            ['Idioma', navigator.language, 'Idioma del navegador']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(metadataData);
        
        // Formatear
        ws['!cols'] = [
            { width: 25 }, // METADATO
            { width: 40 }, // VALOR
            { width: 35 }  // DESCRIPCIÓN
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'METADATOS');
    },

    /**
     * Genera nombre de archivo para export
     * @returns {string} - Nombre del archivo
     */
    generateFilename() {
        const params = ParametersModule.getCurrentParameters();
        const date = new Date().toISOString().split('T')[0];
        const time = new Date().toTimeString().split(' ')[0].replace(/:/g, '');
        
        const tableName = params.tablaDDV || 'tabla';
        const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
        
        return `cuadre_${tableName}_${periods}_${date}_${time}.xlsx`;
    },

    /**
     * Muestra vista previa del export
     * @param {Object} workbook - Workbook exportado
     */
    showExportPreview(workbook) {
        const preview = document.getElementById('exportPreview');
        if (!preview) return;
        
        const sheetNames = workbook.SheetNames;
        const totalSheets = sheetNames.length;
        
        let previewHTML = `
            <div class="export-preview">
                <h4>📊 Export Completado</h4>
                <div class="export-stats">
                    <div class="stat-card">
                        <div class="stat-number">${totalSheets}</div>
                        <div class="stat-label">Hojas Generadas</div>
                    </div>
                </div>
                <div class="sheets-list">
                    <h5>Hojas incluidas:</h5>
                    <ul>
        `;
        
        sheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
            const rows = range.e.r + 1;
            const cols = range.e.c + 1;
            
            previewHTML += `<li><strong>${sheetName}</strong> - ${rows} filas x ${cols} columnas</li>`;
        });
        
        previewHTML += `
                    </ul>
                </div>
                <div class="export-actions">
                    <button class="btn btn-secondary" onclick="ExportModule.exportToExcel()">🔄 Exportar Nuevamente</button>
                </div>
            </div>
        `;
        
        preview.innerHTML = previewHTML;
    },

    /**
     * Exporta queries individuales
     * @param {string} format - Formato de export ('sql', 'txt', 'json')
     */
    exportQueries(format = 'sql') {
        const queries = QueryModule.getGeneratedQueries();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar');
            return;
        }
        
        let content = '';
        let filename = '';
        let mimeType = '';
        
        switch (format) {
            case 'sql':
                content = this.generateSQLFile(queries);
                filename = `queries_ratificacion_${new Date().toISOString().split('T')[0]}.sql`;
                mimeType = 'text/sql';
                break;
                
            case 'json':
                content = JSON.stringify(queries, null, 2);
                filename = `queries_ratificacion_${new Date().toISOString().split('T')[0]}.json`;
                mimeType = 'application/json';
                break;
                
            case 'txt':
                content = this.generateTextFile(queries);
                filename = `queries_ratificacion_${new Date().toISOString().split('T')[0]}.txt`;
                mimeType = 'text/plain';
                break;
                
            default:
                throw new Error(`Formato no soportado: ${format}`);
        }
        
        this.downloadFile(content, filename, mimeType);
    },

    /**
     * Genera archivo SQL con todos los queries
     * @param {Object} queries - Queries generados
     * @returns {string} - Contenido del archivo SQL
     */
    generateSQLFile(queries) {
        const params = ParametersModule.getCurrentParameters();
        const date = new Date().toISOString();
        
        let content = `-- =====================================================\n`;
        content += `-- QUERIES DE RATIFICACIÓN DDV vs EDV\n`;
        content += `-- Generado: ${date}\n`;
        content += `-- Tabla DDV: ${params.esquemaDDV}.${params.tablaDDV}\n`;
        content += `-- Tabla EDV: ${params.esquemaEDV}.${params.tablaEDV}\n`;
        content += `-- Períodos: ${params.periodos}\n`;
        content += `-- =====================================================\n\n`;
        
        Object.entries(queries).forEach(([key, query]) => {
            const title = this.getQueryTitle(key);
            content += `-- ${title}\n`;
            content += `-- ${'-'.repeat(title.length)}\n`;
            content += query;
            content += '\n\n';
        });
        
        return content;
    },

    /**
     * Genera archivo de texto con información completa
     * @param {Object} queries - Queries generados
     * @returns {string} - Contenido del archivo de texto
     */
    generateTextFile(queries) {
        const params = ParametersModule.getCurrentParameters();
        const tableStructure = TableAnalysisModule.getTableStructure();
        
        let content = `REPORTE DE CUADRE DDV vs EDV\n`;
        content += `${'='.repeat(50)}\n\n`;
        
        content += `CONFIGURACIÓN:\n`;
        content += `- Esquema DDV: ${params.esquemaDDV}\n`;
        content += `- Tabla DDV: ${params.tablaDDV}\n`;
        content += `- Esquema EDV: ${params.esquemaEDV}\n`;
        content += `- Tabla EDV: ${params.tablaEDV}\n`;
        content += `- Períodos: ${params.periodos}\n\n`;
        
        content += `ESTRUCTURA DE TABLA:\n`;
        content += `- Total campos: ${tableStructure.length}\n`;
        content += `- Campos COUNT: ${tableStructure.filter(f => f.aggregateFunction === 'count').length}\n`;
        content += `- Campos SUM: ${tableStructure.filter(f => f.aggregateFunction === 'sum').length}\n\n`;
        
        content += `QUERIES GENERADOS:\n`;
        Object.entries(queries).forEach(([key, query]) => {
            const title = this.getQueryTitle(key);
            content += `\n${title}:\n`;
            content += `${'-'.repeat(title.length + 1)}\n`;
            content += query;
            content += '\n';
        });
        
        return content;
    },

    /**
     * Obtiene título descriptivo para un query
     * @param {string} queryKey - Clave del query
     * @returns {string} - Título descriptivo
     */
    getQueryTitle(queryKey) {
        const titles = {
            'universos': 'QUERY DE UNIVERSOS',
            'agrupados': 'QUERY DE AGRUPADOS',
            'minus1': 'QUERY MINUS (EDV - DDV)',
            'minus2': 'QUERY MINUS (DDV - EDV)'
        };
        
        return titles[queryKey] || queryKey.toUpperCase();
    },

    /**
     * Exporta todos los queries en un único archivo TXT
     */
    exportAllQueriesTXT() {
        const queries = QueryModule.getGeneratedQueries();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar. Primero genera los queries en la pestaña correspondiente.');
            return;
        }
        
        const params = ParametersModule.getCurrentParameters();
        const tableStructure = TableAnalysisModule.getTableStructure();
        
        let content = `REPORTE COMPLETO DE CUADRE DDV vs EDV\n`;
        content += `${'='.repeat(60)}\n\n`;
        
        // Información del proyecto
        content += `INFORMACIÓN DEL PROYECTO:\n`;
        content += `${'-'.repeat(30)}\n`;
        content += `Fecha de generación: ${new Date().toLocaleString('es-ES')}\n`;
        content += `Herramienta: Generador de Queries de Ratificación v2\n`;
        content += `\n`;
        content += `CONFIGURACIÓN:\n`;
        content += `- Esquema DDV: ${params.esquemaDDV}\n`;
        content += `- Tabla DDV: ${params.tablaDDV}\n`;
        content += `- Esquema EDV: ${params.esquemaEDV}\n`;
        content += `- Tabla EDV: ${params.tablaEDV}\n`;
        content += `- Períodos: ${params.periodos}\n`;
        content += `\n`;
        
        // Información de la estructura
        content += `ESTRUCTURA DE TABLA:\n`;
        content += `${'-'.repeat(25)}\n`;
        content += `- Total campos: ${tableStructure.length}\n`;
        content += `- Campos COUNT: ${tableStructure.filter(f => f.aggregateFunction === 'count').length}\n`;
        content += `- Campos SUM: ${tableStructure.filter(f => f.aggregateFunction === 'sum').length}\n`;
        if (tableStructure.length > 0) {
            content += `- Porcentaje numérico: ${Math.round((tableStructure.filter(f => f.aggregateFunction === 'sum').length / tableStructure.length) * 100)}%\n`;
        }
        content += `\n`;
        
        // Resumen de queries
        content += `QUERIES GENERADOS:\n`;
        content += `${'-'.repeat(20)}\n`;
        Object.entries(queries).forEach(([key, query]) => {
            const title = this.getQueryTitle(key);
            const lines = query.split('\n').length;
            const chars = query.length;
            content += `- ${title}: ${lines} líneas, ${chars} caracteres\n`;
        });
        content += `\n`;
        
        // Instrucciones de uso
        content += `INSTRUCCIONES DE USO:\n`;
        content += `${'-'.repeat(25)}\n`;
        content += `1. Copiar el query deseado completo\n`;
        content += `2. Pegar en tu editor SQL preferido\n`;
        content += `3. Ejecutar en el motor de base de datos correspondiente\n`;
        content += `4. Analizar los resultados para identificar diferencias\n`;
        content += `5. Documentar hallazgos para seguimiento y corrección\n`;
        content += `\n`;
        
        // Queries completos
        const queryDescriptions = {
            universos: 'Compara el número total de registros entre DDV y EDV',
            agrupados: 'Compara métricas agregadas por cada campo',
            minus1: 'Registros que están en EDV pero NO en DDV',
            minus2: 'Registros que están en DDV pero NO en EDV'
        };
        
        Object.entries(queries).forEach(([key, query]) => {
            const title = this.getQueryTitle(key);
            const description = queryDescriptions[key];
            
            content += `\n\n${'#'.repeat(80)}\n`;
            content += `${title}\n`;
            content += `${'#'.repeat(80)}\n\n`;
            content += `DESCRIPCIÓN:\n${description}\n\n`;
            content += `PARÁMETROS UTILIZADOS:\n`;
            content += `- Esquema DDV: ${params.esquemaDDV}\n`;
            content += `- Tabla DDV: ${params.tablaDDV}\n`;
            content += `- Esquema EDV: ${params.esquemaEDV}\n`;
            content += `- Tabla EDV: ${params.tablaEDV}\n`;
            content += `- Períodos: ${params.periodos}\n\n`;
            content += `QUERY SQL:\n`;
            content += `${'-'.repeat(40)}\n`;
            content += query;
            content += `\n${'-'.repeat(40)}\n`;
        });
        
        // Footer
        content += `\n\n${'='.repeat(80)}\n`;
        content += `FIN DEL REPORTE - Generado por: Generador de Queries de Ratificación v2\n`;
        content += `Fecha: ${new Date().toISOString()}\n`;
        content += `${'='.repeat(80)}`;
        
        // Descargar archivo
        const filename = `reporte_completo_queries_${new Date().toISOString().split('T')[0]}.txt`;
        this.downloadFile(content, filename, 'text/plain');
        
        UIModule.showNotification(
            `📋 Reporte completo descargado: ${filename}`,
            'success',
            4000
        );
    },

    /**
     * Descarga archivo con contenido específico
     * @param {string} content - Contenido del archivo
     * @param {string} filename - Nombre del archivo
     * @param {string} mimeType - Tipo MIME del archivo
     */
    downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }
};