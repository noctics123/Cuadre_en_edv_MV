/**
 * Módulo para exportación a Excel y otros formatos
 * VERSIÓN ACTUALIZADA CON EXCELJS Y DISEÑO PROFESIONAL
 */
const ExportModule = {
    
    /**
     * Exporta Excel específico para V3 - Formato Cuadre EDV (DISEÑO PROFESIONAL CON EXCELJS)
     */
    async exportCuadreEDV() {
        const queries = QueryModule.getGeneratedQueries();
        const params = ParametersModule.getCurrentParameters();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar. Primero genera los queries en la pestaña correspondiente.');
            return;
        }

        try {
            // Importar ExcelJS dinámicamente
            const ExcelJS = await this.loadExcelJS();
            
            // Crear workbook
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Cuadre DDV vs EDV', {
                pageSetup: { paperSize: 9, orientation: 'landscape' }
            });

            // Configurar anchos de columna (A=15, B-K=20-24 cada una)
            worksheet.columns = [
                { width: 15 }, // A - Etiquetas
                { width: 22 }, // B - Contenido
                { width: 22 }, // C
                { width: 22 }, // D
                { width: 22 }, // E
                { width: 22 }, // F
                { width: 22 }, // G
                { width: 22 }, // H
                { width: 22 }, // I
                { width: 22 }, // J
                { width: 22 }  // K
            ];

            let currentRow = 1;

            // 1. TÍTULO PRINCIPAL
            currentRow = this.addMainTitle(worksheet, currentRow);
            
            // Congelar paneles en fila 2
            worksheet.views = [{ state: 'frozen', ySplit: 2 }];

            // 2. SECCIÓN UNIVERSOS
            currentRow = await this.addUniversosSection(worksheet, currentRow, queries.universos, params);

            // 3. SECCIÓN AGRUPADOS  
            currentRow = await this.addAgrupadosSection(worksheet, currentRow, queries.agrupados, params);

            // 4. SECCIÓN MINUS
            currentRow = await this.addMinusSection(worksheet, currentRow, queries.minus1, queries.minus2, params);

            // Generar archivo y descargar
            const tableName = params.tablaDDV || 'TABLA';
            const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
            const filename = `cuadre_${tableName.toUpperCase()}_${periods}.xlsx`;

            const buffer = await workbook.xlsx.writeBuffer();
            this.downloadExcelBuffer(buffer, filename);

            if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
                UIModule.showNotification(`Excel de cuadre generado: ${filename}`, 'success', 5000);
            }

        } catch (error) {
            console.error('Error generando Excel:', error);
            alert('Error al generar Excel: ' + error.message);
        }
    },

    /**
     * Carga ExcelJS dinámicamente
     */
    async loadExcelJS() {
        if (window.ExcelJS) {
            return window.ExcelJS;
        }

        // Cargar ExcelJS desde CDN
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js';
        document.head.appendChild(script);

        return new Promise((resolve, reject) => {
            script.onload = () => resolve(window.ExcelJS);
            script.onerror = reject;
        });
    },

    /**
     * Agrega título principal
     */
    addMainTitle(worksheet, currentRow) {
        // Combinar celdas A1:K1
        worksheet.mergeCells(`A${currentRow}:K${currentRow}`);
        
        const titleCell = worksheet.getCell(`A${currentRow}`);
        titleCell.value = 'Generador de Queries de Ratificación v2';
        
        // Aplicar estilo al título
        titleCell.style = {
            font: { 
                size: 18, 
                bold: true, 
                color: { argb: 'FFFFFFFF' } 
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF6B46C1' } // Púrpura
            },
            alignment: { 
                horizontal: 'center', 
                vertical: 'middle' 
            }
        };
        
        // Altura de fila
        worksheet.getRow(currentRow).height = 40;
        
        return currentRow + 2; // Saltar una fila
    },

    /**
     * Agrega sección UNIVERSOS
     */
    async addUniversosSection(worksheet, currentRow, queryUniversos, params) {
        // H2 - Título de sección
        currentRow = this.addSectionTitle(worksheet, currentRow, 'UNIVERSOS');
        
        // Subtítulo "Código"
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código');
        
        // Bloque de código SQL
        currentRow = this.addCodeBlock(worksheet, currentRow, queryUniversos);
        
        // Subtítulo "Resultado"
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        
        // Tabla de resultado (ejemplo)
        currentRow = this.addUniversosResultTable(worksheet, currentRow, params);
        
        return currentRow + 2; // Espacio entre secciones
    },

    /**
     * Agrega sección AGRUPADOS
     */
    async addAgrupadosSection(worksheet, currentRow, queryAgrupados, params) {
        // H2 - Título de sección
        currentRow = this.addSectionTitle(worksheet, currentRow, 'AGRUPADOS');
        
        // Subtítulo "Código"
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código');
        
        // Bloque de código SQL
        currentRow = this.addCodeBlock(worksheet, currentRow, queryAgrupados);
        
        // Subtítulo "Resultado"
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        
        // Tabla de resultado (ejemplo)
        currentRow = this.addAgrupadosResultTable(worksheet, currentRow, params);
        
        return currentRow + 2;
    },

    /**
     * Agrega sección MINUS
     */
    async addMinusSection(worksheet, currentRow, queryMinus1, queryMinus2, params) {
        // H2 - Título de sección
        currentRow = this.addSectionTitle(worksheet, currentRow, 'MINUS');
        
        // MINUS 1
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código MINUS 1 (EDV - DDV)');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryMinus1);
        
        // MINUS 2
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código MINUS 2 (DDV - EDV)');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryMinus2);
        
        // Subtítulo "Resultado"
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        
        // Tabla de resultado (ejemplo)
        currentRow = this.addMinusResultTable(worksheet, currentRow, params);
        
        return currentRow + 2;
    },

    /**
     * Agrega título de sección (H2)
     */
    addSectionTitle(worksheet, currentRow, title) {
        const titleCell = worksheet.getCell(`A${currentRow}`);
        titleCell.value = title;
        
        // Aplicar estilo H2
        titleCell.style = {
            font: { 
                size: 14, 
                bold: true, 
                color: { argb: 'FFEBCB8B' } // Amarillo suave
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF2E3440' } // Dark
            },
            alignment: { 
                horizontal: 'left', 
                vertical: 'middle' 
            }
        };
        
        worksheet.getRow(currentRow).height = 30;
        
        return currentRow + 1;
    },

    /**
     * Agrega subtítulo
     */
    addSubtitle(worksheet, currentRow, subtitle) {
        const subtitleCell = worksheet.getCell(`A${currentRow}`);
        subtitleCell.value = subtitle;
        
        subtitleCell.style = {
            font: { 
                size: 12, 
                bold: true, 
                color: { argb: 'FF2E3440' } 
            },
            alignment: { 
                horizontal: 'left', 
                vertical: 'middle' 
            }
        };
        
        return currentRow + 1;
    },

    /**
     * Agrega bloque de código SQL con manejo de límite de 32,767 caracteres
     */
    addCodeBlock(worksheet, currentRow, sqlCode) {
        if (!sqlCode) {
            // Código no disponible
            worksheet.mergeCells(`B${currentRow}:K${currentRow}`);
            const cell = worksheet.getCell(`B${currentRow}`);
            cell.value = '-- Query no disponible';
            this.applyCodeStyle(cell);
            return currentRow + 1;
        }

        // Dividir el SQL en trozos seguros (límite 32,760 caracteres)
        const chunks = this.splitSQLIntoChunks(sqlCode, 32760);
        
        chunks.forEach((chunk, index) => {
            // Combinar celdas B:K para el chunk
            worksheet.mergeCells(`B${currentRow}:K${currentRow}`);
            const cell = worksheet.getCell(`B${currentRow}`);
            
            // Agregar saltos de línea cada 120 caracteres para mejor legibilidad
            const formattedChunk = this.addLineBreaks(chunk, 120);
            cell.value = formattedChunk;
            
            // Aplicar estilo de código
            this.applyCodeStyle(cell);
            
            // Etiqueta para chunks adicionales
            if (index > 0) {
                const labelCell = worksheet.getCell(`A${currentRow}`);
                labelCell.value = 'Código (cont.)';
                labelCell.style = {
                    font: { size: 10, italic: true, color: { argb: 'FF6C7B7F' } },
                    alignment: { horizontal: 'left', vertical: 'top' }
                };
            }
            
            // Altura de fila aumentada para acomodar texto wrapped
            const lineCount = formattedChunk.split('\n').length;
            worksheet.getRow(currentRow).height = Math.max(60, lineCount * 15);
            
            currentRow++;
        });
        
        return currentRow;
    },

    /**
     * Aplica estilo de código a una celda
     */
    applyCodeStyle(cell) {
        cell.style = {
            font: { 
                name: 'Consolas',
                size: 10, 
                color: { argb: 'FFECEFF4' } 
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF3B4252' } 
            },
            alignment: { 
                horizontal: 'left', 
                vertical: 'top',
                wrapText: true 
            },
            border: {
                top: { style: 'thin', color: { argb: 'FF4C566A' } },
                left: { style: 'thin', color: { argb: 'FF4C566A' } },
                bottom: { style: 'thin', color: { argb: 'FF4C566A' } },
                right: { style: 'thin', color: { argb: 'FF4C566A' } }
            }
        };
    },

    /**
     * Divide SQL en chunks seguros para Excel
     */
    splitSQLIntoChunks(sql, maxChars) {
        if (sql.length <= maxChars) {
            return [sql];
        }
        
        const chunks = [];
        let currentPos = 0;
        
        while (currentPos < sql.length) {
            let chunkEnd = currentPos + maxChars;
            
            // Si no es el último chunk, buscar un salto de línea cercano
            if (chunkEnd < sql.length) {
                const nearlineBreak = sql.lastIndexOf('\n', chunkEnd);
                if (nearlineBreak > currentPos + maxChars * 0.8) {
                    chunkEnd = nearlineBreak + 1;
                }
            }
            
            chunks.push(sql.substring(currentPos, chunkEnd));
            currentPos = chunkEnd;
        }
        
        return chunks;
    },

    /**
     * Agrega saltos de línea cada N caracteres
     */
    addLineBreaks(text, lineLength) {
        const lines = text.split('\n');
        const result = [];
        
        lines.forEach(line => {
            if (line.length <= lineLength) {
                result.push(line);
            } else {
                // Dividir líneas largas en múltiples líneas
                let pos = 0;
                while (pos < line.length) {
                    result.push(line.substring(pos, pos + lineLength));
                    pos += lineLength;
                }
            }
        });
        
        return result.join('\n');
    },

    /**
     * Agrega tabla de resultado para UNIVERSOS
     */
    addUniversosResultTable(worksheet, currentRow, params) {
        // Encabezados
        const headers = ['codmes', 'numreg_ddv', 'numreg_edv', 'diff_numreg', 'status'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2); // Empezar en columna B
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
        // Datos de ejemplo
        const exampleData = [
            [202505, 2765145, 2765145, 0, '✅ IGUALES'],
            [202506, 2758763, 2758763, 0, '✅ IGUALES'],
            [202507, 2787328, 2787328, 0, '✅ IGUALES']
        ];
        
        exampleData.forEach(row => {
            row.forEach((value, index) => {
                const cell = worksheet.getCell(currentRow, index + 2);
                cell.value = value;
                this.applyDataStyle(cell, index === 3); // Destacar columna diff
            });
            currentRow++;
        });
        
        return currentRow;
    },

    /**
     * Agrega tabla de resultado para AGRUPADOS
     */
    addAgrupadosResultTable(worksheet, currentRow, params) {
        // Encabezados
        const headers = ['capa', 'codmes', 'count_campos', 'sum_campos'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2);
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
        // Datos de ejemplo
        const exampleData = [
            ['EDV', 202505, 150, 1250000.50],
            ['DDV', 202505, 150, 1250000.50],
            ['EDV', 202506, 148, 1180000.25],
            ['DDV', 202506, 148, 1180000.25]
        ];
        
        exampleData.forEach(row => {
            row.forEach((value, index) => {
                const cell = worksheet.getCell(currentRow, index + 2);
                cell.value = value;
                this.applyDataStyle(cell);
            });
            currentRow++;
        });
        
        return currentRow;
    },

    /**
     * Agrega tabla de resultado para MINUS
     */
    addMinusResultTable(worksheet, currentRow, params) {
        // Encabezados
        const headers = ['query', 'registros_encontrados', 'status'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2);
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
        // Datos de ejemplo
        const exampleData = [
            ['MINUS 1 (EDV - DDV)', 0, '✅ Sin diferencias'],
            ['MINUS 2 (DDV - EDV)', 0, '✅ Sin diferencias']
        ];
        
        exampleData.forEach(row => {
            row.forEach((value, index) => {
                const cell = worksheet.getCell(currentRow, index + 2);
                cell.value = value;
                this.applyDataStyle(cell);
            });
            currentRow++;
        });
        
        return currentRow;
    },

    /**
     * Aplica estilo a encabezados de tabla
     */
    applyHeaderStyle(cell) {
        cell.style = {
            font: { 
                size: 11, 
                bold: true, 
                color: { argb: 'FF2E3440' } 
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD8DEE9' } // Gris claro
            },
            alignment: { 
                horizontal: 'center', 
                vertical: 'middle' 
            },
            border: {
                top: { style: 'thin', color: { argb: 'FF4C566A' } },
                left: { style: 'thin', color: { argb: 'FF4C566A' } },
                bottom: { style: 'thin', color: { argb: 'FF4C566A' } },
                right: { style: 'thin', color: { argb: 'FF4C566A' } }
            }
        };
    },

    /**
     * Aplica estilo a datos de tabla
     */
    applyDataStyle(cell, isDiff = false) {
        cell.style = {
            font: { 
                size: 10, 
                color: { argb: 'FF2E3440' },
                bold: isDiff
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: isDiff ? 'FFFFF8DB' : 'FFFFFFFF' } // Amarillo suave para diff
            },
            alignment: { 
                horizontal: typeof cell.value === 'number' ? 'right' : 'left', 
                vertical: 'middle' 
            },
            border: {
                top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
            },
            numFmt: typeof cell.value === 'number' && cell.value > 1000 ? '#,##0' : undefined
        };
    },

    /**
     * Descarga buffer de Excel
     */
    downloadExcelBuffer(buffer, filename) {
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
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
            ['📊 RESUMEN DE CUADRE DDV vs EDV', '', ''],
            ['', '', ''],
            ['📊 INFORMACIÓN GENERAL', '', ''],
            ['Tabla DDV', `${params.esquemaDDV}.${params.tablaDDV}`, ''],
            ['Tabla EDV', `${params.esquemaEDV}.${params.tablaEDV}`, ''],
            ['Períodos', params.periodos, ''],
            ['Total Campos', tableStructure.length, ''],
            ['Campos COUNT', tableStructure.filter(f => f.aggregateFunction === 'count').length, ''],
            ['Campos SUM', tableStructure.filter(f => f.aggregateFunction === 'sum').length, ''],
            ['', '', ''],
            ['🔍 QUERIES GENERADOS', '', ''],
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
     * Agrega hoja de queries (con manejo de límite de caracteres)
     * @param {Object} wb - Workbook
     */
    addQueriesSheet(wb) {
        const queries = QueryModule.getGeneratedQueries();
        
        const queryData = [
            ['TIPO_QUERY', 'DESCRIPCION', 'QUERY_SQL', 'LINEAS', 'CARACTERES'],
            [
                'UNIVERSOS',
                'Compara número total de registros entre DDV y EDV',
                this.truncateForExcel(queries.universos || ''),
                queries.universos ? queries.universos.split('\n').length : 0,
                queries.universos ? queries.universos.length : 0
            ],
            [
                'AGRUPADOS',
                'Compara métricas agregadas por cada campo',
                this.truncateForExcel(queries.agrupados || ''),
                queries.agrupados ? queries.agrupados.split('\n').length : 0,
                queries.agrupados ? queries.agrupados.length : 0
            ],
            [
                'MINUS_EDV_DDV',
                'Registros que están en EDV pero NO en DDV',
                this.truncateForExcel(queries.minus1 || ''),
                queries.minus1 ? queries.minus1.split('\n').length : 0,
                queries.minus1 ? queries.minus1.length : 0
            ],
            [
                'MINUS_DDV_EDV',
                'Registros que están en DDV pero NO en EDV',
                this.truncateForExcel(queries.minus2 || ''),
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
        
        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification(
                `Reporte completo descargado: ${filename}`,
                'success',
                4000
            );
        }
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
    },

    /**
     * Exporta Excel específico para V3 - Formato Cuadre EDV (DISEÑO MEJORADO)
     */
    exportCuadreEDV() {
        const queries = QueryModule.getGeneratedQueries();
        const params = ParametersModule.getCurrentParameters();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar. Primero genera los queries en la pestaña correspondiente.');
            return;
        }
        
        // Generar nombre de archivo
        const tableName = params.tablaDDV || 'TABLA';
        const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
        const filename = `cuadre_${tableName.toUpperCase()}_${periods}.xlsx`;
        
        const wb = XLSX.utils.book_new();
        
        // Crear las 3 pestañas con diseño mejorado
        this.createUniversoSheetDesign(wb, queries.universos, params);
        this.createAgrupadoSheetDesign(wb, queries.agrupados, params);
        this.createMinusSheetDesign(wb, queries.minus1, queries.minus2, params);
        
        // Descargar
        XLSX.writeFile(wb, filename);
        
        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification(`Excel de cuadre generado: ${filename}`, 'success', 5000);
        }
    },

    /**
     * Crea hoja Universo con DISEÑO PROFESIONAL
     * @param {Object} wb - Workbook
     * @param {string} queryUniverso - Query de universos
     * @param {Object} params - Parámetros
     */
    createUniversoSheetDesign(wb, queryUniverso, params) {
        const data = [
            // ENCABEZADO PRINCIPAL
            ['CUADRE DDV vs EDV - ANÁLISIS DE UNIVERSOS', '', '', '', ''],
            ['', '', '', '', ''],
            
            // INFORMACIÓN DEL PROYECTO
            ['INFORMACIÓN DEL PROYECTO', '', '', '', ''],
            ['Tabla DDV:', `${params.esquemaDDV}.${params.tablaDDV}`, '', '', ''],
            ['Tabla EDV:', `${params.esquemaEDV}.${params.tablaEDV}`, '', '', ''],
            ['Períodos:', params.periodos, '', '', ''],
            ['Fecha:', new Date().toLocaleDateString('es-ES'), '', '', ''],
            ['', '', '', '', ''],
            
            // DESCRIPCIÓN
            ['DESCRIPCIÓN', '', '', '', ''],
            ['Este query compara el número total de registros entre las tablas DDV y EDV', '', '', '', ''],
            ['para verificar que ambas tengan la misma cantidad de datos por período.', '', '', '', ''],
            ['', '', '', '', ''],
            
            // QUERY SQL CON FORMATO
            ['QUERY SQL', '', '', '', ''],
            ['Código:', '', '', '', ''],
            ...this.formatQueryForDisplay(queryUniverso),
            ['', '', '', '', ''],
            
            // INTERPRETACIÓN DE RESULTADOS
            ['INTERPRETACIÓN DE RESULTADOS', '', '', '', ''],
            ['• diff_numreg = 0: Las tablas tienen la misma cantidad de registros ✓', '', '', '', ''],
            ['• diff_numreg > 0: Faltan registros en EDV (Revisar proceso) ⚠️', '', '', '', ''],
            ['• diff_numreg < 0: Sobran registros en EDV (Revisar duplicados) ⚠️', '', '', '', ''],
            ['', '', '', '', ''],
            
            // ACCIONES RECOMENDADAS
            ['ACCIONES SEGÚN RESULTADO', '', '', '', ''],
            ['1. Si diff_numreg = 0: Continuar con siguiente validación', '', '', '', ''],
            ['2. Si diff_numreg ≠ 0: Investigar causa de la diferencia', '', '', '', ''],
            ['3. Validar filtros de período aplicados', '', '', '', ''],
            ['4. Revisar proceso de carga de datos', '', '', '', '']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // APLICAR DISEÑO PROFESIONAL
        this.applyUniversoStyling(ws);
        
        XLSX.utils.book_append_sheet(wb, ws, 'Universo');
    },

    /**
     * Crea hoja Agrupado con DISEÑO PROFESIONAL
     * @param {Object} wb - Workbook
     * @param {string} queryAgrupado - Query agrupado
     * @param {Object} params - Parámetros
     */
    createAgrupadoSheetDesign(wb, queryAgrupado, params) {
        const data = [
            // ENCABEZADO PRINCIPAL
            ['CUADRE DDV vs EDV - ANÁLISIS AGRUPADO POR CAMPOS', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // INFORMACIÓN DEL PROYECTO
            ['INFORMACIÓN DEL PROYECTO', '', '', '', '', ''],
            ['Tabla DDV:', `${params.esquemaDDV}.${params.tablaDDV}`, '', '', '', ''],
            ['Tabla EDV:', `${params.esquemaEDV}.${params.tablaEDV}`, '', '', '', ''],
            ['Períodos:', params.periodos, '', '', '', ''],
            ['Fecha:', new Date().toLocaleDateString('es-ES'), '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // DESCRIPCIÓN
            ['DESCRIPCIÓN', '', '', '', '', ''],
            ['Este query compara las métricas agregadas (COUNT/SUM) campo por campo', '', '', '', '', ''],
            ['entre las tablas DDV y EDV para identificar diferencias específicas.', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // ESTRUCTURA DEL RESULTADO
            ['ESTRUCTURA DEL RESULTADO', '', '', '', '', ''],
            ['Columna', 'Descripción', 'Valores Esperados', '', '', ''],
            ['capa', 'Identifica la fuente (DDV/EDV)', 'DDV, EDV', '', '', ''],
            ['codmes', 'Período analizado', params.periodos, '', '', ''],
            ['campos_count', 'Conteos por campo', 'Números enteros', '', '', ''],
            ['campos_sum', 'Sumas por campo', 'Números decimales', '', '', ''],
            ['', '', '', '', '', ''],
            
            // QUERY SQL
            ['QUERY SQL', '', '', '', '', ''],
            ['Código:', '', '', '', '', ''],
            ...this.formatQueryForDisplay(queryAgrupado),
            ['', '', '', '', '', ''],
            
            // ANÁLISIS RECOMENDADO
            ['ANÁLISIS RECOMENDADO', '', '', '', '', ''],
            ['1. Ordenar por codmes y capa para comparación lado a lado', '', '', '', '', ''],
            ['2. Verificar que cada campo tenga valores idénticos entre DDV y EDV', '', '', '', '', ''],
            ['3. Identificar campos con diferencias para investigación detallada', '', '', '', '', ''],
            ['4. Documentar cualquier discrepancia encontrada', '', '', '', '', '']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // APLICAR DISEÑO PROFESIONAL
        this.applyAgrupadoStyling(ws);
        
        XLSX.utils.book_append_sheet(wb, ws, 'Agrupado');
    },

    /**
     * Crea hoja Minus con DISEÑO PROFESIONAL
     * @param {Object} wb - Workbook
     * @param {string} queryMinus1 - Query MINUS 1
     * @param {string} queryMinus2 - Query MINUS 2
     * @param {Object} params - Parámetros
     */
    createMinusSheetDesign(wb, queryMinus1, queryMinus2, params) {
        const data = [
            // ENCABEZADO PRINCIPAL
            ['CUADRE DDV vs EDV - ANÁLISIS DE DIFERENCIAS (MINUS)', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // INFORMACIÓN DEL PROYECTO
            ['INFORMACIÓN DEL PROYECTO', '', '', '', '', ''],
            ['Tabla DDV:', `${params.esquemaDDV}.${params.tablaDDV}`, '', '', '', ''],
            ['Tabla EDV:', `${params.esquemaEDV}.${params.tablaEDV}`, '', '', '', ''],
            ['Períodos:', params.periodos, '', '', '', ''],
            ['Fecha:', new Date().toLocaleDateString('es-ES'), '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // DESCRIPCIÓN
            ['DESCRIPCIÓN', '', '', '', '', ''],
            ['Los queries MINUS identifican registros que están en una tabla pero no en la otra.', '', '', '', '', ''],
            ['Ayudan a detectar datos faltantes o excedentes entre DDV y EDV.', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // QUERY MINUS 1
            ['QUERY MINUS 1: EDV - DDV', '', '', '', '', ''],
            ['Encuentra registros que están en EDV pero NO en DDV', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['Código:', '', '', '', '', ''],
            ...this.formatQueryForDisplay(queryMinus1, 'MINUS1'),
            ['', '', '', '', '', ''],
            
            // QUERY MINUS 2
            ['QUERY MINUS 2: DDV - EDV', '', '', '', '', ''],
            ['Encuentra registros que están en DDV pero NO en EDV', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['Código:', '', '', '', '', ''],
            ...this.formatQueryForDisplay(queryMinus2, 'MINUS2'),
            ['', '', '', '', '', ''],
            
            // INTERPRETACIÓN
            ['INTERPRETACIÓN DE RESULTADOS', '', '', '', '', ''],
            ['• Si ambos queries devuelven 0 registros: Las tablas son idénticas ✓', '', '', '', '', ''],
            ['• Si MINUS 1 devuelve registros: Hay datos en EDV que faltan en DDV ⚠️', '', '', '', '', ''],
            ['• Si MINUS 2 devuelve registros: Hay datos en DDV que faltan en EDV ⚠️', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            
            // ACCIONES CORRECTIVAS
            ['ACCIONES CORRECTIVAS', '', '', '', '', ''],
            ['1. Ejecutar ambos queries por separado', '', '', '', '', ''],
            ['2. Analizar los registros devueltos en detalle', '', '', '', '', ''],
            ['3. Verificar procesos de ETL y transformación de datos', '', '', '', '', ''],
            ['4. Coordinar con el equipo técnico para resolver diferencias', '', '', '', '', '']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // APLICAR DISEÑO PROFESIONAL
        this.applyMinusStyling(ws);
        
        XLSX.utils.book_append_sheet(wb, ws, 'Minus');
    },

    /**
     * Formatea query para mejor visualización en Excel
     * @param {string} query - Query SQL
     * @param {string} prefix - Prefijo para identificar secciones
     * @returns {Array} - Líneas formateadas para Excel
     */
    formatQueryForDisplay(query, prefix = '') {
        if (!query) return [['-- Query no disponible', '', '', '', '']];
        
        const lines = query.split('\n');
        const formattedLines = [];
        
        lines.forEach((line, index) => {
            // Limitar longitud de línea para Excel
            const cleanLine = line.trim();
            if (cleanLine.length > 0) {
                // Dividir líneas muy largas
                if (cleanLine.length > 120) {
                    const chunks = cleanLine.match(/.{1,120}/g) || [cleanLine];
                    chunks.forEach((chunk, chunkIndex) => {
                        formattedLines.push([
                            chunkIndex === 0 ? cleanLine.substring(0, 20) + '...' : '...',
                            chunk,
                            '', '', ''
                        ]);
                    });
                } else {
                    formattedLines.push([cleanLine, '', '', '', '']);
                }
            }
        });
        
        return formattedLines;
    },

    /**
     * Aplica estilos profesionales a la hoja Universo
     * @param {Object} ws - Worksheet
     */
    applyUniversoStyling(ws) {
        // Configurar anchos de columna
        ws['!cols'] = [
            { width: 25 }, // Columna A - Títulos
            { width: 40 }, // Columna B - Contenido
            { width: 20 }, // Columna C - Extra
            { width: 15 }, // Columna D - Extra
            { width: 15 }  // Columna E - Extra
        ];
        
        // Configurar combinación de celdas para títulos
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Título principal
            { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } }, // Información del proyecto
            { s: { r: 8, c: 0 }, e: { r: 8, c: 4 } }, // Descripción
            { s: { r: 14, c: 0 }, e: { r: 14, c: 4 } }, // Query SQL
            { s: { r: 20, c: 0 }, e: { r: 20, c: 4 } }, // Interpretación
            { s: { r: 26, c: 0 }, e: { r: 26, c: 4 } }  // Acciones
        ];
    },

    /**
     * Aplica estilos profesionales a la hoja Agrupado
     * @param {Object} ws - Worksheet
     */
    applyAgrupadoStyling(ws) {
        // Configurar anchos de columna
        ws['!cols'] = [
            { width: 25 }, // Columna A - Títulos
            { width: 40 }, // Columna B - Contenido
            { width: 25 }, // Columna C - Valores
            { width: 15 }, // Columna D - Extra
            { width: 15 }, // Columna E - Extra
            { width: 15 }  // Columna F - Extra
        ];
        
        // Configurar combinación de celdas para títulos
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }, // Título principal
            { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } }, // Información del proyecto
            { s: { r: 8, c: 0 }, e: { r: 8, c: 5 } }, // Descripción
            { s: { r: 12, c: 0 }, e: { r: 12, c: 5 } }, // Estructura
            { s: { r: 20, c: 0 }, e: { r: 20, c: 5 } }, // Query SQL
            { s: { r: 25, c: 0 }, e: { r: 25, c: 5 } }  // Análisis
        ];
    },

    /**
     * Aplica estilos profesionales a la hoja Minus
     * @param {Object} ws - Worksheet
     */
    applyMinusStyling(ws) {
        // Configurar anchos de columna
        ws['!cols'] = [
            { width: 30 }, // Columna A - Títulos
            { width: 50 }, // Columna B - Contenido
            { width: 20 }, // Columna C - Extra
            { width: 15 }, // Columna D - Extra
            { width: 15 }, // Columna E - Extra
            { width: 15 }  // Columna F - Extra
        ];
        
        // Configurar combinación de celdas para títulos
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }, // Título principal
            { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } }, // Información del proyecto
            { s: { r: 8, c: 0 }, e: { r: 8, c: 5 } }, // Descripción
            { s: { r: 12, c: 0 }, e: { r: 12, c: 5 } }, // MINUS 1
            { s: { r: 20, c: 0 }, e: { r: 20, c: 5 } }, // MINUS 2
            { s: { r: 28, c: 0 }, e: { r: 28, c: 5 } }, // Interpretación
            { s: { r: 34, c: 0 }, e: { r: 34, c: 5 } }  // Acciones
        ];
    },

    /**
     * Divide query complejo inteligentemente
     * @param {string} query - Query a dividir
     * @returns {Array<string>} - Partes del query
     */
    splitComplexQuery(query) {
        if (!query || query.length <= 30000) {
            return [query];
        }
        
        const parts = [];
        const lines = query.split('\n');
        let currentPart = '';
        let fieldCount = 0;
        
        for (const line of lines) {
            // Si es una línea con campos count() o sum()
            if (line.includes('count(') || line.includes('sum(')) {
                fieldCount++;
                
                // Cada 10 campos, hacer nueva parte
                if (fieldCount % 10 === 0) {
                    parts.push(currentPart + line);
                    currentPart = '';
                    continue;
                }
            }
            
            // Si agregar esta línea excede 30k, partir
            if ((currentPart + '\n' + line).length > 30000 && currentPart.length > 0) {
                parts.push(currentPart);
                currentPart = line;
            } else {
                currentPart += (currentPart ? '\n' : '') + line;
            }
        }
        
        if (currentPart) {
            parts.push(currentPart);
        }
        
        return parts;
    },

    /**
     * Divide query en partes lógicas
     * @param {string} query - Query a dividir
     * @param {string} keyword - Palabra clave de referencia
     * @returns {Array<string>} - Partes del query
     */
    splitQueryIntoLogicalParts(query, keyword) {
        if (!query) return [''];
        
        const lines = query.split('\n');
        const parts = [];
        let currentPart = '';
        
        for (const line of lines) {
            if ((currentPart + '\n' + line).length > 30000 && currentPart.length > 0) {
                parts.push(currentPart.trim());
                currentPart = line;
            } else {
                currentPart += (currentPart ? '\n' : '') + line;
            }
        }
        
        if (currentPart.trim()) {
            parts.push(currentPart.trim());
        }
        
        return parts.length > 0 ? parts : [query];
    }
};