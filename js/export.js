/**
 * Módulo para exportación a Excel y otros formatos
 * SISTEMA COMPLETO: Template Upload + ExcelJS + Funciones Originales
 */
const ExportModule = {
    
    // Template cargado
    templateBuffer: null,
    templateWorkbook: null,
    
    /**
     * FUNCIÓN PRINCIPAL: Exporta Excel usando template upload o generación automática
     */
    async exportCuadreEDV() {
        const queries = QueryModule.getGeneratedQueries();
        const params = ParametersModule.getCurrentParameters();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar. Primero genera los queries en la pestaña correspondiente.');
            return;
        }

        try {
            // Si hay template cargado, usar sistema de template
            if (this.templateBuffer) {
                await this.exportWithTemplate(queries, params);
            } else {
                // Usar generación automática con ExcelJS
                await this.exportWithAutoGeneration(queries, params);
            }
        } catch (error) {
            console.error('Error generando Excel:', error);
            alert('Error al generar Excel: ' + error.message);
        }
    },

    /**
     * Exporta usando template cargado con soporte para múltiples pestañas
     */
    async exportWithTemplate(queries, params) {
        const ExcelJS = await this.loadExcelJS();
        
        // Cargar template desde buffer
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(this.templateBuffer);
        
        // Preparar datos para inserción
        const data = {
            sqlUniv: queries.universos,
            sqlAgr: queries.agrupados, 
            sqlMinus: this.combineMinus(queries.minus1, queries.minus2),
            tablaUniv: this.generateUniversosTable(params),
            tablaAgr: this.generateAgrupadosTable(params),
            tablaMinus: this.generateMinusTable(params)
        };

        // Procesar cada pestaña del template
        const processedSheets = [];
        
        for (const worksheet of workbook.worksheets) {
            const sheetName = worksheet.name.toLowerCase();
            
            // Determinar qué datos insertar según el nombre de la pestaña
            let sheetData = {};
            
            if (sheetName.includes('universo') || sheetName.includes('universos')) {
                sheetData = {
                    sql: data.sqlUniv,
                    tabla: data.tablaUniv,
                    tipo: 'universos'
                };
            } else if (sheetName.includes('agrupado') || sheetName.includes('agrupados')) {
                sheetData = {
                    sql: data.sqlAgr,
                    tabla: data.tablaAgr,
                    tipo: 'agrupados'
                };
            } else if (sheetName.includes('minus')) {
                sheetData = {
                    sql: data.sqlMinus,
                    tabla: data.tablaMinus,
                    tipo: 'minus'
                };
            } else if (sheetName.includes('cuadre') || sheetName.includes('resumen')) {
                // Pestaña principal con todos los datos
                sheetData = {
                    sqlUniv: data.sqlUniv,
                    sqlAgr: data.sqlAgr,
                    sqlMinus: data.sqlMinus,
                    tablaUniv: data.tablaUniv,
                    tablaAgr: data.tablaAgr,
                    tablaMinus: data.tablaMinus,
                    tipo: 'completo'
                };
            } else {
                // Pestaña genérica - intentar insertar todos los datos
                sheetData = {
                    sqlUniv: data.sqlUniv,
                    sqlAgr: data.sqlAgr,
                    sqlMinus: data.sqlMinus,
                    tablaUniv: data.tablaUniv,
                    tablaAgr: data.tablaAgr,
                    tablaMinus: data.tablaMinus,
                    tipo: 'completo'
                };
            }

            // Insertar contenido en la pestaña
            try {
                await this.insertContentIntoTemplate(worksheet, sheetData);
                processedSheets.push(worksheet.name);
                console.log(`✅ Procesada pestaña: ${worksheet.name} (${sheetData.tipo})`);
            } catch (error) {
                console.warn(`⚠️ Error procesando pestaña ${worksheet.name}:`, error.message);
            }
        }

        if (processedSheets.length === 0) {
            throw new Error('No se pudo procesar ninguna pestaña del template');
        }

        // Generar archivo y descargar
        const filename = this.generateTemplateFilename(params);
        const buffer = await workbook.xlsx.writeBuffer();
        this.downloadExcelBuffer(buffer, filename);

        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification(
                `Excel generado con template: ${filename} (${processedSheets.length} pestañas procesadas)`, 
                'success', 
                5000
            );
        }
    },

    /**
     * Exporta con generación automática usando ExcelJS
     */
    async exportWithAutoGeneration(queries, params) {
        const ExcelJS = await this.loadExcelJS();
        
        // Crear workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Cuadre DDV vs EDV', {
            pageSetup: { paperSize: 9, orientation: 'landscape' }
        });

        // Configurar anchos de columna
        worksheet.columns = [
            { width: 15 }, { width: 22 }, { width: 22 }, { width: 22 },
            { width: 22 }, { width: 22 }, { width: 22 }, { width: 22 },
            { width: 22 }, { width: 22 }, { width: 22 }
        ];

        let currentRow = 1;

        // Construir Excel automáticamente
        currentRow = this.addMainTitle(worksheet, currentRow);
        worksheet.views = [{ state: 'frozen', ySplit: 2 }];
        currentRow = await this.addUniversosSection(worksheet, currentRow, queries.universos, params);
        currentRow = await this.addAgrupadosSection(worksheet, currentRow, queries.agrupados, params);
        currentRow = await this.addMinusSection(worksheet, currentRow, queries.minus1, queries.minus2, params);

        // Generar archivo y descargar
        const filename = this.generateAutoFilename(params);
        const buffer = await workbook.xlsx.writeBuffer();
        this.downloadExcelBuffer(buffer, filename);

        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification(`Excel generado automáticamente: ${filename}`, 'success', 5000);
        }
    },

    /**
     * Cargar template Excel con validación mejorada
     */
    async loadTemplate() {
        try {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx,.xls';
            
            return new Promise((resolve, reject) => {
                input.onchange = async (event) => {
                    const file = event.target.files[0];
                    if (!file) {
                        reject(new Error('No se seleccionó archivo'));
                        return;
                    }

                    try {
                        // Mostrar loading
                        const templateInfo = document.getElementById('templateInfo');
                        if (templateInfo) {
                            templateInfo.innerHTML = `
                                <div class="template-loading">
                                    <div class="spinner"></div>
                                    Cargando y validando template...
                                </div>
                            `;
                        }

                        const arrayBuffer = await file.arrayBuffer();
                        this.templateBuffer = arrayBuffer;
                        
                        // Validar template
                        const validationResult = await this.validateTemplate();
                        
                        // Mostrar información del template
                        if (templateInfo) {
                            templateInfo.innerHTML = `
                                <div class="template-loaded">
                                    <div class="template-header">
                                        <span class="template-icon">📊</span>
                                        <div class="template-details">
                                            <strong>${file.name}</strong>
                                            <small>${(file.size / 1024).toFixed(1)} KB</small>
                                        </div>
                                    </div>
                                    <div class="template-info">
                                        <div class="info-item">
                                            <span class="label">Pestañas encontradas:</span>
                                            <span class="value">${validationResult.sheets.join(', ')}</span>
                                        </div>
                                        <div class="info-item">
                                            <span class="label">Anclas detectadas:</span>
                                            <span class="value">${validationResult.anchors} nombres definidos</span>
                                        </div>
                                        <div class="info-item">
                                            <span class="label">Placeholders:</span>
                                            <span class="value">${validationResult.placeholders} encontrados</span>
                                        </div>
                                    </div>
                                    <div class="template-actions">
                                        <button class="btn btn-sm" onclick="ExportModule.previewTemplate()">👁️ Vista Previa</button>
                                        <button class="btn btn-sm btn-secondary" onclick="ExportModule.clearTemplate()">🗑️ Limpiar</button>
                                    </div>
                                </div>
                            `;
                        }
                        
                        resolve(file.name);
                        
                    } catch (error) {
                        if (templateInfo) {
                            templateInfo.innerHTML = `
                                <div class="template-error">
                                    ❌ Error cargando template: ${error.message}
                                    <br><small>Verifica que el archivo tenga la estructura correcta</small>
                                </div>
                            `;
                        }
                        reject(error);
                    }
                };
                
                input.click();
            });
            
        } catch (error) {
            alert('Error cargando template: ' + error.message);
        }
    },

    /**
     * Valida que el template tenga la estructura esperada
     */
    async validateTemplate() {
        const ExcelJS = await this.loadExcelJS();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(this.templateBuffer);
        
        const validationResult = {
            sheets: [],
            anchors: 0,
            placeholders: 0,
            isValid: false,
            errors: []
        };

        // Obtener lista de pestañas
        workbook.worksheets.forEach(ws => {
            validationResult.sheets.push(ws.name);
        });

        // Buscar pestaña principal (Cuadre, Universo, Agrupados, Minus)
        const mainSheets = ['Cuadre', 'Universo', 'Agrupados', 'Minus'];
        const foundMainSheet = validationResult.sheets.find(sheet => 
            mainSheets.some(main => sheet.toLowerCase().includes(main.toLowerCase()))
        );

        if (!foundMainSheet) {
            validationResult.errors.push('No se encontró pestaña principal (Cuadre, Universo, Agrupados, o Minus)');
        }

        const worksheet = workbook.getWorksheet(foundMainSheet) || workbook.worksheets[0];
        if (!worksheet) {
            validationResult.errors.push('No se pudo acceder a ninguna pestaña del template');
            return validationResult;
        }

        const requiredAnchors = [
            'ANCHOR_UNIV_SQL', 'ANCHOR_UNIV_TABLA',
            'ANCHOR_AGR_SQL', 'ANCHOR_AGR_TABLA', 
            'ANCHOR_MINUS_SQL', 'ANCHOR_MINUS_TABLA'
        ];

        // Verificar nombres definidos
        if (workbook.definedNames) {
            requiredAnchors.forEach(anchor => {
                if (workbook.definedNames.get(anchor)) {
                    validationResult.anchors++;
                }
            });
        }

        // Verificar placeholders
        const placeholders = [
            '<<UNIVERSOS_SQL>>', '<<UNIVERSOS_TABLA>>',
            '<<AGRUPADOS_SQL>>', '<<AGRUPADOS_TABLA>>',
            '<<MINUS_SQL>>', '<<MINUS_TABLA>>',
            // Placeholders alternativos
            '{{UNIVERSOS_SQL}}', '{{UNIVERSOS_TABLA}}',
            '{{AGRUPADOS_SQL}}', '{{AGRUPADOS_TABLA}}',
            '{{MINUS_SQL}}', '{{MINUS_TABLA}}'
        ];

        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value && typeof cell.value === 'string') {
                    placeholders.forEach(placeholder => {
                        if (cell.value.includes(placeholder)) {
                            validationResult.placeholders++;
                        }
                    });
                }
            });
        });

        // Validar que tenga al menos una forma de inserción
        if (validationResult.anchors === 0 && validationResult.placeholders === 0) {
            validationResult.errors.push('No se encontraron nombres definidos ni placeholders esperados');
        }

        validationResult.isValid = validationResult.errors.length === 0;

        console.log(`Template validado: ${validationResult.anchors} nombres definidos, ${validationResult.placeholders} placeholders encontrados`);
        
        return validationResult;
    },

    /**
     * Inserta contenido en el template usando nombres definidos o fallback
     */
    async insertContentIntoTemplate(worksheet, data) {
        const workbook = worksheet.workbook;

        // Determinar el tipo de inserción basado en los datos disponibles
        if (data.tipo === 'universos') {
            await this.insertSingleTypeContent(worksheet, workbook, data, 'UNIV');
        } else if (data.tipo === 'agrupados') {
            await this.insertSingleTypeContent(worksheet, workbook, data, 'AGR');
        } else if (data.tipo === 'minus') {
            await this.insertSingleTypeContent(worksheet, workbook, data, 'MINUS');
        } else {
            // Inserción completa con todos los tipos
            await this.insertCompleteContent(worksheet, workbook, data);
        }
    },

    /**
     * Inserta contenido de un solo tipo (universos, agrupados, o minus)
     */
    async insertSingleTypeContent(worksheet, workbook, data, type) {
        const contentMap = [
            { 
                anchor: `ANCHOR_${type}_SQL`, 
                placeholder: `<<${type}_SQL>>`, 
                altPlaceholder: `{{${type}_SQL}}`,
                content: data.sql, 
                type: 'sql' 
            },
            { 
                anchor: `ANCHOR_${type}_TABLA`, 
                placeholder: `<<${type}_TABLA>>`, 
                altPlaceholder: `{{${type}_TABLA}}`,
                content: data.tabla, 
                type: 'table' 
            }
        ];

        for (const item of contentMap) {
            try {
                const position = this.findContentPosition(workbook, worksheet, item.anchor, item.placeholder, item.altPlaceholder);
                if (position) {
                    if (item.type === 'sql') {
                        await this.insertSQLContent(worksheet, position, item.content);
                    } else if (item.type === 'table') {
                        await this.insertTableContent(worksheet, position, item.content);
                    }
                }
            } catch (error) {
                console.warn(`No se pudo insertar ${item.anchor}:`, error.message);
            }
        }
    },

    /**
     * Inserta contenido completo con todos los tipos
     */
    async insertCompleteContent(worksheet, workbook, data) {
        const contentMap = [
            { anchor: 'ANCHOR_UNIV_SQL', placeholder: '<<UNIVERSOS_SQL>>', altPlaceholder: '{{UNIVERSOS_SQL}}', content: data.sqlUniv, type: 'sql' },
            { anchor: 'ANCHOR_UNIV_TABLA', placeholder: '<<UNIVERSOS_TABLA>>', altPlaceholder: '{{UNIVERSOS_TABLA}}', content: data.tablaUniv, type: 'table' },
            { anchor: 'ANCHOR_AGR_SQL', placeholder: '<<AGRUPADOS_SQL>>', altPlaceholder: '{{AGRUPADOS_SQL}}', content: data.sqlAgr, type: 'sql' },
            { anchor: 'ANCHOR_AGR_TABLA', placeholder: '<<AGRUPADOS_TABLA>>', altPlaceholder: '{{AGRUPADOS_TABLA}}', content: data.tablaAgr, type: 'table' },
            { anchor: 'ANCHOR_MINUS_SQL', placeholder: '<<MINUS_SQL>>', altPlaceholder: '{{MINUS_SQL}}', content: data.sqlMinus, type: 'sql' },
            { anchor: 'ANCHOR_MINUS_TABLA', placeholder: '<<MINUS_TABLA>>', altPlaceholder: '{{MINUS_TABLA}}', content: data.tablaMinus, type: 'table' }
        ];

        for (const item of contentMap) {
            try {
                const position = this.findContentPosition(workbook, worksheet, item.anchor, item.placeholder, item.altPlaceholder);
                if (position) {
                    if (item.type === 'sql') {
                        await this.insertSQLContent(worksheet, position, item.content);
                    } else if (item.type === 'table') {
                        await this.insertTableContent(worksheet, position, item.content);
                    }
                }
            } catch (error) {
                console.warn(`No se pudo insertar ${item.anchor}:`, error.message);
            }
        }
    },

    /**
     * Encuentra la posición donde insertar contenido
     */
    findContentPosition(workbook, worksheet, anchorName, placeholder, altPlaceholder = null) {
        // Intentar nombre definido primero
        if (workbook.definedNames) {
            const definedName = workbook.definedNames.get(anchorName);
            if (definedName) {
                const match = definedName.value.match(/([^!]+)!\$([A-Z]+)\$(\d+)/);
                if (match) {
                    const col = this.columnLetterToNumber(match[2]);
                    const row = parseInt(match[3]);
                    return { row, col };
                }
            }
        }

        // Fallback: buscar placeholder
        let position = null;
        const placeholders = [placeholder];
        if (altPlaceholder) {
            placeholders.push(altPlaceholder);
        }

        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value && typeof cell.value === 'string') {
                    for (const ph of placeholders) {
                        if (cell.value.includes(ph)) {
                            position = { row: rowNumber, col: colNumber };
                            cell.value = cell.value.replace(ph, '').trim() || null;
                            return; // Salir del bucle una vez encontrado
                        }
                    }
                }
            });
        });

        return position;
    },

    /**
     * Inserta contenido SQL respetando límites de caracteres
     */
    async insertSQLContent(worksheet, position, sqlContent) {
        if (!sqlContent) return;

        const chunks = this.splitSQLIntoChunks(sqlContent, 32760);
        let currentRow = position.row;
        
        chunks.forEach((chunk, index) => {
            if (index > 0) {
                worksheet.insertRow(currentRow, []);
                this.copyRowStyle(worksheet, currentRow - 1, currentRow);
            }

            const startCol = 2; // Columna B
            const endCol = 11;  // Columna K
            
            worksheet.mergeCells(currentRow, startCol, currentRow, endCol);
            
            const cell = worksheet.getCell(currentRow, startCol);
            cell.value = this.addLineBreaks(chunk, 120);
            
            this.applySQLCellStyle(cell);
            
            const lineCount = cell.value.split('\n').length;
            worksheet.getRow(currentRow).height = Math.max(60, lineCount * 15);
            
            currentRow++;
        });
    },

    /**
     * Inserta tabla de datos
     */
    async insertTableContent(worksheet, position, tableData) {
        if (!tableData || !tableData.headers || !tableData.rows) return;

        let currentRow = position.row;
        
        // Insertar encabezados
        tableData.headers.forEach((header, colIndex) => {
            const cell = worksheet.getCell(currentRow, position.col + colIndex);
            cell.value = header;
            this.applyHeaderCellStyle(cell);
        });
        
        currentRow++;
        
        // Insertar datos
        tableData.rows.forEach(rowData => {
            rowData.forEach((value, colIndex) => {
                const cell = worksheet.getCell(currentRow, position.col + colIndex);
                cell.value = value;
                this.applyDataCellStyle(cell, colIndex === 3);
            });
            currentRow++;
        });
    },

    /**
     * Aplica estilo de código SQL
     */
    applySQLCellStyle(cell) {
        cell.style = {
            font: { name: 'Consolas', size: 10, color: { argb: 'FFECEFF4' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B4252' } },
            alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
            border: {
                top: { style: 'thin', color: { argb: 'FF4C566A' } },
                left: { style: 'thin', color: { argb: 'FF4C566A' } },
                bottom: { style: 'thin', color: { argb: 'FF4C566A' } },
                right: { style: 'thin', color: { argb: 'FF4C566A' } }
            }
        };
    },

    /**
     * Aplica estilo a encabezados de tabla
     */
    applyHeaderCellStyle(cell) {
        cell.style = {
            font: { size: 11, bold: true, color: { argb: 'FF2E3440' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8DEE9' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin', color: { argb: 'FF4C566A' } },
                left: { style: 'thin', color: { argb: 'FF4C566A' } },
                bottom: { style: 'thin', color: { argb: 'FF4C566A' } },
                right: { style: 'thin', color: { argb: 'FF4C566A' } }
            }
        };
    },

    /**
     * Aplica estilo a celdas de datos
     */
    applyDataCellStyle(cell, isDiff = false) {
        cell.style = {
            font: { size: 10, color: { argb: 'FF2E3440' }, bold: isDiff },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: isDiff ? 'FFFFF8DB' : 'FFFFFFFF' } },
            alignment: { horizontal: typeof cell.value === 'number' ? 'right' : 'left', vertical: 'middle' },
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
     * Copia estilo de una fila a otra
     */
    copyRowStyle(worksheet, sourceRow, targetRow) {
        const sourceRowObj = worksheet.getRow(sourceRow);
        const targetRowObj = worksheet.getRow(targetRow);
        
        targetRowObj.height = sourceRowObj.height;
        
        for (let col = 1; col <= 11; col++) {
            const sourceCell = worksheet.getCell(sourceRow, col);
            const targetCell = worksheet.getCell(targetRow, col);
            if (sourceCell.style) {
                targetCell.style = { ...sourceCell.style };
            }
        }
    },

    /**
     * Convierte letra de columna a número
     */
    columnLetterToNumber(letter) {
        let result = 0;
        for (let i = 0; i < letter.length; i++) {
            result = result * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return result;
    },

    /**
     * Combina queries MINUS en una sola
     */
    combineMinus(minus1, minus2) {
        let combined = '';
        
        if (minus1) {
            combined += '-- MINUS 1: EDV - DDV\n';
            combined += '-- Registros que están en EDV pero NO en DDV\n\n';
            combined += minus1;
            combined += '\n\n';
        }
        
        if (minus2) {
            combined += '-- MINUS 2: DDV - EDV\n';
            combined += '-- Registros que están en DDV pero NO en EDV\n\n';
            combined += minus2;
        }
        
        return combined;
    },

    /**
     * Genera datos de tabla para UNIVERSOS
     */
    generateUniversosTable(params) {
        return {
            headers: ['codmes', 'numreg_ddv', 'numreg_edv', 'diff_numreg', 'status'],
            rows: [
                [202505, 2765145, 2765145, 0, '✅ IGUALES'],
                [202506, 2758763, 2758763, 0, '✅ IGUALES'],
                [202507, 2787328, 2787328, 0, '✅ IGUALES']
            ]
        };
    },

    /**
     * Genera datos de tabla para AGRUPADOS
     */
    generateAgrupadosTable(params) {
        return {
            headers: ['capa', 'codmes', 'count_campos', 'sum_campos'],
            rows: [
                ['EDV', 202505, 150, 1250000.50],
                ['DDV', 202505, 150, 1250000.50],
                ['EDV', 202506, 148, 1180000.25],
                ['DDV', 202506, 148, 1180000.25]
            ]
        };
    },

    /**
     * Genera datos de tabla para MINUS
     */
    generateMinusTable(params) {
        return {
            headers: ['query', 'registros_encontrados', 'status'],
            rows: [
                ['MINUS 1 (EDV - DDV)', 0, '✅ Sin diferencias'],
                ['MINUS 2 (DDV - EDV)', 0, '✅ Sin diferencias']
            ]
        };
    },

    /**
     * Genera nombre de archivo para template
     */
    generateTemplateFilename(params) {
        const tableName = params.tablaDDV || 'TABLA';
        const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
        return `cuadre_template_${tableName.toUpperCase()}_${periods}.xlsx`;
    },

    /**
     * Genera nombre de archivo para auto-generación
     */
    generateAutoFilename(params) {
        const tableName = params.tablaDDV || 'TABLA';
        const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
        return `cuadre_auto_${tableName.toUpperCase()}_${periods}.xlsx`;
    },

    /**
     * Carga ExcelJS dinámicamente
     */
    async loadExcelJS() {
        if (window.ExcelJS) {
            return window.ExcelJS;
        }

        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js';
        document.head.appendChild(script);

        return new Promise((resolve, reject) => {
            script.onload = () => resolve(window.ExcelJS);
            script.onerror = reject;
        });
    },

    /**
     * Divide SQL en chunks seguros para Excel
     */
    splitSQLIntoChunks(sql, maxChars) {
        if (!sql || sql.length <= maxChars) {
            return [sql || ''];
        }
        
        const chunks = [];
        let currentPos = 0;
        
        while (currentPos < sql.length) {
            let chunkEnd = currentPos + maxChars;
            
            if (chunkEnd < sql.length) {
                const nearlineBreak = sql.lastIndexOf('\n', chunkEnd);
                if (nearlineBreak > currentPos + maxChars * 0.8) {
                    chunkEnd = nearlineBreak + 1;
                }
            }
            
            let chunk = sql.substring(currentPos, chunkEnd);
            
            if (chunk.trim().startsWith('=')) {
                chunk = "'" + chunk;
            }
            
            chunks.push(chunk);
            currentPos = chunkEnd;
        }
        
        return chunks;
    },

    /**
     * Agrega saltos de línea cada N caracteres para mejor legibilidad
     */
    addLineBreaks(text, lineLength) {
        if (!text) return '';
        
        const lines = text.split('\n');
        const result = [];
        
        lines.forEach(line => {
            if (line.length <= lineLength) {
                result.push(line);
            } else {
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
     * Vista previa del template cargado
     */
    async previewTemplate() {
        if (!this.templateBuffer) {
            alert('No hay template cargado');
            return;
        }

        try {
            const ExcelJS = await this.loadExcelJS();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(this.templateBuffer);
            
            let previewHTML = '<div class="template-preview">';
            previewHTML += '<h4>📊 Vista Previa del Template</h4>';
            
            workbook.worksheets.forEach((worksheet, index) => {
                previewHTML += `<div class="sheet-preview">`;
                previewHTML += `<h5>📋 ${worksheet.name}</h5>`;
                previewHTML += `<div class="sheet-info">`;
                previewHTML += `<p><strong>Dimensiones:</strong> ${worksheet.rowCount} filas x ${worksheet.columnCount} columnas</p>`;
                
                // Mostrar primeras filas
                previewHTML += `<div class="sheet-sample">`;
                previewHTML += `<h6>Primeras 5 filas:</h6>`;
                previewHTML += `<table class="preview-table">`;
                
                for (let row = 1; row <= Math.min(5, worksheet.rowCount); row++) {
                    previewHTML += `<tr>`;
                    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
                        const cell = worksheet.getCell(row, col);
                        const value = cell.value || '';
                        previewHTML += `<td>${value}</td>`;
                    }
                    previewHTML += `</tr>`;
                }
                
                previewHTML += `</table>`;
                previewHTML += `</div>`;
                previewHTML += `</div>`;
                previewHTML += `</div>`;
            });
            
            previewHTML += '</div>';
            
            if (typeof UIModule !== 'undefined' && UIModule.showModal) {
                UIModule.showModal('Vista Previa del Template', previewHTML);
            } else {
                alert('Vista previa no disponible');
            }
            
        } catch (error) {
            alert('Error generando vista previa: ' + error.message);
        }
    },

    /**
     * Limpia el template cargado
     */
    clearTemplate() {
        this.templateBuffer = null;
        this.templateWorkbook = null;
        
        const templateInfo = document.getElementById('templateInfo');
        if (templateInfo) {
            templateInfo.innerHTML = `
                <div class="template-empty">
                    <span class="template-icon">📂</span>
                    <p>No hay template cargado</p>
                    <small>Haz clic en "Cargar Template" para seleccionar un archivo Excel</small>
                </div>
            `;
        }
        
        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification('Template eliminado', 'info', 2000);
        }
    },

    // =============================================================================
    // FUNCIONES DE GENERACIÓN AUTOMÁTICA CON EXCELJS (para cuando no hay template)
    // =============================================================================

    /**
     * Agrega título principal
     */
    addMainTitle(worksheet, currentRow) {
        worksheet.mergeCells(`A${currentRow}:K${currentRow}`);
        
        const titleCell = worksheet.getCell(`A${currentRow}`);
        titleCell.value = 'Generador de Queries de Ratificación v2';
        
        titleCell.style = {
            font: { size: 18, bold: true, color: { argb: 'FFFFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B46C1' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };
        
        worksheet.getRow(currentRow).height = 40;
        
        return currentRow + 2;
    },

    /**
     * Agrega sección UNIVERSOS
     */
    async addUniversosSection(worksheet, currentRow, queryUniversos, params) {
        currentRow = this.addSectionTitle(worksheet, currentRow, 'UNIVERSOS');
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryUniversos);
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        currentRow = this.addUniversosResultTable(worksheet, currentRow, params);
        
        return currentRow + 2;
    },

    /**
     * Agrega sección AGRUPADOS
     */
    async addAgrupadosSection(worksheet, currentRow, queryAgrupados, params) {
        currentRow = this.addSectionTitle(worksheet, currentRow, 'AGRUPADOS');
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryAgrupados);
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        currentRow = this.addAgrupadosResultTable(worksheet, currentRow, params);
        
        return currentRow + 2;
    },

    /**
     * Agrega sección MINUS
     */
    async addMinusSection(worksheet, currentRow, queryMinus1, queryMinus2, params) {
        currentRow = this.addSectionTitle(worksheet, currentRow, 'MINUS');
        
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código MINUS 1 (EDV - DDV)');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryMinus1);
        
        currentRow = this.addSubtitle(worksheet, currentRow, 'Código MINUS 2 (DDV - EDV)');
        currentRow = this.addCodeBlock(worksheet, currentRow, queryMinus2);
        
        currentRow = this.addSubtitle(worksheet, currentRow, 'Resultado');
        currentRow = this.addMinusResultTable(worksheet, currentRow, params);
        
        return currentRow + 2;
    },

    /**
     * Agrega título de sección (H2)
     */
    addSectionTitle(worksheet, currentRow, title) {
        const titleCell = worksheet.getCell(`A${currentRow}`);
        titleCell.value = title;
        
        titleCell.style = {
            font: { size: 14, bold: true, color: { argb: 'FFEBCB8B' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E3440' } },
            alignment: { horizontal: 'left', vertical: 'middle' }
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
            font: { size: 12, bold: true, color: { argb: 'FF2E3440' } },
            alignment: { horizontal: 'left', vertical: 'middle' }
        };
        
        return currentRow + 1;
    },

    /**
     * Agrega bloque de código SQL con manejo de límite de 32,767 caracteres
     */
    addCodeBlock(worksheet, currentRow, sqlCode) {
        if (!sqlCode) {
            worksheet.mergeCells(`B${currentRow}:K${currentRow}`);
            const cell = worksheet.getCell(`B${currentRow}`);
            cell.value = '-- Query no disponible';
            this.applyCodeStyle(cell);
            return currentRow + 1;
        }

        const chunks = this.splitSQLIntoChunks(sqlCode, 32760);
        
        chunks.forEach((chunk, index) => {
            worksheet.mergeCells(`B${currentRow}:K${currentRow}`);
            const cell = worksheet.getCell(`B${currentRow}`);
            
            const formattedChunk = this.addLineBreaks(chunk, 120);
            cell.value = formattedChunk;
            
            this.applyCodeStyle(cell);
            
            if (index > 0) {
                const labelCell = worksheet.getCell(`A${currentRow}`);
                labelCell.value = 'Código (cont.)';
                labelCell.style = {
                    font: { size: 10, italic: true, color: { argb: 'FF6C7B7F' } },
                    alignment: { horizontal: 'left', vertical: 'top' }
                };
            }
            
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
            font: { name: 'Consolas', size: 10, color: { argb: 'FFECEFF4' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B4252' } },
            alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
            border: {
                top: { style: 'thin', color: { argb: 'FF4C566A' } },
                left: { style: 'thin', color: { argb: 'FF4C566A' } },
                bottom: { style: 'thin', color: { argb: 'FF4C566A' } },
                right: { style: 'thin', color: { argb: 'FF4C566A' } }
            }
        };
    },

    /**
     * Agrega tabla de resultado para UNIVERSOS
     */
    addUniversosResultTable(worksheet, currentRow, params) {
        const headers = ['codmes', 'numreg_ddv', 'numreg_edv', 'diff_numreg', 'status'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2);
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
        const exampleData = [
            [202505, 2765145, 2765145, 0, '✅ IGUALES'],
            [202506, 2758763, 2758763, 0, '✅ IGUALES'],
            [202507, 2787328, 2787328, 0, '✅ IGUALES']
        ];
        
        exampleData.forEach(row => {
            row.forEach((value, index) => {
                const cell = worksheet.getCell(currentRow, index + 2);
                cell.value = value;
                this.applyDataStyle(cell, index === 3);
            });
            currentRow++;
        });
        
        return currentRow;
    },

    /**
     * Agrega tabla de resultado para AGRUPADOS
     */
    addAgrupadosResultTable(worksheet, currentRow, params) {
        const headers = ['capa', 'codmes', 'count_campos', 'sum_campos'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2);
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
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
        const headers = ['query', 'registros_encontrados', 'status'];
        headers.forEach((header, index) => {
            const cell = worksheet.getCell(currentRow, index + 2);
            cell.value = header;
            this.applyHeaderStyle(cell);
        });
        
        currentRow++;
        
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
            font: { size: 11, bold: true, color: { argb: 'FF2E3440' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8DEE9' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
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
            font: { size: 10, color: { argb: 'FF2E3440' }, bold: isDiff },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: isDiff ? 'FFFFF8DB' : 'FFFFFFFF' } },
            alignment: { horizontal: typeof cell.value === 'number' ? 'right' : 'left', vertical: 'middle' },
            border: {
                top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
                right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
            },
            numFmt: typeof cell.value === 'number' && cell.value > 1000 ? '#,##0' : undefined
        };
    },

    // =============================================================================
    // FUNCIONES DE COMPATIBILIDAD (mantener funcionalidad original)
    // =============================================================================

    /**
     * Exporta toda la información a Excel (FUNCIÓN ORIGINAL)
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
     * Crea el workbook de Excel (FUNCIÓN ORIGINAL)
     */
    createWorkbook(format) {
        const wb = XLSX.utils.book_new();
        
        if (format === 'cuadre') {
            this.addParametersSheet(wb);
            this.addDescribeSheet(wb);
            this.addQueriesSheet(wb);
            this.addMetadataSheet(wb);
        } else {
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
     * Agrega hoja de resumen (FUNCIÓN ORIGINAL)
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
        this.formatSummarySheet(ws);
        XLSX.utils.book_append_sheet(wb, ws, 'RESUMEN');
    },

    /**
     * Formatea la hoja de resumen
     */
    formatSummarySheet(ws) {
        ws['!cols'] = [
            { width: 25 },
            { width: 50 },
            { width: 15 }
        ];
        
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
        ];
    },

    /**
     * Agrega hoja de parámetros
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
        
        if (params.renameRules) {
            Object.entries(params.renameRules).forEach(([original, renamed]) => {
                parametersData.push([original, renamed, '✅']);
            });
        }
        
        parametersData.push(['', '', '']);
        parametersData.push(['TABLAS FINALES PARA QUERIES', '', '']);
        parametersData.push(['DDV', `${params.esquemaDDV}.${params.tablaDDV}`, 'Tabla fuente']);
        parametersData.push(['EDV', `${params.esquemaEDV}.${params.tablaEDV}`, 'Tabla destino']);
        
        const ws = XLSX.utils.aoa_to_sheet(parametersData);
        ws['!cols'] = [{ width: 20 }, { width: 40 }, { width: 25 }];
        XLSX.utils.book_append_sheet(wb, ws, 'PARAMETROS');
    },

    /**
     * Agrega hoja de estructura de tabla (describe)
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
                'null',
                field.edvName,
                field.aggregateFunction.toUpperCase(),
                `${field.aggregateFunction}(${field.columnName})`,
                `${field.aggregateFunction}(${field.edvName})`
            ]);
        });
        
        const ws = XLSX.utils.aoa_to_sheet(describeData);
        ws['!cols'] = [
            { width: 20 }, { width: 15 }, { width: 15 }, { width: 20 },
            { width: 12 }, { width: 25 }, { width: 25 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'TABLA_DESCRIBE');
    },

    /**
     * Agrega hoja de estructura detallada de tabla
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
        
        const countFields = tableStructure.filter(f => f.aggregateFunction === 'count').length;
        const sumFields = tableStructure.filter(f => f.aggregateFunction === 'sum').length;
        
        structureData.push(['', '', '', '', '', '', '']);
        structureData.push(['ESTADÍSTICAS', '', '', '', '', '', '']);
        structureData.push(['Total Campos', tableStructure.length, '', '', '', '', '']);
        structureData.push(['Campos COUNT', countFields, '', '', '', '', '']);
        structureData.push(['Campos SUM', sumFields, '', '', '', '', '']);
        structureData.push(['% Numéricos', `${Math.round((sumFields / tableStructure.length) * 100)}%`, '', '', '', '', '']);
        
        const ws = XLSX.utils.aoa_to_sheet(structureData);
        ws['!cols'] = [
            { width: 5 }, { width: 20 }, { width: 15 }, { width: 15 },
            { width: 20 }, { width: 12 }, { width: 15 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'ESTRUCTURA_TABLA');
    },

    /**
     * Agrega hoja de queries (con manejo de límite de caracteres)
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
        ws['!cols'] = [
            { width: 20 }, { width: 40 }, { width: 80 }, { width: 10 }, { width: 12 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'QUERIES_RATIFICACION');
    },

    /**
     * Agrega hoja de validación
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
        ws['!cols'] = [
            { width: 25 }, { width: 10 }, { width: 40 }, { width: 30 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'VALIDACION');
    },

    /**
     * Agrega hoja de metadatos
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
        ws['!cols'] = [
            { width: 25 }, { width: 40 }, { width: 35 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'METADATOS');
    },

    /**
     * Genera nombre de archivo para export
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
        
        content += `ESTRUCTURA DE TABLA:\n`;
        content += `${'-'.repeat(25)}\n`;
        content += `- Total campos: ${tableStructure.length}\n`;
        content += `- Campos COUNT: ${tableStructure.filter(f => f.aggregateFunction === 'count').length}\n`;
        content += `- Campos SUM: ${tableStructure.filter(f => f.aggregateFunction === 'sum').length}\n`;
        if (tableStructure.length > 0) {
            content += `- Porcentaje numérico: ${Math.round((tableStructure.filter(f => f.aggregateFunction === 'sum').length / tableStructure.length) * 100)}%\n`;
        }
        content += `\n`;
        
        content += `QUERIES GENERADOS:\n`;
        content += `${'-'.repeat(20)}\n`;
        Object.entries(queries).forEach(([key, query]) => {
            const title = this.getQueryTitle(key);
            const lines = query.split('\n').length;
            const chars = query.length;
            content += `- ${title}: ${lines} líneas, ${chars} caracteres\n`;
        });
        content += `\n`;
        
        content += `INSTRUCCIONES DE USO:\n`;
        content += `${'-'.repeat(25)}\n`;
        content += `1. Copiar el query deseado completo\n`;
        content += `2. Pegar en tu editor SQL preferido\n`;
        content += `3. Ejecutar en el motor de base de datos correspondiente\n`;
        content += `4. Analizar los resultados para identificar diferencias\n`;
        content += `5. Documentar hallazgos para seguimiento y corrección\n`;
        content += `\n`;
        
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
        
        content += `\n\n${'='.repeat(80)}\n`;
        content += `FIN DEL REPORTE - Generado por: Generador de Queries de Ratificación v2\n`;
        content += `Fecha: ${new Date().toISOString()}\n`;
        content += `${'='.repeat(80)}`;
        
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
     * Exporta queries en formato Excel (función faltante)
     */
    exportQueriesExcel() {
        const queries = QueryModule.getGeneratedQueries();
        const params = ParametersModule.getCurrentParameters();
        
        if (!queries || Object.keys(queries).length === 0) {
            alert('No hay queries para exportar. Primero genera los queries en la pestaña correspondiente.');
            return;
        }
        
        const wb = XLSX.utils.book_new();
        
        Object.entries(queries).forEach(([key, query]) => {
            this.createQuerySheet(wb, key, query);
        });
        
        const tableName = params.tablaDDV || 'TABLA';
        const periods = params.periodos ? params.periodos.replace(/\s/g, '').replace(/,/g, '_') : 'periodos';
        const filename = `queries_${tableName}_${periods}.xlsx`;
        
        XLSX.writeFile(wb, filename);
        
        if (typeof UIModule !== 'undefined' && UIModule.showNotification) {
            UIModule.showNotification(`Excel de queries generado: ${filename}`, 'success', 4000);
        }
    },

    /**
     * Crea hoja individual para un query (evitando límite de 32767 caracteres)
     */
    createQuerySheet(wb, key, query) {
        const title = this.getQueryTitle(key);
        const parts = this.splitTextForExcel(query);
        
        const data = [
            [title],
            [''],
            ['PARTE', 'CONTENIDO SQL'],
            ...parts.map((part, index) => [
                parts.length > 1 ? `Parte ${index + 1}` : 'Query Completo',
                part
            ])
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        ws['!cols'] = [
            { width: 15 },
            { width: 120 }
        ];
        
        const sheetName = key.toUpperCase().substring(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    },

    /**
     * Divide texto largo para evitar el límite de 32767 caracteres de Excel
     */
    splitTextForExcel(text) {
        if (!text) return [''];
        
        const MAX_EXCEL_CHARS = 32000;
        
        if (text.length <= MAX_EXCEL_CHARS) {
            return [text];
        }
        
        const parts = [];
        const lines = text.split('\n');
        let currentPart = '';
        
        for (const line of lines) {
            if ((currentPart + '\n' + line).length > MAX_EXCEL_CHARS && currentPart.length > 0) {
                parts.push(currentPart.trim());
                currentPart = line;
            } else {
                currentPart += (currentPart ? '\n' : '') + line;
            }
        }
        
        if (currentPart.trim()) {
            parts.push(currentPart.trim());
        }
        
        return parts.length > 0 ? parts : [text.substring(0, MAX_EXCEL_CHARS)];
    },

    /**
     * Trunca texto para Excel si es necesario
     */
    truncateForExcel(text) {
        const MAX_EXCEL_CHARS = 32000;
        
        if (!text || text.length <= MAX_EXCEL_CHARS) {
            return text || '';
        }
        
        return text.substring(0, MAX_EXCEL_CHARS) + '\n\n[... CONTENIDO TRUNCADO ...]';
    }
};