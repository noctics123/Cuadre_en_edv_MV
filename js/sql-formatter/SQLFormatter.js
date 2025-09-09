class SQLFormatter {
    constructor(options = {}) {
        this.parser = new SQLParser();
        this.fieldFormatter = new FieldFormatter({
            maxCharsPerLine: options.maxCharsPerLine || 32000,
            excelMaxChars: options.excelMaxChars || 32767,
            indentSize: options.indentSize || 4
        });
        this.queryBuilder = new QueryBuilder({
            indentSize: options.indentSize || 4,
            preserveFormatting: options.preserveFormatting !== false,
            addBlankLines: options.addBlankLines !== false
        });

        // Configuración
        this.maxCharsPerLine = options.maxCharsPerLine || 32000;
        this.excelMaxChars = options.excelMaxChars || 32767;
        this.indentSize = options.indentSize || 4;
    }

    /**
     * Actualiza la configuración del formateador
     */
    updateSettings(settings) {
        if (settings.maxCharsPerLine) this.maxCharsPerLine = settings.maxCharsPerLine;
        if (settings.excelMaxChars) this.excelMaxChars = settings.excelMaxChars;
        if (settings.indentSize) this.indentSize = settings.indentSize;

        // Actualizar en los módulos
        this.fieldFormatter.updateSettings(settings);
        this.queryBuilder.updateSettings(settings);
    }

    /**
     * Formatea una consulta SQL - SOLO campos, preserva esqueleto
     */
    formatSQL(query, isForExcel = false) {
        try {
            if (!query || typeof query !== 'string') {
                throw new Error('La consulta debe ser una cadena de texto válida');
            }

            const trimmedQuery = query.trim();
            if (!trimmedQuery) {
                throw new Error('La consulta no puede estar vacía');
            }

            // NUEVO ENFOQUE: Solo formatear campos en SELECT, preservar todo lo demás
            const formattedQuery = this.formatOnlySelectFields(trimmedQuery, isForExcel);
            const stats = this.calculateBasicStats(trimmedQuery, formattedQuery, isForExcel);

            return {
                success: true,
                query: formattedQuery,
                stats: stats
            };

        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * Formatea SOLO los campos dentro de SELECT, preservando el resto
     */
    formatOnlySelectFields(query, isForExcel = false) {
        let formattedQuery = query;
        
        const selectBlocks = this.findSelectFieldBlocks(query);
        
        for (const block of selectBlocks) {
            const fields = this.parser.extractFields(block.fieldsContent);
            
            if (this.shouldFormatFields(fields, block.fieldsContent)) {
                const formattedFields = this.fieldFormatter.formatFields(fields, isForExcel);
                
                let newSelectBlock;
                if (block.type === 'subquery') {
                    newSelectBlock = '(\n  SELECT\n' + this.addExtraIndent(formattedFields, '  ');
                } else {
                    newSelectBlock = 'SELECT\n' + formattedFields;
                }
                
                formattedQuery = formattedQuery.replace(block.fullMatch, newSelectBlock);
            }
        }
        
        return formattedQuery;
    }

    /**
     * Encuentra bloques de campos SELECT
     */
    findSelectFieldBlocks(query) {
        const blocks = [];
        
        const mainRegex = /SELECT\s+((?:(?!SELECT|FROM|UNION|MINUS|WHERE|GROUP|ORDER|HAVING|LIMIT)\S|\s)*?)(?=\s+(?:FROM|UNION|MINUS|\)|WHERE|GROUP|ORDER|HAVING|LIMIT|$))/gi;
        let match;

        while ((match = mainRegex.exec(query)) !== null) {
            const fieldsContent = match[1].trim();
            
            if (fieldsContent && this.isValidFieldContent(fieldsContent)) {
                blocks.push({
                    fullMatch: match[0],
                    fieldsContent: fieldsContent,
                    startIndex: match.index,
                    type: 'main'
                });
            }
        }

        const subqueryRegex = /\(\s*SELECT\s+((?:(?!SELECT|FROM|\)).)*?)(?=\s+FROM|\))/gi;
        while ((match = subqueryRegex.exec(query)) !== null) {
            const fieldsContent = match[1].trim();
            
            if (fieldsContent && this.isValidFieldContent(fieldsContent)) {
                const fullSelectMatch = query.substring(match.index).match(/\(\s*SELECT\s+[^)]*FROM/i);
                if (fullSelectMatch) {
                    blocks.push({
                        fullMatch: fullSelectMatch[0],
                        fieldsContent: fieldsContent,
                        startIndex: match.index,
                        type: 'subquery'
                    });
                }
            }
        }

        return blocks.sort((a, b) => a.startIndex - b.startIndex);
    }

    /**
     * Valida contenido de campos
     */
    isValidFieldContent(content) {
        const upper = content.toUpperCase().trim();
        if (upper === '*') return false;
        
        const invalidKeywords = ['FROM', 'WHERE', 'GROUP BY', 'ORDER BY', 'HAVING', 'UNION', 'MINUS'];
        for (const keyword of invalidKeywords) {
            if (upper === keyword || upper.startsWith(keyword + ' ')) {
                return false;
            }
        }
        
        return /[a-zA-Z_]/.test(content) && (content.includes(',') || content.includes('(') || content.split('\n').length > 1);
    }

    /**
     * Determina si formatear campos
     */
    shouldFormatFields(fields, originalContent) {
        // Si tiene pocos campos simples, no formatear
        if (fields.length <= 4 && originalContent.length < 200) {
            return false;
        }
        
        // Si los campos ya están en una línea y caben bien, no formatear
        if (originalContent.split('\n').length <= 2 && originalContent.length < 500) {
            return false;
        }
        
        return fields.length > 4 || originalContent.length > 300;
    }

    /**
     * Agrega indentación extra
     */
    addExtraIndent(text, extraIndent) {
        return text.split('\n').map(line => {
            if (line.trim()) {
                return extraIndent + line;
            }
            return line;
        }).join('\n');
    }

    /**
     * Estadísticas básicas
     */
    calculateBasicStats(originalQuery, formattedQuery, isForExcel) {
        const originalLines = originalQuery.split('\n').length;
        const formattedLines = formattedQuery.split('\n').length;
        const fieldCount = (formattedQuery.match(/\w+\s*\(/g) || []).length + 
                          (formattedQuery.match(/,\s*\w+/g) || []).length;
        const reduction = originalLines > 0 
            ? Math.round(((originalLines - formattedLines) / originalLines) * 100) 
            : 0;

        return {
            fieldCount: fieldCount,
            lineCount: formattedLines,
            charCount: formattedQuery.length,
            reductionPercent: Math.max(0, reduction),
            originalLines: originalLines,
            maxCharsUsed: isForExcel ? this.excelMaxChars : this.maxCharsPerLine,
            isExcelOptimized: isForExcel
        };
    }

    /**
     * Prepara para Excel
     */
    prepareForExcel(query) {
        const result = this.formatSQL(query, true);
        if (!result.success) return result;

        const lines = result.query.split('\n');
        const excelData = [];
        
        lines.forEach(line => {
            if (line.trim()) {
                if (line.length > this.excelMaxChars) {
                    const chunks = this.fieldFormatter.splitLineForExcel(line);
                    chunks.forEach(chunk => excelData.push([chunk]));
                } else {
                    excelData.push([line]);
                }
            }
        });

        return {
            success: true,
            data: excelData,
            stats: result.stats,
            rowCount: excelData.length
        };
    }
}