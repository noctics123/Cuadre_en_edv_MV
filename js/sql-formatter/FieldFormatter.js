class FieldFormatter {
    constructor(options = {}) {
        this.maxCharsPerLine = options.maxCharsPerLine || 30000;
        this.excelMaxChars = options.excelMaxChars || 32500;
        this.indentSize = options.indentSize || 4;
        this.fieldSeparator = '    '; // 4 espacios
    }

    /**
     * Formatea una lista de campos horizontalmente
     */
    formatFields(fields, isForExcel = false) {
        if (!fields || fields.length === 0) {
            return '';
        }

        const validFields = this.validateAndCleanFields(fields);
        if (validFields.length === 0) {
            return '';
        }

        const maxChars = isForExcel ? this.excelMaxChars : this.maxCharsPerLine;
        const indent = ' '.repeat(this.indentSize);
        
        return this.arrangeFieldsAggressively(validFields, maxChars, indent);
    }

    /**
     * Valida y limpia campos
     */
    validateAndCleanFields(fields) {
        return fields.filter(field => {
            const trimmed = field.trim();
            return trimmed && this.isActualField(trimmed);
        });
    }

    /**
     * Verifica si es un campo válido
     */
    isActualField(fieldStr) {
        const upper = fieldStr.toUpperCase().trim();
        
        // No debe ser solo asterisco
        if (upper === '*') return false;
        
        // No debe contener palabras clave SQL sueltas
        const invalidKeywords = ['FROM', 'WHERE', 'GROUP BY', 'ORDER BY', 'HAVING', 'UNION', 'MINUS'];
        for (const keyword of invalidKeywords) {
            if (upper === keyword || upper.startsWith(keyword + ' ')) {
                return false;
            }
        }
        
        // Debe parecer un campo válido
        return /[a-zA-Z_]/.test(fieldStr) && (fieldStr.includes(',') || fieldStr.includes('(') || fieldStr.split('\n').length > 1 || fieldStr.length > 0);
    }

    /**
     * Método de empaquetado agresivo para maximizar campos por línea
     */
    arrangeFieldsAggressively(fields, maxChars, indent) {
        if (fields.length === 0) return '';
        
        const lines = [];
        let currentLine = '';
        const availableChars = maxChars - indent.length;
        const separator = this.fieldSeparator;
        
        for (let i = 0; i < fields.length; i++) {
            const field = fields[i].trim();
            const isLast = i === fields.length - 1;
            const fieldWithComma = isLast ? field : field + ',';
            
            // Calcular si cabe en la línea actual
            let proposedLine;
            if (currentLine === '') {
                proposedLine = fieldWithComma;
            } else {
                proposedLine = currentLine + separator + fieldWithComma;
            }

            // Si cabe, agregarlo; si no, crear nueva línea
            if (proposedLine.length <= availableChars) {
                currentLine = proposedLine;
            } else {
                // Línea llena, guardarla e iniciar nueva
                if (currentLine !== '') {
                    if (!currentLine.endsWith(',') && !isLast) {
                        currentLine += ',';
                    }
                    lines.push(indent + currentLine);
                }
                currentLine = fieldWithComma;
            }

            // Si es el último campo, agregar la línea final
            if (isLast && currentLine !== '') {
                lines.push(indent + currentLine);
            }
        }

        return lines.join('\n');
    }

    /**
     * Formatea para Excel dividiendo líneas largas
     */
    formatForExcel(fields) {
        const formattedText = this.formatFields(fields, true);
        const lines = formattedText.split('\n');
        const excelRows = [];

        lines.forEach(line => {
            if (line.trim()) {
                if (line.length > this.excelMaxChars) {
                    const chunks = this.splitLineForExcel(line);
                    chunks.forEach(chunk => excelRows.push([chunk]));
                } else {
                    excelRows.push([line]);
                }
            }
        });

        return excelRows;
    }

    /**
     * Divide línea para Excel
     */
    splitLineForExcel(line) {
        const chunks = [];
        let currentChunk = '';
        const words = line.trim().split(/(\s+|,\s*)/);

        for (const word of words) {
            const testChunk = currentChunk ? currentChunk + word : word;
            
            if (testChunk.length <= this.excelMaxChars) {
                currentChunk = testChunk;
            } else {
                if (currentChunk) {
                    chunks.push(currentChunk);
                }
                currentChunk = word;
            }
        }

        if (currentChunk) {
            chunks.push(currentChunk);
        }

        return chunks.length > 0 ? chunks : [line];
    }
}