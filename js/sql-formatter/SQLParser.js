class SQLParser {
    constructor() {
        this.keywords = [
            'SELECT', 'DISTINCT', 'FROM', 'WHERE', 'GROUP BY', 'HAVING', 
            'ORDER BY', 'LIMIT', 'OFFSET', 'UNION', 'UNION ALL', 'INTERSECT', 
            'EXCEPT', 'MINUS', 'MINUS ALL', 'WITH', 'INSERT', 'UPDATE', 'DELETE', 
            'CREATE', 'ALTER', 'DROP', 'JOIN', 'LEFT JOIN', 'RIGHT JOIN', 
            'INNER JOIN', 'OUTER JOIN', 'FULL JOIN', 'CROSS JOIN', 'ON', 'USING', 
            'AS', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'AND', 'OR', 'NOT', 
            'IN', 'EXISTS', 'BETWEEN', 'FILTER'
        ];
    }

    /**
     * Limpia la consulta removiendo comentarios y normalizando espacios
     */
    cleanQuery(query) {
        // Remover comentarios de línea
        query = query.replace(/--.*$/gm, '');
        
        // Remover comentarios de bloque
        query = query.replace(/\/\*[\s\S]*?\*\//g, '');
        
        // Normalizar espacios múltiples pero preservar saltos de línea importantes
        query = query.replace(/[ \t]+/g, ' ');
        query = query.replace(/\n\s*\n/g, '\n');
        
        return query.trim();
    }

    /**
     * Parsea una consulta SQL en cláusulas estructuradas
     */
    parseQuery(query) {
        const cleanQuery = this.cleanQuery(query);
        const clauses = this.extractClausesAdvanced(cleanQuery);
        
        return {
            original: query,
            cleaned: cleanQuery,
            clauses: clauses,
            hasSubqueries: this.hasSubqueries(cleanQuery)
        };
    }

    /**
     * Extrae campos de una cláusula que los contenga
     */
    extractFields(content) {
        const fields = [];
        let currentField = '';
        let parenthesesCount = 0;
        let inQuotes = false;
        let quoteChar = '';

        for (let i = 0; i < content.length; i++) {
            const char = content[i];
            const prevChar = i > 0 ? content[i - 1] : '';

            // Manejo de comillas
            if ((char === '"' || char === "'") && prevChar !== '\\') {
                if (!inQuotes) {
                    inQuotes = true;
                    quoteChar = char;
                } else if (char === quoteChar) {
                    inQuotes = false;
                    quoteChar = '';
                }
            }

            if (!inQuotes) {
                if (char === '(') {
                    parenthesesCount++;
                } else if (char === ')') {
                    parenthesesCount--;
                } else if (char === ',' && parenthesesCount === 0) {
                    // Fin del campo actual
                    if (currentField.trim()) {
                        fields.push(currentField.trim());
                    }
                    currentField = '';
                    continue;
                }
            }

            currentField += char;
        }

        // Agregar el último campo
        if (currentField.trim()) {
            fields.push(currentField.trim());
        }

        return fields;
    }

    /**
     * Detecta si la consulta tiene subconsultas
     */
    hasSubqueries(query) {
        let parenthesesLevel = 0;
        let inQuotes = false;
        let quoteChar = '';
        
        for (let i = 0; i < query.length; i++) {
            const char = query[i];
            const prevChar = i > 0 ? query[i - 1] : '';

            // Manejo de comillas
            if ((char === '"' || char === "'") && prevChar !== '\\') {
                if (!inQuotes) {
                    inQuotes = true;
                    quoteChar = char;
                } else if (char === quoteChar) {
                    inQuotes = false;
                    quoteChar = '';
                }
            }

            if (!inQuotes) {
                if (char === '(') {
                    parenthesesLevel++;
                    // Verificar si hay SELECT después del paréntesis
                    const remaining = query.substring(i + 1).trim().toUpperCase();
                    if (remaining.startsWith('SELECT')) {
                        return true;
                    }
                } else if (char === ')') {
                    parenthesesLevel--;
                }
            }
        }

        return false;
    }
}