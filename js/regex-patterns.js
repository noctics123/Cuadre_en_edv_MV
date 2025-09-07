/**
 * Patrones regex para detectar y analizar CREATE TABLE statements
 * Incluye soporte para el formato específico con TBLPROPERTIES sin punto y coma
 */
const RegexPatterns = {
    
    /**
     * Patrones para extraer CREATE TABLE completos del texto
     */
    createTablePatterns: [
        // Patrón principal para formato con TBLPROPERTIES (sin punto y coma)
        /CREATE\s+TABLE\s+[\s\S]*?\)\s*USING\s+[\s\S]*?TBLPROPERTIES\s*\([^)]*\)/gi,
        
        // Patrón para formato estándar con punto y coma
        /CREATE\s+TABLE\s+[\s\S]*?\)[^;]*;/gi,
        
        // Patrón para CREATE TABLE hasta el siguiente CREATE TABLE o final del archivo
        /CREATE\s+TABLE\s+[\s\S]*?(?=CREATE\s+TABLE|$)/gi,
        
        // Patrón para CREATE TABLE con LOCATION pero sin TBLPROPERTIES
        /CREATE\s+TABLE\s+[\s\S]*?\)\s*USING\s+[\s\S]*?LOCATION\s+[^'\s]+['\s][^']*'/gi,
        
        // Patrón básico sin características adicionales
        /CREATE\s+TABLE\s+[^;]+\)/gi
    ],

    /**
     * Patrones para extraer nombre de tabla
     */
    tableNamePatterns: [
        // Con esquema completo: catalog.schema.table
        /CREATE\s+TABLE\s+(?:[`"]?[\w.]+[`"]?\.)*([\w]+)[`"]?\s*\(/i,
        
        // Con esquema: schema.table
        /CREATE\s+TABLE\s+(?:[`"]?[\w]+[`"]?\.)?([\w]+)[`"]?\s*\(/i,
        
        // Solo nombre de tabla
        /CREATE\s+TABLE\s+[`"]?([\w]+)[`"]?\s*\(/i
    ],

    /**
     * Patrones para extraer esquema completo
     */
    schemaPatterns: [
        // Esquema completo: catalog.schema.database
        /CREATE\s+TABLE\s+([`"]?[\w.]+[`"]?)\.[\w]+\s*\(/i,
        
        // Solo esquema: schema
        /CREATE\s+TABLE\s+([`"]?[\w]+[`"]?)\.[\w]+\s*\(/i
    ],

    /**
     * Patrón para extraer contenido de columnas del CREATE TABLE
     * Mejorado para manejar el formato específico con USING delta, PARTITIONED BY, LOCATION y TBLPROPERTIES
     */
    columnsContentPattern: /CREATE\s+TABLE\s+[^\(]+\(\s*([\s\S]*?)\s*\)\s*(?:USING|PARTITIONED|LOCATION|TBLPROPERTIES|$)/i,

    /**
     * Patrones para identificar tipos de datos numéricos
     */
    numericDataTypes: [
        /^(DOUBLE|FLOAT|REAL|NUMERIC)$/i,
        /^DECIMAL(\(\d+,\s*\d+\))?$/i,
        /^NUMBER(\(\d+,\s*\d+\))?$/i
    ],

    /**
     * Patrón para limpiar comentarios SQL
     */
    sqlCommentPattern: /--.*$/gm,

    /**
     * Patrón para detectar esquemas DDV vs EDV
     */
    schemaTypePatterns: {
        ddv: /ddv|matrizvariables/i,
        edv: /edv|trdata/i
    }
};

/**
 * Funciones utilitarias para usar los patrones regex
 */
const RegexUtils = {
    
    /**
     * Extrae todos los CREATE TABLE de un texto usando múltiples patrones
     * @param {string} text - Texto que contiene CREATE TABLE statements
     * @returns {Array<string>} - Array de CREATE TABLE statements encontrados
     */
    extractMultipleCreateTables(text) {
        const allMatches = [];
        
        // Limpiar comentarios primero
        const cleanText = text.replace(RegexPatterns.sqlCommentPattern, '');
        
        // Aplicar cada patrón
        for (const pattern of RegexPatterns.createTablePatterns) {
            const matches = cleanText.match(pattern) || [];
            allMatches.push(...matches);
        }
        
        // Eliminar duplicados y filtrar válidos
        const uniqueMatches = [...new Set(allMatches)];
        return uniqueMatches.filter(match => {
            // Verificar que realmente contenga una definición de tabla válida
            return match.includes('(') && this.extractTableName(match) !== null;
        });
    },

    /**
     * Extrae el nombre de tabla usando patrones robustos
     * @param {string} createStatement - CREATE TABLE statement
     * @returns {string|null} - Nombre de la tabla o null si no se encuentra
     */
    extractTableName(createStatement) {
        for (const pattern of RegexPatterns.tableNamePatterns) {
            const match = createStatement.match(pattern);
            if (match && match[1]) {
                return match[1];
            }
        }
        
        // Último intento: buscar cualquier palabra después de CREATE TABLE
        const generalMatch = createStatement.match(/CREATE\s+TABLE\s+([^\s\(]+)/i);
        if (generalMatch) {
            const fullName = generalMatch[1];
            // Si tiene puntos, tomar la última parte
            const parts = fullName.split('.');
            return parts[parts.length - 1];
        }
        
        return null;
    },

    /**
     * Extrae el esquema completo usando patrones robustos
     * @param {string} createStatement - CREATE TABLE statement
     * @returns {string} - Esquema o 'unknown_schema'
     */
    extractSchemaName(createStatement) {
        for (const pattern of RegexPatterns.schemaPatterns) {
            const match = createStatement.match(pattern);
            if (match && match[1]) {
                // Remover comillas si existen
                return match[1].replace(/[`"]/g, '');
            }
        }
        
        return 'unknown_schema';
    },

    /**
     * Extrae el contenido de las columnas del CREATE TABLE
     * @param {string} sql - CREATE TABLE statement
     * @returns {string} - Contenido de las columnas
     */
    extractColumnsContent(sql) {
        const cleanSql = sql.replace(/\s+/g, ' ').trim();
        const match = cleanSql.match(RegexPatterns.columnsContentPattern);
        
        if (!match) {
            throw new Error('No se pudo encontrar la definición de columnas');
        }
        
        return match[1];
    },

    /**
     * Determina si un tipo de dato es numérico
     * @param {string} dataType - Tipo de dato
     * @returns {boolean} - true si es numérico
     */
    isNumericDataType(dataType) {
        const upperType = dataType.toUpperCase();
        return RegexPatterns.numericDataTypes.some(pattern => pattern.test(upperType));
    },

    /**
     * Determina el tipo de esquema (DDV o EDV)
     * @param {string} schema - Nombre del esquema
     * @returns {string} - 'ddv', 'edv' o 'unknown'
     */
    getSchemaType(schema) {
        if (RegexPatterns.schemaTypePatterns.ddv.test(schema)) {
            return 'ddv';
        }
        if (RegexPatterns.schemaTypePatterns.edv.test(schema)) {
            return 'edv';
        }
        return 'unknown';
    }
};