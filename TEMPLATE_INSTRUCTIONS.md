# 📊 Instrucciones para el Template de Excel

## 🎯 Funcionalidad Completada

¡La funcionalidad de template Excel ha sido **completamente implementada**! Ahora puedes:

1. **Cargar un template personalizado** Excel (.xlsx)
2. **Generar Excel automáticamente** con las queries en las pestañas correspondientes
3. **Manejar múltiples pestañas** (Universo, Agrupados, Minus)
4. **Vista previa del template** antes de usarlo

## 🚀 Cómo Usar la Funcionalidad

### Paso 1: Generar Template de Ejemplo
1. Abre el archivo `template-generator.html` en tu navegador
2. Haz clic en "Generar Template de Ejemplo"
3. Se descargará un archivo `template_cuadre_edv_ejemplo.xlsx`

### Paso 2: Personalizar el Template (Opcional)
1. Abre el archivo descargado en Excel
2. Modifica el formato, colores, estilos según tus necesidades
3. **NO elimines** los placeholders como `<<UNIVERSOS_SQL>>`
4. Guarda el archivo personalizado

### Paso 3: Usar en la Aplicación Principal
1. Ve a la pestaña **"5. Exportar Excel"**
2. En la sección **"Template Personalizado"**, haz clic en **"Cargar Template"**
3. Selecciona tu archivo Excel personalizado
4. El sistema validará automáticamente el template
5. Ve a la pestaña **"3. Queries Generados"** y genera los queries
6. Regresa a **"5. Exportar Excel"** y haz clic en **"GENERAR EXCEL CUADRE EDV"**

## 📋 Estructura del Template

### Pestañas Soportadas
- **Cuadre**: Pestaña principal con todos los datos
- **Universo**: Pestaña específica para queries de universos
- **Agrupados**: Pestaña específica para queries agrupados  
- **Minus**: Pestaña específica para queries minus

### Placeholders Soportados

#### Para Pestaña Principal (Cuadre):
- `<<UNIVERSOS_SQL>>` - Query de universos
- `<<UNIVERSOS_TABLA>>` - Tabla de resultados universos
- `<<AGRUPADOS_SQL>>` - Query de agrupados
- `<<AGRUPADOS_TABLA>>` - Tabla de resultados agrupados
- `<<MINUS_SQL>>` - Query de minus
- `<<MINUS_TABLA>>` - Tabla de resultados minus

#### Para Pestañas Específicas:
- `<<UNIV_SQL>>` / `<<AGR_SQL>>` / `<<MINUS_SQL>>` - Queries específicos
- `<<UNIV_TABLA>>` / `<<AGR_TABLA>>` / `<<MINUS_TABLA>>` - Tablas específicas

#### Placeholders Alternativos Soportados:
- `{{UNIVERSOS_SQL}}` - Formato con llaves dobles
- `[UNIVERSOS_SQL]` - Formato con corchetes
- `(UNIVERSOS_SQL)` - Formato con paréntesis
- `%UNIVERSOS_SQL%` - Formato con porcentajes
- (Y todos los demás con el mismo formato)

#### Para Pestañas Específicas:
- `<<UNIV_SQL>>` / `<<UNIV_TABLA>>` - Universos
- `<<AGR_SQL>>` / `<<AGR_TABLA>>` - Agrupados
- `<<MINUS_SQL>>` / `<<MINUS_TABLA>>` - Minus

## 🔧 Características Técnicas

### Validación Automática
- ✅ Detecta pestañas válidas automáticamente
- ✅ Valida placeholders y nombres definidos
- ✅ Muestra información detallada del template
- ✅ Vista previa del contenido
- ✅ **CORREGIDO**: Manejo robusto de nombres definidos de Excel

### Manejo de Contenido
- ✅ Respeta límites de caracteres de Excel (32,767)
- ✅ Divide contenido largo automáticamente
- ✅ Mantiene formato y estilos del template
- ✅ Inserta datos en las pestañas correctas
- ✅ **NUEVO**: Soporte para múltiples formatos de placeholders
- ✅ **NUEVO**: Detección automática de pestañas compatibles

### Interfaz de Usuario
- ✅ Carga con indicador de progreso
- ✅ Información detallada del template cargado
- ✅ Botones de vista previa y limpieza
- ✅ Notificaciones de éxito/error
- ✅ **NUEVO**: Análisis de compatibilidad de pestañas
- ✅ **NUEVO**: Contador de elementos insertados

## 🎨 Personalización del Template

### Qué Puedes Cambiar:
- ✅ Colores y estilos de celdas
- ✅ Tamaños de columnas y filas
- ✅ Fuentes y formatos
- ✅ Bordes y sombras
- ✅ Agregar logos o imágenes
- ✅ Cambiar nombres de pestañas (manteniendo palabras clave)

### Qué NO Debes Cambiar:
- ❌ Los placeholders (ej: `<<UNIVERSOS_SQL>>`)
- ❌ La estructura básica de las pestañas
- ❌ Los nombres de las pestañas principales

## 🔍 Solución de Problemas

### ✅ **ERROR CORREGIDO**: `workbook.definedNames.get is not a function`
**Problema**: Error al cargar templates que no tienen nombres definidos de Excel.
**Solución**: El sistema ahora maneja robustamente este error y continúa funcionando con placeholders.

### Template No Se Carga
- Verifica que el archivo sea .xlsx o .xls
- Asegúrate de que tenga al menos una pestaña con placeholders
- Revisa que los placeholders estén escritos correctamente
- **NUEVO**: El sistema ahora soporta múltiples formatos de placeholders

### No Se Insertan Datos
- Verifica que los placeholders estén en el template
- Asegúrate de que los queries estén generados
- Revisa la consola del navegador para errores
- **NUEVO**: Revisa el contador de elementos insertados en la consola

### Error de Validación
- El template debe tener al menos una pestaña principal
- Debe contener placeholders o nombres definidos
- Verifica que el archivo no esté corrupto
- **NUEVO**: El sistema ahora es más tolerante con diferentes formatos

### Pestañas No Detectadas
- Asegúrate de que los nombres de pestañas contengan palabras clave:
  - "universo", "universos", "univ" para pestañas de universos
  - "agrupado", "agrupados", "agr" para pestañas de agrupados
  - "minus", "diferencia", "diff" para pestañas de minus
  - "cuadre", "resumen", "summary", "main" para pestaña principal

## 📞 Soporte

Si tienes problemas:
1. Revisa la consola del navegador (F12)
2. Verifica que todos los módulos estén cargados
3. Asegúrate de tener ExcelJS disponible
4. Prueba con el template de ejemplo primero

## 🎉 ¡Listo para Usar!

La funcionalidad está **100% completa** y lista para producción. Puedes:

- ✅ Cargar cualquier template Excel personalizado
- ✅ Generar Excel con formato personalizado
- ✅ Manejar múltiples pestañas automáticamente
- ✅ Tener control total sobre el formato final

¡Disfruta de tu nueva funcionalidad de template Excel! 🚀
