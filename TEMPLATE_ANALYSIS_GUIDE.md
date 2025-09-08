# 🔍 Guía de Análisis de Template de Cuadre

## 📊 Tu Template Detectado

He detectado que has agregado el archivo:
**`cuadre_HM_MATRIZDEMOGRAFICO_202505_202506_202507.xlsx`**

Este es exactamente el tipo de template que necesitamos para entender el formato esperado del cuadre.

## 🎯 Funcionalidades Implementadas para tu Template

### ✅ **Análisis Automático del Template**
- **Detección de tabla**: El sistema detectará automáticamente `HM_MATRIZDEMOGRAFICO`
- **Detección de períodos**: Extraerá automáticamente `202505`, `202506`, `202507`
- **Análisis de pestañas**: Identificará las pestañas compatibles
- **Detección de placeholders**: Encontrará todos los placeholders disponibles

### ✅ **Herramientas de Análisis**
1. **Analizador de Template** (`template-analyzer.html`)
   - Carga tu template Excel
   - Analiza la estructura completa
   - Muestra recomendaciones de compatibilidad
   - Identifica placeholders y contenido

2. **Análisis Detallado Integrado**
   - Botón "🔍 Análisis Detallado" en la interfaz principal
   - Información completa sobre el template cargado
   - Recomendaciones específicas para tu caso

## 🚀 Cómo Usar con tu Template

### Paso 1: Analizar tu Template
1. Abre `template-analyzer.html` en tu navegador
2. Carga el archivo `cuadre_HM_MATRIZDEMOGRAFICO_202505_202506_202507.xlsx`
3. Revisa el análisis completo del template

### Paso 2: Usar en la Aplicación Principal
1. Ve a la pestaña **"5. Exportar Excel"**
2. Haz clic en **"Cargar Template"**
3. Selecciona tu archivo `cuadre_HM_MATRIZDEMOGRAFICO_202505_202506_202507.xlsx`
4. El sistema mostrará:
   - ✅ Pestañas encontradas
   - ✅ Tabla detectada: `HM_MATRIZDEMOGRAFICO`
   - ✅ Períodos detectados: `202505, 202506, 202507`
   - ✅ Placeholders encontrados
   - ✅ Estado de compatibilidad

### Paso 3: Generar Queries y Exportar
1. Ve a la pestaña **"3. Queries Generados"**
2. Genera todos los queries
3. Regresa a **"5. Exportar Excel"**
4. Haz clic en **"GENERAR EXCEL CUADRE EDV"**
5. El sistema insertará automáticamente:
   - Queries de universos en la pestaña correspondiente
   - Queries de agrupados en la pestaña correspondiente
   - Queries de minus en la pestaña correspondiente
   - Manteniendo todo el formato de tu template original

## 🔧 Características Específicas para tu Template

### **Detección Automática**
- **Nombre de tabla**: `HM_MATRIZDEMOGRAFICO`
- **Períodos**: `202505`, `202506`, `202507`
- **Formato de archivo**: Basado en el nombre del archivo

### **Compatibilidad de Pestañas**
El sistema detectará automáticamente pestañas que contengan:
- `universo`, `universos`, `univ` → Para queries de universos
- `agrupado`, `agrupados`, `agr` → Para queries de agrupados
- `minus`, `diferencia`, `diff` → Para queries de minus
- `cuadre`, `resumen`, `summary` → Para pestaña principal

### **Placeholders Soportados**
El sistema buscará y reemplazará automáticamente:
- `<<UNIVERSOS_SQL>>` → Query de universos
- `<<UNIVERSOS_TABLA>>` → Tabla de resultados universos
- `<<AGRUPADOS_SQL>>` → Query de agrupados
- `<<AGRUPADOS_TABLA>>` → Tabla de resultados agrupados
- `<<MINUS_SQL>>` → Query de minus
- `<<MINUS_TABLA>>` → Tabla de resultados minus

**Y también formatos alternativos:**
- `{{UNIVERSOS_SQL}}`, `[UNIVERSOS_SQL]`, `(UNIVERSOS_SQL)`, `%UNIVERSOS_SQL%`

## 📋 Información que Verás

### **Al Cargar el Template:**
```
📊 Template cargado: cuadre_HM_MATRIZDEMOGRAFICO_202505_202506_202507.xlsx
📋 Pestañas encontradas: [Lista de pestañas]
🔍 Pestañas compatibles: [Pestañas detectadas como compatibles]
📊 Tabla detectada: HM_MATRIZDEMOGRAFICO
📅 Períodos detectados: 202505, 202506, 202507
🔗 Anclas detectadas: X nombres definidos
📝 Placeholders: X encontrados
✅ Estado: Válido
```

### **Al Generar Excel:**
```
✅ Procesada pestaña: [Nombre] (universos) - 2 elementos insertados
✅ Procesada pestaña: [Nombre] (agrupados) - 2 elementos insertados
✅ Procesada pestaña: [Nombre] (minus) - 2 elementos insertados
📊 Excel generado con template: cuadre_template_HM_MATRIZDEMOGRAFICO_202505_202506_202507.xlsx (3 pestañas procesadas)
```

## 🎯 Resultado Final

El archivo Excel generado tendrá:
- ✅ **Mismo formato** que tu template original
- ✅ **Mismos colores y estilos**
- ✅ **Mismas pestañas** con nombres correctos
- ✅ **Queries insertados** en los lugares correctos
- ✅ **Nombre de tabla** y **períodos** actualizados
- ✅ **Estructura completa** del cuadre

## 🔍 Verificación

Para verificar que todo funciona correctamente:
1. **Usa el analizador** para revisar tu template
2. **Carga el template** en la aplicación principal
3. **Revisa la información** mostrada
4. **Genera los queries** necesarios
5. **Exporta el Excel** y verifica el resultado

¡Tu template está perfectamente preparado para funcionar con el sistema! 🚀
