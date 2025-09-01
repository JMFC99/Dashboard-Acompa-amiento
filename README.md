# Documentación de Herramientas Qualtrics

## Conversión de Excel de Qualtrics a Formato Procesado

Esta documentación describe las herramientas disponibles para convertir archivos Excel descargados de Qualtrics a formatos procesados para análisis.

### 1. Procesamiento Interno

Convierte datos de encuestas internas de Qualtrics a formato Excel procesado.

**Script:** `qualtrics_internal_to_excel.py`

#### Parámetros

| Parámetro | Descripción |
|-----------|-------------|
| `-i, --input` | Ruta del archivo Excel de entrada (descargado de Qualtrics) |
| `-o, --output` | Ruta del archivo Excel de salida procesado |
| `-v, --verbose` | Mostrar información detallada del progreso |

#### Ejemplos de uso

```bash
# Usar archivos predeterminados
python qualtrics_internal_to_excel.py

# Especificar archivos personalizados
python qualtrics_internal_to_excel.py -i data/survey.xlsx -o results/processed.xlsx -v

# Solo entrada personalizada
python qualtrics_internal_to_excel.py -i mi_encuesta.xlsx -v
```

### 2. Procesamiento Externo

Convierte datos de encuestas externas de Qualtrics a formato Excel procesado.

**Script:** `qualtrics_external_to_excel.py`

#### Parámetros

| Parámetro | Descripción |
|-----------|-------------|
| `-i, --input` | Ruta del archivo Excel de entrada (descargado de Qualtrics) |
| `-o, --output` | Ruta del archivo Excel de salida procesado |
| `-v, --verbose` | Mostrar información detallada del progreso |

#### Ejemplos de uso

```bash
# Usar archivos predeterminados
python qualtrics_external_to_excel.py

# Especificar archivos personalizados
python qualtrics_external_to_excel.py -i data/survey.xlsx -o results/processed.xlsx -v

# Solo entrada personalizada
python qualtrics_external_to_excel.py -i encuesta_externa.xlsx -v
```

### 3. Generación de Dashboard

Convierte una base de datos en Excel a un dashboard interactivo.

**Script:** `excel_to_dashboard.py`

#### Parámetros

| Parámetro | Descripción |
|-----------|-------------|
| `-i, --input` | Ruta del archivo Excel con la base de datos |
| `-o, --output` | Nombre del archivo de salida para el dashboard |

#### Sintaxis

```bash
python excel_to_dashboard.py -i "base_datos.xlsx" -o "dashboard_output.xlsx"
```

#### Ejemplos de uso

```bash
# Ejemplo básico
python excel_to_dashboard.py -i "resultados_encuesta.xlsx" -o "dashboard_resultados.xlsx"

# Con rutas completas
python excel_to_dashboard.py -i "data/base_datos_completa.xlsx" -o "dashboards/mi_dashboard.xlsx"
```

## Flujo de Trabajo Recomendado

1. **Descarga** los datos de Qualtrics en formato Excel
2. **Procesa** los datos usando el script apropiado (interno o externo)
3. **Genera** el dashboard usando el archivo procesado como entrada

```bash
# Ejemplo completo del flujo
python qualtrics_internal_to_excel.py -i raw_data.xlsx -o processed_data.xlsx -v
python excel_to_dashboard.py -i processed_data.xlsx -o final_dashboard.xlsx
```

## Notas Importantes

- Asegúrate de que los archivos Excel de entrada estén en el formato correcto de Qualtrics
- Los archivos de salida se sobrescribirán si ya existen
- Usa la opción `-v` para monitorear el progreso en archivos grandes
- Verifica que tienes permisos de escritura en las carpetas de destino