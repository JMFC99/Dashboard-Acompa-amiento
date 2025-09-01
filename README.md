# Documentación de Herramientas Qualtrics

## Instalación y Configuración

### Paso 1: Instalación de Python

#### En Windows

1. **Descargar Python:**
   - Ve a [python.org](https://www.python.org/downloads/)
   - Descarga Python 3.8 o superior
   
2. **Instalar Python:**
   - Ejecuta el instalador
   - **IMPORTANTE**: Marca la casilla "Add Python to PATH"
   - Selecciona "Install Now"

3. **Verificar instalación:**
   ```cmd
   python --version
   pip --version
   ```

#### En macOS

**Opción 1: Usando Homebrew (Recomendado)**
```bash
# Instalar Homebrew si no lo tienes
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Instalar Python
brew install python
```

**Opción 2: Descarga directa**
- Ve a [python.org](https://www.python.org/downloads/macos/)
- Descarga e instala Python 3.8+

**Verificar instalación:**
```bash
python3 --version
pip3 --version
```

#### En Linux (Ubuntu/Debian)

```bash
# Actualizar repositorios
sudo apt update

# Instalar Python y herramientas
sudo apt install python3 python3-pip python3-venv

# Verificar instalación
python3 --version
pip3 --version
```

**En otras distribuciones:**
- **CentOS/RHEL/Fedora**: `sudo yum install python3 python3-pip` o `sudo dnf install python3 python3-pip`
- **Arch Linux**: `sudo pacman -S python python-pip`

### Paso 2: Configuración del Ambiente Virtual

Los ambientes virtuales mantienen las dependencias del proyecto separadas y evitan conflictos.

Se recomienda usar un ambiente virtual para evitar conflictos de dependencias.

#### En Windows

```cmd
# 1. Navegar al directorio del proyecto
cd ruta\al\proyecto

# 2. Crear ambiente virtual
python -m venv qualtrics_env

# 3. Activar ambiente virtual
qualtrics_env\Scripts\activate

# 4. Actualizar pip (recomendado)
python -m pip install --upgrade pip

# 5. Instalar dependencias
pip install -r requirements.txt
```

**Nota:** Verás `(qualtrics_env)` al inicio de tu línea de comandos cuando el ambiente esté activado.

#### En macOS

```bash
# 1. Navegar al directorio del proyecto
cd /ruta/al/proyecto

# 2. Crear ambiente virtual
python3 -m venv qualtrics_env

# 3. Activar ambiente virtual
source qualtrics_env/bin/activate

# 4. Actualizar pip (recomendado)
python -m pip install --upgrade pip

# 5. Instalar dependencias
pip install -r requirements.txt
```

#### En Linux

```bash
# 1. Navegar al directorio del proyecto
cd /ruta/al/proyecto

# 2. Crear ambiente virtual
python3 -m venv qualtrics_env

# 3. Activar ambiente virtual
source qualtrics_env/bin/activate

# 4. Actualizar pip (recomendado)
python -m pip install --upgrade pip

# 5. Instalar dependencias
pip install -r requirements.txt
```

### Paso 3: Verificación de la Instalación

Para verificar que todo se instaló correctamente:

```bash
# Verificar que el ambiente virtual está activo
# Debes ver (qualtrics_env) al inicio de la línea de comandos

# Verificar versión de Python
python --version

# Listar paquetes instalados
pip list

# Ejecutar prueba básica
python -c "import pandas, openpyxl; print('✅ Instalación exitosa - Todos los paquetes funcionan correctamente')"
```

### Paso 4: Uso Diario del Proyecto

#### Activar el ambiente virtual (cada vez que trabajes)

**Windows:**
```cmd
cd ruta\al\proyecto
qualtrics_env\Scripts\activate
```

**macOS/Linux:**
```bash
cd /ruta/al/proyecto
source qualtrics_env/bin/activate
```

#### Desactivar el ambiente virtual (cuando termines)

```bash
deactivate
```

### Configuración Alternativa (Sin Ambiente Virtual)

⚠️ **No recomendado pero disponible si tienes problemas con ambientes virtuales:**

#### Windows
```cmd
pip install -r requirements.txt
```

#### macOS/Linux
```bash
pip3 install -r requirements.txt
```

### Guía Rápida de Instalación

**Resumen para usuarios experimentados:**

```bash
# Clonar/descargar proyecto y navegar al directorio
cd proyecto-qualtrics

# Crear y activar ambiente virtual
python3 -m venv qualtrics_env
source qualtrics_env/bin/activate  # macOS/Linux
# O: qualtrics_env\Scripts\activate  # Windows

# Instalar dependencias
pip install -r requirements.txt

# Verificar instalación
python -c "import pandas, openpyxl; print('✅ Listo para usar')"
```

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

## Solución de Problemas Comunes

### 🚨 Errores de Python

#### "Python no se reconoce como comando"

**Windows:**
- Reinstala Python desde [python.org](https://www.python.org/downloads/)
- **IMPORTANTE**: Marca "Add Python to PATH" durante la instalación
- Reinicia la terminal después de instalar

**macOS:**
- Usa `python3` en lugar de `python`
- Si no funciona: `brew install python`

**Linux:**
```bash
sudo apt install python3 python3-pip python3-venv
```

### 🚨 Errores de Ambiente Virtual

#### "No se puede crear el ambiente virtual"

**Todos los sistemas:**
```bash
# Asegúrate de tener venv instalado
python -m pip install virtualenv

# Usa virtualenv como alternativa
virtualenv qualtrics_env
```

#### "comando 'source' no reconocido" (Windows)

- Estás usando Command Prompt en lugar de PowerShell
- Usa: `qualtrics_env\Scripts\activate.bat`

### 🚨 Errores de Instalación de Paquetes

#### "No module named [nombre_paquete]"

1. Verifica que el ambiente virtual esté activado:
   ```bash
   # Debes ver (qualtrics_env) en tu terminal
   ```

2. Reinstala las dependencias:
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

#### "Permission denied" / Error de permisos

**macOS/Linux:**
- NO uses `sudo` con pip en ambiente virtual
- Si el problema persiste:
  ```bash
  pip install --user -r requirements.txt
  ```

**Windows:**
- Ejecuta la terminal como administrador
- O instala solo para tu usuario:
  ```cmd
  pip install --user -r requirements.txt
  ```

### 🚨 Problemas con requirements.txt

#### "requirements.txt not found"

1. Verifica que estás en el directorio correcto:
   ```bash
   ls requirements.txt  # Linux/macOS
   dir requirements.txt # Windows
   ```

2. Si no existe el archivo, crea uno básico:
   ```txt
   pandas>=1.3.0
   openpyxl>=3.0.0
   xlsxwriter>=3.0.0
   ```

### 🚨 Problemas Específicos por Sistema

#### macOS: "command line tools"
```bash
xcode-select --install
```

#### Linux: Paquetes del sistema faltantes
```bash
sudo apt update
sudo apt install build-essential python3-dev
```

### ✅ Verificación Final

Ejecuta este comando para verificar que todo funciona:

```bash
python -c "
import sys
print(f'✅ Python {sys.version}')
try:
    import pandas as pd
    import openpyxl
    print('✅ Pandas y OpenPyXL funcionan correctamente')
    print('🎉 ¡Todo listo para usar las herramientas de Qualtrics!')
except ImportError as e:
    print(f'❌ Error: {e}')
    print('Ejecuta: pip install -r requirements.txt')
"
```