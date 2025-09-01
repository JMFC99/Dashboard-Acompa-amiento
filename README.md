# Documentación información


### Convertir de excel descargado a qualtrics a formato

#### Interno
`-i/--input`: Especificar entrada del excel. Excel que utilizaremos
`-o/--output`: Especificar la salida del excel
`-v/--verbose`: Ver progreso

"""
# Use default files
python qualtrics_internal_to_excel.py

# Custom input and output
python qualtrics_internal_to_excel.py -i data/survey.xlsx -o results/processed.xlsx -v

"""

#### Externo

`-i/--input`: Especificar entrada del excel. Excel que utilizaremos
`-o/--output`: Especificar la salida del excel
`-v/--verbose`: Ver progreso

"""
# Use default files
python qualtrics_external_to_excel.py

# Custom input and output
python qualtrics_external_to_excel.py -i data/survey.xlsx -o results/processed.xlsx -v

"""

### Base de datos a dashboard

Ejecuta el programa `excel_to_dashboard.py` utilizando la siguiente estructura `python excel_to_dashboard.py -i "base_datos.xlsx" -o "nombre_output.xlsx"`






