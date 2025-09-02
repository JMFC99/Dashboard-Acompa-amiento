import pandas as pd
import argparse
import sys
from pathlib import Path

# ================== CONFIGURATION ==================
DEFAULT_INPUT_FILE = 'ToDocument/Conociendo+el+talento+de+los+coaches_29+de+agosto+de+2025_13.58.xlsx'
DEFAULT_OUTPUT_FILE = "ToDocument/outputfile.xlsx"
DEFAULT_EXISTING_COACHES_FILE = None  # Path to existing coaches list

# Coach type configuration
COACH_TYPE = 'internal'  # 'internal' or 'external'

# ================== COACH TYPE SPECIFIC CONFIGURATIONS ==================
COACH_CONFIGS = {
    'internal': {
        'email_column': "Email\n(Asegurate que el formato de correo sea correcto con @ y un dominio válido, no agregues espacios antes o después del texto)",
        'basic_elements': {
            "Nombre(s):":"Nombre(s)",
            "Apellido Paterno:": 'Apellido Paterno:',
            'Apellido Materno:':'Apellido Materno',
            'Fecha de nacimiento (dd/mm/yyyy)':'Fecha nacimiento',
            "Email\n(Asegurate que el formato de correo sea correcto con @ y un dominio válido, no agregues espacios antes o después del texto)":'Email',
            'País de Residencia:':'País',
            'Estado de Residencia:':'Estado',
            'Celular (+lada) [número telefónico]\n:':'Celular',
            'Género: - Selected Choice':'Género',
            'Respecto al coaching:':'Respecto al coaching',
            '¿Cuántas horas de práctica tienes en coaching? Por favor asegúrate de que marcas las horas reales en las sesiones de coaching (grupal, individual, de equipos),  excluyendo consultoría, mentoría y/o otras prácticas.\n\nFavor de no utilizar comas ni texto, solo incluye las horas estimadas\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000 )':'Horas de práctica',
            '¿Cuántos años de experiencia ACTIVA tienes como coach?\n\nFavor de no utilizar comas ni texto, solo incluye los años estimados\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000 )':'Años de experiencia',
            '¿Cuándo fue la última vez que atendiste un cliente con una sesión de coaching?':'Última vez que atendió cliente',
            '¿Recibes supervisión en tu práctica como coach?':'Recibe supervisión en práctica como coach',
            '¿Has recibido coaching? es decir, ¿has sido coachee?':'Ha sido coachee',
            '¿Cuentas con alguna certificación de coaching?':'Cuenta con certificación',
            '¿Cuándo fue la última vez que recibiste capacitación en temas de coaching?':'Última capacitación en coaching',
            '¿Con cuántas horas de formación en coaching cuentas?\n\nFavor de no utilizar comas ni texto, solo incluye las horas estimadas\n\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000)':'Horas de formación',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Innermetrix':'Innermetrix',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Leadership Circle Profile (LCP)':'LCP',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - PDA':'PDA',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Eneagrama':'Eneagrama',
            'Tipo de contratación que cuentas en el Tec - Selected Choice':"Tipo de contratación",
            "Menciona las empresas en las que has impartido servicios de coaching":"Empresas que impartió coaching",
            "¿Consideras que puedes brindar coaching en inglés?":"¿Puedes brindar coaching en inglés?",
            'CURP: (en caso de ser extranjero, incluye tu número de identidad nacional)':"CURP / DNI"
        },
        'certifications': {
            'ICF': {
                'columns': {
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICF Associate Certified Coach (ACC)  ':'ACC',
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICF Professional Certified Coach (PCC)  ':'PCC',
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICF Master Certified Coach (MCC) ':'MCC',
                },
                'expected_values': {
                    'ACC': 'ICF Associate Certified Coach (ACC)  ',
                    'PCC': 'ICF Professional Certified Coach (PCC)  ',
                    'MCC': 'ICF Master Certified Coach (MCC) '
                },
                'final_column': 'Certificación ICF'
            },
            'EMCC': {
                'columns': {
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - EMCC Foundation  ':'Foundation',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - EMCC Practitioner ':'Practitioner',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - EMCC Senior Practitioner  ':'Senior Practitioner',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - EMCC Master Practitioner  ':'Master Practitioner'
                },
                'expected_values': {
                    'Foundation': 'EMCC Foundation  ',
                    'Practitioner': 'EMCC Practitioner ',
                    'Senior Practitioner': 'EMCC Senior Practitioner  ',
                    'Master Practitioner': 'EMCC Master Practitioner  '
                },
                'final_column': 'Certificación EMCC'
            },
            'ICC': {
                'columns': {
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICC Certificación Internacional Coaching (CIC)  ':'ICC Certificación Internacional Coaching (CIC)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICC Coaching Equipos (CCEQ)  ':'ICC Coaching Equipos (CCEQ)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICC Coaching Negocios (CCN)  ':'ICC Coaching Negocios (CCN)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICC Coaching Ejecutivo (CCEJ)  ':'ICC Coaching Ejecutivo (CCEJ)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida (puedes seleccionar más de una opción) - Selected Choice - ICC Coaching Vida (VIDA)   ':'ICC Coaching Vida (VIDA)',
                },
                'expected_values': {
                    'ICC Certificación Internacional Coaching (CIC)': 'ICC Certificación Internacional Coaching (CIC)  ',
                    'ICC Coaching Equipos (CCEQ)': 'ICC Coaching Equipos (CCEQ)  ',
                    'ICC Coaching Negocios (CCN)': 'ICC Coaching Negocios (CCN)  ',
                    'ICC Coaching Ejecutivo (CCEJ)': 'ICC Coaching Ejecutivo (CCEJ)  ',
                    'ICC Coaching Vida (VIDA)': 'ICC Coaching Vida (VIDA)   '
                },
                'final_column': 'Certificación ICC'
            },
            'WABC': {
                'columns': {
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Registered Corporate Coach (RCC)  ':'WABC Registered Corporate Coach (RCC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Business Coach (CBC)  ':'WABC Certified Business Coach (CBC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Master Business Coach (CMBC)  ':'WABC Certified Master Business Coach (CMBC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Chartered Business Coach (ChBC)  ':'WABC Chartered Business Coach (ChBC)',
                },
                'expected_values': {
                    'WABC Registered Corporate Coach (RCC)': 'WABC Registered Corporate Coach (RCC)  ',
                    'WABC Certified Business Coach (CBC)': 'WABC Certified Business Coach (CBC)  ',
                    'WABC Certified Master Business Coach (CMBC)': 'WABC Certified Master Business Coach (CMBC)  ',
                    'WABC Chartered Business Coach (ChBC)': 'WABC Chartered Business Coach (ChBC)  '
                },
                'final_column': 'Certificación WABC'
            }
        }
    },
    'external': {
        'email_column': 'Email\r\n(Asegurate que el formato de correo sea correcto con @ y un dominio válido, no agregues espacios antes o después del texto)',
        'basic_elements': {
            "Nombre(s):":"Nombre(s)",
            "Apellido Paterno:": 'Apellido Paterno:',
            'Apellido Materno:':'Apellido Materno',
            'Fecha de nacimiento (dd/mm/yyyy)':'Fecha nacimiento',
            'Email\r\n(Asegurate que el formato de correo sea correcto con @ y un dominio válido, no agregues espacios antes o después del texto)':'Email',
            'País de Residencia:':'País',
            'Estado de Residencia:':'Estado',
            'Celular (+lada) [número telefónico]\n:':'Celular',
            'Género: - Selected Choice':'Género',
            'Respecto al coaching:':'Respecto al coaching',
            '¿Cuántas horas de práctica tienes en coaching? Por favor asegúrate de que marcas las horas reales en las sesiones de coaching (grupal, individual, de equipos),  excluyendo consultoría, mentoría y/o otras prácticas.\n\nFavor de no utilizar comas ni texto, solo incluye las horas estimadas\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000 )':'Horas de práctica',
            '¿Cuántos años de experiencia ACTIVA tienes como coach?\n\nFavor de no utilizar comas ni texto, solo incluye los años estimados\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000 )':'Años de experiencia',
            '¿Cuándo fue la última vez que tuviste una sesión de coaching?':'Última vez que atendió cliente',
            '¿Recibes supervisión en tu práctica como coach?':'Recibe supervisión en práctica como coach',
            '¿Has recibido coaching? es decir, ¿has sido coachee?':'Ha sido coachee',
            '¿Cuentas con alguna certificación de coaching?':'Cuenta con certificación',
            '¿Cuándo fue la última vez que recibiste capacitación en temas de coaching?':'Última capacitación en coaching',
            '¿Con cuántas horas de formación en coaching cuentas?\n\nFavor de no utilizar comas ni texto, solo incluye las horas estimadas\n(Correcto: 1500 | 30 | 2000 ; Formato incorrecto 1,500 | 30 horas | + de 1000)':'Horas de formación',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Innermetrix':'Innermetrix',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Leadership Circle Profile (LCP)':'LCP',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - PDA':'PDA',
            'Selecciona el tipo de certificación con la que cuentas (puedes incluir más de uno) - Selected Choice - Eneagrama':'Eneagrama',
            'CURP: (en caso de ser extranjero, incluye tu número de identidad nacional)':"CURP / DNI",
            "¿Te encuentras trabajando en proyectos de coaching en el TEC?":"¿Te encuentras trabajando en proyectos de coaching en el TEC?",
            'Compártenos los proyectos en los que estás trabajando o en los que estás activo actualmente en el TEC.':'Compártenos los proyectos en los que estás trabajando o en los que estás activo actualmente en el TEC.'
        },
        'certifications': {
            'ICF': {
                'columns': {
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Associate Certified Coach (ACC)  ':'ACC',
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Professional Certified Coach (PCC)  ':'PCC',
                    'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Master Certified Coach (MCC) ':'MCC',
                },
                'expected_values': {
                    'ACC': 'ICF Associate Certified Coach (ACC)  ',
                    'PCC': 'ICF Professional Certified Coach (PCC)  ',
                    'MCC': 'ICF Master Certified Coach (MCC) '
                },
                'final_column': 'Certificación ICF'
            },
            'EMCC': {
                'columns': {
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Foundation  ':'Foundation',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Practitioner ':'Practitioner',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Senior Practitioner  ':'Senior Practitioner',
                    'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Master Practitioner  ':'Master Practitioner'
                },
                'expected_values': {
                    'Foundation': 'EMCC Foundation  ',
                    'Practitioner': 'EMCC Practitioner ',
                    'Senior Practitioner': 'EMCC Senior Practitioner  ',
                    'Master Practitioner': 'EMCC Master Practitioner  '
                },
                'final_column': 'Certificación EMCC'
            },
            'ICC': {
                'columns': {
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Certificación Internacional Coaching (CIC)  ':'ICC Certificación Internacional Coaching (CIC)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Equipos (CCEQ)  ':'ICC Coaching Equipos (CCEQ)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Negocios (CCN)  ':'ICC Coaching Negocios (CCN)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Ejecutivo (CCEJ)  ':'ICC Coaching Ejecutivo (CCEJ)',
                    'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Vida (VIDA)   ':'ICC Coaching Vida (VIDA)',
                },
                'expected_values': {
                    'ICC Certificación Internacional Coaching (CIC)': 'ICC Certificación Internacional Coaching (CIC)  ',
                    'ICC Coaching Equipos (CCEQ)': 'ICC Coaching Equipos (CCEQ)  ',
                    'ICC Coaching Negocios (CCN)': 'ICC Coaching Negocios (CCN)  ',
                    'ICC Coaching Ejecutivo (CCEJ)': 'ICC Coaching Ejecutivo (CCEJ)  ',
                    'ICC Coaching Vida (VIDA)': 'ICC Coaching Vida (VIDA)   '
                },
                'final_column': 'Certificación ICC'
            },
            'WABC': {
                'columns': {
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Registered Corporate Coach (RCC)  ':'WABC Registered Corporate Coach (RCC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Business Coach (CBC)  ':'WABC Certified Business Coach (CBC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Master Business Coach (CMBC)  ':'WABC Certified Master Business Coach (CMBC)',
                    'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Chartered Business Coach (ChBC)  ':'WABC Chartered Business Coach (ChBC)',
                },
                'expected_values': {
                    'WABC Registered Corporate Coach (RCC)': 'WABC Registered Corporate Coach (RCC)  ',
                    'WABC Certified Business Coach (CBC)': 'WABC Certified Business Coach (CBC)  ',
                    'WABC Certified Master Business Coach (CMBC)': 'WABC Certified Master Business Coach (CMBC)  ',
                    'WABC Chartered Business Coach (ChBC)': 'WABC Chartered Business Coach (ChBC)  '
                },
                'final_column': 'Certificación WABC'
            }
        }
    }
}

# ================== STATIC MAPPINGS (SHARED BETWEEN INTERNAL AND EXTERNAL) ==================
CATEGORY_MAPPINGS = {
    'tipo_coaching': {
        'columns': {
            'Marca los tipos de coaching en los que te sientes con amplia experiencia - Selected Choice - Coaching personal ':"Coaching personal",
            'Marca los tipos de coaching en los que te sientes con amplia experiencia - Selected Choice - Coaching de equipos ': 'Coaching equipos',
            'Marca los tipos de coaching en los que te sientes con amplia experiencia - Selected Choice - Coaching grupal ':'Coaching grupal',
            'Marca los tipos de coaching en los que te sientes con amplia experiencia - Selected Choice - Coaching de bienestar ':'Coaching bienestar',
            'Marca los tipos de coaching en los que te sientes con amplia experiencia - Selected Choice - Coaching ejecutivo':"Coaching Ejecutivo",
        },
        'expected_values': {
            'Coaching personal': 'Coaching personal ',
            'Coaching equipos': 'Coaching de equipos ',
            'Coaching grupal': 'Coaching grupal ',
            'Coaching bienestar': 'Coaching de bienestar ',
            'Coaching Ejecutivo': 'Coaching ejecutivo'
        },
        'final_column': 'Tipo de coaching'
    },
    'tipo_clientes': {
        'columns': {
            'Selecciona tipo de clientes que has brindado coaching: - Selected Choice - Personal (coaching de vida, de bienestar, etc)' :'Personal',
            'Selecciona tipo de clientes que has brindado coaching: - Selected Choice - Organizacional (contexto de empresas, ONG, Universidades)':'Organizaciones'
        },
        'expected_values': {
            'Personal': 'Personal (coaching de vida, de bienestar, etc)',
            'Organizaciones': 'Organizacional (contexto de empresas, ONG, Universidades)'
        },
        'final_column': 'Tipo de cliente que ha atendido'
    },
    'perfiles_clientes': {
        'columns': {
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Propietarios de Negocios y Emprendedores ':"Propietarios de Negocios y Emprendedores",
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Ejecutivos de Alto Nivel (C-Suite): CEO/CIO/CFO ':'Ejecutivos alto nivel C-Suite:CEO/CFO/CTO',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Vicepresidentes: VSP/EVP':'Vicepresidentes(VSP/EVP)',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Gerentes de Departamento/Directores  ':'Gerentes de departamento / Directores',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Empleados de Alto Potencial ':'Empleados de Alto potencial',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Nuevos Empleados ': 'Nuevos empleados',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Equipos y Grupos ':'Equipos y Grupos',
            '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Nuevos Líderes ': 'Nuevos líderes'
        },
        'expected_values': {
            'Propietarios de Negocios y Emprendedores': 'Propietarios de Negocios y Emprendedores ',
            'Ejecutivos alto nivel C-Suite:CEO/CFO/CTO': 'Ejecutivos de Alto Nivel (C-Suite): CEO/CIO/CFO ',
            'Vicepresidentes(VSP/EVP)': 'Vicepresidentes: VSP/EVP',
            'Gerentes de departamento / Directores': 'Gerentes de Departamento/Directores  ',
            'Empleados de Alto potencial': 'Empleados de Alto Potencial ',
            'Equipos y Grupos': 'Equipos y Grupos ',
            'Nuevos empleados': 'Nuevos Empleados ',
            'Nuevos líderes': 'Nuevos Líderes '
        },
        'final_column': 'Perfil de clientes'
    },
    'tipo_industria': {
        'columns': {
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Comunicaciones, Entretenimiento y Medios ':"Comunicaciones, Entretenimiento y Medios",
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Educación ': 'Educación',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Energía y Servicios Públicos ':'Energía y Servicios Públicos',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Gobierno y Sector Público ':'Gobierno y Sector Público',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Salud, Farmacéutica y Ciencia ':'Salud, Farmacéutica y Ciencia',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Hospitalidad y Ocio ':'Hospitalidad y Ocio',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Manufactura, Ingeniería y Construcción ':'Manufactura, Ingeniería y Construcción',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Servicios Profesionales y Financieros ':'Servicios Profesionales y Financieros',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Retail y Consumo ':'Retail y Consumo',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Tecnología ':'Tecnología',
            'Selecciona la(s) industria(s) que con la(s) que tienes experiencia atendiendo clientes: - Selected Choice - Transporte ':"Transporte"
        },
        'expected_values': {
            'Comunicaciones, Entretenimiento y Medios': 'Comunicaciones, Entretenimiento y Medios ',
            'Educación': 'Educación ',
            'Energía y Servicios Públicos': 'Energía y Servicios Públicos ',
            'Gobierno y Sector Público': 'Gobierno y Sector Público ',
            'Salud, Farmacéutica y Ciencia': 'Salud, Farmacéutica y Ciencia ',
            'Hospitalidad y Ocio': 'Hospitalidad y Ocio ',
            'Manufactura, Ingeniería y Construcción': 'Manufactura, Ingeniería y Construcción ',
            'Servicios Profesionales y Financieros': 'Servicios Profesionales y Financieros ',
            'Retail y Consumo': 'Retail y Consumo ',
            'Tecnología': 'Tecnología ',
            'Transporte': 'Transporte '
        },
        'final_column': 'Tipo industrias cliente'
    }
}

CERTIFICATION_MAPPINGS = {
    'ICF': {
        'columns': {
            'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Associate Certified Coach (ACC)  ':'ACC',
            'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Professional Certified Coach (PCC)  ':'PCC',
            'Si tienes o has tenido credencial de ICF selecciona la que coincida - Selected Choice - ICF Master Certified Coach (MCC) ':'MCC',
        },
        'expected_values': {
            'ACC': 'ICF Associate Certified Coach (ACC)  ',
            'PCC': 'ICF Professional Certified Coach (PCC)  ',
            'MCC': 'ICF Master Certified Coach (MCC) '
        },
        'final_column': 'Certificación ICF'
    },
    'EMCC': {
        'columns': {
            'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Foundation  ':'Foundation',
            'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Practitioner ':'Practitioner',
            'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Senior Practitioner  ':'Senior Practitioner',
            'Si tienes o has tenido credencial de EMCC selecciona la que coincida - Selected Choice - EMCC Master Practitioner  ':'Master Practitioner'
        },
        'expected_values': {
            'Foundation': 'EMCC Foundation  ',
            'Practitioner': 'EMCC Practitioner ',
            'Senior Practitioner': 'EMCC Senior Practitioner  ',
            'Master Practitioner': 'EMCC Master Practitioner  '
        },
        'final_column': 'Certificación EMCC'
    },
    'ICC': {
        'columns': {
            'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Certificación Internacional Coaching (CIC)  ':'ICC Certificación Internacional Coaching (CIC)',
            'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Equipos (CCEQ)  ':'ICC Coaching Equipos (CCEQ)',
            'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Negocios (CCN)  ':'ICC Coaching Negocios (CCN)',
            'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Ejecutivo (CCEJ)  ':'ICC Coaching Ejecutivo (CCEJ)',
            'Si tienes o has tenido credencial de ICC selecciona la que coincida - Selected Choice - ICC Coaching Vida (VIDA)   ':'ICC Coaching Vida (VIDA)',
        },
        'expected_values': {
            'ICC Certificación Internacional Coaching (CIC)': 'ICC Certificación Internacional Coaching (CIC)  ',
            'ICC Coaching Equipos (CCEQ)': 'ICC Coaching Equipos (CCEQ)  ',
            'ICC Coaching Negocios (CCN)': 'ICC Coaching Negocios (CCN)  ',
            'ICC Coaching Ejecutivo (CCEJ)': 'ICC Coaching Ejecutivo (CCEJ)  ',
            'ICC Coaching Vida (VIDA)': 'ICC Coaching Vida (VIDA)   '
        },
        'final_column': 'Certificación ICC'
    },
    'WABC': {
        'columns': {
            'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Registered Corporate Coach (RCC)  ':'WABC Registered Corporate Coach (RCC)',
            'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Business Coach (CBC)  ':'WABC Certified Business Coach (CBC)',
            'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Certified Master Business Coach (CMBC)  ':'WABC Certified Master Business Coach (CMBC)',
            'Si tienes o has tenido credencial de WABC selecciona la que coincida - Selected Choice - WABC Chartered Business Coach (ChBC)  ':'WABC Chartered Business Coach (ChBC)',
        },
        'expected_values': {
            'WABC Registered Corporate Coach (RCC)': 'WABC Registered Corporate Coach (RCC)  ',
            'WABC Certified Business Coach (CBC)': 'WABC Certified Business Coach (CBC)  ',
            'WABC Certified Master Business Coach (CMBC)': 'WABC Certified Master Business Coach (CMBC)  ',
            'WABC Chartered Business Coach (ChBC)': 'WABC Chartered Business Coach (ChBC)  '
        },
        'final_column': 'Certificación WABC'
    }
}

# ================== HELPER FUNCTIONS ==================
def load_existing_coaches(file_path, verbose=False):
    """
    Load existing coaches list from Excel file
    
    Args:
        file_path (str): Path to existing coaches Excel file
        verbose (bool): Print verbose output
    
    Returns:
        set: Set of existing coach emails (normalized)
    """
    if not file_path or not Path(file_path).exists():
        if verbose:
            print("No existing coaches file provided or file not found. Processing all coaches.")
        return set()
    
    try:
        # Try to read from different possible sheets
        possible_sheets = ['Basic information', 'Basic Info', 'Sheet1', 0]
        existing_df = None
        
        for sheet in possible_sheets:
            try:
                existing_df = pd.read_excel(file_path, sheet_name=sheet)
                if 'Email' in existing_df.columns:
                    break
            except:
                continue
        
        if existing_df is None or 'Email' not in existing_df.columns:
            if verbose:
                print(f"Warning: Could not find 'Email' column in {file_path}")
            return set()
        
        # Normalize emails (lowercase, strip whitespace)
        existing_emails = set(
            email.lower().strip() 
            for email in existing_df['Email'].dropna().astype(str)
            if email.lower().strip() != ''
        )
        
        if verbose:
            print(f"Loaded {len(existing_emails)} existing coaches from {file_path}")
        
        return existing_emails
        
    except Exception as e:
        if verbose:
            print(f"Error loading existing coaches file: {str(e)}")
        return set()

def filter_new_coaches(df, existing_emails, email_column='Email', verbose=False):
    """
    Filter dataframe to only include coaches not in existing list
    
    Args:
        df (pandas.DataFrame): Input dataframe
        existing_emails (set): Set of existing coach emails
        email_column (str): Name of email column
        verbose (bool): Print verbose output
    
    Returns:
        tuple: (filtered_df, filtered_count, existing_count)
    """
    if email_column not in df.columns:
        if verbose:
            print(f"Warning: Email column '{email_column}' not found in dataframe")
        return df, len(df), 0
    
    # Normalize emails in dataframe
    df_emails = df[email_column].astype(str).str.lower().str.strip()
    
    # Create mask for new coaches (not in existing list)
    mask = ~df_emails.isin(existing_emails)
    
    # Filter dataframe
    filtered_df = df[mask].reset_index(drop=True)
    
    filtered_count = len(filtered_df)
    existing_count = len(df) - filtered_count
    
    if verbose:
        print(f"Filtering results:")
        print(f"  - Total coaches in input: {len(df)}")
        print(f"  - Already exist (skipped): {existing_count}")
        print(f"  - New coaches to process: {filtered_count}")
    
    return filtered_df, filtered_count, existing_count

def filter_country(element):
    """Standardize country names"""
    mexico_variants = ['México','MEXICO', 'México ', 'Méxcio', 'MÉXICO','México y USA','M','CDMX',
                      'MEXICO / ESPAÑA','Mexico ','Mexico']
    usa_variants = ['Estados unidos ','USA']
    colombia_variants = ['Colombia','COLOMBIA']
    ecuador_variants = ['Ecuador y Colombia']
    costa_rica_variants = ['Costa Rica', 'Costa Rica ']
    
    if element in mexico_variants:
        return 'México'
    elif element in usa_variants:
        return 'Estados Unidos'
    elif element in colombia_variants:
        return 'Colombia'
    elif element in ecuador_variants:
        return 'Ecuador'
    elif element in costa_rica_variants:
        return 'Costa Rica'
    elif pd.isna(element):
        return 'México'
    else:
        return element

def homologate_states(state):
    """Standardize Mexican state names"""
    if pd.isna(state):
        return 'No Especificado'
    
    state = str(state).strip().title()
    
    state_mapping = {
        # Ciudad de México
        'Cdmx': 'Ciudad de México', 'Cdmx ': 'Ciudad de México', 'Cdmx/Bcs': 'Ciudad de México',
        'Cdmx / Queretaro': 'Ciudad de México', 'Ciudad De México': 'Ciudad de México',
        'Ciudad De Mexico': 'Ciudad de México', 'Ciudad De Mexico ': 'Ciudad de México',
        'Df': 'Ciudad de México', "Permanente":'Ciudad de México', 'Cdmx / Querétaro': 'Ciudad de México',
        
        # Estado de México
        'Estado De México': 'Estado de México', 'Estado De Mexico': 'Estado de México',
        'Edomex': 'Estado de México', 'Edo. Méxic': 'Estado de México',
        'Eso. De Mexico ': 'Estado de México', 'Estado De México ': 'Estado de México',
        'Estado De Mexico ': 'Estado de México', 'Estado De México - Toluca Y Cdmx': 'Estado de México',
        'Estado De México (Área Metropolitana)': 'Estado de México', 'Área Metropolitana': 'N/A',
        'Toluca ': 'Estado de México', 'Eso. De Mexico': 'Estado de México',
        'Mexico': 'Estado de México', 'México':'Estado de México',
        
        # Nuevo León
        'Nuevo Leon': 'Nuevo León', 'Nl': 'Nuevo León', 'Monterrey': 'Nuevo León',
        'Nuevo León ': 'Nuevo León',
        
        # Querétaro
        'Queretaro': 'Querétaro', 'Queretaro ': 'Querétaro',
        
        # Multi-location cases
        'Estado De México Y Carolina Del Norte': 'Estado de México',
        'Puebla Y Querétaro': 'Puebla', 'Morelos, Cuernavaca': 'Morelos',
        'Morelos (Cuernavaca)': 'Morelos', 'Quintana Roo/Cdmx': 'Quintana Roo',
        'Ciudad De México / Cataluña': 'Ciudad de México', 'Coahuila - Saltillo':'Coahuila',
        
        # Other states
        'San Luis Potosí ': 'San Luis Potosí', 'Hidalgo ': 'Hidalgo',
        'Jalisco ': 'Jalisco', 'Estado De México ': 'Estado de México',
    }
    
    # Add uppercase variations
    uppercase_mapping = {k.upper(): v for k, v in state_mapping.items()}
    state_mapping.update(uppercase_mapping)
    
    return state_mapping.get(state, state)

def transform_columns(df, id_column='Email', transform_columns=None, 
                     column_type=bool, new_column_name='Tipo de Transformación'):
    """Transform boolean columns into rows"""
    # Ensure Email column exists
    if id_column not in df.columns:
        print(f"  Warning: ID column '{id_column}' not found in dataframe")
        return pd.DataFrame(columns=[id_column, new_column_name])
    
    if transform_columns is None:
        transform_columns = [col for col in df.columns 
                           if df[col].dtype == column_type and col != id_column]
    
    transform_columns = [col for col in transform_columns if col != id_column and col in df.columns]
    
    if not transform_columns:
        print(f"  Warning: No columns to transform")
        return pd.DataFrame(columns=[id_column, new_column_name])
    
    # Melt DataFrame
    transformed_df = df.melt(
        id_vars=[id_column],
        value_vars=transform_columns,
        var_name='Columna Original',
        value_name='Valor'
    )
    
    # Filter True values
    if column_type == bool:
        transformed_df = transformed_df[transformed_df['Valor'] == True]
    else:
        transformed_df = transformed_df[transformed_df['Valor'].notna()]
    
    # Clean column names
    transformed_df[new_column_name] = (transformed_df['Columna Original']
                                       .str.replace('Coaching ', '')
                                       .str.strip())
    
    return transformed_df[[id_column, new_column_name]]

def process_category(fulltable, category_config):
    """Process a category of columns"""
    # Check which columns exist
    required_cols = list(category_config['columns'].keys())
    available_cols = ['Email'] + [col for col in required_cols if col in fulltable.columns]
    
    if len(available_cols) == 1:  # Only Email column found
        print(f"  Warning: No columns found for this category")
        return pd.DataFrame({'Email': fulltable['Email'].unique()})
    
    # Extract and rename columns
    df = fulltable[available_cols].rename(
        columns={col: category_config['columns'][col] for col in available_cols if col != 'Email'}
    ).reset_index(drop=True)
    
    # Convert to boolean based on expected values
    for col, expected_val in category_config['expected_values'].items():
        if col in df.columns:
            df[col] = df[col] == expected_val
    
    return df

def process_certification(fulltable, cert_config):
    """Process certification columns"""
    return process_category(fulltable, cert_config)

def process_external_perfiles_clientes(fulltable, category_config):
    """Special processing for external coaches client profiles - handles combined categories"""
    # Get the base dataframe
    df = process_category(fulltable, category_config)
    
    # Check if we have the separate Director/Manager columns that need combining
    directores_col = '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Directores de Departamento'
    gerencia_col = '¿Con cuál o cuáles perfiles de coachees tienes experiencia atendiendo? - Selected Choice - Gerencia Media'
    
    if directores_col in fulltable.columns and gerencia_col in fulltable.columns:
        if 'Directores de Departamento' not in df.columns:
            df['Directores de Departamento'] = fulltable[directores_col] == 'Directores de Departamento'
        if 'Gerencia Media' not in df.columns:
            df['Gerencia Media'] = fulltable[gerencia_col] == 'Gerencia Media'
        
        # Combine into "Gerentes de departamento / Directores"
        df['Gerentes de departamento / Directores'] = (
            (df.get('Directores de Departamento', False) == True) | 
            (df.get('Gerencia Media', False) == True)
        )
    
    return df

# ================== MAIN PROCESSING ==================
def process_coaches_data(input_file, output_file, existing_coaches_file=None, coach_type='internal', verbose=False):
    """
    Main processing function for coaches data with filtering
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
        existing_coaches_file (str): Path to existing coaches Excel file
        coach_type (str): 'internal' or 'external'
        verbose (bool): Print verbose output
    
    Returns:
        tuple: (coaches_data, all_dataframes, transformed_tables, filter_stats)
    """
    if verbose:
        print(f"Processing {coach_type} coaches data from: {input_file}")
    
    # Get configuration for coach type
    if coach_type not in COACH_CONFIGS:
        raise ValueError(f"Invalid coach type: {coach_type}. Must be 'internal' or 'external'")
    
    config = COACH_CONFIGS[coach_type]
    basic_elements = config['basic_elements']
    email_column_original = config['email_column']
    
    # Load existing coaches for filtering
    existing_emails = load_existing_coaches(existing_coaches_file, verbose)
    
    # Load data
    coaches_fulltable = pd.read_excel(input_file, skiprows=1)
    coaches_fulltable = coaches_fulltable[coaches_fulltable['Finalizado'] == True]
    
    if verbose:
        print(f"Loaded {len(coaches_fulltable)} completed records")
    
    # Rename email column for consistency
    coaches_fulltable.rename(columns={email_column_original: "Email"}, inplace=True)
    
    # Filter out existing coaches BEFORE processing
    coaches_fulltable, new_count, existing_count = filter_new_coaches(
        coaches_fulltable, existing_emails, 'Email', verbose
    )
    
    filter_stats = {
        'new_coaches': new_count,
        'existing_coaches': existing_count,
        'total_input': new_count + existing_count
    }
    
    # If no new coaches to process, return empty results
    if new_count == 0:
        print("✓ No new coaches to process. All coaches already exist in the database.")
        empty_df = pd.DataFrame()
        return empty_df, {}, [], filter_stats
    
    all_dataframes = {}
    transformed_tables = []
    
    # Process basic information
    if verbose:
        print("Processing basic information...")
    
    # Check which columns actually exist in the dataframe
    available_columns = []
    missing_columns = []
    for col in basic_elements.keys():
        if col in coaches_fulltable.columns:
            available_columns.append(col)
        else:
            missing_columns.append(col)
    
    if missing_columns:
        if verbose:
            print(f"  Warning: Missing columns: {missing_columns}")
        
        # Try to find the Email column with a different format
        email_columns = [col for col in coaches_fulltable.columns if 'Email' in col or 'email' in col]
        if email_columns and 'Email' not in coaches_fulltable.columns:
            available_columns.append('Email')  # We already renamed it above
    
    if not available_columns:
        raise ValueError("No valid columns found in the input file. Please check the file format.")
    
    # Only select columns that exist
    existing_basic_columns = [col for col in basic_elements.keys() if col in coaches_fulltable.columns]
    if 'Email' not in existing_basic_columns:
        existing_basic_columns.append('Email')  # Add Email since we renamed it
    
    coaches_data = coaches_fulltable[['Email'] + [col for col in existing_basic_columns if col != 'Email']].reset_index(drop=True)
    
    # Apply country and state standardization if columns exist
    if 'País de Residencia:' in coaches_data.columns:
        coaches_data['País de Residencia:'] = coaches_data['País de Residencia:'].apply(filter_country)
    if 'Estado de Residencia:' in coaches_data.columns:
        coaches_data['Estado de Residencia:'] = coaches_data['Estado de Residencia:'].apply(homologate_states)
    
    # Rename columns - only rename those that exist
    rename_dict = {}
    for old_col in coaches_data.columns:
        if old_col in basic_elements:
            rename_dict[old_col] = basic_elements[old_col]
        elif old_col == 'Email':
            rename_dict[old_col] = 'Email'  # Keep as is
    
    coaches_data.rename(columns=rename_dict, inplace=True)
    all_dataframes['Basic information'] = coaches_data
    
    # Process categories
    if verbose:
        print("Processing coaching categories...")
    
    for category_name, category_config in CATEGORY_MAPPINGS.items():
        if category_name == 'perfiles_clientes' and coach_type == 'external':
            # Special handling for external coaches client profiles
            df = process_external_perfiles_clientes(coaches_fulltable, category_config)
        else:
            df = process_category(coaches_fulltable, category_config)
        
        all_dataframes[f'{category_name.replace("_", " ")}'] = df
        
        # Get valid columns for transformation
        valid_columns = [col for col in category_config['columns'].values() if col in df.columns]
        
        if valid_columns:
            # Transform to long format
            transformed_df = transform_columns(
                df=df,
                id_column='Email',
                transform_columns=valid_columns,
                new_column_name=category_config['final_column']
            )
            transformed_tables.append(transformed_df)
            
            if verbose:
                print(f"  - Processed {category_name}: {len(transformed_df)} records")
        else:
            if verbose:
                print(f"  - Skipped {category_name}: No valid columns found")
    
    # Process certifications
    if verbose:
        print("Processing certifications...")
    for cert_type, cert_config in config['certifications'].items():
        df = process_certification(coaches_fulltable, cert_config)
        all_dataframes[f'{cert_type} certificacion'] = df
        
        # Get valid columns for transformation
        valid_columns = [col for col in cert_config['columns'].values() if col in df.columns]
        
        if valid_columns:
            # Transform to long format
            transformed_df = transform_columns(
                df=df,
                id_column='Email',
                transform_columns=valid_columns,
                new_column_name=cert_config['final_column']
            )
            transformed_tables.append(transformed_df)
            
            if verbose:
                print(f"  - Processed {cert_type}: {len(transformed_df)} records")
        else:
            if verbose:
                print(f"  - Skipped {cert_type}: No valid columns found")
    
    # Save to Excel
    if verbose:
        print(f"\nSaving results to: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Save basic info
        coaches_data.to_excel(writer, sheet_name='Basic information', index=False)
        
        # Save all category dataframes
        for sheet_name, df in all_dataframes.items():
            if sheet_name != 'Basic information':  # Already saved above
                # Ensure sheet name is valid (Excel has limitations)
                safe_sheet_name = sheet_name[:31]  # Excel max sheet name length
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    
    print(f"\n✓ Processing complete. Output saved to {output_file}")
    print(f"✓ New coaches processed: {new_count}")
    print(f"✓ Coaches skipped (already exist): {existing_count}")
    print(f"✓ Total sheets created: {len(all_dataframes)}")
    print(f"✓ Processed {len(transformed_tables)} transformed tables")
    
    return coaches_data, all_dataframes, transformed_tables, filter_stats

def setup_argparse():
    """Setup command line argument parser with filtering support"""
    parser = argparse.ArgumentParser(
        description='Process coaches data from Qualtrics survey export with duplicate filtering',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process internal coaches with filtering
  python %(prog)s --type internal --existing existing_coaches.xlsx
  
  # Process external coaches without filtering
  python %(prog)s --type external -i external_survey.xlsx -o external_output.xlsx
  
  # Process with verbose output and force overwrite
  python %(prog)s --type internal --existing coaches_list.xlsx -v --force
  
  # Check files only
  python %(prog)s --check-only --existing coaches_list.xlsx
        """
    )
    
    parser.add_argument(
        '-i', '--input',
        type=str,
        default=DEFAULT_INPUT_FILE,
        help=f'Input Excel file path (default: {DEFAULT_INPUT_FILE})'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        default=DEFAULT_OUTPUT_FILE,
        help=f'Output Excel file path (default: {DEFAULT_OUTPUT_FILE})'
    )
    
    parser.add_argument(
        '--existing',
        type=str,
        default=DEFAULT_EXISTING_COACHES_FILE,
        help='Path to existing coaches Excel file for filtering'
    )
    
    parser.add_argument(
        '--type',
        type=str,
        choices=['internal', 'external'],
        default='internal',
        help='Type of coaches to process (default: internal)'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output'
    )
    
    parser.add_argument(
        '--check-only',
        action='store_true',
        help='Only check if input file exists, do not process'
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='Overwrite output file if it exists'
    )
    
    parser.add_argument(
        '--no-filter',
        action='store_true',
        help='Skip filtering and process all coaches (ignore existing coaches file)'
    )
    
    return parser

def validate_files(input_file, output_file, existing_file=None, force=False, verbose=False):
    """
    Validate input, output, and existing coaches files
    
    Args:
        input_file (str): Input file path
        output_file (str): Output file path
        existing_file (str): Existing coaches file path
        force (bool): Force overwrite if output exists
        verbose (bool): Print verbose output
    
    Returns:
        bool: True if validation passes
    """
    # Check input file
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"❌ Error: Input file not found: {input_file}")
        return False
    
    if not input_path.suffix.lower() in ['.xlsx', '.xls']:
        print(f"❌ Error: Input file must be an Excel file (.xlsx or .xls)")
        return False
    
    # Check existing coaches file (optional)
    if existing_file:
        existing_path = Path(existing_file)
        if not existing_path.exists():
            print(f"⚠️  Warning: Existing coaches file not found: {existing_file}")
            print("   Processing will continue without filtering.")
        elif not existing_path.suffix.lower() in ['.xlsx', '.xls']:
            print(f"❌ Error: Existing coaches file must be an Excel file (.xlsx or .xls)")
            return False
        elif verbose:
            print(f"✓ Existing coaches file found: {existing_file}")
    
    # Check output file
    output_path = Path(output_file)
    if output_path.exists() and not force:
        response = input(f"⚠️  Output file already exists: {output_file}\n   Overwrite? (y/n): ")
        if response.lower() != 'y':
            print("Operation cancelled.")
            return False
    
    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    return True

def main():
    """Main entry point with argument parsing and filtering"""
    parser = setup_argparse()
    args = parser.parse_args()
    
    # Check-only mode
    if args.check_only:
        input_path = Path(args.input)
        output_path = Path(args.output)
        existing_path = Path(args.existing) if args.existing else None
        
        print("Coaches Processing File Status Check:")
        print(f"  Input file:  {args.input}")
        print(f"    - Exists: {'✓' if input_path.exists() else '✗'}")
        if input_path.exists():
            print(f"    - Size: {input_path.stat().st_size / 1024:.1f} KB")
        
        if args.existing:
            print(f"  Existing coaches file: {args.existing}")
            print(f"    - Exists: {'✓' if existing_path and existing_path.exists() else '✗'}")
            if existing_path and existing_path.exists():
                print(f"    - Size: {existing_path.stat().st_size / 1024:.1f} KB")
        
        print(f"  Output file: {args.output}")
        print(f"    - Exists: {'✓' if output_path.exists() else '✗'}")
        if output_path.exists():
            print(f"    - Size: {output_path.stat().st_size / 1024:.1f} KB")
        
        print(f"  Coach type: {args.type}")
        print(f"  Filtering: {'Disabled' if args.no_filter else 'Enabled'}")
        
        return 0 if input_path.exists() else 1
    
    # Validate files
    existing_file = None if args.no_filter else args.existing
    if not validate_files(args.input, args.output, existing_file, args.force, args.verbose):
        return 1
    
    try:
        # Process the data
        coaches_data, all_dataframes, tables, filter_stats = process_coaches_data(
            args.input, 
            args.output,
            existing_file,
            args.type,
            args.verbose
        )
        
        # Print summary
        print(f"\n📊 Processing Summary:")
        print(f"   Coach type: {args.type}")
        print(f"   Input records: {filter_stats['total_input']}")
        print(f"   New coaches processed: {filter_stats['new_coaches']}")
        print(f"   Existing coaches skipped: {filter_stats['existing_coaches']}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error during processing: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

# ================== EXECUTION ==================
if __name__ == "__main__":
    sys.exit(main())