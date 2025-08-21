import pandas as pd
import re
import sys

def normalizar_id(id_str):
    """Normaliza el ID eliminando puntos y decimales."""
    if pd.isna(id_str):
        return pd.NA
    return re.sub(r'\D', '', str(id_str))

def calcular_tipo_nps(satisfaccion):
    """Calcula el tipo NPS según la satisfacción general."""
    try:
        val = int(satisfaccion)
    except (ValueError, TypeError):
        return pd.NA
    if val >= 9:
        return "Entusiastas"
    elif val >= 7:
        return "Pasivos"
    elif val >= 0:
        return "Detractores"
    return pd.NA

def similitud_tokens(a, b):
    """Compara similitud de tokens entre dos strings."""
    tokens_a = set(re.findall(r'\w+', str(a).lower()))
    tokens_b = set(re.findall(r'\w+', str(b).lower()))
    return len(tokens_a & tokens_b) / max(len(tokens_a | tokens_b), 1)

# Lista de columnas modelo
modelo_columnas = [
    'Unnamed: 0', 'ID', 'FECHA', 'ESTATUS', 'RESULTADO SATISFACCION', 'SATISFACCION GENERAL', 'SATISFACCION COVID',
    'SATISFACCION CONDICIONES', 'COMENTARIOS', 'BENEFICIOS', 'BENEFICIO 1', 'BENEFICIO 1 INTERES', 'BENEFICIO 1 USO',
    'BENEFICIO 2', 'BENEFICIO 1 INTERES.1', 'BENEFICIO 1 USO.1', 'BENEFICIO 3', 'BENEFICIO 1 INTERES.2', 'BENEFICIO 1 USO.2',
    'BENEFICIO 4', 'BENEFICIO 1 INTERES.3', 'BENEFICIO 1 USO.3', 'BENEFICIO 5', 'BENEFICIO 1 INTERES.4', 'BENEFICIO 1 USO.4',
    'BENEFICIO 6', 'BENEFICIO 1 INTERES.5', 'BENEFICIO 1 USO.5', 'BENEFICIO 7', 'BENEFICIO 1 INTERES.6', 'BENEFICIO 1 USO.6',
    'BENEFICIO 8', 'BENEFICIO 1 INTERES.7', 'BENEFICIO 1 USO.7', 'BENEFICIO 9', 'BENEFICIO 1 INTERES.8', 'BENEFICIO 1 USO.8',
    'PREGUNTAS CLIMA', 'NIVEL DE ORGULLO', 'NIVEL DE RECOMENDACIÓN', 'PERCEPCION MISION INSPIRADORA',
    'APORTE A OBJETIVOS ORGANIZACIONALES', 'BENEFICIOS NO MONETARIOS', 'ACTIVIDADES BIENESTAR', 'BALANCE TRABAJO VIDA PERSONAL',
    'SENSIBILIDAD POR LA VIDA PERSONAL', 'PREGUNTAS CLIMA.1', 'RECURSOS REQUERIDOS', 'ACCESO A INFORMACION',
    'SENSACION DE PROGRESO EN EL CARGO', 'CLARIDAD DE CARGOS Y RESPONSABILIDADES', 'ENTRENAMIENTO PUESTO TRABAJO',
    'FORMACION PARA DESARROLLO PERSONAL Y LABORAL', 'EVALUACION DESEMPEÑO', 'CELEBRACION DE EXITOS', 'PREGUNTAS CLIMA.2',
    'COMUNICACION INTERAREAS', 'COMUNICACION CON LIDER', 'TRABAJO INTERAREAS', 'TRABAJO DENTRO DEL AREA',
    'OPORTUNIDAD EN DECISIONES', 'ANALISIS DE DECISIONES', 'LIDERES EJEMPLO', 'LIDERES MENTORES', 'PREGUNTAS CLIMA.3',
    'RESPETO EN EL TRATO', 'DAR PUNTO DE VISTA', 'RELACIONES DE CONFIANZA', 'AUTONOMIA DEL CARGO', 'TRATO JUSTO Y EQUITATIVO',
    'EVITAR INTIMIDACION Y HOSTIGAMIENTO', 'AMOR Y COMPROMISO POR LA ORGANIZACIÓN', 'GUSTO POR EL TRABAJO', 'PREGUNTAS CLIMA.4',
    'CONDICIONES SEGURAS Y COMODAS', 'MANEJO DEL ESTRÉS', 'CONTROLES DE CALIDAD', 'MEJORA CONTINUA', 'CUIDADO DEL MEDIO AMBIENTE',
    'POLITICAS AMBIENTALES', 'CONTRATACION Y PAGO OPORTUNO', 'IMPACTO Y APORTE A LA COMUNIDAD', 'GENERO', 'ESTADO CIVIL',
    'AÑO NACIMIENTO', 'NIVEL ESTUDIOS', 'TIPO VIVIENDA', 'ANTIGÜEDAD', 'TIPO CONTRATO', 'MODALIDAD DE TRABAJO', 'VARIABLE 2',
    'VARIABLE 1', 'APORTE AL HOGAR', 'APORTE SOLO YO', 'APORTE ESPOSO', 'APORTE HIJOS', 'APORTE PADRES', 'APORTE HERMANOS',
    'APORTE OTROS', 'REDUCCION INGRESOS', 'MEDIOS FINANCIACION', 'TEMAS DE FORMACION', 'FORMACION TECNICA',
    'FORMACION PROFESIONAL', 'FORMACION IDIOMAS', 'FORMACION HABILIDADES BASICAs', 'FORMACION DESARROLLO PERSONAL',
    'FORMACION ARTISTICA', 'FORMACION EMPRENDEDORES', 'CON QUIEN VIVE', 'VIVE CON ESPOSO', 'VIVE CON HIJOS', 'VIVE CON PADRES',
    'VIVE CON HERMANOS', 'VIVE CON SOLO', 'VIVE CON FAMILIAR', 'VIVE CON NO FAMILIAR', 'HOGAR', 'HOGAR BEBES',
    'HOGAR PREESCOLAR', 'HOGAR PRIMARIA', 'HOGAR ADOLESCENTES', 'HOGAR JOVENES', 'HOGAR ADULTOS', 'HOGAR NINGUNO',
    'HOGAR ENFERMO CRONICO', 'HOGAR NIÑOS ESTUDIANDO', 'HOGAR OTROS TRABAJAN', 'MEDIO TRANSPORTE', 'MASCOTAS', 'MASCOTAS NO',
    'MASCOTAS GATO', 'MASCOTAS PERRO', 'MASCOTAS PAJARO', 'MASCOTAS PEZ', 'MASCOTAS OTRO', 'NUMERO HIJOS', 'EDAD PRIMER HIJO',
    'EDAD SEGUNDO HIJO', 'EDAD TERCER HIJO', 'EDAD CUARTO HIJO', 'EDAD QUINTO HIJO', 'AFILIADO COLSUBSIDIO', 'TMS',
    'SERVICIOS CAJA', 'PISCILAGO', 'COLSUBSIDIO', 'CATERING Y RESTAURANTES', 'GIMNASIOS Y ZONAS HÚMEDAS', 'VIAJES',
    'RECREACION', 'TORNEOS Y COMPETENCIAS PARA ADULTOS', 'ACTIVIDADES RECREODEPORTIVAS PARA ADULTOS MAYORES', 'TEATRO',
    'SUPERMERCADOS', 'DROGUERÍAS', 'CRÉDITOS Y SEGUROS', 'BIBLIOTECA VIRTUAL', 'BACHILLERATO POR CICLOS',
    'PROGRAMAS FORMACION TECNICA', 'PROYECTOS DE VIVIENDA', 'SERVICIOS ODONTOLÓGICOS', 'CHEQUEOS MÉDICOS',
    'PLAN COMPLEMENTARIO', 'CIRUGÍA ESTÉTICA', 'COMPRA EN', 'COMPRA EN SUPERMECADOS', 'COMPRA DROGUERÍAS', 'SUBSIDIO VIVIENDA',
    'CUPO CREDITO', 'IMPACTO ORGANIZACIÓN', 'PRODUCTIVIDAD', 'SERVICIO EXTERNO', 'SERVICIO INTERNO', 'TIEMPO DEDICADO',
    'TIEMPO DORMIR', 'TIEMPO TRABAJAR', 'TIEMPO DEPORTE', 'TIEMPO ARTE', 'TIEMPO EDUCACION', 'TIEMPO FAMILIA', 'TIEMPO REDES',
    'TIEMPO PANTALLAS', 'HABITIOS', 'ALIMENTACION', 'CONSUMO SUSTANCIAS', 'NIVEL PRECOUPACION', 'PREOCUPACION SALUD FISICA',
    'PREOCUPACION SALUD MENTAL', 'PREOCUPACION SALUD FISICA FAMILIARES', 'PREOCUPACION SALUD MENTAL FAMILIARES',
    'PREOCUPACION SUSTENTO', 'PREOCUPACION FUTURO', 'DEPRESIÓN', 'DEPRESION USTED', 'DEPRESION FAMILIAR', 'DEPRESION AMIGO',
    'ANSIEDAD', 'ANSIEDAD USTED', 'ANSIEDAD FAMILIAR', 'ANSIEDAD AMIGO', 'VARIABLE 3', 'SATISAFACCION MODALIDAD DE TRABAJO',
    'Generación', 'Tipo NPS', 'EMPRESA'
]

# Cargar archivos
sabana = pd.read_excel("1. Sabana Original AFunza 2023.xlsx")
status = pd.read_excel("Alcaldía de Funza - Status Med Clima.xlsx")

# Filtrar participación completa
sabana_filtrada = sabana[sabana["Estado de la participación"] == "Participación completa"].copy()

# Validación de columna IDs / TAN del participante
if "IDs / TAN del participante" not in sabana_filtrada.columns:
    print("ERROR: La columna 'IDs / TAN del participante' no existe en el archivo de la sabana.")
    print("Columnas encontradas:", sabana_filtrada.columns.tolist())
    sys.exit(1)

# Normalizar IDs
sabana_filtrada["ID"] = sabana_filtrada["IDs / TAN del participante"].apply(normalizar_id)
status["ID"] = status["ID"].apply(normalizar_id)

# Mapear columnas por nombre exacto o similitud
def mapear_columnas(sabana_cols, modelo_cols):
    mapping = {}
    for modelo_col in modelo_cols:
        if modelo_col in sabana_cols:
            mapping[modelo_col] = modelo_col
        else:
            similitudes = [(col, similitud_tokens(col, modelo_col)) for col in sabana_cols]
            similitudes.sort(key=lambda x: x[1], reverse=True)
            mejor = similitudes[0]
            if mejor[1] > 0.5:
                mapping[modelo_col] = mejor[0]
            else:
                mapping[modelo_col] = None
    return mapping

mapping = mapear_columnas(sabana_filtrada.columns, modelo_columnas)

# Construir DataFrame final
df_final = pd.DataFrame(columns=modelo_columnas)
for idx, row in sabana_filtrada.iterrows():
    new_row = {}
    for modelo_col in modelo_columnas:
        sabana_col = mapping.get(modelo_col)
        if sabana_col and sabana_col in row:
            new_row[modelo_col] = row[sabana_col]
        else:
            new_row[modelo_col] = pd.NA
    df_final = df_final.append(new_row, ignore_index=True)

# Sobrescribir variables oficiales desde Status
variables_oficiales = ["VARIABLE 1", "VARIABLE 2", "VARIABLE 3", "CARGO", "AREA", "Estado"]
df_final = df_final.merge(
    status[["ID"] + variables_oficiales],
    on="ID", how="left", suffixes=('', '_status')
)
for var in variables_oficiales:
    df_final[var] = df_final[f"{var}_status"].combine_first(df_final[var])
    df_final.drop(columns=[f"{var}_status"], inplace=True)

# Calcular Tipo NPS
df_final["Tipo NPS"] = df_final["SATISFACCION GENERAL"].apply(calcular_tipo_nps)

# Agregar columna EMPRESA
df_final["EMPRESA"] = "ALCALDIA FUNZA"

# Exportar a Excel
df_final.to_excel("Df Procesado Final.xlsx", index=False)
print("Archivo 'Df Procesado Final.xlsx' generado exitosamente.")