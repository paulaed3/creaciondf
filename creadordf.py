import pandas as pd
import re
from pathlib import Path

# ================= Utilidades ================= #

def normalizar_id(id_str):
    if pd.isna(id_str):
        return pd.NA
    solo = re.sub(r'\D', '', str(id_str))
    return solo if solo else pd.NA

def calcular_tipo_nps(valor):
    try:
        v = int(valor)
    except (ValueError, TypeError):
        return pd.NA
    if v >= 9: return "Entusiastas"
    if v >= 7: return "Pasivos"
    if v >= 0: return "Detractores"
    return pd.NA

# versión robusta basada en año numérico (recomendado)
def map_generacion(anio_str: str):
    if pd.isna(anio_str):
        return pd.NA
    s = str(anio_str)
    m = re.search(r'\d{4}', s)
    if not m:
        return pd.NA
    y = int(m.group())
    if 1946 <= y <= 1964: return 'Baby Boomers'
    if 1965 <= y <= 1980: return 'Generación X'
    if 1981 <= y <= 1996: return 'Millennials'
    if 1997 <= y <= 2012: return 'Centennials'
    return pd.NA

def limpiar_area(area: str):
    if pd.isna(area):
        return pd.NA
    a = str(area).strip().upper()
    a = re.sub(r'^SECRETAR[ÍI]A\s+DE\s+', '', a)
    a = a.replace('  ', ' ').strip()
    return a

def safe_get(row, col):
    return row.get(col) if col in row and pd.notna(row.get(col)) else pd.NA

# ================= Definición columnas destino ================= #
OUTPUT_COLUMNS = [
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
    'ANSIEDAD', 'ANSIEDAD USTED', 'ANSIEDAD FAMILIAR', 'ANSIEDAD AMIGO', 'VARIABLE 3', 'SATISFACCIÓN MODALIDAD DE TRABAJO',
    'Generación', 'Tipo NPS', 'EMPRESA'
]

# ================= Mapeo de columnas origen -> destino ================= #
MAPPING_DIRECTO = {
    # Identificación y estado
    'IDs / TAN del participante': 'ID',
    'Fecha y hora': 'FECHA',
    'Estado de la participación': 'ESTATUS',
    # Satisfacción (prefijos posibles)
    '1. En una escala de 1 a 10': 'SATISFACCION GENERAL',
    '2. En una escala de 1 a 10': 'SATISFACCION CONDICIONES',
    '3. ¿Quiere hacer algún comentario': 'COMENTARIOS',
    # Clima (bloque 4)
    'Me siento orgulloso(a)': 'NIVEL DE ORGULLO',
    'Recomendaría a otros trabajar': 'NIVEL DE RECOMENDACIÓN',
    'Me parece inspiradora la misión': 'PERCEPCION MISION INSPIRADORA',
    'Comprendo los objetivos de la entidad': 'APORTE A OBJETIVOS ORGANIZACIONALES',
    'Se reciben beneficios no monetarios': 'BENEFICIOS NO MONETARIOS',
    'Se realizan actividades de bienestar': 'ACTIVIDADES BIENESTAR',
    'Las características de mi trabajo me permiten': 'BALANCE TRABAJO VIDA PERSONAL',
    'La entidad demuestra sensibilidad': 'SENSIBILIDAD POR LA VIDA PERSONAL',
    # Clima (bloque 5)
    'Cuento con los recursos mínimos': 'RECURSOS REQUERIDOS',
    'Puedo acceder a toda la información': 'ACCESO A INFORMACION',
    'Siento que mi cargo me brinda la oportunidad': 'SENSACION DE PROGRESO EN EL CARGO',
    'La formación que brinda la entidad': 'ENTRENAMIENTO PUESTO TRABAJO',
    'El entrenamiento que recibo en mi cargo': 'FORMACION PARA DESARROLLO PERSONAL Y LABORAL',
    'Me evalúan por mi desempeño': 'EVALUACION DESEMPEÑO',
    'En esta entidad celebramos los éxitos': 'CELEBRACION DE EXITOS',
    'Todos en la entidad tienen claro su cargo': 'CLARIDAD DE CARGOS Y RESPONSABILIDADES',
    # Clima (bloque 6)
    'La comunicación entre áreas': 'COMUNICACION INTERAREAS',
    'Mi jefe se comunica de forma clara': 'COMUNICACION CON LIDER',
    'Conozco personas de diferentes áreas': 'TRABAJO INTERAREAS',
    'En mi área se facilita y promueve': 'TRABAJO DENTRO DEL AREA',
    'En la entidad las decisiones se toman': 'OPORTUNIDAD EN DECISIONES',
    'Las decisiones que se toman en la entidad': 'ANALISIS DE DECISIONES',
    'Los líderes en la entidad inspiran': 'LIDERES EJEMPLO',
    'Los jefes forman a sus colaboradores': 'LIDERES MENTORES',
    # Clima (bloque 7)
    'Los jefes y directivos son cordiales': 'RESPETO EN EL TRATO',
    'Puedo dar mi punto de vista y sugerencias': 'DAR PUNTO DE VISTA',
    'Siento que las relaciones en esta entidad': 'RELACIONES DE CONFIANZA',
    'Puedo determinar formas propias': 'AUTONOMIA DEL CARGO',
    'En esta entidad el trato es justo': 'TRATO JUSTO Y EQUITATIVO',
    'En esta entidad se evita que se utilice la intimidación': 'EVITAR INTIMIDACION Y HOSTIGAMIENTO',
    'Yo quiero a la entidad y me siento comprometido': 'AMOR Y COMPROMISO POR LA ORGANIZACIÓN',
    'Me gusta mi trabajo y siento que estoy': 'GUSTO POR EL TRABAJO',
    # Clima (bloque 8)
    'Puedo desempeñar mi trabajo de forma segura': 'CONDICIONES SEGURAS Y COMODAS',
    'Puedo manejar el estrés que se genera': 'MANEJO DEL ESTRÉS',
    'Los procesos de la entidad cuentan con controles': 'CONTROLES DE CALIDAD',
    'Se revisa de forma constante la calidad': 'MEJORA CONTINUA',
    'Se nota una sensibilidad en la entidad por cuidar el impacto': 'CUIDADO DEL MEDIO AMBIENTE',
    'La entidad implementa políticas y/o procedimientos': 'POLITICAS AMBIENTALES',
    'La entidad contrata empleados y proveedores': 'CONTRATACION Y PAGO OPORTUNO',
    'La entidad cuida el impacto que puede tener en la comunidad': 'IMPACTO Y APORTE A LA COMUNIDAD',
    # Demográficos
    '9. Género': 'GENERO',
    '10. Estado civil': 'ESTADO CIVIL',
    '11. Año de nacimiento': 'AÑO NACIMIENTO',
    '11. Año de nacimiento.': 'AÑO NACIMIENTO',
    '12. Nivel de estudios': 'NIVEL ESTUDIOS',
    '13. Tipo de vivienda donde habita': 'TIPO VIVIENDA',
    '27. Antigüedad en la entidad': 'ANTIGÜEDAD',
    '28. Tipo de vinculación': 'TIPO CONTRATO',
    '26. Modalidad de trabajo': 'MODALIDAD DE TRABAJO',
    '30. Área en la que labora': 'VARIABLE 1',   # <- será sobrescrita por maestro
    '29. Ciudad / Región donde vive': 'VARIABLE 3',  # <- será sobrescrita por maestro
    # Hogar y aportes
    '14. ¿Cuántos hijos / hijas tiene?': 'NUMERO HIJOS',
    '15. Indique el rango de edad de su primer hijo/a': 'EDAD PRIMER HIJO',
    '16. Indique el rango de edad de su segundo hijo/a': 'EDAD SEGUNDO HIJO',
    '17. Indique el rango de edad de su tercer hijo/a': 'EDAD TERCER HIJO',
    '18. Indique el rango de edad de su cuarto hijo/a': 'EDAD CUARTO HIJO',
    '19. Indique el rango de edad de su quinto hijo/a': 'EDAD QUINTO HIJO',
    'Sólo yo aporto al hogar': 'APORTE SOLO YO',
    'Esposo(a)/Pareja': 'APORTE ESPOSO',
    'Hijos(as)': 'APORTE HIJOS',
    'Padres': 'APORTE PADRES',
    'Hermanos(as)': 'APORTE HERMANOS',
    'Otros': 'APORTE OTROS',
    # Ingresos / financiación
    '21. ¿En su hogar se ha dado una disminución': 'REDUCCION INGRESOS',
    '22. Si usted tuviera una necesidad de financiación': 'MEDIOS FINANCIACION',
    # Formación interés
    'Formación técnica/tecnológica': 'FORMACION TECNICA',
    'Formación profesional (pregrado/ posgrado)': 'FORMACION PROFESIONAL',
    'Idiomas': 'FORMACION IDIOMAS',
    'Habilidades básicas (Excel, power point, computación)': 'FORMACION HABILIDADES BASICAs',
    'Desarrollo personal': 'FORMACION DESARROLLO PERSONAL',
    'Formación artística': 'FORMACION ARTISTICA',
    'Formación para emprendedores': 'FORMACION EMPRENDEDORES',
    # Salud mental
    '24. La DEPRESIÓN es': 'DEPRESION USTED',
    'Depresión Usted?': 'DEPRESION FAMILIAR',
    'Depresión Algún familiar cercano?': 'DEPRESION AMIGO', 
    'Depresión Algún amigo cercano?': None,
    '25. Los trastornos de ANSIEDAD': 'ANSIEDAD USTED',
    'Ansiedad Usted?': 'ANSIEDAD FAMILIAR',
    'Ansiedad Algún familiar cercano?': 'ANSIEDAD AMIGO',
    'Ansiedad Algún amigo cercano?': None,
    # Colsubsidio / compras / crédito (igual que antes)...
    '31. ¿Es afiliado a Colsubsidio?': 'AFILIADO COLSUBSIDIO',
    '32. ¿Cuenta usted la Tarjeta Multiservicios': 'TMS',
    '33. Como afiliado a Colsubsidio': 'SERVICIOS CAJA',
    'Piscilago': 'PISCILAGO',
    'Hoteles Colsubsidio': 'COLSUBSIDIO',
    'Catering y restaurantes': 'CATERING Y RESTAURANTES',
    'Clubes deportivos Colsubsidio (Deportes, Gimnasios, zonas húmedas)': 'GIMNASIOS Y ZONAS HÚMEDAS',
    'Agencia de Viajes': 'VIAJES',
    'Recreación para niños': 'RECREACION',
    'Torneos y competencias para adultos': 'TORNEOS Y COMPETENCIAS PARA ADULTOS',
    'Actividades recreodeportivas para adultos mayores': 'ACTIVIDADES RECREODEPORTIVAS PARA ADULTOS MAYORES',
    'Teatro': 'TEATRO',
    'Supermercados': 'SUPERMERCADOS',
    'Droguerías': 'DROGUERÍAS',
    'Créditos y seguros': 'CRÉDITOS Y SEGUROS',
    'Biblioteca virtual': 'BIBLIOTECA VIRTUAL',
    'Bachillerato por ciclos': 'BACHILLERATO POR CICLOS',
    'Programas de formación técnica y tecnológica': 'PROGRAMAS FORMACION TECNICA',
    'Proyectos de vivienda': 'PROYECTOS DE VIVIENDA',
    'Servicios odontológicos': 'SERVICIOS ODONTOLÓGICOS',
    'Chequeos médicos': 'CHEQUEOS MÉDICOS',
    'Plan complementario de salud': 'PLAN COMPLEMENTARIO',
    'Cirugía estética': 'CIRUGÍA ESTÉTICA',
    '34. ¿Hace compras en algunos de estos establecimientos?:': 'COMPRA EN',
    'Supermercados Colsubsidio': 'COMPRA EN SUPERMECADOS',
    'Droguerías Colsubsidio': 'COMPRA DROGUERÍAS',
    '35. ¿Alguna vez ha sido beneficiario del subsidio de vivienda': 'SUBSIDIO VIVIENDA',
    '36. Como afiliado a Colsubsidio usted puede tener un cupo de crédito': 'CUPO CREDITO',
}

# ================= Carga de datos ================= #
INPUT1 = 'input1.xlsx'
INPUT2 = 'input2.xlsx'
OUTPUT = 'output.xlsx'

if not Path(INPUT1).exists():
    raise FileNotFoundError(f"No se encuentra {INPUT1}")
if not Path(INPUT2).exists():
    raise FileNotFoundError(f"No se encuentra {INPUT2}")

df_in = pd.read_excel(INPUT1)
df_map = pd.read_excel(INPUT2)

df_in = df_in[df_in.get('Estado de la participación') == 'Participación completa'].copy()


col_cedula = next(
    (c for c in df_map.columns
     if str(c).strip().lower().replace('é','e') == 'cedula'),
    None
)
if col_cedula is None:
    raise KeyError("INPUT2 no tiene la columna 'CEDULA' (o variante).")

df_map = df_map.rename(columns={col_cedula: 'ID'})
df_map['ID'] = df_map['ID'].apply(normalizar_id)

df_map['ID'] = df_map['ID'].apply(normalizar_id)

cols_maestro = ['ID', 'VARIABLE 1', 'VARIABLE 2', 'VARIABLE 3', 'EMPRESA']
cols_maestro = [c for c in cols_maestro if c in df_map.columns]
df_map_sel = df_map[cols_maestro].copy()

if 'VARIABLE 1' in df_map_sel.columns:
    df_map_sel['VARIABLE 1'] = df_map_sel['VARIABLE 1'].apply(limpiar_area)

df_in = df_in.merge(df_map_sel, on='ID', how='left', suffixes=('', '_MAP'))


sin_match = df_in['ID'].isna() | df_in[['VARIABLE 1','VARIABLE 2','VARIABLE 3']].isna().all(axis=1) if set(['VARIABLE 1','VARIABLE 2','VARIABLE 3']).issubset(df_in.columns) else df_in['ID'].isna()
faltantes = df_in.loc[sin_match, 'ID'].dropna().unique().tolist()
if faltantes:
    print(f"[AVISO] {len(faltantes)} ID(s) no encontraron match en INPUT2 (maestro) para VARIABLES 1/2/3.")

salida_rows = []
for idx, row in df_in.iterrows():
    out = {col: pd.NA for col in OUTPUT_COLUMNS}
    out['Unnamed: 0'] = str(len(salida_rows))

    for col_src, col_dst in MAPPING_DIRECTO.items():
        match_val = None
        if col_src in row.index:
            match_val = row[col_src]
        else:
            if col_src not in {'24. La DEPRESIÓN es', '25. Los trastornos de ANSIEDAD'}:
                for candidate in row.index:
                    if str(candidate).startswith(col_src):
                        match_val = row[candidate]
                        break
        if match_val is not None and col_dst is not None:
            out[col_dst] = match_val if pd.notna(match_val) else pd.NA

    out['SATISFACCION COVID'] = pd.NA
    out['RESULTADO SATISFACCION'] = pd.NA


    if 'VARIABLE 1' in df_in.columns:
        out['VARIABLE 1'] = row.get('VARIABLE 1')
    if 'VARIABLE 2' in df_in.columns:
        out['VARIABLE 2'] = row.get('VARIABLE 2')
    if 'VARIABLE 3' in df_in.columns:
        out['VARIABLE 3'] = row.get('VARIABLE 3')

    out['Generación'] = map_generacion(out.get('AÑO NACIMIENTO'))
    out['Tipo NPS'] = calcular_tipo_nps(out.get('SATISFACCION GENERAL'))

    if 'EMPRESA' in df_in.columns and pd.notna(row.get('EMPRESA')):
        out['EMPRESA'] = row.get('EMPRESA')
    else:
        out['EMPRESA'] = 'ALCALDIA FUNZA'  

    out['SATISFACCIÓN MODALIDAD DE TRABAJO'] = pd.NA

    salida_rows.append(out)

df_out = pd.DataFrame(salida_rows, columns=OUTPUT_COLUMNS)

if 'DEPRESION USTED' in df_out.columns:
    df_out['DEPRESIÓN'] = df_out['DEPRESION USTED']
if 'ANSIEDAD USTED' in df_out.columns:
    df_out['ANSIEDAD'] = df_out['ANSIEDAD USTED']

df_out.to_excel(OUTPUT, index=False)
print(f"Archivo '{OUTPUT}' generado con {len(df_out)} filas y {len(df_out.columns)} columnas.")
