import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import re

# Configuración de la página
st.set_page_config(
    page_title="🛠️ Analizador de Respiradoresimport streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import re

# Configuración de la página
st.set_page_config(
    page_title="🛠️ Analizador de Respiradores v3.0",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("🛠️ Analizador Automático de Respiradores v3.0")
st.markdown("**Automatización completa de selección basada en metodología Noria Corporation**")

# Sidebar para configuración técnica
st.sidebar.header("⚙️ Configuración Técnica")
st.sidebar.markdown("---")

# Coeficientes de expansión térmica (Sistema Inglés - Según Excel Aprobado)
st.sidebar.subheader("🧮 Coeficientes de Expansión (Sistema Inglés)")
gamma_aceite = st.sidebar.number_input("γ (Aceite) [F⁻¹]", value=0.0003611, format="%.7f", help="Coeficiente expansión aceite en sistema inglés")
beta_aire = st.sidebar.number_input("β (Aire) [F⁻¹]", value=0.001894, format="%.6f", help="Coeficiente expansión aire en sistema inglés")
factor_seguridad = st.sidebar.number_input("Factor Seguridad CFM", value=1.4, format="%.1f", help="Factor de seguridad aplicado al CFM calculado")

# Parámetros de iteración
st.sidebar.subheader("🔄 Parámetros de Calibración")
porcentajes_aceite = st.sidebar.multiselect(
    "% Aceite a iterar", 
    [25, 30, 35, 40, 45, 50, 55], 
    default=[30, 35, 40, 45, 50],
    help="Rangos de llenado de aceite para calibración del modelo"
)
tiempos_expansion = st.sidebar.multiselect(
    "Tiempos expansión [min]", 
    [6, 8, 10, 12, 15], 
    default=[8, 10, 12],
    help="Tiempos de expansión térmica para diferentes escenarios"
)

st.sidebar.markdown("---")
st.sidebar.info("📋 **Flujo Completo v3.0:**\n1. Análisis preliminar automático\n2. Inputs individuales por equipo\n3. Análisis robustez real\n4. Data Report completo")

# Base de datos de respiradores
RESPIRADORES_DB = {
    # X-Series
    "X-100": {"altura": 6.25, "cfm_max": 10, "serie": "X-Series", "cartucho": "L-143"},
    "X-50": {"altura": 5, "cfm_max": 10, "serie": "X-Series", "cartucho": "L-50"},
    "X-501": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-341"},
    "X-502": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-342"},
    "X-503": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-345"},
    "X-505": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-346"},
    "X-521": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-343"},
    "X-522": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-344"},
    
    # Guardian Series
    "G5S1N": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5S"},
    "G5S1NC": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5SC"},
    "G5S1NG": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5S"},
    "G5S1NGC": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5SC"},
    "G8S1N": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8S"},
    "G8S1NC": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8SC"},
    "G8S1NG": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8S"},
    "G8S1NGC": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8SC"},
    "G12S1N": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12S"},
    "G12S1NC": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12SC"},
    "G12S1NG": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12S"},
    "G12S1NGC": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12SC"},
    
    # D-Series
    "D-100": {"altura": 3.5, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-101": {"altura": 5, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-102": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-103": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-104": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-108": {"altura": 10, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"},
    "D-2": {"altura": 10, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable"}
}

# Estado de la aplicación
if 'analysis_stage' not in st.session_state:
    st.session_state.analysis_stage = 'upload'
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
if 'equipos_candidatos' not in st.session_state:
    st.session_state.equipos_candidatos = None

# Funciones principales
@st.cache_data
def cargar_datos(uploaded_file):
    """Cargar y procesar el archivo Excel"""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None

def identificar_equipos_candidatos(df):
    """Identificar equipos únicos que requieren respiradores"""
    mask = (
        (df['ComponentTemplate'].str.contains('gearbox', case=False, na=False)) |
        (df['ComponentTemplate'].str.contains('oil', case=False, na=False)) |
        (df['(D) Oil Capacity'].fillna(0) > 0) |
        (df['MaintPointTemplate'].str.contains('oil', case=False, na=False))
    )
    candidatos = df[mask].copy()
    
    equipos_unicos = []
    grupos = candidatos.groupby(['Machine', 'Component'])
    
    for (machine, component), grupo in grupos:
        registro_principal = None
        
        registros_con_aceite = grupo[grupo['(D) Oil Capacity'].fillna(0) > 0]
        if len(registros_con_aceite) > 0:
            registro_principal = registros_con_aceite.iloc[0]
        elif len(grupo[grupo['MaintPointTemplate'].str.contains('oil', case=False, na=False)]) > 0:
            registro_principal = grupo[grupo['MaintPointTemplate'].str.contains('oil', case=False, na=False)].iloc[0]
        else:
            registro_principal = grupo.iloc[0]
        
        registro_consolidado = registro_principal.copy()
        
        for idx, row in grupo.iterrows():
            for col in ['(D) Average Relative Humidity', '(D) Contaminant Abrasive Index', 
                       '(D) Contaminant Likelihood', '(D) Vibration', '(D) Oil Mist Evidence on Headspace',
                       '(D) Water Contact Conditions', '(D) Environmental Conditions',
                       '(D) Breather/Fill Port Clearance', '(D) Breather/Fill Port Size']:
                if pd.isna(registro_consolidado.get(col)) and pd.notna(row.get(col)):
                    registro_consolidado[col] = row[col]
        
        equipos_unicos.append(registro_consolidado)
    
    return pd.DataFrame(equipos_unicos), candidatos

def analizar_inputs_faltantes_por_equipo(equipos_df):
    """Analizar qué inputs manuales faltan por cada equipo individual"""
    equipos_con_faltantes = []
    resumen_faltantes = {}
    
    for idx, row in equipos_df.iterrows():
        machine = row['Machine']
        component = row['Component']
        equipo_key = f"{machine}_{component}"
        
        faltantes_equipo = {}
        
        # 1. CRITICIDAD - Siempre falta (client provided)
        faltantes_equipo['criticidad'] = {
            'requerido': True,
            'razon': 'Client provided (no viene en Data Report)'
        }
        
        # 2. HIGH PARTICLE REMOVAL - Siempre falta (Noria expert advice)
        faltantes_equipo['high_particle_removal'] = {
            'requerido': True,
            'razon': 'Noria expert advice (no viene en Data Report)'
        }
        
        # 3. TEMPERATURAS - Faltan 3 de las 4 requeridas según Excel aprobado
        faltantes_equipo['temp_operating_min'] = {
            'requerido': True,
            'razon': 'Solo se tiene Operating Temperature máxima en Data Report'
        }
        
        faltantes_equipo['temp_ambient_min'] = {
            'requerido': True,
            'razon': 'Temperatura ambiente mínima requerida para cálculo ΔT'
        }
        
        faltantes_equipo['temp_ambient_max'] = {
            'requerido': True,
            'razon': 'Temperatura ambiente máxima requerida para cálculo ΔT'
        }
        
        # 4. OPERATION RUN HOURS - Solo si falta
        operation_hours = row.get('(D) Operation Run Hours', '')
        if pd.isna(operation_hours) or str(operation_hours).strip() == '':
            faltantes_equipo['operation_hours'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report'
            }
        
        # 5. OIL DRYNESS LIMIT - Solo si falta
        oil_dryness = row.get('(D) Oil Dryness Limit', '')
        if pd.isna(oil_dryness) or str(oil_dryness).strip() == '':
            faltantes_equipo['oil_dryness_limit'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report'
            }
        
        # 6. RELATIVE HUMIDITY - Solo si falta
        humidity = row.get('(D) Average Relative Humidity', '')
        if pd.isna(humidity) or str(humidity).strip() == '':
            faltantes_equipo['relative_humidity'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report (requerido para selección desecante)'
            }
        
        if faltantes_equipo:
            equipos_con_faltantes.append({
                'machine': machine,
                'component': component,
                'equipo_key': equipo_key,
                'oil_capacity': row.get('(D) Oil Capacity', 0),
                'faltantes': faltantes_equipo
            })
            
            for input_name in faltantes_equipo.keys():
                if input_name not in resumen_faltantes:
                    resumen_faltantes[input_name] = 0
                resumen_faltantes[input_name] += 1
    
    return equipos_con_faltantes, resumen_faltantes

def extraer_valor_numerico(valor_str, tipo="porcentaje"):
    """Extraer valor numérico de strings con rangos o texto"""
    if pd.isna(valor_str):
        return None
    
    valor_str = str(valor_str).lower()
    
    if tipo == "porcentaje":
        match = re.search(r'(\d+\.?\d*)%', valor_str)
        if match:
            return float(match.group(1))
        match = re.search(r'(\d+\.?\d*)', valor_str)
        if match:
            return float(match.group(1))
    elif tipo == "indice":
        if 'heavy' in valor_str or 'mining' in valor_str:
            return 80
        elif 'medium' in valor_str:
            return 50
        elif 'light' in valor_str or 'low' in valor_str:
            return 20
        match = re.search(r'(\d+\.?\d*)', valor_str)
        if match:
            return float(match.group(1))
    
    return None

def extraer_espacio_disponible(clearance_str):
    """Extraer espacio disponible en inches"""
    if pd.isna(clearance_str):
        return None
    
    clearance_str = str(clearance_str).lower()
    
    if 'no port' in clearance_str or 'no available' in clearance_str:
        return 0
    elif '2 to 4' in clearance_str or '2-4' in clearance_str:
        return 4
    elif '4 to 6' in clearance_str or '4-6' in clearance_str:
        return 6
    elif '6 to 8' in clearance_str or '6-8' in clearance_str:
        return 8
    elif '8 to 10' in clearance_str or '8-10' in clearance_str:
        return 10
    elif 'greater than 12' in clearance_str or '>12' in clearance_str:
        return 20
    elif 'greater than 10' in clearance_str or '>10' in clearance_str:
        return 15
    elif 'greater than 6' in clearance_str or '>6' in clearance_str:
        return 12
    elif 'free' in clearance_str or 'unlimited' in clearance_str:
        return 30
    elif 'limited' in clearance_str:
        return 5
    
    match = re.search(r'(\d+\.?\d*)', clearance_str)
    if match:
        return float(match.group(1))
    
    return None

def extraer_temperaturas_completas(temp_string, inputs_manuales_equipo):
    """Extraer temperaturas según metodología Excel aprobada - SISTEMA INGLÉS"""
    
    # 1. Extraer Operating Temperature máxima del Data Report
    operating_max = None
    if pd.notna(temp_string):
        temp_str = str(temp_string).lower()
        match_f = re.search(r'(\d+\.?\d*)°?f', temp_str)
        if match_f:
            operating_max = float(match_f.group(1))
        else:
            match_c = re.search(r'(\d+\.?\d*)°?c', temp_str)
            if match_c:
                temp_c = float(match_c.group(1))
                operating_max = temp_c * 9/5 + 32
            else:
                match_num = re.search(r'(\d+\.?\d*)', temp_str)
                if match_num:
                    operating_max = float(match_num.group(1))
    
    if operating_max is None:
        operating_max = 180
    
    # 2. Obtener temperaturas de inputs manuales
    operating_min = inputs_manuales_equipo.get('temp_operating_min', 77)
    ambient_min = inputs_manuales_equipo.get('temp_ambient_min', 50)
    ambient_max = inputs_manuales_equipo.get('temp_ambient_max', 100)
    
    # 3. Aplicar lógica del Excel aprobado
    T_max = max(operating_max, ambient_max)
    T_min = min(operating_min, ambient_min)
    delta_t = T_max - T_min
    
    return T_min, T_max, delta_t, {
        'operating_max': operating_max,
        'operating_min': operating_min,
        'ambient_max': ambient_max,
        'ambient_min': ambient_min
    }

def determinar_respirador_por_cfm(cfm, series_permitidas, requiere_cartucho_reemplazable):
    """Determinar qué respirador sería seleccionado para un CFM específico"""
    candidatos = []
    for modelo, specs in RESPIRADORES_DB.items():
        if specs['serie'] in series_permitidas and specs['cfm_max'] >= cfm:
            if requiere_cartucho_reemplazable and specs['serie'] == 'D-Series':
                continue
            candidatos.append((modelo, specs))
    
    if not candidatos:
        return None
    
    orden_serie = {'Guardian': 3, 'X-Series': 2, 'D-Series': 1}
    mejor = min(candidatos, key=lambda x: (
        -orden_serie.get(x[1]['serie'], 0),
        abs(x[1]['cfm_max'] - cfm * 1.2),
        x[1]['altura']
    ))
    
    return mejor[0]

def calcular_cfm_con_robustez_real(row, porcentajes, tiempos, gamma, beta, factor_seguridad, inputs_manuales_equipo, criticidad_impacto):
    """Calcular CFM con análisis de robustez REAL basado en selección de respirador"""
    oil_capacity_gal = row.get('(D) Oil Capacity', 0)
    
    if pd.isna(oil_capacity_gal) or oil_capacity_gal <= 0:
        return None
    
    # Extraer temperaturas con metodología correcta del Excel
    temp_min, temp_max, delta_t, temp_breakdown = extraer_temperaturas_completas(
        row.get('(D) Operating Temperature'), inputs_manuales_equipo
    )
    
    # Obtener dimensiones
    length = row.get('(D) Length', np.nan)
    width = row.get('(D) Width', np.nan)
    height = row.get('(D) Height', np.nan)
    
    # Determinar método de cálculo
    metodo = "Estimativo"
    if pd.notna(length) and pd.notna(width) and pd.notna(height) and all(x > 0 for x in [length, width, height]):
        metodo = "Directo"
    elif pd.notna(length) and pd.notna(height) and length > 0 and height > 0:
        metodo = "Inverso"
    
    # Determinar restricciones para selección de respirador
    series_permitidas = criticidad_impacto['series_preferidas']
    requiere_cartucho_reemplazable = (
        criticidad_impacto['cartucho_obligatorio'] or 
        inputs_manuales_equipo.get('high_particle_removal') == 'Yes'
    )
    
    # Calcular TODAS las iteraciones para análisis de robustez REAL
    iteraciones_detalladas = []
    respiradores_sugeridos = []
    
    for pct in porcentajes:
        for tiempo in tiempos:
            if metodo == "Directo":
                volumen_total_gal = (length * width * height) * 2.64172e-7
            elif metodo == "Inverso":
                volumen_total_gal = oil_capacity_gal / (pct / 100)
            else:
                volumen_total_gal = oil_capacity_gal / (pct / 100)
            
            volumen_aceite_gal = volumen_total_gal * (pct / 100)
            volumen_aire_gal = volumen_total_gal - volumen_aceite_gal
            
            # Expansión térmica (Sistema Inglés)
            delta_vo = gamma * volumen_aceite_gal * delta_t
            delta_va = beta * volumen_aire_gal * delta_t
            v_exp_total = (delta_vo + delta_va) * factor_seguridad
            
            # Convertir a ft³ y calcular CFM
            v_exp_ft3 = v_exp_total * 0.133681
            cfm = v_exp_ft3 / tiempo
            
            # CLAVE: Determinar qué respirador sería seleccionado para este CFM
            respirador_sugerido = determinar_respirador_por_cfm(cfm, series_permitidas, requiere_cartucho_reemplazable)
            respiradores_sugeridos.append(respirador_sugerido)
            
            iteraciones_detalladas.append({
                'porcentaje': pct,
                'tiempo': tiempo,
                'volumen_total': volumen_total_gal,
                'volumen_aceite': volumen_aceite_gal,
                'volumen_aire': volumen_aire_gal,
                'cfm': cfm,
                'respirador_sugerido': respirador_sugerido or 'Ninguno compatible'
            })
    
    # Calcular estadísticas CFM
    cfms = [it['cfm'] for it in iteraciones_detalladas]
    cfm_min = min(cfms)
    cfm_max = max(cfms)
    cfm_promedio = np.mean(cfms)
    variacion_cfm = (cfm_max - cfm_min) / cfm_promedio * 100
    
    # ANÁLISIS DE ROBUSTEZ REAL
    respiradores_unicos = [r for r in set(respiradores_sugeridos) if r is not None]
    es_robusto_real = len(respiradores_unicos) <= 1
    
    # Contar frecuencia de cada respirador sugerido
    frecuencia_respiradores = {}
    total_validos = 0
    for resp in respiradores_sugeridos:
        if resp is not None:
            frecuencia_respiradores[resp] = frecuencia_respiradores.get(resp, 0) + 1
            total_validos += 1
    
    return {
        'metodo': metodo,
        'temp_min': temp_min,
        'temp_max': temp_max,
        'delta_t': delta_t,
        'temp_breakdown': temp_breakdown,
        'cfm_min': cfm_min,
        'cfm_max': cfm_max,
        'cfm_promedio': cfm_promedio,
        'cfm_rango': f"{cfm_min:.3f} - {cfm_max:.3f}",
        'variacion_porcentual': variacion_cfm,
        'iteraciones_detalladas': iteraciones_detalladas,
        'total_iteraciones': len(iteraciones_detalladas),
        'respiradores_sugeridos': respiradores_unicos,
        'frecuencia_respiradores': frecuencia_respiradores,
        'es_robusto_real': es_robusto_real,
        'total_iteraciones_validas': total_validos
    }

def filtrar_respiradores_por_espacio(espacio_disponible, cfm_requerido):
    """Filtrar respiradores que realmente caben en el espacio disponible"""
    if espacio_disponible is None:
        espacio_disponible = 8
    
    respiradores_compatibles = []
    
    for modelo, specs in RESPIRADORES_DB.items():
        if specs['altura'] <= espacio_disponible and specs['cfm_max'] >= cfm_requerido:
            respiradores_compatibles.append({
                'modelo': modelo,
                'altura': specs['altura'],
                'cfm_max': specs['cfm_max'],
                'serie': specs['serie'],
                'cartucho': specs['cartucho'],
                'margen_espacio': espacio_disponible - specs['altura']
            })
    
    orden_serie = {'Guardian': 3, 'X-Series': 2, 'D-Series': 1}
    respiradores_compatibles.sort(key=lambda x: (orden_serie.get(x['serie'], 0), -x['altura']))
    
    return respiradores_compatibles

def aplicar_criterios_seleccion(row, inputs_manuales_equipo):
    """Aplicar criterios de selección según Excel aprobado"""
    criterios = {
        'filtro_particulas': False,
        'desecante': False,
        'desecante_alta_capacidad': False,
        'valvula_check': False,
        'resistente_vibracion': False,
        'camara_expansion': False,
        'control_niebla': False,
        'cartucho_reemplazable': False
    }
    
    # Extraer datos del Data Report
    humidity = extraer_valor_numerico(row.get('(D) Average Relative Humidity', ''), "porcentaje")
    contaminant_index = extraer_valor_numerico(row.get('(D) Contaminant Abrasive Index', ''), "indice")
    contaminant_likelihood = str(row.get('(D) Contaminant Likelihood', '')).lower()
    vibration = str(row.get('(D) Vibration', '')).lower()
    water_contact = str(row.get('(D) Water Contact Conditions', '')).lower()
    oil_mist = row.get('(D) Oil Mist Evidence on Headspace', False)
    
    # Usar inputs manuales si faltan datos
    if humidity is None and 'relative_humidity' in inputs_manuales_equipo:
        humidity = inputs_manuales_equipo['relative_humidity']
    elif humidity is None:
        humidity = 50
    
    if contaminant_index is None:
        contaminant_index = 40
    
    # 1. FILTRO DE PARTÍCULAS - Siempre requerido según Excel
    criterios['filtro_particulas'] = True
    
    # 2. CARTUCHO REEMPLAZABLE
    if contaminant_index > 75 and 'high' in contaminant_likelihood:
        criterios['cartucho_reemplazable'] = True
    
    # HIGH PARTICLE REMOVAL = CARTUCHOS REEMPLAZABLES OBLIGATORIOS
    if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
        criterios['cartucho_reemplazable'] = True
    
    # 3. DESECANTE
    if humidity > 75:
        criterios['desecante'] = True
        criterios['desecante_alta_capacidad'] = True
        criterios['valvula_check'] = True
    elif humidity > 50:
        criterios['desecante'] = True
    
    # 4. CHECK VALVE adicional
    if 'water' in water_contact and 'no water' not in water_contact:
        criterios['valvula_check'] = True
        criterios['desecante'] = True
    
    # 5. RESISTENTE A VIBRACIÓN
    if 'ips' in vibration:
        try:
            vib_value = float(re.search(r'(\d+\.?\d*)', vibration).group(1))
            if vib_value > 0.4:
                criterios['resistente_vibracion'] = True
        except:
            pass
    elif '>0.4' in vibration or 'high' in vibration:
        criterios['resistente_vibracion'] = True
    
    # 6. CÁMARA DE EXPANSIÓN - NOT recommended según Excel
    criterios['camara_expansion'] = False
    
    # 7. CONTROL DE NIEBLA
    if oil_mist == True or str(oil_mist).lower() == 'yes':
        criterios['control_niebla'] = True
    
    return criterios

def determinar_criticidad_impacto(criticidad_str):
    """Determinar impacto de criticidad en la selección - LÓGICA TEMPORAL"""
    if 'A' in criticidad_str or 'critica' in criticidad_str.lower():
        return {
            'nivel': 'A (Crítica)',
            'prioridad': 'Crítica',
            'series_preferidas': ['Guardian'],
            'cartucho_obligatorio': True,
            'modificador_conservador': True
        }
    elif 'B' in criticidad_str or 'intermedia' in criticidad_str.lower():
        return {
            'nivel': 'B (Intermedia)',
            'prioridad': 'Media',
            'series_preferidas': ['Guardian', 'X-Series'],
            'cartucho_obligatorio': False,
            'modificador_conservador': False
        }
    else:
        return {
            'nivel': 'C (Baja)',
            'prioridad': 'Baja',
            'series_preferidas': ['Guardian', 'X-Series', 'D-Series'],
            'cartucho_obligatorio': False,
            'modificador_conservador': False
        }

def seleccionar_respirador_final(respiradores_compatibles, criterios, criticidad_impacto, cfm_promedio, inputs_manuales_equipo):
    """Seleccionar el respirador final basado en criterios y criticidad"""
    if not respiradores_compatibles:
        return None
    
    series_permitidas = criticidad_impacto['series_preferidas']
    candidatos_finales = [r for r in respiradores_compatibles if r['serie'] in series_permitidas]
    
    if not candidatos_finales:
        candidatos_finales = respiradores_compatibles
    
    # Aplicar criterios específicos
    if criterios['cartucho_reemplazable'] or criticidad_impacto['cartucho_obligatorio']:
        candidatos_finales = [r for r in candidatos_finales if r['serie'] != 'D-Series']
    
    # HIGH PARTICLE REMOVAL = EXCLUIR D-SERIES
    if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
        candidatos_finales = [r for r in candidatos_finales if r['serie'] != 'D-Series']
    
    if criterios['desecante_alta_capacidad']:
        candidatos_check = [r for r in candidatos_finales if 'C' in r['modelo']]
        if candidatos_check:
            candidatos_finales = candidatos_check
    
    if not candidatos_finales:
        candidatos_finales = respiradores_compatibles
    
    orden_serie = {'Guardian': 3, 'X-Series': 2, 'D-Series': 1}
    
    mejor_candidato = min(candidatos_finales, key=lambda x: (
        -orden_serie.get(x['serie'], 0),
        abs(x['cfm_max'] - cfm_promedio * 1.2),
        x['altura']
    ))
    
    return mejor_candidato

def generar_modificaciones_requeridas(row, respiradores_compatibles, cfm_requerido, inputs_manuales_equipo, criticidad_impacto):
    """Generar lista de modificaciones requeridas y recomendación post-modificaciones"""
    modificaciones = []
    
    clearance = str(row.get('(D) Breather/Fill Port Clearance', '')).lower()
    port_size = str(row.get('(D) Breather/Fill Port Size', '')).lower()
    port_available = str(row.get('(D) Port Availability', '')).lower()
    espacio_disponible = extraer_espacio_disponible(clearance)
    
    # 1. Verificar puerto disponible
    if 'no port' in clearance or 'no' in port_available or espacio_disponible == 0:
        if '1/2' in port_size or '0.5' in port_size:
            modificaciones.append("🔧 Instalar puerto 1/2\" NPT en tapa del equipo")
        elif '1\"' in port_size or '1 inch' in port_size:
            modificaciones.append("🔧 Instalar puerto 1\" NPT en tapa del equipo")
        elif '2\"' in port_size or '2 inch' in port_size:
            modificaciones.append("🔧 Instalar puerto 2\" NPT en tapa del equipo")
        else:
            modificaciones.append("🔧 Instalar puerto 1\" NPT en tapa del equipo (recomendado)")
    
    # 2. Verificar espacio suficiente y generar recomendación post-modificaciones
    recomendacion_post_modificaciones = None
    
    if not respiradores_compatibles:
        # Filtrar respiradores según restricciones
        todos_respiradores = list(RESPIRADORES_DB.items())
        if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
            todos_respiradores = [(k, v) for k, v in todos_respiradores if v['serie'] != 'D-Series']
        
        series_permitidas = criticidad_impacto['series_preferidas']
        respiradores_filtrados = [(k, v) for k, v in todos_respiradores if v['serie'] in series_permitidas]
        
        if not respiradores_filtrados:
            respiradores_filtrados = todos_respiradores
        
        respiradores_adecuados = [(k, v) for k, v in respiradores_filtrados if v['cfm_max'] >= cfm_requerido]
        
        if respiradores_adecuados:
            respirador_recomendado = min(respiradores_adecuados, key=lambda x: x[1]['altura'])
            modelo_rec = respirador_recomendado[0]
            specs_rec = respirador_recomendado[1]
            
            recomendacion_post_modificaciones = {
                'modelo': modelo_rec,
                'serie': specs_rec['serie'],
                'cartucho': specs_rec['cartucho'],
                'altura': specs_rec['altura'],
                'cfm_capacidad': specs_rec['cfm_max']
            }
            
            if espacio_disponible is not None and espacio_disponible > 0:
                modificaciones.append(f"📏 Crear espacio mínimo {specs_rec['altura']}\" vertical (actual: {espacio_disponible}\")")
            else:
                modificaciones.append(f"📏 Crear espacio mínimo {specs_rec['altura']}\" vertical para instalación")
        else:
            modelo_minimo = min(RESPIRADORES_DB.values(), key=lambda x: x['altura'])
            modificaciones.append(f"📏 Crear espacio mínimo {modelo_minimo['altura']}\" vertical para instalación")
            modificaciones.append(f"⚡ CFM requerido ({cfm_requerido:.2f}) excede capacidad disponible")
        
        modificaciones.append("🔄 ALTERNATIVA: Instalación remota con línea de conexión")
    
    # 3. Verificar CFM adecuado
    if respiradores_compatibles:
        max_cfm_disponible = max([r['cfm_max'] for r in respiradores_compatibles])
        if cfm_requerido > max_cfm_disponible:
            modificaciones.append(f"⚡ CFM requerido ({cfm_requerido:.2f}) excede capacidad máxima disponible ({max_cfm_disponible})")
            modificaciones.append("🔄 Considerar: Múltiples respiradores o modelo de mayor capacidad")
    
    # 4. Recomendaciones de acceso
    if respiradores_compatibles and any('Guardian' in r['serie'] for r in respiradores_compatibles):
        modificaciones.append("🔧 Asegurar acceso para mantenimiento de cartuchos reemplazables")
    elif recomendacion_post_modificaciones and recomendacion_post_modificaciones['serie'] == 'Guardian':
        modificaciones.append("🔧 Asegurar acceso para mantenimiento de cartuchos reemplazables")
    
    # 5. Consideraciones ambientales
    env_conditions = str(row.get('(D) Environmental Conditions', '')).lower()
    if 'outside' in env_conditions or 'exterior' in env_conditions:
        modificaciones.append("🌧️ Considerar protección adicional contra intemperie")
    
    return modificaciones, recomendacion_post_modificaciones

def determinar_tipo_sistema(row):
    """Determinar si el sistema es estático o circulatorio"""
    flow_rate = row.get('(D) Flow Rate', np.nan)
    if pd.notna(flow_rate) and flow_rate > 0:
        return "Circulatorio"
    else:
        return "Estático (Splash)"

def procesar_equipo_completo(row, inputs_manuales_equipo, porcentajes, tiempos, gamma, beta, factor_seguridad):
    """Procesar un equipo individual con análisis completo"""
    machine = row['Machine']
    component = row['Component']
    
    # Determinar criticidad
    criticidad_str = inputs_manuales_equipo.get('criticidad', 'B (Intermedia)')
    criticidad_impacto = determinar_criticidad_impacto(criticidad_str)
    
    # Determinar tipo de sistema
    tipo_sistema = determinar_tipo_sistema(row)
    
    # Calcular CFM con análisis de robustez REAL
    cfm_data = calcular_cfm_con_robustez_real(row, porcentajes, tiempos, gamma, beta, factor_seguridad, inputs_manuales_equipo, criticidad_impacto)
    
    if not cfm_data:
        return {
            'machine': machine,
            'component': component,
            'error': 'Datos insuficientes de Oil Capacity',
            'cfm_data': None,
            'recomendacion': None,
            'modificaciones': ['⚠️ Faltan datos de Oil Capacity para análisis']
        }
    
    # Aplicar criterios de selección
    criterios = aplicar_criterios_seleccion(row, inputs_manuales_equipo)
    
    # Filtrar respiradores por espacio disponible
    espacio_disponible = extraer_espacio_disponible(row.get('(D) Breather/Fill Port Clearance'))
    cfm_promedio = cfm_data['cfm_promedio']
    respiradores_compatibles = filtrar_respiradores_por_espacio(espacio_disponible, cfm_promedio)
    
    # Generar modificaciones requeridas y recomendación post-modificaciones
    modificaciones, recomendacion_post_modificaciones = generar_modificaciones_requeridas(
        row, respiradores_compatibles, cfm_promedio, inputs_manuales_equipo, criticidad_impacto
    )
    
    # Seleccionar respirador final
    respirador_seleccionado = seleccionar_respirador_final(
        respiradores_compatibles, criterios, criticidad_impacto, cfm_promedio, inputs_manuales_equipo
    )
    
    # Análisis de robustez REAL basado en selección de respirador
    if cfm_data['es_robusto_real']:
        if len(cfm_data['respiradores_sugeridos']) == 1:
            respirador_dominante = cfm_data['respiradores_sugeridos'][0]
            robustez = f"✅ ROBUSTO - Todas las iteraciones convergen a {respirador_dominante}"
        else:
            robustez = f"✅ ROBUSTO - Recomendación estable (mismo respirador en todas las condiciones)"
    else:
        freq = cfm_data['frecuencia_respiradores']
        total = cfm_data['total_iteraciones_validas']
        respiradores_desc = []
        for resp, count in sorted(freq.items(), key=lambda x: x[1], reverse=True):
            porcentaje = (count / total) * 100
            respiradores_desc.append(f"{resp} ({porcentaje:.0f}%)")
        
        respiradores_texto = ", ".join(respiradores_desc)
        robustez = f"🔍 VARIABLE - Iteraciones sugieren: {respiradores_texto}"
    
    # Generar recomendación final
    if respirador_seleccionado:
        recomendacion = {
            'modelo_recomendado': respirador_seleccionado['modelo'],
            'serie_recomendada': respirador_seleccionado['serie'],
            'cartucho': respirador_seleccionado['cartucho'],
            'altura_requerida': f"{respirador_seleccionado['altura']}\"",
            'cfm_capacidad': respirador_seleccionado['cfm_max'],
            'cfm_requerido': f"{cfm_promedio:.2f}",
            'cfm_rango': cfm_data['cfm_rango'],
            'prioridad': criticidad_impacto['prioridad'],
            'criticidad': criticidad_impacto['nivel'],
            'robustez': robustez,
            'margen_cfm': f"{((respirador_seleccionado['cfm_max'] / cfm_promedio - 1) * 100):.1f}%",
            'espacio_disponible': f"{espacio_disponible}\"" if espacio_disponible else "No determinado",
            'margen_espacio': f"{respirador_seleccionado['margen_espacio']:.1f}\"" if espacio_disponible else "N/A",
            'recomendacion_post_modificaciones': None
        }
    else:
        recomendacion = {
            'modelo_recomendado': 'REQUIERE MODIFICACIONES',
            'serie_recomendada': 'N/A',
            'cartucho': 'N/A',
            'altura_requerida': 'Ver modificaciones',
            'cfm_capacidad': 'N/A',
            'cfm_requerido': f"{cfm_promedio:.2f}",
            'cfm_rango': cfm_data['cfm_rango'],
            'prioridad': 'CRÍTICA - Modificaciones requeridas',
            'criticidad': criticidad_impacto['nivel'],
            'robustez': robustez,
            'margen_cfm': 'N/A',
            'espacio_disponible': f"{espacio_disponible}\"" if espacio_disponible else "No determinado",
            'margen_espacio': 'Insuficiente',
            'recomendacion_post_modificaciones': recomendacion_post_modificaciones
        }
    
    return {
        'machine': machine,
        'component': component,
        'oil_capacity': row.get('(D) Oil Capacity', 0),
        'tipo_sistema': tipo_sistema,
        'cfm_data': cfm_data,
        'criterios': criterios,
        'criticidad_impacto': criticidad_impacto,
        'respiradores_compatibles': respiradores_compatibles,
        'respirador_seleccionado': respirador_seleccionado,
        'recomendacion': recomendacion,
        'modificaciones': modificaciones,
        'inputs_utilizados': inputs_manuales_equipo
    }

# INTERFAZ PRINCIPAL
def main():
    
    # Mostrar información sobre versión
    with st.expander("🔄 Analizador de Respiradores v3.0 - Reescritura Completa", expanded=False):
        st.markdown("""
        **✅ Versión 3.0 - Implementación Completa y Optimizada:**
        - **Flujo inteligente**: Análisis preliminar → Inputs por equipo → Análisis completo
        - **Robustez REAL**: Basada en selección de respirador, no solo variación CFM
        - **4 temperaturas**: Operating Min/Max + Ambient Min/Max según Excel aprobado
        - **Filtro espacio correcto**: Solo recomienda lo que realmente cabe
        - **High particle removal** → Cartuchos reemplazables obligatorios (excluye D-Series)
        - **Modificaciones + recomendación post-modificaciones**
        - **Data Report completo** + nuevas columnas integradas
        - **Base datos respiradores** completa (X-Series, Guardian, D-Series)
        
        **🤖 Nota sobre IA:**
        ⚠️ *Evaluar uso de IA para completar automáticamente temperaturas ambiente por ubicación geográfica.*
        """)
    
    # Nota prominente sobre IA
    st.info("🤖 **Validar**: Uso de IA para auto-completar datos ambientales por ubicación geográfica. Actualmente manual para mayor precisión.")
    
    # ETAPA 1: UPLOAD Y ANÁLISIS PRELIMINAR
    if st.session_state.analysis_stage == 'upload':
        st.header("📤 Subir Data Report")
        uploaded_file = st.file_uploader(
            "Selecciona tu archivo Excel",
            type=['xlsx', 'xls'],
            help="Sube el data report en formato Excel para análisis preliminar"
        )
        
        if uploaded_file is not None:
            with st.spinner("Cargando y analizando datos..."):
                df = cargar_datos(uploaded_file)
                
                if df is not None:
                    st.session_state.df_original = df
                    
                    equipos_unicos, todos_candidatos = identificar_equipos_candidatos(df)
                    st.session_state.equipos_candidatos = equipos_unicos
                    
                    equipos_con_faltantes, resumen_faltantes = analizar_inputs_faltantes_por_equipo(equipos_unicos)
                    st.session_state.equipos_con_faltantes = equipos_con_faltantes
                    st.session_state.resumen_faltantes = resumen_faltantes
                    
                    st.success(f"✅ Análisis preliminar completado: {len(equipos_unicos)} equipos candidatos")
                    
                    st.subheader("📋 Resumen del Análisis Preliminar")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Registros Total", len(df))
                    with col2:
                        st.metric("Equipos Candidatos", len(equipos_unicos))
                    with col3:
                        st.metric("Equipos con Inputs Faltantes", len(equipos_con_faltantes))
                    
                    with st.expander("👀 Equipos Candidatos Identificados", expanded=True):
                        display_cols = ['Machine', 'Component', '(D) Oil Capacity', '(D) Breather/Fill Port Clearance']
                        available_cols = [col for col in display_cols if col in equipos_unicos.columns]
                        st.dataframe(equipos_unicos[available_cols], use_container_width=True)
                    
                    if resumen_faltantes:
                        st.subheader("⚠️ Resumen de Inputs Manuales Requeridos")
                        for input_name, count in resumen_faltantes.items():
                            st.warning(f"**{input_name.replace('_', ' ').title()}**: {count} equipos requieren este input")
                        
                        if st.button("➡️ Continuar a Solicitud de Inputs por Equipo", type="primary"):
                            st.session_state.analysis_stage = 'inputs'
                            st.rerun()
                    else:
                        st.info("🎉 Todos los datos necesarios están disponibles en el Data Report")
                        if st.button("➡️ Continuar a Análisis Final", type="primary"):
                            st.session_state.analysis_stage = 'analysis'
                            st.rerun()
    
    # ETAPA 2: SOLICITAR INPUTS FALTANTES POR EQUIPO
    elif st.session_state.analysis_stage == 'inputs':
        st.header("📝 Proporcionar Datos Faltantes por Equipo")
        
        st.info(f"Se requieren inputs manuales para {len(st.session_state.equipos_con_faltantes)} equipos")
        
        inputs_manuales_por_equipo = {}
        
        for equipo_info in st.session_state.equipos_con_faltantes:
            machine = equipo_info['machine']
            component = equipo_info['component']
            equipo_key = equipo_info['equipo_key']
            oil_capacity = equipo_info['oil_capacity']
            faltantes = equipo_info['faltantes']
            
            st.subheader(f"🔧 {machine} - {component}")
            st.write(f"**Oil Capacity**: {oil_capacity} gal")
            
            inputs_equipo = {}
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Criticidad
                if 'criticidad' in faltantes:
                    criticidad = st.selectbox(
                        "🎯 Criticidad",
                        ["A (Crítica)", "B (Intermedia)", "C (Baja)"],
                        index=1,
                        key=f"crit_{equipo_key}",
                        help="Nivel de criticidad del equipo para operación"
                    )
                    inputs_equipo['criticidad'] = criticidad
                
                # High Particle Removal
                if 'high_particle_removal' in faltantes:
                    high_particle = st.selectbox(
                        "🔧 High Particle Removal",
                        ["No", "Yes"],
                        key=f"particle_{equipo_key}",
                        help="Basado en recomendación experta Noria - Yes = Cartuchos reemplazables obligatorios"
                    )
                    inputs_equipo['high_particle_removal'] = high_particle
                
                # Operation Hours
                if 'operation_hours' in faltantes:
                    operation_hours = st.selectbox(
                        "⏰ Operation Hours",
                        ["Continuous", "Stand By", "Intermittent"],
                        key=f"hours_{equipo_key}",
                        help="Patrón operativo del equipo"
                    )
                    inputs_equipo['operation_hours'] = operation_hours
            
            with col2:
                # Oil Dryness Limit
                if 'oil_dryness_limit' in faltantes:
                    oil_dryness = st.selectbox(
                        "💧 Oil Dryness Limit",
                        ["< 200 ppm", "200-500 ppm", "> 500 ppm"],
                        index=1,
                        key=f"dryness_{equipo_key}",
                        help="Límite de sequedad del aceite requerido"
                    )
                    inputs_equipo['oil_dryness_limit'] = oil_dryness
                
                # Relative Humidity
                if 'relative_humidity' in faltantes:
                    humidity = st.number_input(
                        "💨 Relative Humidity (%)",
                        min_value=0,
                        max_value=100,
                        value=50,
                        key=f"humidity_{equipo_key}",
                        help="Humedad relativa promedio del ambiente donde está el equipo"
                    )
                    inputs_equipo['relative_humidity'] = humidity
            
            # Sección de temperaturas (siempre requeridas según Excel)
            st.markdown("**🌡️ Temperaturas Requeridas** *(Data Report solo tiene Operating Temperature máxima)*")
            
            col_temp1, col_temp2, col_temp3 = st.columns(3)
            
            with col_temp1:
                if 'temp_operating_min' in faltantes:
                    temp_op_min = st.number_input(
                        "🌡️ Min Operating Temp (°F)",
                        min_value=0,
                        max_value=300,
                        value=77,
                        key=f"temp_op_min_{equipo_key}",
                        help="Temperatura mínima de operación del equipo"
                    )
                    inputs_equipo['temp_operating_min'] = temp_op_min
            
            with col_temp2:
                if 'temp_ambient_min' in faltantes:
                    temp_amb_min = st.number_input(
                        "🌡️ Min Ambient Temp (°F)",
                        min_value=0,
                        max_value=150,
                        value=50,
                        key=f"temp_amb_min_{equipo_key}",
                        help="Temperatura ambiente mínima donde está el equipo"
                    )
                    inputs_equipo['temp_ambient_min'] = temp_amb_min
            
            with col_temp3:
                if 'temp_ambient_max' in faltantes:
                    temp_amb_max = st.number_input(
                        "🌡️ Max Ambient Temp (°F)",
                        min_value=0,
                        max_value=150,
                        value=100,
                        key=f"temp_amb_max_{equipo_key}",
                        help="Temperatura ambiente máxima donde está el equipo"
                    )
                    inputs_equipo['temp_ambient_max'] = temp_amb_max
            
            inputs_manuales_por_equipo[equipo_key] = inputs_equipo
            st.markdown("---")
        
        if st.button("🚀 Ejecutar Análisis Completo v3.0", type="primary"):
            st.session_state.inputs_manuales_por_equipo = inputs_manuales_por_equipo
            st.session_state.analysis_stage = 'analysis'
            st.rerun()
    
    # ETAPA 3: ANÁLISIS COMPLETO
    elif st.session_state.analysis_stage == 'analysis':
        st.header("📊 Análisis Completo v3.0")
        
        with st.spinner("Ejecutando análisis completo con robustez real..."):
            resultados_completos = []
            
            for idx, row in st.session_state.equipos_candidatos.iterrows():
                machine = row['Machine']
                component = row['Component']
                equipo_key = f"{machine}_{component}"
                
                inputs_manuales_equipo = st.session_state.inputs_manuales_por_equipo.get(equipo_key, {})
                
                resultado = procesar_equipo_completo(
                    row, inputs_manuales_equipo, porcentajes_aceite, tiempos_expansion,
                    gamma_aceite, beta_aire, factor_seguridad
                )
                
                resultados_completos.append(resultado)
        
        st.success("✅ Análisis completo v3.0 terminado!")
        
        # Mostrar métricas generales
        total_equipos = len(resultados_completos)
        equipos_con_recomendacion = len([r for r in resultados_completos if r.get('recomendacion') and r['recomendacion']['modelo_recomendado'] != 'REQUIERE MODIFICACIONES'])
        equipos_requieren_modificaciones = total_equipos - equipos_con_recomendacion
        equipos_robustos_real = len([r for r in resultados_completos if r.get('cfm_data') and r['cfm_data']['es_robusto_real']])
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Equipos", total_equipos)
        with col2:
            st.metric("Con Recomendación Directa", equipos_con_recomendacion)
        with col3:
            st.metric("Requieren Modificaciones", equipos_requieren_modificaciones)
        with col4:
            st.metric("Robustos (Mismo Respirador)", equipos_robustos_real)
        
        # Mostrar resultados por equipo
        st.subheader("🔧 Resultados Detallados por Equipo")
        
        for resultado in resultados_completos:
            if 'error' in resultado:
                with st.expander(f"❌ {resultado['component']} - {resultado['machine']}", expanded=False):
                    st.error(resultado['error'])
                    for mod in resultado['modificaciones']:
                        st.write(mod)
                continue
            
            # Determinar el ícono según el resultado
            if resultado['recomendacion']['modelo_recomendado'] == 'REQUIERE MODIFICACIONES':
                icono = "🔧"
                expanded = True
            elif not resultado['cfm_data']['es_robusto_real']:
                icono = "⚠️"
                expanded = False
            else:
                icono = "✅"
                expanded = False
            
            with st.expander(f"{icono} {resultado['component']} - {resultado['machine']}", expanded=expanded):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📋 Datos del Equipo:**")
                    st.write(f"• **Oil Capacity:** {resultado['oil_capacity']} gal")
                    st.write(f"• **Tipo Sistema:** {resultado['tipo_sistema']}")
                    st.write(f"• **Criticidad:** {resultado['recomendacion']['criticidad']}")
                    
                    if resultado['cfm_data']:
                        st.write(f"• **CFM Requerido:** {resultado['recomendacion']['cfm_requerido']}")
                        st.write(f"• **CFM Rango:** {resultado['recomendacion']['cfm_rango']}")
                        st.write(f"• **ΔT:** {resultado['cfm_data']['delta_t']:.1f}°F")
                        st.write(f"• **T_max:** {resultado['cfm_data']['temp_max']:.1f}°F")
                        st.write(f"• **T_min:** {resultado['cfm_data']['temp_min']:.1f}°F")
                        st.write(f"• **Espacio Disponible:** {resultado['recomendacion']['espacio_disponible']}")
                
                with col2:
                    st.markdown("**🏆 Recomendación Final:**")
                    rec = resultado['recomendacion']
                    if rec['modelo_recomendado'] != 'REQUIERE MODIFICACIONES':
                        st.write(f"• **Modelo:** {rec['modelo_recomendado']}")
                        st.write(f"• **Serie:** {rec['serie_recomendada']}")
                        st.write(f"• **Cartucho:** {rec['cartucho']}")
                        st.write(f"• **Altura:** {rec['altura_requerida']}")
                        st.write(f"• **Margen CFM:** {rec['margen_cfm']}")
                        st.write(f"• **Margen Espacio:** {rec['margen_espacio']}")
                    else:
                        st.error("**REQUIERE MODIFICACIONES PARA INSTALACIÓN**")
                        
                        if rec.get('recomendacion_post_modificaciones'):
                            post_rec = rec['recomendacion_post_modificaciones']
                            st.success("**✅ RECOMENDADO DESPUÉS DE MODIFICACIONES:**")
                            st.write(f"• **Modelo:** {post_rec['modelo']} ({post_rec['serie']})")
                            st.write(f"• **Cartucho:** {post_rec['cartucho']}")
                            st.write(f"• **Altura:** {post_rec['altura']}\"")
                            st.write(f"• **CFM Capacidad:** {post_rec['cfm_capacidad']} (requerido: {rec['cfm_requerido']})")
                    
                    st.write(f"• **Prioridad:** {rec['prioridad']}")
                
                # Análisis de robustez REAL
                st.markdown("**🔍 Análisis de Robustez:**")
                if resultado['cfm_data']['es_robusto_real']:
                    st.success(rec['robustez'])
                    st.caption("La recomendación es estable - todas las iteraciones convergen al mismo respirador.")
                else:
                    st.warning(rec['robustez'])
                    st.caption("Las iteraciones sugieren múltiples respiradores - revisar opciones disponibles.")
                    
                    # Mostrar desglose de opciones cuando es variable
                    if resultado['cfm_data']['frecuencia_respiradores']:
                        st.markdown("**🎯 Distribución de Respiradores Sugeridos:**")
                        freq = resultado['cfm_data']['frecuencia_respiradores']
                        total = resultado['cfm_data']['total_iteraciones_validas']
                        
                        for resp, count in sorted(freq.items(), key=lambda x: x[1], reverse=True):
                            porcentaje = (count / total) * 100
                            if resp in RESPIRADORES_DB:
                                specs = RESPIRADORES_DB[resp]
                                st.write(f"• **{resp}** ({specs['serie']}) - {porcentaje:.0f}% de iteraciones - CFM: {specs['cfm_max']} - Altura: {specs['altura']}\"")
                            else:
                                st.write(f"• **{resp}** - {porcentaje:.0f}% de iteraciones")
                
                # Mostrar iteraciones detalladas si hay variabilidad
                if not resultado['cfm_data']['es_robusto_real'] or resultado['cfm_data']['variacion_porcentual'] > 25:
                    if st.checkbox(f"📊 Ver iteraciones detalladas (Variación CFM: {resultado['cfm_data']['variacion_porcentual']:.1f}%)", key=f"iter_detail_{resultado['machine']}_{resultado['component']}"):
                        st.markdown("**Detalle de todas las iteraciones:**")
                        
                        if 'temp_breakdown' in resultado['cfm_data']:
                            temp_breakdown = resultado['cfm_data']['temp_breakdown']
                            st.markdown("**🌡️ Temperaturas utilizadas en el cálculo:**")
                            temp_col1, temp_col2 = st.columns(2)
                            with temp_col1:
                                st.write(f"• Operating Max: {temp_breakdown['operating_max']:.1f}°F (Data Report)")
                                st.write(f"• Operating Min: {temp_breakdown['operating_min']:.1f}°F (Input manual)")
                            with temp_col2:
                                st.write(f"• Ambient Max: {temp_breakdown['ambient_max']:.1f}°F (Input manual)")
                                st.write(f"• Ambient Min: {temp_breakdown['ambient_min']:.1f}°F (Input manual)")
                            st.write(f"**→ T_max = {resultado['cfm_data']['temp_max']:.1f}°F | T_min = {resultado['cfm_data']['temp_min']:.1f}°F | ΔT = {resultado['cfm_data']['delta_t']:.1f}°F**")
                            st.markdown("---")
                        
                        iteraciones_df = pd.DataFrame(resultado['cfm_data']['iteraciones_detalladas'])
                        st.dataframe(iteraciones_df, use_container_width=True)
                
                # Modificaciones requeridas
                if resultado['modificaciones']:
                    st.markdown("**🔧 Modificaciones Requeridas:**")
                    for modificacion in resultado['modificaciones']:
                        if "ALTERNATIVA" in modificacion:
                            st.info(modificacion)
                        elif "⚡" in modificacion or "🔄" in modificacion:
                            st.warning(modificacion)
                        else:
                            st.write(modificacion)
                
                # Respiradores compatibles (solo si hay opciones)
                if len(resultado['respiradores_compatibles']) > 1:
                    if st.checkbox(f"🔍 Ver otras opciones compatibles ({len(resultado['respiradores_compatibles'])-1} adicionales)", key=f"opciones_{resultado['machine']}_{resultado['component']}"):
                        st.markdown("**Otras opciones que caben en el espacio disponible:**")
                        for resp in resultado['respiradores_compatibles']:
                            if resp['modelo'] != resultado['recomendacion'].get('modelo_recomendado', ''):
                                st.write(f"• **{resp['modelo']}** ({resp['serie']}) - Altura: {resp['altura']}\" - CFM: {resp['cfm_max']} - Margen: {resp['margen_espacio']:.1f}\"")
        
        # Generar descarga del Data Report completo
        st.header("💾 Descargar Data Report Completo + Análisis")
        
        df_completo = st.session_state.df_original.copy()
        
        # Agregar nuevas columnas
        nuevas_columnas = {
            'Respirador_Requerido': 'No analizado',
            'Criticidad': '',
            'Tipo_Sistema': '',
            'CFM_Requerido': '',
            'CFM_Rango': '',
            'Delta_T_F': '',
            'T_max_F': '',
            'T_min_F': '',
            'Operating_Max_F': '',
            'Operating_Min_F': '',
            'Ambient_Max_F': '',
            'Ambient_Min_F': '',
            'Modelo_Recomendado': '',
            'Serie_Recomendada': '',
            'Cartucho_Recomendado': '',
            'Altura_Requerida': '',
            'Espacio_Disponible': '',
            'Margen_Espacio': '',
            'Margen_CFM': '',
            'Prioridad': '',
            'Robustez': '',
            'Variacion_CFM_Pct': '',
            'Total_Iteraciones': '',
            'Modelo_Post_Modificaciones': '',
            'Serie_Post_Modificaciones': '',
            'Cartucho_Post_Modificaciones': '',
            'Altura_Post_Modificaciones': '',
            'CFM_Capacidad_Post_Modificaciones': '',
            'Modificaciones_Requeridas': '',
            'Status_Analisis': 'No analizado'
        }
        
        for col in nuevas_columnas:
            df_completo[col] = nuevas_columnas[col]
        
        # Actualizar registros analizados
        for resultado in resultados_completos:
            if 'error' in resultado:
                mask = (df_completo['Machine'] == resultado['machine']) & (df_completo['Component'] == resultado['component'])
                indices = df_completo[mask].index
                
                if len(indices) > 0:
                    idx_principal = indices[0]
                    df_completo.loc[idx_principal, 'Respirador_Requerido'] = 'Error en análisis'
                    df_completo.loc[idx_principal, 'Status_Analisis'] = resultado['error']
                continue
            
            rec = resultado['recomendacion']
            
            mask = (df_completo['Machine'] == resultado['machine']) & (df_completo['Component'] == resultado['component'])
            indices = df_completo[mask].index
            
            if len(indices) > 0:
                registros_candidatos = df_completo.loc[indices]
                idx_principal = indices[0]
                
                registros_con_aceite = registros_candidatos[registros_candidatos['(D) Oil Capacity'].fillna(0) > 0]
                if len(registros_con_aceite) > 0:
                    idx_principal = registros_con_aceite.index[0]
                
                # Actualizar con resultados del análisis
                df_completo.loc[idx_principal, 'Respirador_Requerido'] = 'Sí'
                df_completo.loc[idx_principal, 'Criticidad'] = rec['criticidad']
                df_completo.loc[idx_principal, 'Tipo_Sistema'] = resultado['tipo_sistema']
                df_completo.loc[idx_principal, 'CFM_Requerido'] = rec['cfm_requerido']
                df_completo.loc[idx_principal, 'CFM_Rango'] = rec['cfm_rango']
                
                # Información de temperaturas
                if resultado['cfm_data']:
                    df_completo.loc[idx_principal, 'Delta_T_F'] = f"{resultado['cfm_data']['delta_t']:.1f}"
                    df_completo.loc[idx_principal, 'T_max_F'] = f"{resultado['cfm_data']['temp_max']:.1f}"
                    df_completo.loc[idx_principal, 'T_min_F'] = f"{resultado['cfm_data']['temp_min']:.1f}"
                    
                    temp_breakdown = resultado['cfm_data']['temp_breakdown']
                    df_completo.loc[idx_principal, 'Operating_Max_F'] = f"{temp_breakdown['operating_max']:.1f}"
                    df_completo.loc[idx_principal, 'Operating_Min_F'] = f"{temp_breakdown['operating_min']:.1f}"
                    df_completo.loc[idx_principal, 'Ambient_Max_F'] = f"{temp_breakdown['ambient_max']:.1f}"
                    df_completo.loc[idx_principal, 'Ambient_Min_F'] = f"{temp_breakdown['ambient_min']:.1f}"
                
                df_completo.loc[idx_principal, 'Modelo_Recomendado'] = rec['modelo_recomendado']
                df_completo.loc[idx_principal, 'Serie_Recomendada'] = rec['serie_recomendada']
                df_completo.loc[idx_principal, 'Cartucho_Recomendado'] = rec['cartucho']
                df_completo.loc[idx_principal, 'Altura_Requerida'] = rec['altura_requerida']
                df_completo.loc[idx_principal, 'Espacio_Disponible'] = rec['espacio_disponible']
                df_completo.loc[idx_principal, 'Margen_Espacio'] = rec['margen_espacio']
                df_completo.loc[idx_principal, 'Margen_CFM'] = rec['margen_cfm']
                df_completo.loc[idx_principal, 'Prioridad'] = rec['prioridad']
                df_completo.loc[idx_principal, 'Robustez'] = rec['robustez']
                df_completo.loc[idx_principal, 'Variacion_CFM_Pct'] = resultado['cfm_data']['variacion_porcentual'] if resultado['cfm_data'] else 'N/A'
                df_completo.loc[idx_principal, 'Total_Iteraciones'] = resultado['cfm_data']['total_iteraciones'] if resultado['cfm_data'] else 'N/A'
                df_completo.loc[idx_principal, 'Status_Analisis'] = 'Completado'
                
                # Información post-modificaciones si existe
                if rec.get('recomendacion_post_modificaciones'):
                    post_rec = rec['recomendacion_post_modificaciones']
                    df_completo.loc[idx_principal, 'Modelo_Post_Modificaciones'] = post_rec['modelo']
                    df_completo.loc[idx_principal, 'Serie_Post_Modificaciones'] = post_rec['serie']
                    df_completo.loc[idx_principal, 'Cartucho_Post_Modificaciones'] = post_rec['cartucho']
                    df_completo.loc[idx_principal, 'Altura_Post_Modificaciones'] = f"{post_rec['altura']}\""
                    df_completo.loc[idx_principal, 'CFM_Capacidad_Post_Modificaciones'] = post_rec['cfm_capacidad']
                
                # Modificaciones requeridas
                if resultado['modificaciones']:
                    modificaciones_texto = " | ".join(resultado['modificaciones'])
                    df_completo.loc[idx_principal, 'Modificaciones_Requeridas'] = modificaciones_texto
                
                # Marcar registros secundarios
                for idx_sec in indices:
                    if idx_sec != idx_principal:
                        df_completo.loc[idx_sec, 'Respirador_Requerido'] = f'Ver registro principal ({resultado["component"]})'
                        df_completo.loc[idx_sec, 'Status_Analisis'] = 'Registro secundario'
        
        # Crear archivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_completo.to_excel(writer, sheet_name='Data_Report_Completo_v3.0', index=False)
            
            # Hoja de parámetros utilizados
            parametros_data = [
                ['Coeficiente γ (aceite)', gamma_aceite, 'F⁻¹'],
                ['Coeficiente β (aire)', beta_aire, 'F⁻¹'],
                ['Factor seguridad CFM', factor_seguridad, 'multiplicador'],
                ['Porcentajes aceite iterados', ', '.join(map(str, porcentajes_aceite)), '%'],
                ['Tiempos expansión iterados', ', '.join(map(str, tiempos_expansion)), 'min'],
                ['Total registros en data report', len(df_completo), 'registros'],
                ['Equipos candidatos identificados', len(st.session_state.equipos_candidatos), 'equipos'],
                ['Equipos analizados', len(resultados_completos), 'equipos'],
                ['Equipos con recomendación directa', equipos_con_recomendacion, 'equipos'],
                ['Equipos que requieren modificaciones', equipos_requieren_modificaciones, 'equipos'],
                ['Resultados robustos (mismo respirador)', equipos_robustos_real, 'equipos']
            ]
            parametros_df = pd.DataFrame(parametros_data, columns=['Parámetro', 'Valor', 'Unidad'])
            parametros_df.to_excel(writer, sheet_name='Parametros_Analisis_v3.0', index=False)
        
        output.seek(0)
        
        total_registros = len(df_completo)
        registros_analizados = len([r for r in resultados_completos if 'error' not in r])
        
        st.info(f"""
        📊 **Archivo generado v3.0:**
        - **Total registros**: {total_registros} (Data Report completo)
        - **Equipos analizados**: {registros_analizados} con recomendaciones de respiradores
        - **Nuevas columnas**: {len(nuevas_columnas)} agregadas al Data Report original
        - **Robustez real**: Basada en selección de respirador, no solo variación CFM
        - **Temperaturas**: 4 temperaturas según metodología Excel aprobada
        - **🤖 IA**: Temperaturas ambiente manuales - evaluar automatización por ubicación
        """)
        
        st.download_button(
            label="📥 Descargar Data Report Completo + Análisis v3.0",
            data=output.getvalue(),
            file_name=f"data_report_completo_respiradores_v3_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Mostrar estadísticas finales
        with st.expander("📊 Estadísticas del Análisis v3.0", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Distribución por Criticidad:**")
                criticidades = [r['recomendacion']['criticidad'] for r in resultados_completos if r.get('recomendacion')]
                if criticidades:
                    criticidad_counts = pd.Series(criticidades).value_counts()
                    st.bar_chart(criticidad_counts)
                
                st.write("**Distribución por Prioridad:**")
                prioridades = [r['recomendacion']['prioridad'] for r in resultados_completos if r.get('recomendacion')]
                if prioridades:
                    prioridad_counts = pd.Series(prioridades).value_counts()
                    st.bar_chart(prioridad_counts)
            
            with col2:
                st.write("**Series Recomendadas:**")
                series = [r['recomendacion']['serie_recomendada'] for r in resultados_completos if r.get('recomendacion') and r['recomendacion']['modelo_recomendado'] != 'REQUIERE MODIFICACIONES']
                if series:
                    serie_counts = pd.Series(series).value_counts()
                    st.bar_chart(serie_counts)
                
                st.write("**Robustez del Análisis:**")
                st.caption("Evalúa si la recomendación es estable (mismo respirador) ante variaciones operativas")
                variaciones = []
                for r in resultados_completos:
                    if r.get('cfm_data'):
                        if r['cfm_data']['es_robusto_real']:
                            variaciones.append('✅ Robusto (Mismo respirador)')
                        else:
                            variaciones.append('🔍 Variable (Múltiples opciones)')
                
                if variaciones:
                    robustez_counts = pd.Series(variaciones).value_counts()
                    st.bar_chart(robustez_counts)
        
        # Botón para reiniciar
        if st.button("🔄 Reiniciar Análisis Completo"):
            st.session_state.analysis_stage = 'upload'
            st.session_state.df_original = None
            st.session_state.equipos_candidatos = None
            st.session_state.equipos_con_faltantes = []
            st.session_state.inputs_manuales_por_equipo = {}
            st.rerun()

if __name__ == "__main__":
    main()

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p><strong>🛠️ Analizador de Respiradores v3.0 - Implementación Completa</strong></p>
    <p>Robustez Real + 4 Temperaturas + Filtro Espacio + Modificaciones | Metodología Noria Corporation</p>
    <p><em>⚠️ Lógica de criticidad implementada temporalmente - pendiente validación final con equipo técnico</em></p>
</div>
""", unsafe_allow_html=True) v2.2",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("🛠️ Analizador Automático de Respiradores v2.2")
st.markdown("**Automatización de selección basada en metodología Noria Corporation**")

# Sidebar para configuración técnica (siempre visible)
st.sidebar.header("⚙️ Configuración Técnica")
st.sidebar.markdown("---")

# Coeficientes de expansión térmica (Sistema Inglés - Según Excel Aprobado)
st.sidebar.subheader("🧮 Coeficientes de Expansión (Sistema Inglés)")
gamma_aceite = st.sidebar.number_input("γ (Aceite) [F⁻¹]", value=0.0003611, format="%.7f", help="Coeficiente expansión aceite en sistema inglés")
beta_aire = st.sidebar.number_input("β (Aire) [F⁻¹]", value=0.001894, format="%.6f", help="Coeficiente expansión aire en sistema inglés")
factor_seguridad = st.sidebar.number_input("Factor Seguridad CFM", value=1.4, format="%.1f", help="Factor de seguridad aplicado al CFM calculado")

# Parámetros de iteración (Según metodología aprobada)
st.sidebar.subheader("🔄 Parámetros de Calibración")
porcentajes_aceite = st.sidebar.multiselect(
    "% Aceite a iterar", 
    [25, 30, 35, 40, 45, 50, 55], 
    default=[30, 35, 40, 45, 50],
    help="Rangos de llenado de aceite para calibración del modelo"
)
tiempos_expansion = st.sidebar.multiselect(
    "Tiempos expansión [min]", 
    [6, 8, 10, 12, 15], 
    default=[8, 10, 12],
    help="Tiempos de expansión térmica para diferentes escenarios"
)

st.sidebar.markdown("---")
st.sidebar.info("📋 **Flujo Inteligente v2.2:**\n1. Análisis preliminar automático\n2. Inputs individuales por equipo\n3. Filtro espacio real\n4. Modificaciones requeridas\n5. Análisis de robustez")

# Base de datos de respiradores con especificaciones
RESPIRADORES_DB = {
    # X-Series
    "X-100": {"altura": 6.25, "cfm_max": 10, "serie": "X-Series", "cartucho": "L-143", "precio_relativo": 1},
    "X-50": {"altura": 5, "cfm_max": 10, "serie": "X-Series", "cartucho": "L-50", "precio_relativo": 1},
    "X-501": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-341", "precio_relativo": 2},
    "X-502": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-342", "precio_relativo": 2},
    "X-503": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-345", "precio_relativo": 2},
    "X-505": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-346", "precio_relativo": 2},
    "X-521": {"altura": 7, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-343", "precio_relativo": 2},
    "X-522": {"altura": 10, "cfm_max": 20, "serie": "X-Series", "cartucho": "A-344", "precio_relativo": 2},
    
    # Guardian Series
    "G5S1N": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5S", "precio_relativo": 3},
    "G5S1NC": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5SC", "precio_relativo": 3},
    "G5S1NG": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5S", "precio_relativo": 3},
    "G5S1NGC": {"altura": 9, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC5SC", "precio_relativo": 3},
    "G8S1N": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8S", "precio_relativo": 4},
    "G8S1NC": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8SC", "precio_relativo": 4},
    "G8S1NG": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8S", "precio_relativo": 4},
    "G8S1NGC": {"altura": 12, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC8SC", "precio_relativo": 4},
    "G12S1N": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12S", "precio_relativo": 5},
    "G12S1NC": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12SC", "precio_relativo": 5},
    "G12S1NG": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12S", "precio_relativo": 5},
    "G12S1NGC": {"altura": 16, "cfm_max": 25, "serie": "Guardian", "cartucho": "GRC12SC", "precio_relativo": 5},
    
    # D-Series
    "D-100": {"altura": 3.5, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-101": {"altura": 5, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-102": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-103": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-104": {"altura": 8, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-108": {"altura": 10, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1},
    "D-2": {"altura": 10, "cfm_max": 20, "serie": "D-Series", "cartucho": "Desechable", "precio_relativo": 1}
}

# Estado de la aplicación
if 'analysis_stage' not in st.session_state:
    st.session_state.analysis_stage = 'upload'
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
if 'equipos_candidatos' not in st.session_state:
    st.session_state.equipos_candidatos = None
if 'inputs_faltantes' not in st.session_state:
    st.session_state.inputs_faltantes = {}

# Funciones principales
@st.cache_data
def cargar_datos(uploaded_file):
    """Cargar y procesar el archivo Excel"""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None

def identificar_equipos_candidatos(df):
    """Identificar equipos únicos que requieren respiradores"""
    # Filtrar registros candidatos
    mask = (
        (df['ComponentTemplate'].str.contains('gearbox', case=False, na=False)) |
        (df['ComponentTemplate'].str.contains('oil', case=False, na=False)) |
        (df['(D) Oil Capacity'].fillna(0) > 0) |
        (df['MaintPointTemplate'].str.contains('oil', case=False, na=False))
    )
    candidatos = df[mask].copy()
    
    # Agrupar por Machine + Component para evitar duplicados
    equipos_unicos = []
    grupos = candidatos.groupby(['Machine', 'Component'])
    
    for (machine, component), grupo in grupos:
        # Buscar el registro principal (con datos de aceite o el más completo)
        registro_principal = None
        
        # Prioridad 1: Registro con Oil Capacity
        registros_con_aceite = grupo[grupo['(D) Oil Capacity'].fillna(0) > 0]
        if len(registros_con_aceite) > 0:
            registro_principal = registros_con_aceite.iloc[0]
        
        # Prioridad 2: Registro con MaintPointTemplate que contenga 'oil'
        elif len(grupo[grupo['MaintPointTemplate'].str.contains('oil', case=False, na=False)]) > 0:
            registro_principal = grupo[grupo['MaintPointTemplate'].str.contains('oil', case=False, na=False)].iloc[0]
        
        # Prioridad 3: Primer registro del grupo
        else:
            registro_principal = grupo.iloc[0]
        
        # Consolidar datos de todo el grupo en el registro principal
        registro_consolidado = registro_principal.copy()
        
        # Buscar datos faltantes en otros registros del mismo grupo
        for idx, row in grupo.iterrows():
            for col in ['(D) Average Relative Humidity', '(D) Contaminant Abrasive Index', 
                       '(D) Contaminant Likelihood', '(D) Vibration', '(D) Oil Mist Evidence on Headspace',
                       '(D) Water Contact Conditions', '(D) Environmental Conditions',
                       '(D) Breather/Fill Port Clearance', '(D) Breather/Fill Port Size']:
                if pd.isna(registro_consolidado.get(col)) and pd.notna(row.get(col)):
                    registro_consolidado[col] = row[col]
        
        equipos_unicos.append(registro_consolidado)
    
    return pd.DataFrame(equipos_unicos), candidatos

def analizar_inputs_faltantes_por_equipo(equipos_df):
    """Analizar qué inputs manuales faltan por cada equipo individual"""
    equipos_con_faltantes = []
    resumen_faltantes = {}
    
    for idx, row in equipos_df.iterrows():
        machine = row['Machine']
        component = row['Component']
        equipo_key = f"{machine}_{component}"
        
        faltantes_equipo = {}
        
        # 1. CRITICIDAD - Siempre falta (client provided)
        faltantes_equipo['criticidad'] = {
            'requerido': True,
            'razon': 'Client provided (no viene en Data Report)',
            'valor_actual': None
        }
        
        # 2. HIGH PARTICLE REMOVAL - Siempre falta (Noria expert advice)
        faltantes_equipo['high_particle_removal'] = {
            'requerido': True,
            'razon': 'Noria expert advice (no viene en Data Report)',
            'valor_actual': None
        }
        
        # 3. OPERATION RUN HOURS - Falta si no está en data report
        operation_hours = row.get('(D) Operation Run Hours', '')
        if pd.isna(operation_hours) or str(operation_hours).strip() == '':
            faltantes_equipo['operation_hours'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report',
                'valor_actual': None
            }
        
        # 4. OIL DRYNESS LIMIT - Falta si no está en data report
        oil_dryness = row.get('(D) Oil Dryness Limit', '')
        if pd.isna(oil_dryness) or str(oil_dryness).strip() == '':
            faltantes_equipo['oil_dryness_limit'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report',
                'valor_actual': None
            }
        
        # 5. RELATIVE HUMIDITY - Solo si falta
        humidity = row.get('(D) Average Relative Humidity', '')
        if pd.isna(humidity) or str(humidity).strip() == '':
            faltantes_equipo['relative_humidity'] = {
                'requerido': True,
                'razon': 'No encontrado en Data Report (requerido para selección desecante)',
                'valor_actual': None
            }
        
        if faltantes_equipo:
            equipos_con_faltantes.append({
                'machine': machine,
                'component': component,
                'equipo_key': equipo_key,
                'oil_capacity': row.get('(D) Oil Capacity', 0),
                'faltantes': faltantes_equipo
            })
            
            # Agregar al resumen
            for input_name in faltantes_equipo.keys():
                if input_name not in resumen_faltantes:
                    resumen_faltantes[input_name] = 0
                resumen_faltantes[input_name] += 1
    
    return equipos_con_faltantes, resumen_faltantes

def extraer_espacio_disponible(clearance_str):
    """Extraer espacio disponible en inches"""
    if pd.isna(clearance_str):
        return None
    
    clearance_str = str(clearance_str).lower()
    
    # Patrones de búsqueda
    if 'no port' in clearance_str or 'no available' in clearance_str:
        return 0  # Sin puerto
    elif '2 to 4' in clearance_str or '2-4' in clearance_str:
        return 4  # Máximo del rango
    elif '4 to 6' in clearance_str or '4-6' in clearance_str:
        return 6
    elif '6 to 8' in clearance_str or '6-8' in clearance_str:
        return 8
    elif '8 to 10' in clearance_str or '8-10' in clearance_str:
        return 10
    elif 'greater than 12' in clearance_str or '>12' in clearance_str:
        return 20  # Espacio libre
    elif 'greater than 10' in clearance_str or '>10' in clearance_str:
        return 15
    elif 'greater than 6' in clearance_str or '>6' in clearance_str:
        return 12
    elif 'free' in clearance_str or 'unlimited' in clearance_str:
        return 30  # Espacio libre
    elif 'limited' in clearance_str:
        return 5  # Espacio muy limitado
    
    # Buscar números específicos
    match = re.search(r'(\d+\.?\d*)', clearance_str)
    if match:
        return float(match.group(1))
    
    return None  # No se pudo determinar

def filtrar_respiradores_por_espacio(espacio_disponible, cfm_requerido):
    """Filtrar respiradores que realmente caben en el espacio disponible"""
    if espacio_disponible is None:
        # Si no se puede determinar el espacio, asumir espacio limitado
        espacio_disponible = 8
    
    respiradores_compatibles = []
    
    for modelo, specs in RESPIRADORES_DB.items():
        # Verificar si cabe físicamente
        if specs['altura'] <= espacio_disponible:
            # Verificar si maneja el CFM requerido
            if specs['cfm_max'] >= cfm_requerido:
                respiradores_compatibles.append({
                    'modelo': modelo,
                    'altura': specs['altura'],
                    'cfm_max': specs['cfm_max'],
                    'serie': specs['serie'],
                    'cartucho': specs['cartucho'],
                    'precio_relativo': specs['precio_relativo'],
                    'margen_espacio': espacio_disponible - specs['altura']
                })
    
    # Ordenar por serie (Guardian > X-Series > D-Series) y luego por altura
    orden_serie = {'Guardian': 3, 'X-Series': 2, 'D-Series': 1}
    respiradores_compatibles.sort(key=lambda x: (orden_serie.get(x['serie'], 0), -x['altura']))
    
    return respiradores_compatibles

def generar_modificaciones_requeridas(row, respiradores_compatibles, cfm_requerido, inputs_manuales_equipo, criticidad_impacto):
    """Generar lista de modificaciones requeridas y recomendación post-modificaciones"""
    modificaciones = []
    
    # Extraer datos del equipo
    clearance = str(row.get('(D) Breather/Fill Port Clearance', '')).lower()
    port_size = str(row.get('(D) Breather/Fill Port Size', '')).lower()
    port_available = str(row.get('(D) Port Availability', '')).lower()
    espacio_disponible = extraer_espacio_disponible(clearance)
    
    # 1. Verificar puerto disponible
    if 'no port' in clearance or 'no' in port_available or espacio_disponible == 0:
        if '1/2' in port_size or '0.5' in port_size:
            modificaciones.append("🔧 Instalar puerto 1/2\" NPT en tapa del equipo")
        elif '1\"' in port_size or '1 inch' in port_size:
            modificaciones.append("🔧 Instalar puerto 1\" NPT en tapa del equipo")
        elif '2\"' in port_size or '2 inch' in port_size:
            modificaciones.append("🔧 Instalar puerto 2\" NPT en tapa del equipo")
        else:
            modificaciones.append("🔧 Instalar puerto 1\" NPT en tapa del equipo (recomendado)")
    
    # 2. Verificar espacio suficiente y generar recomendación post-modificaciones
    recomendacion_post_modificaciones = None
    
    if not respiradores_compatibles:
        # No hay respiradores que quepan - necesita más espacio
        
        # Filtrar respiradores según high particle removal
        todos_respiradores = list(RESPIRADORES_DB.items())
        if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
            # Excluir D-Series para high particle removal
            todos_respiradores = [(k, v) for k, v in todos_respiradores if v['serie'] != 'D-Series']
        
        # Filtrar por criticidad
        series_permitidas = criticidad_impacto['series_preferidas']
        respiradores_filtrados = [(k, v) for k, v in todos_respiradores if v['serie'] in series_permitidas]
        
        if not respiradores_filtrados:
            respiradores_filtrados = todos_respiradores
        
        # Encontrar el más pequeño que maneje el CFM
        respiradores_adecuados = [(k, v) for k, v in respiradores_filtrados if v['cfm_max'] >= cfm_requerido]
        
        if respiradores_adecuados:
            # Ordenar por altura (más pequeño primero)
            respirador_recomendado = min(respiradores_adecuados, key=lambda x: x[1]['altura'])
            modelo_rec = respirador_recomendado[0]
            specs_rec = respirador_recomendado[1]
            
            recomendacion_post_modificaciones = {
                'modelo': modelo_rec,
                'serie': specs_rec['serie'],
                'cartucho': specs_rec['cartucho'],
                'altura': specs_rec['altura'],
                'cfm_capacidad': specs_rec['cfm_max']
            }
            
            if espacio_disponible is not None and espacio_disponible > 0:
                modificaciones.append(f"📏 Crear espacio mínimo {specs_rec['altura']}\" vertical (actual: {espacio_disponible}\")")
            else:
                modificaciones.append(f"📏 Crear espacio mínimo {specs_rec['altura']}\" vertical para instalación")
        else:
            # Ningún respirador maneja el CFM requerido
            modelo_minimo = min(RESPIRADORES_DB.values(), key=lambda x: x['altura'])
            modificaciones.append(f"📏 Crear espacio mínimo {modelo_minimo['altura']}\" vertical para instalación")
            modificaciones.append(f"⚡ CFM requerido ({cfm_requerido:.2f}) excede capacidad disponible")
        
        # Opción de instalación remota
        modificaciones.append("🔄 ALTERNATIVA: Instalación remota con línea de conexión")
    
    # 3. Verificar CFM adecuado
    if respiradores_compatibles:
        max_cfm_disponible = max([r['cfm_max'] for r in respiradores_compatibles])
        if cfm_requerido > max_cfm_disponible:
            modificaciones.append(f"⚡ CFM requerido ({cfm_requerido:.2f}) excede capacidad máxima disponible ({max_cfm_disponible})")
            modificaciones.append("🔄 Considerar: Múltiples respiradores o modelo de mayor capacidad")
    
    # 4. Recomendaciones de acceso
    if respiradores_compatibles and any('Guardian' in r['serie'] for r in respiradores_compatibles):
        modificaciones.append("🔧 Asegurar acceso para mantenimiento de cartuchos reemplazables")
    elif recomendacion_post_modificaciones and recomendacion_post_modificaciones['serie'] == 'Guardian':
        modificaciones.append("🔧 Asegurar acceso para mantenimiento de cartuchos reemplazables")
    
    # 5. Consideraciones ambientales
    env_conditions = str(row.get('(D) Environmental Conditions', '')).lower()
    if 'outside' in env_conditions or 'exterior' in env_conditions:
        modificaciones.append("🌧️ Considerar protección adicional contra intemperie")
    
    return modificaciones, recomendacion_post_modificaciones

def extraer_valor_numerico(valor_str, tipo="porcentaje"):
    """Extraer valor numérico de strings con rangos o texto"""
    if pd.isna(valor_str):
        return None
    
    valor_str = str(valor_str).lower()
    
    if tipo == "porcentaje":
        # Buscar números seguidos de %
        match = re.search(r'(\d+\.?\d*)%', valor_str)
        if match:
            return float(match.group(1))
        # Buscar solo números si no hay %
        match = re.search(r'(\d+\.?\d*)', valor_str)
        if match:
            return float(match.group(1))
    elif tipo == "indice":
        # Para contamination index, buscar rangos o números
        if 'heavy' in valor_str or 'mining' in valor_str:
            return 80  # Asumimos alto
        elif 'medium' in valor_str:
            return 50
        elif 'light' in valor_str or 'low' in valor_str:
            return 20
        # Buscar números
        match = re.search(r'(\d+\.?\d*)', valor_str)
        if match:
            return float(match.group(1))
    
    return None

def extraer_temperaturas(temp_string):
    """Extraer temperaturas mínima y máxima de string - SISTEMA INGLÉS"""
    if pd.isna(temp_string):
        return 77, 180, 103  # °F por defecto
    
    # Buscar patrones de temperatura en °F y °C
    temp_pattern_f = r'(\d+\.?\d*)°F.*?(\d+\.?\d*)°F'
    temp_pattern_c = r'(\d+\.?\d*)°C.*?(\d+\.?\d*)°C'
    
    match_f = re.search(temp_pattern_f, str(temp_string))
    match_c = re.search(temp_pattern_c, str(temp_string))
    
    if match_f:
        temp_min = float(match_f.group(1))
        temp_max = float(match_f.group(2))
    elif match_c:
        # Convertir °C a °F
        temp_min_c = float(match_c.group(1))
        temp_max_c = float(match_c.group(2))
        temp_min = temp_min_c * 9/5 + 32
        temp_max = temp_max_c * 9/5 + 32
    else:
        temp_min = 77   # °F (~25°C)
        temp_max = 180  # °F (~82°C)
    
    delta_t = temp_max - temp_min
    return temp_min, temp_max, delta_t

def calcular_cfm_con_iteraciones(row, porcentajes, tiempos, gamma, beta, factor_seguridad):
    """Calcular CFM con análisis de robustez de iteraciones"""
    oil_capacity_gal = row.get('(D) Oil Capacity', 0)
    
    if pd.isna(oil_capacity_gal) or oil_capacity_gal <= 0:
        return None
    
    # Extraer temperaturas (en °F)
    temp_min, temp_max, delta_t = extraer_temperaturas(
        row.get('(D) Operating Temperature')
    )
    
    # Obtener dimensiones (convertir a sistema inglés si es necesario)
    length = row.get('(D) Length', np.nan)
    width = row.get('(D) Width', np.nan)
    height = row.get('(D) Height', np.nan)
    
    # Determinar método de cálculo
    metodo = "Estimativo"
    if pd.notna(length) and pd.notna(width) and pd.notna(height) and all(x > 0 for x in [length, width, height]):
        metodo = "Directo"
    elif pd.notna(length) and pd.notna(height) and length > 0 and height > 0:
        metodo = "Inverso"
    
    # Calcular TODAS las iteraciones para análisis de robustez
    iteraciones_detalladas = []
    
    for pct in porcentajes:
        for tiempo in tiempos:
            # Calcular volumen total según método
            if metodo == "Directo":
                volumen_total_gal = (length * width * height) * 2.64172e-7
            elif metodo == "Inverso":
                volumen_total_gal = oil_capacity_gal / (pct / 100)
            else:
                volumen_total_gal = oil_capacity_gal / (pct / 100)
            
            volumen_aceite_gal = volumen_total_gal * (pct / 100)
            volumen_aire_gal = volumen_total_gal - volumen_aceite_gal
            
            # Expansión térmica (Sistema Inglés)
            delta_vo = gamma * volumen_aceite_gal * delta_t
            delta_va = beta * volumen_aire_gal * delta_t
            v_exp_total = (delta_vo + delta_va) * factor_seguridad
            
            # Convertir a ft³ y calcular CFM
            v_exp_ft3 = v_exp_total * 0.133681
            cfm = v_exp_ft3 / tiempo
            
            iteraciones_detalladas.append({
                'porcentaje': pct,
                'tiempo': tiempo,
                'volumen_total': volumen_total_gal,
                'volumen_aceite': volumen_aceite_gal,
                'volumen_aire': volumen_aire_gal,
                'cfm': cfm
            })
    
    # Calcular estadísticas
    cfms = [it['cfm'] for it in iteraciones_detalladas]
    cfm_min = min(cfms)
    cfm_max = max(cfms)
    cfm_promedio = np.mean(cfms)
    variacion_cfm = (cfm_max - cfm_min) / cfm_promedio * 100
    
    return {
        'metodo': metodo,
        'temp_min': temp_min,
        'temp_max': temp_max,
        'delta_t': delta_t,
        'cfm_min': cfm_min,
        'cfm_max': cfm_max,
        'cfm_promedio': cfm_promedio,
        'cfm_rango': f"{cfm_min:.3f} - {cfm_max:.3f}",
        'variacion_porcentual': variacion_cfm,
        'iteraciones_detalladas': iteraciones_detalladas,
        'total_iteraciones': len(iteraciones_detalladas)
    }

def determinar_tipo_sistema(row):
    """Determinar si el sistema es estático o circulatorio"""
    flow_rate = row.get('(D) Flow Rate', np.nan)
    if pd.notna(flow_rate) and flow_rate > 0:
        return "Circulatorio"
    else:
        return "Estático (Splash)"

def aplicar_criterios_seleccion(row, inputs_manuales_equipo):
    """Aplicar criterios de selección según Excel aprobado"""
    criterios = {
        'filtro_particulas': False,
        'desecante': False,
        'desecante_alta_capacidad': False,
        'valvula_check': False,
        'resistente_vibracion': False,
        'camara_expansion': False,
        'control_niebla': False,
        'cartucho_reemplazable': False
    }
    
    # Extraer y procesar datos del Data Report
    humidity = extraer_valor_numerico(row.get('(D) Average Relative Humidity', ''), "porcentaje")
    contaminant_index = extraer_valor_numerico(row.get('(D) Contaminant Abrasive Index', ''), "indice")
    contaminant_likelihood = str(row.get('(D) Contaminant Likelihood', '')).lower()
    vibration = str(row.get('(D) Vibration', '')).lower()
    water_contact = str(row.get('(D) Water Contact Conditions', '')).lower()
    oil_mist = row.get('(D) Oil Mist Evidence on Headspace', False)
    
    # Usar inputs manuales específicos del equipo si los datos faltan
    if humidity is None and 'relative_humidity' in inputs_manuales_equipo:
        humidity = inputs_manuales_equipo['relative_humidity']
    elif humidity is None:
        humidity = 50  # % por defecto
    
    if contaminant_index is None:
        contaminant_index = 40  # Índice medio por defecto
    
    # 1. FILTRO DE PARTÍCULAS (Según Excel aprobado)
    # Siempre requerido según la lógica del Excel
    criterios['filtro_particulas'] = True
    
    # 2. CARTUCHO REEMPLAZABLE (Part_Cartr)
    # CI >75 AND CL=High → Required, Part_Cartr
    if contaminant_index > 75 and 'high' in contaminant_likelihood:
        criterios['cartucho_reemplazable'] = True
    
    # High particle removal capacity (input manual específico del equipo)
    if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
        criterios['cartucho_reemplazable'] = True  # High particle removal = cartuchos reemplazables obligatorios
    
    # 3. DESECANTE (Según rangos específicos del Excel)
    if humidity > 75:
        criterios['desecante'] = True
        criterios['desecante_alta_capacidad'] = True
        criterios['valvula_check'] = True  # HR >75% → Desiccant (High capacity)+Check valve
    elif humidity > 50:
        criterios['desecante'] = True
    # HR <30% → Not required (no desecante)
    
    # 4. CHECK VALVE adicional
    # WWSH + → Desiccant +Check valve
    if 'water' in water_contact and 'no water' not in water_contact:
        criterios['valvula_check'] = True
        criterios['desecante'] = True
    
    # 5. RESISTENTE A VIBRACIÓN
    # Vib > 0.4 ips → Vibration resistant breather
    if 'ips' in vibration:
        try:
            vib_value = float(re.search(r'(\d+\.?\d*)', vibration).group(1))
            if vib_value > 0.4:
                criterios['resistente_vibracion'] = True
        except:
            pass
    elif '>0.4' in vibration or 'high' in vibration:
        criterios['resistente_vibracion'] = True
    
    # 6. CÁMARA DE EXPANSIÓN
    # NOT recommended (según Excel)
    criterios['camara_expansion'] = False
    
    # 7. CONTROL DE NIEBLA
    # Mist + → Vent guard
    if oil_mist == True or str(oil_mist).lower() == 'yes':
        criterios['control_niebla'] = True
    
    return criterios

def determinar_criticidad_impacto(criticidad_str):
    """Determinar impacto de criticidad en la selección - LÓGICA TEMPORAL"""
    if 'A' in criticidad_str or 'critica' in criticidad_str.lower():
        return {
            'nivel': 'A (Crítica)',
            'prioridad': 'Crítica',
            'series_preferidas': ['Guardian'],
            'cartucho_obligatorio': True,
            'modificador_conservador': True
        }
    elif 'B' in criticidad_str or 'intermedia' in criticidad_str.lower():
        return {
            'nivel': 'B (Intermedia)',
            'prioridad': 'Media',
            'series_preferidas': ['Guardian', 'X-Series'],
            'cartucho_obligatorio': False,
            'modificador_conservador': False
        }
    else:  # Criticidad C
        return {
            'nivel': 'C (Baja)',
            'prioridad': 'Baja',
            'series_preferidas': ['Guardian', 'X-Series', 'D-Series'],
            'cartucho_obligatorio': False,
            'modificador_conservador': False
        }

def seleccionar_respirador_final(respiradores_compatibles, criterios, criticidad_impacto, cfm_promedio, inputs_manuales_equipo):
    """Seleccionar el respirador final basado en criterios y criticidad"""
    if not respiradores_compatibles:
        return None
    
    # Filtrar por series permitidas según criticidad
    series_permitidas = criticidad_impacto['series_preferidas']
    candidatos_finales = [r for r in respiradores_compatibles if r['serie'] in series_permitidas]
    
    if not candidatos_finales:
        # Si no hay candidatos de las series preferidas, tomar el mejor disponible
        candidatos_finales = respiradores_compatibles
    
    # Aplicar criterios específicos
    if criterios['cartucho_reemplazable'] or criticidad_impacto['cartucho_obligatorio']:
        # Preferir Guardian y X-Series (no D-Series para cartuchos reemplazables)
        candidatos_finales = [r for r in candidatos_finales if r['serie'] != 'D-Series']
    
    # HIGH PARTICLE REMOVAL = CARTUCHOS REEMPLAZABLES OBLIGATORIOS
    if inputs_manuales_equipo.get('high_particle_removal') == 'Yes':
        # Excluir D-Series cuando se requiere high particle removal
        candidatos_finales = [r for r in candidatos_finales if r['serie'] != 'D-Series']
    
    if criterios['desecante_alta_capacidad']:
        # Buscar modelos con check valve integrada
        candidatos_check = [r for r in candidatos_finales if 'C' in r['modelo']]
        if candidatos_check:
            candidatos_finales = candidatos_check
    
    if not candidatos_finales:
        # Si no quedan candidatos, regresar el mejor compatible
        candidatos_finales = respiradores_compatibles
    
    # Seleccionar el mejor candidato
    # Prioridad: 1) Serie según criticidad 2) CFM adecuado 3) Menor altura para instalación
    orden_serie = {'Guardian': 3, 'X-Series': 2, 'D-Series': 1}
    
    mejor_candidato = min(candidatos_finales, key=lambda x: (
        -orden_serie.get(x['serie'], 0),  # Serie preferida primero
        abs(x['cfm_max'] - cfm_promedio * 1.2),  # CFM cercano al requerido + margen
        x['altura']  # Menor altura para facilitar instalación
    ))
    
    return mejor_candidato

def procesar_equipo_individual(row, inputs_manuales_equipo, porcentajes, tiempos, gamma, beta, factor_seguridad):
    """Procesar un equipo individual con todos los análisis"""
    machine = row['Machine']
    component = row['Component']
    
    # Determinar criticidad
    criticidad_str = inputs_manuales_equipo.get('criticidad', 'B (Intermedia)')
    criticidad_impacto = determinar_criticidad_impacto(criticidad_str)
    
    # Determinar tipo de sistema
    tipo_sistema = determinar_tipo_sistema(row)
    
    # Calcular CFM con análisis de robustez
    cfm_data = calcular_cfm_con_iteraciones(row, porcentajes, tiempos, gamma, beta, factor_seguridad)
    
    if not cfm_data:
        return {
            'machine': machine,
            'component': component,
            'error': 'Datos insuficientes de Oil Capacity',
            'cfm_data': None,
            'recomendacion': None,
            'modificaciones': ['⚠️ Faltan datos de Oil Capacity para análisis']
        }
    
    # Aplicar criterios de selección
    criterios = aplicar_criterios_seleccion(row, inputs_manuales_equipo)
    
    # Filtrar respiradores por espacio disponible
    espacio_disponible = extraer_espacio_disponible(row.get('(D) Breather/Fill Port Clearance'))
    cfm_promedio = cfm_data['cfm_promedio']
    respiradores_compatibles = filtrar_respiradores_por_espacio(espacio_disponible, cfm_promedio)
    
    # Generar modificaciones requeridas y recomendación post-modificaciones
    modificaciones, recomendacion_post_modificaciones = generar_modificaciones_requeridas(
        row, respiradores_compatibles, cfm_promedio, inputs_manuales_equipo, criticidad_impacto
    )
    
    # Seleccionar respirador final
    respirador_seleccionado = seleccionar_respirador_final(
        respiradores_compatibles, criterios, criticidad_impacto, cfm_promedio, inputs_manuales_equipo
    )
    
    # Análisis de robustez
    variacion_cfm = cfm_data['variacion_porcentual']
    if variacion_cfm < 15:
        robustez = f"✅ ROBUSTO - Variación CFM: {variacion_cfm:.1f}%"
    elif variacion_cfm < 30:
        robustez = f"⚠️ MODERADO - Variación CFM: {variacion_cfm:.1f}%"
    else:
        robustez = f"🔍 VARIABLE - Variación CFM: {variacion_cfm:.1f}% (requiere datos más precisos)"
    
    # Generar recomendación final
    if respirador_seleccionado:
        recomendacion = {
            'modelo_recomendado': respirador_seleccionado['modelo'],
            'serie_recomendada': respirador_seleccionado['serie'],
            'cartucho': respirador_seleccionado['cartucho'],
            'altura_requerida': f"{respirador_seleccionado['altura']}\"",
            'cfm_capacidad': respirador_seleccionado['cfm_max'],
            'cfm_requerido': f"{cfm_promedio:.2f}",
            'cfm_rango': cfm_data['cfm_rango'],
            'prioridad': criticidad_impacto['prioridad'],
            'criticidad': criticidad_impacto['nivel'],
            'robustez': robustez,
            'margen_cfm': f"{((respirador_seleccionado['cfm_max'] / cfm_promedio - 1) * 100):.1f}%",
            'espacio_disponible': f"{espacio_disponible}\"" if espacio_disponible else "No determinado",
            'margen_espacio': f"{respirador_seleccionado['margen_espacio']:.1f}\"" if espacio_disponible else "N/A",
            'recomendacion_post_modificaciones': None
        }
    else:
        recomendacion = {
            'modelo_recomendado': 'REQUIERE MODIFICACIONES',
            'serie_recomendada': 'N/A',
            'cartucho': 'N/A',
            'altura_requerida': 'Ver modificaciones',
            'cfm_capacidad': 'N/A',
            'cfm_requerido': f"{cfm_promedio:.2f}",
            'cfm_rango': cfm_data['cfm_rango'],
            'prioridad': 'CRÍTICA - Modificaciones requeridas',
            'criticidad': criticidad_impacto['nivel'],
            'robustez': robustez,
            'margen_cfm': 'N/A',
            'espacio_disponible': f"{espacio_disponible}\"" if espacio_disponible else "No determinado",
            'margen_espacio': 'Insuficiente',
            'recomendacion_post_modificaciones': recomendacion_post_modificaciones
        }
    
    return {
        'machine': machine,
        'component': component,
        'oil_capacity': row.get('(D) Oil Capacity', 0),
        'tipo_sistema': tipo_sistema,
        'cfm_data': cfm_data,
        'criterios': criterios,
        'criticidad_impacto': criticidad_impacto,
        'respiradores_compatibles': respiradores_compatibles,
        'respirador_seleccionado': respirador_seleccionado,
        'recomendacion': recomendacion,
        'modificaciones': modificaciones,
        'inputs_utilizados': inputs_manuales_equipo
    }

# INTERFAZ PRINCIPAL
def main():
    
    # Mostrar información sobre versión
    with st.expander("🔄 Cambios en v2.2 - Implementación Final Completa", expanded=False):
        st.markdown("""
        **✅ Implementación Final:**
        - **Inputs individuales por equipo**: Todos los parámetros faltantes específicos por equipo
        - **Filtro espacio real**: Solo recomienda respiradores que realmente caben
        - **Modificaciones específicas**: Lista detallada de modificaciones requeridas por equipo
        - **Base de datos completa**: Catálogo integrado de X-Series, Guardian y D-Series
        - **Análisis de robustez**: Validación de variabilidad CFM en recomendaciones
        - **Output individual**: Recomendación + modificaciones por cada equipo
        
        **🎯 Flujo Final:**
        1. Análisis preliminar automático
        2. Inputs específicos por equipo faltante  
        3. Filtrado por espacio disponible real
        4. Recomendación final + modificaciones requeridas
        """)
    
    # ETAPA 1: UPLOAD Y ANÁLISIS PRELIMINAR
    if st.session_state.analysis_stage == 'upload':
        st.header("📤 Subir Data Report")
        uploaded_file = st.file_uploader(
            "Selecciona tu archivo Excel",
            type=['xlsx', 'xls'],
            help="Sube el data report en formato Excel para análisis preliminar"
        )
        
        if uploaded_file is not None:
            with st.spinner("Cargando y analizando datos..."):
                df = cargar_datos(uploaded_file)
                
                if df is not None:
                    st.session_state.df_original = df
                    
                    # Identificar equipos candidatos
                    equipos_unicos, todos_candidatos = identificar_equipos_candidatos(df)
                    st.session_state.equipos_candidatos = equipos_unicos
                    
                    # Analizar qué inputs faltan por equipo
                    equipos_con_faltantes, resumen_faltantes = analizar_inputs_faltantes_por_equipo(equipos_unicos)
                    st.session_state.equipos_con_faltantes = equipos_con_faltantes
                    st.session_state.resumen_faltantes = resumen_faltantes
                    
                    st.success(f"✅ Análisis preliminar completado: {len(equipos_unicos)} equipos candidatos")
                    
                    # Mostrar resumen del análisis preliminar
                    st.subheader("📋 Resumen del Análisis Preliminar")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Registros Total", len(df))
                    with col2:
                        st.metric("Equipos Candidatos", len(equipos_unicos))
                    with col3:
                        st.metric("Equipos con Inputs Faltantes", len(equipos_con_faltantes))
                    
                    # Mostrar equipos candidatos
                    with st.expander("👀 Equipos Candidatos Identificados", expanded=True):
                        display_cols = ['Machine', 'Component', '(D) Oil Capacity', '(D) Breather/Fill Port Clearance']
                        available_cols = [col for col in display_cols if col in equipos_unicos.columns]
                        st.dataframe(equipos_unicos[available_cols], use_container_width=True)
                    
                    # Mostrar resumen de inputs faltantes
                    if resumen_faltantes:
                        st.subheader("⚠️ Resumen de Inputs Manuales Requeridos")
                        for input_name, count in resumen_faltantes.items():
                            st.warning(f"**{input_name.replace('_', ' ').title()}**: {count} equipos requieren este input")
                        
                        if st.button("➡️ Continuar a Solicitud de Inputs por Equipo", type="primary"):
                            st.session_state.analysis_stage = 'inputs'
                            st.rerun()
                    else:
                        st.info("🎉 Todos los datos necesarios están disponibles en el Data Report")
                        if st.button("➡️ Continuar a Análisis Final", type="primary"):
                            st.session_state.analysis_stage = 'analysis'
                            st.rerun()
    
    # ETAPA 2: SOLICITAR INPUTS FALTANTES POR EQUIPO
    elif st.session_state.analysis_stage == 'inputs':
        st.header("📝 Proporcionar Datos Faltantes por Equipo")
        
        st.info(f"Se requieren inputs manuales para {len(st.session_state.equipos_con_faltantes)} equipos")
        
        inputs_manuales_por_equipo = {}
        
        # Procesar cada equipo con inputs faltantes
        for equipo_info in st.session_state.equipos_con_faltantes:
            machine = equipo_info['machine']
            component = equipo_info['component']
            equipo_key = equipo_info['equipo_key']
            oil_capacity = equipo_info['oil_capacity']
            faltantes = equipo_info['faltantes']
            
            st.subheader(f"🔧 {machine} - {component}")
            st.write(f"**Oil Capacity**: {oil_capacity} gal")
            
            inputs_equipo = {}
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Criticidad
                if 'criticidad' in faltantes:
                    criticidad = st.selectbox(
                        "🎯 Criticidad",
                        ["A (Crítica)", "B (Intermedia)", "C (Baja)"],
                        index=1,
                        key=f"crit_{equipo_key}",
                        help="Nivel de criticidad del equipo para operación"
                    )
                    inputs_equipo['criticidad'] = criticidad
                
                # High Particle Removal
                if 'high_particle_removal' in faltantes:
                    high_particle = st.selectbox(
                        "🔧 High Particle Removal",
                        ["No", "Yes"],
                        key=f"particle_{equipo_key}",
                        help="Basado en recomendación experta Noria"
                    )
                    inputs_equipo['high_particle_removal'] = high_particle
                
                # Operation Hours
                if 'operation_hours' in faltantes:
                    operation_hours = st.selectbox(
                        "⏰ Operation Hours",
                        ["Continuous", "Stand By", "Intermittent"],
                        key=f"hours_{equipo_key}",
                        help="Patrón operativo del equipo"
                    )
                    inputs_equipo['operation_hours'] = operation_hours
            
            with col2:
                # Oil Dryness Limit
                if 'oil_dryness_limit' in faltantes:
                    oil_dryness = st.selectbox(
                        "💧 Oil Dryness Limit",
                        ["< 200 ppm", "200-500 ppm", "> 500 ppm"],
                        index=1,
                        key=f"dryness_{equipo_key}",
                        help="Límite de sequedad del aceite requerido"
                    )
                    inputs_equipo['oil_dryness_limit'] = oil_dryness
                
                # Relative Humidity
                if 'relative_humidity' in faltantes:
                    humidity = st.number_input(
                        "💨 Relative Humidity (%)",
                        min_value=0,
                        max_value=100,
                        value=50,
                        key=f"humidity_{equipo_key}",
                        help="Humedad relativa promedio del ambiente donde está el equipo"
                    )
                    inputs_equipo['relative_humidity'] = humidity
            
            inputs_manuales_por_equipo[equipo_key] = inputs_equipo
            st.markdown("---")
        
        # Botón para continuar
        if st.button("🚀 Ejecutar Análisis Completo Final", type="primary"):
            st.session_state.inputs_manuales_por_equipo = inputs_manuales_por_equipo
            st.session_state.analysis_stage = 'analysis'
            st.rerun()
    
    # ETAPA 3: ANÁLISIS COMPLETO FINAL
    elif st.session_state.analysis_stage == 'analysis':
        st.header("📊 Análisis Completo Final")
        
        with st.spinner("Ejecutando análisis completo con todos los criterios..."):
            # Procesar cada equipo individualmente
            resultados_completos = []
            
            for idx, row in st.session_state.equipos_candidatos.iterrows():
                machine = row['Machine']
                component = row['Component']
                equipo_key = f"{machine}_{component}"
                
                # Obtener inputs manuales específicos del equipo
                inputs_manuales_equipo = st.session_state.inputs_manuales_por_equipo.get(equipo_key, {})
                
                # Procesar equipo individual
                resultado = procesar_equipo_individual(
                    row, inputs_manuales_equipo, porcentajes_aceite, tiempos_expansion,
                    gamma_aceite, beta_aire, factor_seguridad
                )
                
                resultados_completos.append(resultado)
        
        st.success("✅ Análisis completo final terminado!")
        
        # Mostrar métricas generales
        total_equipos = len(resultados_completos)
        equipos_con_recomendacion = len([r for r in resultados_completos if r.get('recomendacion') and r['recomendacion']['modelo_recomendado'] != 'REQUIERE MODIFICACIONES'])
        equipos_requieren_modificaciones = total_equipos - equipos_con_recomendacion
        equipos_robustos = len([r for r in resultados_completos if r.get('cfm_data') and r['cfm_data']['variacion_porcentual'] < 15])
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Equipos", total_equipos)
        with col2:
            st.metric("Con Recomendación Directa", equipos_con_recomendacion)
        with col3:
            st.metric("Requieren Modificaciones", equipos_requieren_modificaciones)
        with col4:
            st.metric("Resultados Robustos", equipos_robustos)
        
        # Mostrar resultados por equipo
        st.subheader("🔧 Resultados Detallados por Equipo")
        
        for resultado in resultados_completos:
            if 'error' in resultado:
                with st.expander(f"❌ {resultado['component']} - {resultado['machine']}", expanded=False):
                    st.error(resultado['error'])
                    for mod in resultado['modificaciones']:
                        st.write(mod)
                continue
            
            # Determinar el ícono según el resultado
            if resultado['recomendacion']['modelo_recomendado'] == 'REQUIERE MODIFICACIONES':
                icono = "🔧"
                expanded = True
            elif resultado['cfm_data']['variacion_porcentual'] > 30:
                icono = "⚠️"
                expanded = False
            else:
                icono = "✅"
                expanded = False
            
            with st.expander(f"{icono} {resultado['component']} - {resultado['machine']}", expanded=expanded):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📋 Datos del Equipo:**")
                    st.write(f"• **Oil Capacity:** {resultado['oil_capacity']} gal")
                    st.write(f"• **Tipo Sistema:** {resultado['tipo_sistema']}")
                    st.write(f"• **Criticidad:** {resultado['recomendacion']['criticidad']}")
                    
                    if resultado['cfm_data']:
                        st.write(f"• **CFM Requerido:** {resultado['recomendacion']['cfm_requerido']}")
                        st.write(f"• **CFM Rango:** {resultado['recomendacion']['cfm_rango']}")
                        st.write(f"• **Espacio Disponible:** {resultado['recomendacion']['espacio_disponible']}")
                
                with col2:
                    st.markdown("**🏆 Recomendación Final:**")
                    rec = resultado['recomendacion']
                    if rec['modelo_recomendado'] != 'REQUIERE MODIFICACIONES':
                        st.write(f"• **Modelo:** {rec['modelo_recomendado']}")
                        st.write(f"• **Serie:** {rec['serie_recomendada']}")
                        st.write(f"• **Cartucho:** {rec['cartucho']}")
                        st.write(f"• **Altura:** {rec['altura_requerida']}")
                        st.write(f"• **Margen CFM:** {rec['margen_cfm']}")
                        st.write(f"• **Margen Espacio:** {rec['margen_espacio']}")
                    else:
                        st.error("**REQUIERE MODIFICACIONES PARA INSTALACIÓN**")
                        
                        # Mostrar recomendación post-modificaciones si existe
                        if rec.get('recomendacion_post_modificaciones'):
                            post_rec = rec['recomendacion_post_modificaciones']
                            st.success("**✅ RECOMENDADO DESPUÉS DE MODIFICACIONES:**")
                            st.write(f"• **Modelo:** {post_rec['modelo']} ({post_rec['serie']})")
                            st.write(f"• **Cartucho:** {post_rec['cartucho']}")
                            st.write(f"• **Altura:** {post_rec['altura']}\"")
                            st.write(f"• **CFM Capacidad:** {post_rec['cfm_capacidad']} (requerido: {rec['cfm_requerido']})")
                    
                    st.write(f"• **Prioridad:** {rec['prioridad']}")
                
                # Análisis de robustez
                st.markdown("**🔍 Análisis de Robustez:**")
                if resultado['cfm_data']['variacion_porcentual'] < 15:
                    st.success(rec['robustez'])
                elif resultado['cfm_data']['variacion_porcentual'] < 30:
                    st.warning(rec['robustez'])
                else:
                    st.error(rec['robustez'])
                    if st.checkbox(f"📊 Ver iteraciones detalladas (Variación: {resultado['cfm_data']['variacion_porcentual']:.1f}%)", key=f"iter_detail_{resultado['machine']}_{resultado['component']}"):
                        st.markdown("**Detalle de todas las iteraciones:**")
                        iteraciones_df = pd.DataFrame(resultado['cfm_data']['iteraciones_detalladas'])
                        st.dataframe(iteraciones_df, use_container_width=True)
                
                # Modificaciones requeridas
                if resultado['modificaciones']:
                    st.markdown("**🔧 Modificaciones Requeridas:**")
                    for modificacion in resultado['modificaciones']:
                        if "ALTERNATIVA" in modificacion:
                            st.info(modificacion)
                        elif "⚡" in modificacion or "🔄" in modificacion:
                            st.warning(modificacion)
                        else:
                            st.write(modificacion)
                
                # Respiradores compatibles (solo si hay opciones)
                if len(resultado['respiradores_compatibles']) > 1:
                    if st.checkbox(f"🔍 Ver otras opciones compatibles ({len(resultado['respiradores_compatibles'])-1} adicionales)", key=f"opciones_{resultado['machine']}_{resultado['component']}"):
                        st.markdown("**Otras opciones que caben en el espacio disponible:**")
                        for resp in resultado['respiradores_compatibles']:
                            if resp['modelo'] != resultado['recomendacion'].get('modelo_recomendado', ''):
                                st.write(f"• **{resp['modelo']}** ({resp['serie']}) - Altura: {resp['altura']}\" - CFM: {resp['cfm_max']} - Margen: {resp['margen_espacio']:.1f}\"")
        
        # Generar descarga
        st.header("💾 Descargar Resultados Completos")
        
        # Preparar DataFrame COMPLETO (Data Report original + nuevas columnas)
        df_completo = st.session_state.df_original.copy()
        
        # Agregar nuevas columnas con valores por defecto
        nuevas_columnas = {
            'Respirador_Requerido': 'No analizado',
            'Criticidad': '',
            'Tipo_Sistema': '',
            'CFM_Requerido': '',
            'CFM_Rango': '',
            'Modelo_Recomendado': '',
            'Serie_Recomendada': '',
            'Cartucho_Recomendado': '',
            'Altura_Requerida': '',
            'Espacio_Disponible': '',
            'Margen_Espacio': '',
            'Margen_CFM': '',
            'Prioridad': '',
            'Robustez': '',
            'Variacion_CFM_Pct': '',
            'Total_Iteraciones': '',
            'Modelo_Post_Modificaciones': '',
            'Serie_Post_Modificaciones': '',
            'Cartucho_Post_Modificaciones': '',
            'Altura_Post_Modificaciones': '',
            'CFM_Capacidad_Post_Modificaciones': '',
            'Modificaciones_Requeridas': '',
            'Status_Analisis': 'No analizado'
        }
        
        for col in nuevas_columnas:
            df_completo[col] = nuevas_columnas[col]
        
        # Actualizar registros analizados con resultados
        for resultado in resultados_completos:
            if 'error' in resultado:
                # Buscar registro en el DataFrame original
                mask = (df_completo['Machine'] == resultado['machine']) & (df_completo['Component'] == resultado['component'])
                indices = df_completo[mask].index
                
                if len(indices) > 0:
                    idx_principal = indices[0]  # Tomar el primer match
                    df_completo.loc[idx_principal, 'Respirador_Requerido'] = 'Error en análisis'
                    df_completo.loc[idx_principal, 'Status_Analisis'] = resultado['error']
                continue
            
            rec = resultado['recomendacion']
            
            # Buscar registro principal en el DataFrame original
            mask = (df_completo['Machine'] == resultado['machine']) & (df_completo['Component'] == resultado['component'])
            indices = df_completo[mask].index
            
            if len(indices) > 0:
                # Actualizar registro principal (priorizar el que tiene Oil Capacity)
                registros_candidatos = df_completo.loc[indices]
                idx_principal = indices[0]  # Por defecto el primer match
                
                # Buscar el que tiene Oil Capacity si hay múltiples
                registros_con_aceite = registros_candidatos[registros_candidatos['(D) Oil Capacity'].fillna(0) > 0]
                if len(registros_con_aceite) > 0:
                    idx_principal = registros_con_aceite.index[0]
                
                # Actualizar con resultados del análisis
                df_completo.loc[idx_principal, 'Respirador_Requerido'] = 'Sí'
                df_completo.loc[idx_principal, 'Criticidad'] = rec['criticidad']
                df_completo.loc[idx_principal, 'Tipo_Sistema'] = resultado['tipo_sistema']
                df_completo.loc[idx_principal, 'CFM_Requerido'] = rec['cfm_requerido']
                df_completo.loc[idx_principal, 'CFM_Rango'] = rec['cfm_rango']
                df_completo.loc[idx_principal, 'Modelo_Recomendado'] = rec['modelo_recomendado']
                df_completo.loc[idx_principal, 'Serie_Recomendada'] = rec['serie_recomendada']
                df_completo.loc[idx_principal, 'Cartucho_Recomendado'] = rec['cartucho']
                df_completo.loc[idx_principal, 'Altura_Requerida'] = rec['altura_requerida']
                df_completo.loc[idx_principal, 'Espacio_Disponible'] = rec['espacio_disponible']
                df_completo.loc[idx_principal, 'Margen_Espacio'] = rec['margen_espacio']
                df_completo.loc[idx_principal, 'Margen_CFM'] = rec['margen_cfm']
                df_completo.loc[idx_principal, 'Prioridad'] = rec['prioridad']
                df_completo.loc[idx_principal, 'Robustez'] = rec['robustez']
                df_completo.loc[idx_principal, 'Variacion_CFM_Pct'] = resultado['cfm_data']['variacion_porcentual'] if resultado['cfm_data'] else 'N/A'
                df_completo.loc[idx_principal, 'Total_Iteraciones'] = resultado['cfm_data']['total_iteraciones'] if resultado['cfm_data'] else 'N/A'
                df_completo.loc[idx_principal, 'Status_Analisis'] = 'Completado'
                
                # Información post-modificaciones si existe
                if rec.get('recomendacion_post_modificaciones'):
                    post_rec = rec['recomendacion_post_modificaciones']
                    df_completo.loc[idx_principal, 'Modelo_Post_Modificaciones'] = post_rec['modelo']
                    df_completo.loc[idx_principal, 'Serie_Post_Modificaciones'] = post_rec['serie']
                    df_completo.loc[idx_principal, 'Cartucho_Post_Modificaciones'] = post_rec['cartucho']
                    df_completo.loc[idx_principal, 'Altura_Post_Modificaciones'] = f"{post_rec['altura']}\""
                    df_completo.loc[idx_principal, 'CFM_Capacidad_Post_Modificaciones'] = post_rec['cfm_capacidad']
                
                # Modificaciones requeridas
                if resultado['modificaciones']:
                    modificaciones_texto = " | ".join(resultado['modificaciones'])
                    df_completo.loc[idx_principal, 'Modificaciones_Requeridas'] = modificaciones_texto
                
                # Marcar registros secundarios del mismo equipo
                for idx_sec in indices:
                    if idx_sec != idx_principal:
                        df_completo.loc[idx_sec, 'Respirador_Requerido'] = f'Ver registro principal ({resultado["component"]})'
                        df_completo.loc[idx_sec, 'Status_Analisis'] = 'Registro secundario'
        
        # Preparar datos adicionales para hojas separadas
        resumen_data = []
        iteraciones_data = []
        modificaciones_data = []
        
        for resultado in resultados_completos:
            if 'error' in resultado:
                continue
            
            rec = resultado['recomendacion']
            # Información básica para hoja de resumen
            row_data = {
                'Machine': resultado['machine'],
                'Component': resultado['component'],
                'Oil_Capacity_gal': resultado['oil_capacity'],
                'Tipo_Sistema': resultado['tipo_sistema'],
                'Criticidad': rec['criticidad'],
                'CFM_Requerido': rec['cfm_requerido'],
                'CFM_Rango': rec['cfm_rango'],
                'Modelo_Recomendado': rec['modelo_recomendado'],
                'Serie': rec['serie_recomendada'],
                'Cartucho': rec['cartucho'],
                'Altura_Requerida': rec['altura_requerida'],
                'Espacio_Disponible': rec['espacio_disponible'],
                'Margen_Espacio': rec['margen_espacio'],
                'Margen_CFM': rec['margen_cfm'],
                'Prioridad': rec['prioridad'],
                'Robustez': rec['robustez'],
                'Variacion_CFM_Pct': resultado['cfm_data']['variacion_porcentual'] if resultado['cfm_data'] else 'N/A',
                'Total_Iteraciones': resultado['cfm_data']['total_iteraciones'] if resultado['cfm_data'] else 'N/A'
            }
            
            # Agregar información post-modificaciones si existe
            if rec.get('recomendacion_post_modificaciones'):
                post_rec = rec['recomendacion_post_modificaciones']
                row_data.update({
                    'Modelo_Post_Modificaciones': post_rec['modelo'],
                    'Serie_Post_Modificaciones': post_rec['serie'],
                    'Cartucho_Post_Modificaciones': post_rec['cartucho'],
                    'Altura_Post_Modificaciones': f"{post_rec['altura']}\"",
                    'CFM_Capacidad_Post_Modificaciones': post_rec['cfm_capacidad']
                })
            else:
                row_data.update({
                    'Modelo_Post_Modificaciones': 'N/A',
                    'Serie_Post_Modificaciones': 'N/A',
                    'Cartucho_Post_Modificaciones': 'N/A',
                    'Altura_Post_Modificaciones': 'N/A',
                    'CFM_Capacidad_Post_Modificaciones': 'N/A'
                })
            
            resumen_data.append(row_data)
            
            # Agregar modificaciones
            if resultado['modificaciones']:
                for mod in resultado['modificaciones']:
                    modificaciones_data.append({
                        'Machine': resultado['machine'],
                        'Component': resultado['component'],
                        'Modificacion': mod
                    })
            
            # Agregar iteraciones para casos variables
            if resultado.get('cfm_data') and resultado['cfm_data']['variacion_porcentual'] > 30:
                for iter_data in resultado['cfm_data']['iteraciones_detalladas']:
                    iteraciones_data.append({
                        'Machine': resultado['machine'],
                        'Component': resultado['component'],
                        'Porcentaje_Aceite': iter_data['porcentaje'],
                        'Tiempo_min': iter_data['tiempo'],
                        'CFM': iter_data['cfm'],
                        'Volumen_Total_gal': iter_data['volumen_total'],
                        'Volumen_Aceite_gal': iter_data['volumen_aceite'],
                        'Volumen_Aire_gal': iter_data['volumen_aire']
                    })
        
        # Crear archivo Excel completo
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # HOJA PRINCIPAL: Data Report completo con nuevas columnas
            df_completo.to_excel(writer, sheet_name='Data_Report_Completo_v2.2', index=False)
            
            # Hoja de resumen solo equipos con respiradores
            if resumen_data:
                resumen_df = pd.DataFrame(resumen_data)
                resumen_df.to_excel(writer, sheet_name='Resumen_Respiradores', index=False)
            
            # Hoja de modificaciones
            if modificaciones_data:
                modificaciones_df = pd.DataFrame(modificaciones_data)
                modificaciones_df.to_excel(writer, sheet_name='Modificaciones_Requeridas', index=False)
            
            # Hoja de iteraciones para casos variables
            if iteraciones_data:
                iteraciones_df = pd.DataFrame(iteraciones_data)
                iteraciones_df.to_excel(writer, sheet_name='Iteraciones_Variables', index=False)
            
            # Hoja de parámetros utilizados
            parametros_data = [
                ['Coeficiente γ (aceite)', gamma_aceite, 'F⁻¹'],
                ['Coeficiente β (aire)', beta_aire, 'F⁻¹'],
                ['Factor seguridad CFM', factor_seguridad, 'multiplicador'],
                ['Porcentajes aceite iterados', ', '.join(map(str, porcentajes_aceite)), '%'],
                ['Tiempos expansión iterados', ', '.join(map(str, tiempos_expansion)), 'min'],
                ['Total registros en data report', len(df_completo), 'registros'],
                ['Equipos candidatos identificados', len(st.session_state.equipos_candidatos), 'equipos'],
                ['Equipos analizados', len(resultados_completos), 'equipos'],
                ['Equipos con recomendación directa', equipos_con_recomendacion, 'equipos'],
                ['Equipos que requieren modificaciones', equipos_requieren_modificaciones, 'equipos'],
                ['Resultados robustos (var <15%)', equipos_robustos, 'equipos']
            ]
            parametros_df = pd.DataFrame(parametros_data, columns=['Parámetro', 'Valor', 'Unidad'])
            parametros_df.to_excel(writer, sheet_name='Parametros_Analisis', index=False)
        
        output.seek(0)
        
        # Información del archivo
        total_registros = len(df_completo)
        registros_analizados = len([r for r in resultados_completos if 'error' not in r])
        
        st.info(f"""
        📊 **Archivo generado:**
        - **Total registros**: {total_registros} (Data Report completo)
        - **Equipos analizados**: {registros_analizados} con recomendaciones de respiradores
        - **Nuevas columnas**: {len(nuevas_columnas)} agregadas al Data Report original
        """)
        
        st.download_button(
            label="📥 Descargar Data Report Completo + Análisis Respiradores v2.2",
            data=output.getvalue(),
            file_name=f"data_report_completo_respiradores_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Mostrar estadísticas finales
        with st.expander("📊 Estadísticas del Análisis Final", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Distribución por Criticidad:**")
                criticidades = [r['recomendacion']['criticidad'] for r in resultados_completos if r.get('recomendacion')]
                if criticidades:
                    criticidad_counts = pd.Series(criticidades).value_counts()
                    st.bar_chart(criticidad_counts)
                
                st.write("**Distribución por Prioridad:**")
                prioridades = [r['recomendacion']['prioridad'] for r in resultados_completos if r.get('recomendacion')]
                if prioridades:
                    prioridad_counts = pd.Series(prioridades).value_counts()
                    st.bar_chart(prioridad_counts)
            
            with col2:
                st.write("**Series Recomendadas:**")
                series = [r['recomendacion']['serie_recomendada'] for r in resultados_completos if r.get('recomendacion') and r['recomendacion']['modelo_recomendado'] != 'REQUIERE MODIFICACIONES']
                if series:
                    serie_counts = pd.Series(series).value_counts()
                    st.bar_chart(serie_counts)
                
                st.write("**Robustez del Análisis:**")
                variaciones = []
                for r in resultados_completos:
                    if r.get('cfm_data'):
                        var = r['cfm_data']['variacion_porcentual']
                        if var < 15:
                            variaciones.append('Robusto (<15%)')
                        elif var < 30:
                            variaciones.append('Moderado (15-30%)')
                        else:
                            variaciones.append('Variable (>30%)')
                
                if variaciones:
                    robustez_counts = pd.Series(variaciones).value_counts()
                    st.bar_chart(robustez_counts)
        
        # Botón para reiniciar
        if st.button("🔄 Reiniciar Análisis Completo"):
            st.session_state.analysis_stage = 'upload'
            st.session_state.df_original = None
            st.session_state.equipos_candidatos = None
            st.session_state.inputs_faltantes = {}
            st.session_state.equipos_con_faltantes = []
            st.session_state.inputs_manuales_por_equipo = {}
            st.rerun()

if __name__ == "__main__":
    main()

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p><strong>🛠️ Analizador de Respiradores v2.2 - Implementación Final</strong></p>
    <p>Flujo Inteligente + Inputs Individuales + Filtro Espacio + Modificaciones | Metodología Noria Corporation</p>
    <p><em>⚠️ Lógica de criticidad implementada temporalmente - pendiente validación final con equipo técnico</em></p>
</div>
""", unsafe_allow_html=True)
