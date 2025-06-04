import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import re

# Configuración de la página
st.set_page_config(
    page_title="🛠️ Analizador de Respiradores",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("🛠️ Analizador Automático de Respiradores")
st.markdown("**Automatización de selección de respiradores basado en expansión térmica y criterios de selección**")

# Sidebar para configuración
st.sidebar.header("⚙️ Configuración")
st.sidebar.markdown("---")

# Coeficientes de expansión térmica
st.sidebar.subheader("🧮 Coeficientes de Expansión")
gamma_aceite = st.sidebar.number_input("γ (Aceite) [1/°C]", value=0.00063, format="%.5f")
beta_aire = st.sidebar.number_input("β (Aire) [1/°C]", value=0.003667, format="%.6f")
temp_ambiente = st.sidebar.number_input("Temperatura ambiente [°C]", value=28.0)

# Parámetros de iteración
st.sidebar.subheader("🔄 Parámetros de Iteración")
porcentajes_aceite = st.sidebar.multiselect(
    "% Aceite a iterar", 
    [25, 30, 35, 40, 45, 50, 55], 
    default=[30, 35, 40, 45, 50]
)
tiempos_expansion = st.sidebar.multiselect(
    "Tiempos expansión [min]", 
    [6, 8, 10, 12, 15], 
    default=[8, 10, 12]
)

st.sidebar.markdown("---")
st.sidebar.info("📋 **Instrucciones:**\n1. Sube tu data report (Excel)\n2. El sistema analizará automáticamente\n3. Descarga el archivo con recomendaciones")

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

def extraer_temperaturas(temp_string, temp_ambiente):
    """Extraer temperaturas mínima y máxima de string"""
    if pd.isna(temp_string):
        return temp_ambiente, temp_ambiente + 50, 50
    
    # Buscar patrones de temperatura en °F y °C
    temp_pattern_f = r'(\d+\.?\d*)°F.*?(\d+\.?\d*)°F'
    temp_pattern_c = r'(\d+\.?\d*)°C.*?(\d+\.?\d*)°C'
    
    match_f = re.search(temp_pattern_f, str(temp_string))
    match_c = re.search(temp_pattern_c, str(temp_string))
    
    if match_c:
        temp_min = float(match_c.group(1))
        temp_max = float(match_c.group(2))
    elif match_f:
        # Convertir °F a °C
        temp_min_f = float(match_f.group(1))
        temp_max_f = float(match_f.group(2))
        temp_min = (temp_min_f - 32) * 5/9
        temp_max = (temp_max_f - 32) * 5/9
    else:
        temp_min = temp_ambiente
        temp_max = temp_ambiente + 50
    
    delta_t = temp_max - temp_min
    return temp_min, temp_max, delta_t

def calcular_cfm_equipo(row, porcentajes, tiempos, gamma, beta, temp_ambiente):
    """Calcular CFM para un equipo específico"""
    oil_capacity = row.get('(D) Oil Capacity', 0)
    
    if pd.isna(oil_capacity) or oil_capacity <= 0:
        return None
    
    # Extraer temperaturas
    temp_min, temp_max, delta_t = extraer_temperaturas(
        row.get('(D) Operating Temperature'), temp_ambiente
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
    
    # Calcular iteraciones
    resultados = []
    for pct in porcentajes:
        if metodo == "Directo":
            volumen_total = (length * width * height) / 1_000_000_000  # mm³ a litros
        elif metodo == "Inverso":
            volumen_total = oil_capacity / (pct / 100)
            area = (length * height) / 1_000_000  # mm² a m²
            width_calc = (volumen_total / 1000) / area * 1000  # mm
        else:
            volumen_total = oil_capacity / (pct / 100)
            width_calc = np.nan
        
        volumen_aceite = volumen_total * (pct / 100)
        volumen_aire = volumen_total - volumen_aceite
        
        # Expansión térmica
        delta_vo = gamma * volumen_aceite * delta_t
        delta_va = beta * volumen_aire * delta_t
        v_exp = delta_vo + delta_va
        v_exp_ft3 = v_exp * 0.0353147  # litros a ft³
        
        # CFM para diferentes tiempos
        cfms = [v_exp_ft3 / t for t in tiempos]
        
        resultado = {
            'porcentaje': pct,
            'volumen_total': volumen_total,
            'volumen_aceite': volumen_aceite,
            'volumen_aire': volumen_aire,
            'v_exp': v_exp,
            'cfm_valores': cfms
        }
        
        if metodo == "Inverso":
            resultado['width_calculado'] = width_calc
        
        resultados.append(resultado)
    
    # Calcular rango CFM
    all_cfms = [cfm for res in resultados for cfm in res['cfm_valores']]
    cfm_min = min(all_cfms)
    cfm_max = max(all_cfms)
    
    return {
        'metodo': metodo,
        'temp_min': temp_min,
        'temp_max': temp_max,
        'delta_t': delta_t,
        'cfm_min': cfm_min,
        'cfm_max': cfm_max,
        'cfm_rango': f"{cfm_min:.3f} - {cfm_max:.3f}",
        'iteraciones': resultados
    }

def determinar_tipo_sistema(row):
    """Determinar si el sistema es estático o circulatorio"""
    flow_rate = row.get('(D) Flow Rate', np.nan)
    if pd.notna(flow_rate) and flow_rate > 0:
        return "Circulatorio"
    else:
        return "Estático"

def aplicar_criterios_seleccion(row):
    """Aplicar diagrama de criterios de selección"""
    criterios = {
        'filtro_particulas': False,
        'desecante': False,
        'valvula_check': False,
        'resistente_vibracion': False,
        'camara_expansion': False,
        'control_niebla': False,
        'mayor_capacidad': False
    }
    
    # Extraer datos
    humidity = row.get('(D) Average Relative Humidity', '')
    contaminant_likelihood = str(row.get('(D) Contaminant Likelihood', '')).lower()
    contaminant_index = str(row.get('(D) Contaminant Abrasive Index', '')).lower()
    vibration = str(row.get('(D) Vibration', ''))
    env_conditions = str(row.get('(D) Environmental Conditions', '')).lower()
    water_contact = str(row.get('(D) Water Contact Conditions', '')).lower()
    oil_mist = row.get('(D) Oil Mist Evidence on Headspace', False)
    
    # Extraer valor numérico de humedad
    if humidity:
        try:
            if '%' in str(humidity):
                humidity_val = float(str(humidity).split('%')[0])
            else:
                humidity_val = float(str(humidity).split('(')[0]) if '(' in str(humidity) else float(humidity)
        except:
            humidity_val = 50  # valor por defecto
    else:
        humidity_val = 50
    
    # 1. FILTRO DE PARTÍCULAS
    ambiente_arido = 'heavy' in contaminant_index or 'mining' in contaminant_index
    equipo_interior = 'inside' in env_conditions
    ambiente_sucio = 'high' in contaminant_likelihood
    humedad_baja = humidity_val <= 50
    
    if ambiente_arido or equipo_interior or ambiente_sucio or humedad_baja:
        criterios['filtro_particulas'] = True
    
    # 2. DESECANTE
    humedad_alta = humidity_val > 75
    if humedad_alta and equipo_interior:
        criterios['desecante'] = True
    
    # 3. VÁLVULA CHECK
    humedad_muy_alta = humidity_val > 85
    exterior = 'outside' in env_conditions
    lavado = 'water' in water_contact and 'no water' not in water_contact
    
    if humedad_muy_alta or exterior or lavado:
        criterios['valvula_check'] = True
    
    # 4. RESISTENTE A VIBRACIÓN
    alta_vibracion = '<0.2' not in vibration and vibration != ''
    if alta_vibracion:
        criterios['resistente_vibracion'] = True
    
    # 5. CÁMARA DE EXPANSIÓN
    # Se determina con CFM bajo en la función principal
    
    # 6. CONTROL DE NIEBLA
    if oil_mist == True:
        criterios['control_niebla'] = True
    
    # 7. MAYOR CAPACIDAD
    ambiente_severo = 'heavy' in contaminant_index or 'mining' in contaminant_index
    mucho_polvo = 'high' in contaminant_likelihood
    
    if ambiente_severo or mucho_polvo:
        criterios['mayor_capacidad'] = True
    
    return criterios

def recomendar_respirador(cfm_data, criterios, row):
    """Generar recomendación específica de respirador"""
    if not cfm_data:
        return {
            'modelo_recomendado': 'Datos insuficientes',
            'tipo_respirador': 'No determinado',
            'cartucho': 'N/A',
            'justificacion': 'Faltan datos de Oil Capacity',
            'cfm_requerido': 'N/A',
            'espacio_requerido': 'N/A',
            'prioridad': 'Baja'
        }
    
    # Analizar espacio disponible
    clearance = str(row.get('(D) Breather/Fill Port Clearance', '')).lower()
    port_size = str(row.get('(D) Breather/Fill Port Size', ''))
    contaminant_likelihood = str(row.get('(D) Contaminant Likelihood', '')).lower()
    
    # Determinar si requiere cartucho reemplazable
    requiere_reemplazable = 'high' in contaminant_likelihood
    
    # CFM requerido
    cfm_max = cfm_data['cfm_max']
    cfm_rango = cfm_data['cfm_rango']
    
    # Determinar cámara de expansión
    criterios['camara_expansion'] = cfm_max < 1.0  # Bajo flujo
    
    # Generar tipo de respirador
    componentes_requeridos = []
    if criterios['filtro_particulas']:
        componentes_requeridos.append("Filtro partículas")
    if criterios['desecante']:
        componentes_requeridos.append("Desecante")
    if criterios['valvula_check']:
        componentes_requeridos.append("Válvula check")
    if criterios['camara_expansion']:
        componentes_requeridos.append("Cámara expansión")
    if criterios['mayor_capacidad']:
        componentes_requeridos.append("Mayor capacidad")
    
    tipo_respirador = " + ".join(componentes_requeridos) if componentes_requeridos else "Respirador básico"
    
    # Recomendación de modelo
    modelo = "Pendiente definir"
    cartucho = "N/A"
    espacio_requerido = "A determinar"
    prioridad = "Media"
    
    if requiere_reemplazable:
        if 'no port' in clearance:
            modelo = "X-503 (post-instalación puerto)"
            cartucho = "A-345 (reemplazable)"
            espacio_requerido = ">8 inches + Puerto 1\" NPT"
            prioridad = "Alta - Requiere modificación"
        elif '2 to' in clearance and '4' in clearance:
            modelo = "X-100"
            cartucho = "L-143 (reemplazable)"
            espacio_requerido = "6-8 inches"
            prioridad = "Alta"
        elif 'greater than 6' in clearance:
            modelo = "X-505 o Guardian G5S1N"
            cartucho = "A-346 o GRC5S (reemplazable)"
            espacio_requerido = ">10 inches"
            prioridad = "Media"
        elif 'flange' in clearance.lower():
            modelo = "X-100 (adaptador requerido)"
            cartucho = "L-143 (reemplazable)"
            espacio_requerido = "Según flange existente"
            prioridad = "Alta"
        else:
            modelo = "X-503"
            cartucho = "A-345 (reemplazable)"
            espacio_requerido = "8-10 inches"
            prioridad = "Media"
    else:
        modelo = "D-Series (desechable)"
        cartucho = "Unidad completa"
        prioridad = "Baja"
    
    # Justificación
    justificacion_partes = []
    if requiere_reemplazable:
        justificacion_partes.append("High contamination → Cartucho reemplazable obligatorio")
    if criterios['filtro_particulas']:
        justificacion_partes.append("Ambiente industrial → Filtración partículas")
    if criterios['camara_expansion']:
        justificacion_partes.append(f"Bajo CFM ({cfm_rango}) → Cámara expansión")
    if criterios['mayor_capacidad']:
        justificacion_partes.append("Ambiente severo → Mayor capacidad")
    
    justificacion = " | ".join(justificacion_partes) if justificacion_partes else "Análisis estándar"
    
    return {
        'modelo_recomendado': modelo,
        'tipo_respirador': tipo_respirador,
        'cartucho': cartucho,
        'justificacion': justificacion,
        'cfm_requerido': cfm_rango,
        'espacio_requerido': espacio_requerido,
        'prioridad': prioridad
    }

def procesar_data_report(df, porcentajes, tiempos, gamma, beta, temp_ambiente):
    """Función principal para procesar todo el data report"""
    
    # Crear copia del DataFrame original COMPLETO
    df_resultado = df.copy()
    
    # Identificar equipos candidatos únicos y todos los registros candidatos
    equipos_unicos, todos_candidatos = identificar_equipos_candidatos(df)
    
    # Inicializar nuevas columnas en TODO el DataFrame
    nuevas_columnas = {
        'Respirador_Requerido': 'No',
        'Tipo_Sistema': '',
        'CFM_Calculado': '',
        'Modelo_Respirador': '',
        'Tipo_Respirador': '',
        'Cartucho_Recomendado': '',
        'Justificacion': '',
        'Espacio_Requerido': '',
        'Prioridad': '',
        'Status_Analisis': 'No analizado'
    }
    
    for col in nuevas_columnas:
        df_resultado[col] = nuevas_columnas[col]
    
    # Crear mapeo de Machine+Component a índice principal
    equipos_procesados = {}
    resultados_detallados = []
    
    # Procesar cada equipo único
    for idx, row in equipos_unicos.iterrows():
        machine = row['Machine']
        component = row['Component'] 
        key = f"{machine}_{component}"
        
        # Determinar tipo de sistema
        tipo_sistema = determinar_tipo_sistema(row)
        
        # Calcular CFM
        cfm_data = calcular_cfm_equipo(row, porcentajes, tiempos, gamma, beta, temp_ambiente)
        
        # Aplicar criterios de selección
        criterios = aplicar_criterios_seleccion(row)
        
        # Generar recomendación
        recomendacion = recomendar_respirador(cfm_data, criterios, row)
        
        # Buscar el índice original en el DataFrame para este equipo
        # Priorizar registro con Oil Capacity
        registros_equipo = df_resultado[
            (df_resultado['Machine'] == machine) & 
            (df_resultado['Component'] == component)
        ]
        
        # Encontrar el índice del registro principal
        idx_principal = None
        if len(registros_equipo[registros_equipo['(D) Oil Capacity'].fillna(0) > 0]) > 0:
            idx_principal = registros_equipo[registros_equipo['(D) Oil Capacity'].fillna(0) > 0].index[0]
        elif len(registros_equipo[registros_equipo['MaintPointTemplate'].str.contains('oil', case=False, na=False)]) > 0:
            idx_principal = registros_equipo[registros_equipo['MaintPointTemplate'].str.contains('oil', case=False, na=False)].index[0]
        else:
            idx_principal = registros_equipo.index[0]
        
        # Actualizar SOLO el registro principal con las recomendaciones
        df_resultado.loc[idx_principal, 'Respirador_Requerido'] = 'Sí'
        df_resultado.loc[idx_principal, 'Tipo_Sistema'] = tipo_sistema
        df_resultado.loc[idx_principal, 'CFM_Calculado'] = recomendacion['cfm_requerido']
        df_resultado.loc[idx_principal, 'Modelo_Respirador'] = recomendacion['modelo_recomendado']
        df_resultado.loc[idx_principal, 'Tipo_Respirador'] = recomendacion['tipo_respirador']
        df_resultado.loc[idx_principal, 'Cartucho_Recomendado'] = recomendacion['cartucho']
        df_resultado.loc[idx_principal, 'Justificacion'] = recomendacion['justificacion']
        df_resultado.loc[idx_principal, 'Espacio_Requerido'] = recomendacion['espacio_requerido']
        df_resultado.loc[idx_principal, 'Prioridad'] = recomendacion['prioridad']
        df_resultado.loc[idx_principal, 'Status_Analisis'] = 'Completado' if cfm_data else 'Datos insuficientes'
        
        # Marcar registros secundarios del mismo equipo
        for idx_sec in registros_equipo.index:
            if idx_sec != idx_principal:
                df_resultado.loc[idx_sec, 'Respirador_Requerido'] = f'Ver registro principal ({component})'
                df_resultado.loc[idx_sec, 'Status_Analisis'] = 'Registro secundario'
        
        # Guardar para resultados detallados
        equipos_procesados[key] = idx_principal
        
        resultado_detallado = {
            'idx': idx_principal,
            'Machine': machine,
            'Component': component,
            'Oil_Capacity': row.get('(D) Oil Capacity', 0),
            'tipo_sistema': tipo_sistema,
            'cfm_data': cfm_data,
            'criterios': criterios,
            'recomendacion': recomendacion
        }
        resultados_detallados.append(resultado_detallado)
    
    # Crear DataFrame de equipos únicos para mostrar resumen
    equipos_con_respirador = df_resultado[df_resultado['Respirador_Requerido'] == 'Sí']
    
    return df_resultado, resultados_detallados, equipos_con_respirador

def mostrar_resultados_detallados(resultados_detallados):
    """Mostrar resultados detallados del análisis"""
    
    st.header("📊 Resultados Detallados del Análisis")
    
    # Métricas generales
    total_equipos = len(resultados_detallados)
    equipos_completos = len([r for r in resultados_detallados if r['cfm_data']])
    equipos_faltantes = total_equipos - equipos_completos
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Equipos", total_equipos)
    with col2:
        st.metric("Análisis Completo", equipos_completos)
    with col3:
        st.metric("Datos Faltantes", equipos_faltantes)
    
    # Mostrar cada equipo
    for i, resultado in enumerate(resultados_detallados):
        with st.expander(f"🔧 {resultado['Component']} - {resultado['Machine']}", expanded=False):
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**📋 Datos Básicos:**")
                st.write(f"• **Oil Capacity:** {resultado['Oil_Capacity']} litros")
                st.write(f"• **Tipo Sistema:** {resultado['tipo_sistema']}")
                
                if resultado['cfm_data']:
                    st.write(f"• **Rango CFM:** {resultado['cfm_data']['cfm_rango']}")
                    st.write(f"• **Método Cálculo:** {resultado['cfm_data']['metodo']}")
                    st.write(f"• **ΔT:** {resultado['cfm_data']['delta_t']:.1f}°C")
            
            with col2:
                st.markdown("**🏆 Recomendación:**")
                rec = resultado['recomendacion']
                st.write(f"• **Modelo:** {rec['modelo_recomendado']}")
                st.write(f"• **Cartucho:** {rec['cartucho']}")
                st.write(f"• **Prioridad:** {rec['prioridad']}")
                st.write(f"• **Espacio:** {rec['espacio_requerido']}")
            
            # Criterios aplicados
            criterios = resultado['criterios']
            criterios_activos = [k.replace('_', ' ').title() for k, v in criterios.items() if v]
            if criterios_activos:
                st.markdown("**✅ Componentes Requeridos:**")
                st.write(" • ".join(criterios_activos))
            
            # Justificación
            st.markdown("**💡 Justificación:**")
            st.info(rec['justificacion'])

# INTERFAZ PRINCIPAL
def main():
    
    # Upload de archivo
    st.header("📤 Subir Data Report")
    uploaded_file = st.file_uploader(
        "Selecciona tu archivo Excel",
        type=['xlsx', 'xls'],
        help="Sube el data report en formato Excel para análisis automático"
    )
    
    if uploaded_file is not None:
        # Cargar datos
        with st.spinner("Cargando datos..."):
            df = cargar_datos(uploaded_file)
        
        if df is not None:
            st.success(f"✅ Archivo cargado exitosamente: {len(df)} registros")
            
            # Mostrar preview
            with st.expander("👀 Vista previa de datos", expanded=False):
                st.dataframe(df.head())
            
            # Botón para procesar
            if st.button("🚀 Ejecutar Análisis Completo", type="primary"):
                
                with st.spinner("Ejecutando análisis automático..."):
                    df_resultado, resultados_detallados, equipos_candidatos = procesar_data_report(
                        df, porcentajes_aceite, tiempos_expansion, 
                        gamma_aceite, beta_aire, temp_ambiente
                    )
                
                st.success("✅ Análisis completado!")
                
                # Mostrar resumen
                st.header("📈 Resumen de Resultados")
                
                # Equipos que requieren respiradores
                equipos_con_respirador = df_resultado[df_resultado['Respirador_Requerido'] == 'Sí']
                registros_secundarios = df_resultado[df_resultado['Status_Analisis'] == 'Registro secundario']
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Registros Total", len(df_resultado))
                with col2:
                    st.metric("Equipos Únicos c/Respirador", len(equipos_con_respirador))
                with col3:
                    st.metric("Análisis Completo", len(equipos_con_respirador[equipos_con_respirador['Status_Analisis'] == 'Completado']))
                with col4:
                    st.metric("Registros Secundarios", len(registros_secundarios))
                
                # Tabla resumen de equipos con respiradores
                if len(equipos_con_respirador) > 0:
                    st.subheader("🔧 Equipos que Requieren Respiradores")
                    
                    # Seleccionar columnas relevantes para mostrar
                    columnas_mostrar = [
                        'Machine', 'Component', '(D) Oil Capacity', 
                        'CFM_Calculado', 'Modelo_Respirador', 'Cartucho_Recomendado',
                        'Prioridad', 'Status_Analisis'
                    ]
                    
                    df_display = equipos_con_respirador[columnas_mostrar].copy()
                    st.dataframe(df_display, use_container_width=True)
                    
                    # Mostrar información adicional
                    st.info(f"📊 **Total de equipos únicos que requieren respiradores: {len(equipos_con_respirador)}**")
                    
                    # Contar registros secundarios
                    registros_secundarios = len(df_resultado[df_resultado['Status_Analisis'] == 'Registro secundario'])
                    if registros_secundarios > 0:
                        st.info(f"ℹ️ **Registros secundarios (mismo equipo, diferentes MaintPoints): {registros_secundarios}**")
                
                # Mostrar resultados detallados
                if resultados_detallados:
                    mostrar_resultados_detallados(resultados_detallados)
                
                # Generar archivo para descarga
                st.header("💾 Descargar Resultados")
                
                # Crear buffer para Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, sheet_name='Data_Report_Analizado', index=False)
                    
                    # Crear hoja de resumen
                    if len(equipos_con_respirador) > 0:
                        resumen_df = equipos_con_respirador[columnas_mostrar].copy()
                        resumen_df.to_excel(writer, sheet_name='Resumen_Respiradores', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Descargar Data Report con Recomendaciones",
                    data=output.getvalue(),
                    file_name=f"data_report_respiradores_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Mostrar estadísticas finales
                with st.expander("📊 Estadísticas del Análisis", expanded=False):
                    if len(equipos_con_respirador) > 0:
                        st.write("**Distribución por Prioridad:**")
                        prioridad_counts = equipos_con_respirador['Prioridad'].value_counts()
                        st.bar_chart(prioridad_counts)
                        
                        st.write("**Modelos Recomendados:**")
                        modelo_counts = equipos_con_respirador['Modelo_Respirador'].value_counts()
                        st.bar_chart(modelo_counts)
                        
                        st.write("**Tipos de Sistema:**")
                        sistema_counts = equipos_con_respirador['Tipo_Sistema'].value_counts()
                        st.bar_chart(sistema_counts)
                        
                        # Mostrar resumen de registros
                        total_registros = len(df_resultado)
                        registros_principales = len(equipos_con_respirador)
                        registros_secundarios = len(df_resultado[df_resultado['Status_Analisis'] == 'Registro secundario'])
                        registros_no_analizados = total_registros - registros_principales - registros_secundarios
                        
                        st.write("**Distribución de Registros:**")
                        registro_data = {
                            'Equipos principales': registros_principales,
                            'Registros secundarios': registros_secundarios,
                            'Sin análisis': registros_no_analizados
                        }
                        st.bar_chart(registro_data)

if __name__ == "__main__":
    main()

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p><strong>🛠️ Analizador de Respiradores v1.0</strong></p>
    <p>Automatización de selección basada en expansión térmica y criterios técnicos</p>
</div>
""", unsafe_allow_html=True)