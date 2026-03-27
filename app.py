import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Tablero Pedagógico Universal", layout="wide")
st.title("📊 Tablero de Control - SIAGIE (Inicial, Primaria y Secundaria) por HWEP")
st.markdown("### El App detecta automáticamente el nivel y procesa Actas PDF o Registros Excel.")

# --- BARRA LATERAL ---
st.sidebar.header("📂 Cargar Documentos")
uploaded_files = st.sidebar.file_uploader(
    "Sube Actas PDF o Excel", 
    type=["csv", "xlsx", "xls", "pdf"], 
    accept_multiple_files=True
)

# --- MOTOR DE ANÁLISIS DE PDF (MULTI-NIVEL) ---
def procesar_acta_siagie(pdf_file):
    """Extrae notas adaptándose a Inicial, Primaria y Secundaria"""
    datos_extraidos = []
    
    # --- DEFINICIÓN DE ESQUEMAS (PLANTILLAS) ---
    
    # 1. INICIAL
    esquema_inicial_3 = [("PERSONAL SOCIAL", "Construye su identidad"), ("PERSONAL SOCIAL", "Convive y participa"), ("RELIGIÓN", "Id. Religiosa"), ("PSICOMOTRIZ", "Motricidad"), ("COMUNICACIÓN", "Oral"), ("COMUNICACIÓN", "Lee"), ("COMUNICACIÓN", "Crea proyectos"), ("MATEMÁTICA", "Cantidad"), ("MATEMÁTICA", "Forma"), ("CIENCIA", "Indaga")]
    esquema_inicial_4 = esquema_inicial_3[:7] + [("COMUNICACIÓN", "Escribe")] + esquema_inicial_3[7:]
    esquema_inicial_5 = esquema_inicial_4 + [("TIC", "TIC"), ("GEST. AUTÓNOMA", "Autonomía")]

    # 2. PRIMARIA
    esquema_prim_A = [
        ("PERSONAL SOCIAL", "Construye su identidad"), ("PERSONAL SOCIAL", "Convive y participa"), ("PERSONAL SOCIAL", "Construye interpretaciones"), ("PERSONAL SOCIAL", "Gestiona espacio"), ("PERSONAL SOCIAL", "Gestiona recursos"),
        ("EDUC. FÍSICA", "Se desenvuelve"), ("EDUC. FÍSICA", "Asume vida saludable"), ("EDUC. FÍSICA", "Interactúa"),
        ("COMUNICACIÓN", "Se comunica oralmente"), ("COMUNICACIÓN", "Lee textos"), ("COMUNICACIÓN", "Escribe textos"),
        ("ARTE Y CULTURA", "Aprecia"), ("ARTE Y CULTURA", "Crea proyectos")
    ]
    esquema_prim_B = [
        ("MATEMÁTICA", "Cantidad"), ("MATEMÁTICA", "Regularidad"), ("MATEMÁTICA", "Forma"), ("MATEMÁTICA", "Datos"),
        ("CIENCIA Y TEC.", "Indaga"), ("CIENCIA Y TEC.", "Explica"), ("CIENCIA Y TEC.", "Diseña"),
        ("RELIGIÓN", "Construye"), ("RELIGIÓN", "Asume"),
        ("TIC", "TIC"), ("GEST. AUTÓNOMA", "Autonomía")
    ]

    # 3. SECUNDARIA (NUEVO)
    # Bloque A: DPCC(2), CCSS(3), EPT(1), EF(3), COMU(3), ARTE(2) -> 14 Notas
    esquema_sec_A = [
        ("DPCC", "Construye su identidad"), ("DPCC", "Convive y participa"),
        ("CIENCIAS SOCIALES", "Construye interpretaciones"), ("CIENCIAS SOCIALES", "Gestiona espacio"), ("CIENCIAS SOCIALES", "Gestiona recursos"),
        ("EPT", "Gestiona proyectos"),
        ("EDUC. FÍSICA", "Se desenvuelve"), ("EDUC. FÍSICA", "Asume vida saludable"), ("EDUC. FÍSICA", "Interactúa"),
        ("COMUNICACIÓN", "Se comunica oralmente"), ("COMUNICACIÓN", "Lee textos"), ("COMUNICACIÓN", "Escribe textos"),
        ("ARTE Y CULTURA", "Aprecia"), ("ARTE Y CULTURA", "Crea proyectos")
    ]
    # Bloque B: INGLÉS(3), MATE(4), CYT(3), RELIGIÓN(2) -> 12 Notas (Castellano suele estar vacío)
    esquema_sec_B = [
        ("INGLÉS", "Se comunica"), ("INGLÉS", "Lee textos"), ("INGLÉS", "Escribe textos"),
        ("MATEMÁTICA", "Cantidad"), ("MATEMÁTICA", "Regularidad"), ("MATEMÁTICA", "Forma"), ("MATEMÁTICA", "Datos"),
        ("CIENCIA Y TEC.", "Indaga"), ("CIENCIA Y TEC.", "Explica"), ("CIENCIA Y TEC.", "Diseña"),
        ("RELIGIÓN", "Construye"), ("RELIGIÓN", "Asume")
    ]

    with pdfplumber.open(pdf_file) as pdf:
        # Detectar Nivel
        first_text = pdf.pages[0].extract_text().upper()
        es_inicial = "INICIAL" in first_text
        es_secundaria = "SECUNDARIA" in first_text
        
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            for line in lines:
                # Buscar alumnos
                if re.search(r'^\d+\s+(D\s*N\s*I|.*?)\s+\d+', line):
                    partes = line.split()
                    # Extraer notas válidas (incluyendo EXO para Religión/Inglés)
                    notas = [p for p in partes if p in ["AD", "A", "B", "C", "EXO"]]
                    
                    esquema_seleccionado = []
                    
                    # --- LÓGICA DE SELECCIÓN ---
                    if es_inicial:
                        if len(notas) == 10: esquema_seleccionado = esquema_inicial_3
                        elif len(notas) == 11: esquema_seleccionado = esquema_inicial_4
                        elif len(notas) >= 12: esquema_seleccionado = esquema_inicial_5
                        
                    elif es_secundaria:
                        # Detectar Bloque por encabezado de página o por cantidad de notas
                        if "DESARROLLO PERSONAL" in text or len(notas) == 14:
                            esquema_seleccionado = esquema_sec_A
                        elif "INGLÉS" in text or "MATEMÁTICA" in text:
                            # Ajuste: Secundaria B tiene 12 notas. Si hay Castellano serían 15.
                            if len(notas) == 12:
                                esquema_seleccionado = esquema_sec_B
                            elif len(notas) == 15: # Si hay Castellano
                                esquema_extra = [("CASTELLANO", "Oral"), ("CASTELLANO", "Lee"), ("CASTELLANO", "Escribe")]
                                esquema_seleccionado = esquema_extra + esquema_sec_B
                                
                    else: # Primaria
                        if "PERSONAL SOCIAL" in text:
                            esquema_seleccionado = esquema_prim_A
                        elif "MATEMÁTICA" in text:
                            esquema_seleccionado = esquema_prim_B

                    # Guardar datos
                    if esquema_seleccionado and len(notas) == len(esquema_seleccionado):
                        for i, nota in enumerate(notas):
                            area, comp = esquema_seleccionado[i]
                            datos_extraidos.append({
                                "Área": area,
                                "Competencia": comp,
                                "Nota": nota
                            })

    # Convertir a DataFrame
    if datos_extraidos:
        df_raw = pd.DataFrame(datos_extraidos)
        resumen = []
        grupos = df_raw.groupby(["Área", "Competencia"])
        for (area, comp), grupo in grupos:
            conteo = grupo["Nota"].value_counts()
            resumen.append({
                "Área": area,
                "Cód": "ACTA",
                "Competencia": comp,
                "AD": conteo.get("AD", 0),
                "A": conteo.get("A", 0),
                "B": conteo.get("B", 0),
                "C": conteo.get("C", 0),
                "EXO": conteo.get("EXO", 0),
                "Total": conteo.sum()
            })
        return pd.DataFrame(resumen)
    
    return pd.DataFrame()

# --- FUNCIONES AUXILIARES EXCEL (Mismo motor de siempre) ---
def buscar_fila_cabecera(df):
    limit = min(40, len(df))
    for i in range(limit):
        fila = df.iloc[i].astype(str).tolist()
        if any(c.strip() == 'NL' for c in fila): return i
    return -1

def procesar_excel_csv(df, nombre_archivo, nombre_hoja="Main"):
    resultados_hoja = []
    idx_cabecera = buscar_fila_cabecera(df)
    if idx_cabecera == -1: return [] 

    try:
        header_codes = df.iloc[idx_cabecera - 1]
        header_types = df.iloc[idx_cabecera]
        nombre_full = (nombre_archivo + " " + nombre_hoja).upper()
        
        area = "OTRO"
        if "COMU" in nombre_full: area = "COMUNICACIÓN"
        elif "MATE" in nombre_full: area = "MATEMÁTICA"
        elif "CIENC" in nombre_full: area = "CIENCIA Y TEC."
        elif "PPSS" in nombre_full or "SOCIAL" in nombre_full: area = "PERSONAL SOCIAL"
        elif "ART" in nombre_full: area = "ARTE Y CULTURA"
        elif "RELIG" in nombre_full or "EREL" in nombre_full: area = "RELIGIÓN"
        elif "TIC" in nombre_full: area = "TIC"
        elif "GEST" in nombre_full or "AUTO" in nombre_full: area = "GEST. AUTÓNOMA"
        elif "EFIS" in nombre_full: area = "EDUC. FÍSICA"

        mapa_competencias = {}
        all_text = pd.concat([df[col].astype(str) for col in df.columns])
        for val in all_text:
            match = re.search(r'^\s*(\d+)\s*=\s*(.+)', val)
            if match:
                mapa_competencias[match.group(1).zfill(2)] = match.group(2).strip()

        indices_nl = [i for i, x in enumerate(header_types) if str(x).strip() == 'NL']
        df_data = df.iloc[idx_cabecera + 1:]
        df_alumnos = df_data[df_data.iloc[:, 0].astype(str).str.match(r'^\d+$')]
        
        for col_idx in indices_nl:
            raw_code = str(header_codes[col_idx]).strip()
            comp_code = raw_code.zfill(2) if raw_code.isdigit() else raw_code
            nombre_comp = mapa_competencias.get(comp_code, f"Competencia {comp_code}").split('\n')[0][:150]
            conteo = df_alumnos.iloc[:, col_idx].value_counts()
            
            resultados_hoja.append({
                "Área": area,
                "Cód": comp_code,
                "Competencia": nombre_comp,
                "AD": conteo.get("AD", 0),
                "A": conteo.get("A", 0),
                "B": conteo.get("B", 0),
                "C": conteo.get("C", 0),
                "EXO": 0,
                "Total": conteo.sum()
            })
    except: pass
    return resultados_hoja

# --- LÓGICA PRINCIPAL ---
def procesar_todo(files_list):
    todos_resultados = []
    for uploaded_file in files_list:
        name = uploaded_file.name
        
        if name.lower().endswith('.pdf'):
            try:
                df = procesar_acta_siagie(uploaded_file)
                if not df.empty: todos_resultados.extend(df.to_dict('records'))
            except: pass
        elif name.lower().endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, header=None)
                todos_resultados.extend(procesar_excel_csv(df, name, "CSV"))
            except: pass
        else:
            try:
                xls_dict = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                for sheet_name, df_sheet in xls_dict.items():
                    res = procesar_excel_csv(df_sheet, name, sheet_name)
                    if res: todos_resultados.extend(res)
            except: pass

    return pd.DataFrame(todos_resultados)

# --- INTERFAZ ---
if uploaded_files:
    with st.spinner('Procesando documentos (Inicial / Primaria / Secundaria)...'):
        df_resumen = procesar_todo(uploaded_files)
    
    if not df_resumen.empty:
        df_resumen = df_resumen.sort_values(by=["Área", "Competencia"])
        
        st.subheader("📋 Resumen Consolidado")
        st.dataframe(df_resumen, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
        st.download_button("📥 Descargar Excel", data=buffer.getvalue(), file_name="Reporte_Final.xlsx", mime="application/vnd.ms-excel")
        
        st.divider()
        st.markdown("### 📊 Análisis Visual (Torta + Barras)")
        
        areas = df_resumen["Área"].unique()
        for area in areas:
            with st.expander(f"📘 {area}", expanded=True):
                df_area = df_resumen[df_resumen["Área"] == area]
                for _, row in df_area.iterrows():
                    st.markdown(f"**{row['Competencia']}**")
                    col_torta, col_barra = st.columns([1, 2])
                    
                    valores = [row["AD"], row["A"], row["B"], row["C"]]
                    etiquetas = ["AD", "A", "B", "C"]
                    colores = ["#2E86C1", "#28B463", "#F1C40F", "#E74C3C"]
                    
                    with col_torta:
                        v_clean = [v for v in valores if v > 0]
                        l_clean = [l for v, l in zip(valores, etiquetas) if v > 0]
                        c_clean = [c for v, c in zip(valores, colores) if v > 0]
                        if v_clean:
                            fig, ax = plt.subplots(figsize=(3,3))
                            ax.pie(v_clean, labels=l_clean, autopct='%1.1f%%', colors=c_clean, startangle=90)
                            st.pyplot(fig)
                            plt.close(fig)
                    
                    with col_barra:
                        if sum(valores) > 0:
                            fig, ax = plt.subplots(figsize=(6,3))
                            bars = ax.bar(etiquetas, valores, color=colores)
                            ax.bar_label(bars)
                            ax.set_title("N° Alumnos")
                            st.pyplot(fig)
                            plt.close(fig)
                    st.divider()
    else:
        st.warning("⚠️ No se encontraron datos. Verifica tus archivos.")
else:
    st.info("📂 Sube tus archivos (Actas PDF o Registros Excel) para comenzar.")