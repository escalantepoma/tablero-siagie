import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Tablero Pedagógico Universal", layout="wide")
st.title("ESTADISTICA SIAGIE - HWEP")
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
    datos_extraidos = []
    
    esquema_inicial_3 = [("PERSONAL SOCIAL", "Construye su identidad"), ("PERSONAL SOCIAL", "Convive y participa"), ("RELIGIÓN", "Id. Religiosa"), ("PSICOMOTRIZ", "Motricidad"), ("COMUNICACIÓN", "Oral"), ("COMUNICACIÓN", "Lee"), ("COMUNICACIÓN", "Crea proyectos"), ("MATEMÁTICA", "Cantidad"), ("MATEMÁTICA", "Forma"), ("CIENCIA", "Indaga")]
    esquema_inicial_4 = esquema_inicial_3[:7] + [("COMUNICACIÓN", "Escribe")] + esquema_inicial_3[7:]
    esquema_inicial_5 = esquema_inicial_4 + [("TIC", "TIC"), ("GEST. AUTÓNOMA", "Autonomía")]

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

    esquema_sec_A = [
        ("DPCC", "Construye su identidad"), ("DPCC", "Convive y participa"),
        ("CIENCIAS SOCIALES", "Construye interpretaciones"), ("CIENCIAS SOCIALES", "Gestiona espacio"), ("CIENCIAS SOCIALES", "Gestiona recursos"),
        ("EPT", "Gestiona proyectos"),
        ("EDUC. FÍSICA", "Se desenvuelve"), ("EDUC. FÍSICA", "Asume vida saludable"), ("EDUC. FÍSICA", "Interactúa"),
        ("COMUNICACIÓN", "Se comunica oralmente"), ("COMUNICACIÓN", "Lee textos"), ("COMUNICACIÓN", "Escribe textos"),
        ("ARTE Y CULTURA", "Aprecia"), ("ARTE Y CULTURA", "Crea proyectos")
    ]
    esquema_sec_B = [
        ("INGLÉS", "Se comunica"), ("INGLÉS", "Lee textos"), ("INGLÉS", "Escribe textos"),
        ("MATEMÁTICA", "Cantidad"), ("MATEMÁTICA", "Regularidad"), ("MATEMÁTICA", "Forma"), ("MATEMÁTICA", "Datos"),
        ("CIENCIA Y TEC.", "Indaga"), ("CIENCIA Y TEC.", "Explica"), ("CIENCIA Y TEC.", "Diseña"),
        ("RELIGIÓN", "Construye"), ("RELIGIÓN", "Asume")
    ]

    with pdfplumber.open(pdf_file) as pdf:
        first_text = pdf.pages[0].extract_text().upper()
        es_inicial = "INICIAL" in first_text
        es_secundaria = "SECUNDARIA" in first_text
        
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            for line in lines:
                if re.search(r'^\d+\s+(D\s*N\s*I|.*?)\s+\d+', line):
                    partes = line.split()
                    notas = [p for p in partes if p in ["AD", "A", "B", "C", "EXO"]]
                    esquema_seleccionado = []
                    
                    if es_inicial:
                        if len(notas) == 10: esquema_seleccionado = esquema_inicial_3
                        elif len(notas) == 11: esquema_seleccionado = esquema_inicial_4
                        elif len(notas) >= 12: esquema_seleccionado = esquema_inicial_5
                        
                    elif es_secundaria:
                        if "DESARROLLO PERSONAL" in text or len(notas) == 14:
                            esquema_seleccionado = esquema_sec_A
                        elif "INGLÉS" in text or "MATEMÁTICA" in text:
                            if len(notas) == 12:
                                esquema_seleccionado = esquema_sec_B
                            elif len(notas) == 15: 
                                esquema_extra = [("CASTELLANO", "Oral"), ("CASTELLANO", "Lee"), ("CASTELLANO", "Escribe")]
                                esquema_seleccionado = esquema_extra + esquema_sec_B
                                
                    else: 
                        if "PERSONAL SOCIAL" in text:
                            esquema_seleccionado = esquema_prim_A
                        elif "MATEMÁTICA" in text:
                            esquema_seleccionado = esquema_prim_B

                    if esquema_seleccionado and len(notas) == len(esquema_seleccionado):
                        for i, nota in enumerate(notas):
                            area, comp = esquema_seleccionado[i]
                            datos_extraidos.append({
                                "Área": area,
                                "Competencia": comp,
                                "Nota": nota
                            })

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

# --- FUNCIONES AUXILIARES EXCEL ---
def buscar_fila_cabecera(df):
    limit = min(40, len(df))
    for i in range(limit):
        fila = df.iloc[i].astype(str).str.strip().str.upper().tolist()
        if 'NL' in fila or 'NIVEL DE LOGRO' in fila: 
            return i
    return -1

def procesar_excel_csv(df, nombre_archivo, nombre_hoja="Main"):
    if nombre_hoja.strip().upper() in ["GENERALIDADES", "PARAMETROS", "PARÁMETROS", "MAIN"]:
        return []

    idx_cabecera = buscar_fila_cabecera(df)
    if idx_cabecera == -1: return [] 

    resultados_hoja = []
    try:
        fila_nl = df.iloc[idx_cabecera].astype(str).str.strip().str.upper().tolist()
        nombre_full = nombre_hoja.upper()
        area = "OTRO"
        if "COMU" in nombre_full: area = "COMUNICACIÓN"
        elif "MATE" in nombre_full: area = "MATEMÁTICA"
        elif "CIENC" in nombre_full or "CYT" in nombre_full: area = "CIENCIA Y TEC."
        elif "PPSS" in nombre_full or "SOCIAL" in nombre_full: area = "PERSONAL SOCIAL"
        elif "ART" in nombre_full: area = "ARTE Y CULTURA"
        elif "RELIG" in nombre_full or "EREL" in nombre_full: area = "RELIGIÓN"
        elif "TIC" in nombre_full: area = "TIC"
        elif "GEST" in nombre_full or "AUTO" in nombre_full: area = "GEST. AUTÓNOMA"
        elif "EFIS" in nombre_full: area = "EDUC. FÍSICA"
        elif "INGLES" in nombre_full or "INGLÉS" in nombre_full: area = "INGLÉS"
        elif "CAST" in nombre_full: area = "CASTELLANO"

        mapa_competencias = {}
        for col in df.columns:
            for val in df[col]:
                val_clean = str(val).replace('\n', ' ').strip()
                match = re.search(r'^(?:C|c)?0*(\d+)\s*=\s*(.+)', val_clean)
                if match:
                    num = int(match.group(1))
                    texto = str(match.group(2)).strip()
                    mapa_competencias[num] = texto

        indices_nl = [i for i, x in enumerate(fila_nl) if x in ['NL', 'NIVEL DE LOGRO']]
        df_data = df.iloc[idx_cabecera + 1:]
        df_alumnos = df_data[df_data.iloc[:, 0].astype(str).str.strip().str.match(r'^\d+$')]
        
        for i, col_idx in enumerate(indices_nl):
            raw_code = str(df.iloc[idx_cabecera - 1, col_idx]).strip()
            num_comp = i + 1
            match_num = re.search(r'\d+', raw_code)
            if match_num: num_comp = int(match_num.group(0))

            if num_comp in mapa_competencias:
                comp_name = mapa_competencias[num_comp]
            else:
                comp_name = raw_code
                if str(comp_name).lower() == 'nan' or str(comp_name) == '':
                    comp_name = str(df.iloc[idx_cabecera - 2, col_idx]).strip()
                if str(comp_name).lower() == 'nan' or str(comp_name) == '':
                    comp_name = f"Competencia {num_comp}"

            comp_name = str(comp_name).replace('\n', ' ').strip()[:150] 
            notas_columna = df_alumnos.iloc[:, col_idx].astype(str).str.strip().str.upper()
            conteo = notas_columna.value_counts()
            
            resultados_hoja.append({
                "Área": area,
                "Cód": f"C{num_comp}",
                "Competencia": f"C{num_comp} - {comp_name}",
                "AD": conteo.get("AD", 0),
                "A": conteo.get("A", 0),
                "B": conteo.get("B", 0),
                "C": conteo.get("C", 0),
                "EXO": conteo.get("EXO", 0),
                "Total": conteo.get("AD", 0) + conteo.get("A", 0) + conteo.get("B", 0) + conteo.get("C", 0)
            })
    except Exception as e: 
        st.error(f"⚠️ Error leyendo la hoja {nombre_hoja}. Motivo: {e}")
    return resultados_hoja

# --- MOTOR DE EXPORTACIÓN (AHORA SOPORTA 2D Y 3D) ---
def generar_excel_con_graficos(df_resultados, estilo_3d=True):
    """Genera el Excel insertando los gráficos dependiendo del estilo elegido"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultados.to_excel(writer, sheet_name='Consolidado_Bimestre', index=False)
        workbook  = writer.book
        worksheet = writer.sheets['Consolidado_Bimestre']
        
        worksheet.set_column('A:A', 22)  
        worksheet.set_column('B:B', 8)   
        worksheet.set_column('C:C', 55)  
        worksheet.set_column('D:H', 10)  
        worksheet.set_column('I:I', 12)  
        
        fila_grafico = 1
        columna_grafico = 'K'
        
        for idx, row in df_resultados.iterrows():
            categorias = ['AD', 'A', 'B', 'C']
            valores = [row['AD'], row['A'], row['B'], row['C']]
            colores = ['#2ed573', '#1e90ff', '#ffa502', '#ff4757'] 
            
            if estilo_3d:
                # --- VERSIÓN 3D ---
                fig = plt.figure(figsize=(6, 3.5))
                ax = fig.add_subplot(111, projection='3d')
                
                xpos = [0, 1, 2, 3]
                ypos = [0, 0, 0, 0]
                zpos = [0, 0, 0, 0]
                dx = [0.65, 0.65, 0.65, 0.65]
                dy = [0.65, 0.65, 0.65, 0.65]
                dz = valores
                
                ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=colores, shade=True, alpha=0.95)
                
                max_val = max(valores) if max(valores) > 0 else 1
                for i in range(len(valores)):
                    ax.text(xpos[i] + 0.325, ypos[i] + 0.325, dz[i] + (max_val * 0.05), 
                            str(valores[i]), color='black', fontsize=10, fontweight='bold', 
                            ha='center', va='bottom')
                
                ax.set_xticks([0.3, 1.3, 2.3, 3.3])
                ax.set_xticklabels(categorias, fontsize=10, fontweight='bold')
                ax.set_yticks([]) 
                
                ax.set_title(f"[MODO 3D] {row['Cód']} - {row['Área']}", fontsize=11, fontweight='bold', pad=15)
                ax.set_zlabel("Alumnos", fontsize=9)
                ax.view_init(elev=25, azim=-45)
                
                ax.xaxis.pane.fill = False
                ax.yaxis.pane.fill = False
                ax.zaxis.pane.fill = False
                ax.grid(False) 
                plt.tight_layout()
            
            else:
                # --- VERSIÓN 2D ---
                fig, ax = plt.subplots(figsize=(5.5, 3))
                bars = ax.bar(categorias, valores, color=colores, edgecolor='none', width=0.6)
                
                # Números en 2D
                ax.bar_label(bars, fontsize=10, fontweight='bold', padding=3)
                
                ax.set_title(f"[MODO 2D] {row['Cód']} - {row['Área']}", fontsize=11, fontweight='bold', pad=10)
                ax.set_ylabel("Alumnos", fontsize=9)
                
                # Limpiar bordes 2D
                for spine in ['top', 'right']:
                    ax.spines[spine].set_visible(False)
                plt.tight_layout()

            # Guardar y adjuntar al Excel (común para 2D y 3D)
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=100)
            img_buffer.seek(0)
            plt.close(fig) 
            
            celda_destino = f"{columna_grafico}{fila_grafico + 1}"
            worksheet.insert_image(celda_destino, f"grafico_{idx}.png", {'image_data': img_buffer})
            
            # Espaciado de filas según el estilo
            fila_grafico += 18 if estilo_3d else 16
            
    return output.getvalue()

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

# --- INTERFAZ, DESCARGA Y VISUALIZACIÓN WEB ---
if uploaded_files:
    with st.spinner('Procesando documentos (Inicial / Primaria / Secundaria)...'):
        df_resumen = procesar_todo(uploaded_files)
    
    if not df_resumen.empty:
        df_resumen = df_resumen.sort_values(by=["Área", "Competencia"])
        
        st.subheader("📋 Resumen Consolidado")
        st.dataframe(df_resumen, use_container_width=True)
        
        st.divider()
        
        # --- SELECTOR DE ESTILO VISUAL ---
        st.markdown("### ⚙️ Preferencias de Gráficos")
        opcion_grafico = st.radio(
            "Selecciona el estilo de visualización (Se aplica a la Web y al Excel):",
            ["✨ Gráficos en 3D (Volumen y Sombras)", "📊 Gráficos en 2D (Planos y Minimalistas)"],
            horizontal=True
        )
        es_estilo_3d = "3D" in opcion_grafico
        
        # --- DESCARGA EXCEL CON EL ESTILO ELEGIDO ---
        with st.spinner("Generando Excel con los gráficos seleccionados..."):
            excel_data = generar_excel_con_graficos(df_resumen, estilo_3d=es_estilo_3d)

        label_boton = "📥 Descargar Excel con Gráficos 3D" if es_estilo_3d else "📥 Descargar Excel con Gráficos 2D"
        st.download_button(
            label=label_boton,
            data=excel_data,
            file_name="Reporte_SIAGIE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.divider()
        st.markdown("### 📈 Análisis Visual (Previsualización Web)")
        
        # --- MOTOR DE RENDERIZADO WEB DUAL ---
        areas = df_resumen["Área"].unique()
        for area in areas:
            with st.expander(f"📘 {area}", expanded=True):
                df_area = df_resumen[df_resumen["Área"] == area]
                for _, row in df_area.iterrows():
                    st.markdown(f"**{row['Competencia']}**")
                    col_torta, col_barra = st.columns([1, 2])
                    
                    valores = [row["AD"], row["A"], row["B"], row["C"]]
                    etiquetas = ["AD", "A", "B", "C"]
                    colores = ["#2ed573", "#1e90ff", "#ffa502", "#ff4757"] 
                    
                    with col_torta:
                        v_clean = [v for v in valores if v > 0]
                        l_clean = [l for v, l in zip(valores, etiquetas) if v > 0]
                        c_clean = [c for v, c in zip(valores, colores) if v > 0]
                        
                        if v_clean:
                            fig, ax = plt.subplots(figsize=(3,3))
                            if es_estilo_3d:
                                separacion = [0.08] * len(v_clean)
                                ax.pie(v_clean, labels=l_clean, autopct='%1.1f%%', colors=c_clean, 
                                       startangle=90, shadow=True, explode=separacion)
                            else:
                                ax.pie(v_clean, labels=l_clean, autopct='%1.1f%%', colors=c_clean, startangle=90)
                            
                            st.pyplot(fig)
                            plt.close(fig)
                    
                    with col_barra:
                        if sum(valores) > 0:
                            if es_estilo_3d:
                                fig = plt.figure(figsize=(6, 3.5))
                                ax = fig.add_subplot(111, projection='3d')
                                
                                xpos = [0, 1, 2, 3]
                                ypos = [0, 0, 0, 0]
                                zpos = [0, 0, 0, 0]
                                dx = [0.65, 0.65, 0.65, 0.65]
                                dy = [0.65, 0.65, 0.65, 0.65]
                                dz = valores
                                
                                ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=colores, shade=True, alpha=0.95)
                                
                                max_val_web = max(valores) if max(valores) > 0 else 1
                                for i in range(len(valores)):
                                    ax.text(xpos[i] + 0.325, ypos[i] + 0.325, dz[i] + (max_val_web * 0.05), 
                                            str(valores[i]), color='black', fontsize=10, fontweight='bold', 
                                            ha='center', va='bottom')

                                ax.set_xticks([0.3, 1.3, 2.3, 3.3])
                                ax.set_xticklabels(etiquetas, fontsize=10, fontweight='bold')
                                ax.set_yticks([]) 
                                ax.set_zlabel("N° Alumnos", fontsize=9)
                                ax.view_init(elev=25, azim=-45)
                                ax.xaxis.pane.fill = False
                                ax.yaxis.pane.fill = False
                                ax.zaxis.pane.fill = False
                                ax.grid(False)
                                
                            else:
                                fig, ax = plt.subplots(figsize=(6, 3))
                                bars = ax.bar(etiquetas, valores, color=colores)
                                ax.bar_label(bars, padding=3, fontweight='bold')
                                ax.set_title("N° Alumnos", pad=10)
                                for spine in ['top', 'right']:
                                    ax.spines[spine].set_visible(False)
                            
                            plt.tight_layout()
                            st.pyplot(fig)
                            plt.close(fig)
                    st.divider()
    else:
        st.warning("⚠️ No se encontraron datos. Verifica tus archivos.")
else:
    st.info("📂 Sube tus archivos (Actas PDF o Registros Excel) para comenzar.")
