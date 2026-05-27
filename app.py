import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import matplotlib.pyplot as plt
import traceback 

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Tablero Pedagógico Universal", layout="wide")
st.title("📊 ESTADISTICA SIAGIE - HWEP")
st.markdown("### Procesador Inteligente de Actas (PDF) y Registros (Excel)")

# --- BARRA LATERAL ---
st.sidebar.header("📂 Cargar Documentos")
uploaded_files = st.sidebar.file_uploader(
    "Sube Actas PDF o Registros Excel", 
    type=["csv", "xlsx", "xls", "pdf"], 
    accept_multiple_files=True
)

# Diccionario de niveles de logro
MAPA_LOGROS = {
    "AD": "Logro Destacado",
    "A": "Logro Esperado",
    "B": "En Proceso",
    "C": "En Inicio",
    "EXO": "Exonerado"
}

# ==========================================
# 0. FUNCIONES ACELERADORAS (Caché de Imágenes y Excel)
# ==========================================
@st.cache_data(show_spinner=False)
def obtener_torta_img(valores, etiquetas, colores, estilo_3d):
    fig, ax = plt.subplots(figsize=(3,3))
    v_clean = [v for v in valores if v > 0]
    l_clean = [l for v, l in zip(valores, etiquetas) if v > 0]
    c_clean = [c for v, c in zip(valores, colores) if v > 0]
    
    if v_clean:
        if estilo_3d:
            separacion = [0.08] * len(v_clean)
            ax.pie(v_clean, labels=l_clean, autopct='%1.1f%%', colors=c_clean, startangle=90, shadow=True, explode=separacion)
        else:
            ax.pie(v_clean, labels=l_clean, autopct='%1.1f%%', colors=c_clean, startangle=90)
            
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def obtener_barra_img(valores, etiquetas, colores, estilo_3d):
    if estilo_3d:
        fig = plt.figure(figsize=(6, 3.5))
        ax = fig.add_subplot(111, projection='3d')
        ax.bar3d([0,1,2,3], [0,0,0,0], [0,0,0,0], [0.65]*4, [0.65]*4, valores, color=colores, shade=True, alpha=0.95)
        max_val_web = max(valores) if max(valores) > 0 else 1
        for i in range(len(valores)):
            ax.text(i + 0.325, 0.325, valores[i] + (max_val_web * 0.05), str(valores[i]), color='black', fontsize=10, fontweight='bold', ha='center', va='bottom')
        ax.set_xticks([0.3, 1.3, 2.3, 3.3]); ax.set_xticklabels(etiquetas, fontsize=10, fontweight='bold')
        ax.set_yticks([]); ax.set_zlabel("N° Alumnos", fontsize=9)
        ax.view_init(elev=25, azim=-45); ax.grid(False)
        ax.xaxis.pane.fill = False; ax.yaxis.pane.fill = False; ax.zaxis.pane.fill = False
    else:
        fig, ax = plt.subplots(figsize=(6, 3))
        bars = ax.bar(etiquetas, valores, color=colores)
        max_val_web = max(valores) if max(valores) > 0 else 1
        for i, v in enumerate(valores):
            ax.text(i, v + (max_val_web * 0.05), str(v), color='black', fontsize=10, fontweight='bold', ha='center', va='bottom')
        ax.set_title("N° Alumnos", pad=10)
        for spine in ['top', 'right']: ax.spines[spine].set_visible(False)
        
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def preparar_matriz_panoramica(df_alumnos):
    matriz = df_alumnos.pivot_table(
        index='Estudiante', 
        columns=['Área', 'Competencia'], 
        values='Nota', 
        aggfunc='first'
    )
    
    total_estudiantes = len(df_alumnos['Estudiante'].unique())
    
    def formatear_stat(counts):
        df_stat = pd.DataFrame({'Cantidad': counts, 'Porcentaje': (counts / total_estudiantes * 100)})
        return df_stat.apply(lambda x: f"N° {int(x['Cantidad'])} ({x['Porcentaje']:.1f}%)", axis=1)

    c_counts = df_alumnos[df_alumnos['Nota'] == 'C'].groupby('Área')['Nota'].count()
    top_c = formatear_stat(c_counts).sort_values(ascending=False).head(3)
    logro_counts = df_alumnos[df_alumnos['Nota'].isin(['A', 'AD'])].groupby('Área')['Nota'].count()
    top_logro = formatear_stat(logro_counts).sort_values(ascending=False).head(3)
    
    # --- EXPORTACIÓN CON COLORES ---
    excel_bytes = io.BytesIO()
    with pd.ExcelWriter(excel_bytes, engine='xlsxwriter') as writer:
        matriz.to_excel(writer, sheet_name='Matriz_Logros')
        workbook = writer.book
        worksheet = writer.sheets['Matriz_Logros']
        
        # Definir formatos de color
        fmt_ad = workbook.add_format({'bg_color': '#1e90ff', 'font_color': 'white', 'bold': True})
        fmt_a = workbook.add_format({'bg_color': '#2ed573', 'font_color': 'black', 'bold': True})
        fmt_b = workbook.add_format({'bg_color': '#ffa502', 'font_color': 'black', 'bold': True})
        fmt_c = workbook.add_format({'bg_color': '#ff4757', 'font_color': 'white', 'bold': True})
        
        # Aplicar formato condicional en el rango de datos de la matriz
        # (Asumiendo que los datos empiezan en la fila 3, col 2)
        rango = f"B3:ZZ{len(matriz)+2}"
        worksheet.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': '"AD"', 'format': fmt_ad})
        worksheet.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': '"A"', 'format': fmt_a})
        worksheet.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': '"B"', 'format': fmt_b})
        worksheet.conditional_format(rango, {'type': 'cell', 'criteria': '==', 'value': '"C"', 'format': fmt_c})
    
    excel_bytes.seek(0)
    return matriz, excel_bytes.getvalue(), top_c, top_logro
# ==========================================
# EXTRA: DETECTOR DE METADATOS (MODO TURBO)
# ==========================================
def extraer_metadatos_siagie(xls_dict):
    nivel = "No detectado"
    grado = "No detectado"
    seccion = "No detectada"
    
    hojas_prioritarias = [k for k in xls_dict.keys() if "GEN" in k.upper() or "PAR" in k.upper()]
    orden_búsqueda = hojas_prioritarias + [k for k in xls_dict.keys() if k not in hojas_prioritarias]
    
    for sheet_name in orden_búsqueda:
        df = xls_dict[sheet_name]
        max_r = min(15, df.shape[0])
        max_c = min(10, df.shape[1])
        
        for r in range(max_r):
            for c in range(max_c):
                val = str(df.iloc[r, c]).strip().upper()
                if not val or val == 'NAN': continue
                
                if "NIVEL" in val and nivel == "No detectado":
                    match = re.search(r'NIVEL\s*:\s*(.+)', val)
                    if match and match.group(1).strip():
                        nivel = match.group(1).strip()
                    else:
                        for next_c in range(c + 1, min(c + 4, df.shape[1])):
                            nv = str(df.iloc[r, next_c]).strip().upper()
                            if nv and nv != 'NAN' and nv != ':':
                                if nv.isdigit() and (next_c + 1) < df.shape[1]:
                                    nnv = str(df.iloc[r, next_c + 1]).strip().upper()
                                    if nnv and nnv != 'NAN': nv = nnv
                                nivel = nv
                                break

                if "GRADO" in val and grado == "No detectado":
                    match = re.search(r'GRADO\s*:\s*(.+)', val)
                    if match and match.group(1).strip():
                        grado = match.group(1).strip()
                    else:
                        for next_c in range(c + 1, min(c + 4, df.shape[1])):
                            nv = str(df.iloc[r, next_c]).strip().upper()
                            if nv and nv != 'NAN' and nv != ':':
                                if nv.isdigit() and (next_c + 1) < df.shape[1]:
                                    nnv = str(df.iloc[r, next_c + 1]).strip().upper()
                                    if nnv and nnv != 'NAN': nv = nnv
                                grado = nv
                                break
                                
                if "SECCI" in val and seccion == "No detectada":
                    match = re.search(r'SECCI[ÓO]N\s*:\s*(.+)', val)
                    if match and match.group(1).strip():
                        seccion = match.group(1).strip()
                    else:
                        for next_c in range(c + 1, min(c + 4, df.shape[1])):
                            nv = str(df.iloc[r, next_c]).strip().upper()
                            if nv and nv != 'NAN' and nv != ':':
                                if nv.isdigit() and (next_c + 1) < df.shape[1]:
                                    nnv = str(df.iloc[r, next_c + 1]).strip().upper()
                                    if nnv and nnv != 'NAN': nv = nnv
                                seccion = nv
                                break
                                
        if nivel != "No detectado" and grado != "No detectado" and seccion != "No detectada":
            break
            
    return nivel.replace('"', ''), grado.replace('"', ''), seccion.replace('"', '')

# ==========================================
# 1. MOTOR DE ANÁLISIS DE PDF (ACTAS)
# ==========================================
def procesar_acta_siagie(pdf_file):
    datos_extraidos = []
    alumnos_extraidos = []
    
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
                    
                    partes_letras = [p for p in partes if p.isalpha() and len(p) > 1 and p not in ["AD", "EXO", "DNI"]]
                    nombre_alumno = " ".join(partes_letras) if partes_letras else f"Estudiante Anónimo"

                    esquema_seleccionado = []
                    if es_inicial:
                        if len(notas) == 10: esquema_seleccionado = esquema_inicial_3
                        elif len(notas) == 11: esquema_seleccionado = esquema_inicial_4
                        elif len(notas) >= 12: esquema_seleccionado = esquema_inicial_5
                    elif es_secundaria:
                        if "DESARROLLO PERSONAL" in text or len(notas) == 14: esquema_seleccionado = esquema_sec_A
                        elif "INGLÉS" in text or "MATEMÁTICA" in text:
                            if len(notas) == 12: esquema_seleccionado = esquema_sec_B
                            elif len(notas) == 15: esquema_seleccionado = [("CASTELLANO", "Oral"), ("CASTELLANO", "Lee"), ("CASTELLANO", "Escribe")] + esquema_sec_B
                    else: 
                        if "PERSONAL SOCIAL" in text: esquema_seleccionado = esquema_prim_A
                        elif "MATEMÁTICA" in text: esquema_seleccionado = esquema_prim_B

                    if esquema_seleccionado and len(notas) == len(esquema_seleccionado):
                        for i, nota in enumerate(notas):
                            area, comp = esquema_seleccionado[i]
                            datos_extraidos.append({"Área": area, "Competencia": comp, "Nota": nota})
                            alumnos_extraidos.append({
                                "Estudiante": nombre_alumno,
                                "Área": area,
                                "Cód": f"C{i+1}",
                                "Competencia": comp,
                                "Nota": nota,
                                "Nivel de Logro": MAPA_LOGROS.get(nota, "")
                            })

    df_resumen = pd.DataFrame()
    if datos_extraidos:
        df_raw = pd.DataFrame(datos_extraidos)
        resumen = []
        for (area, comp), group in df_raw.groupby(["Área", "Competencia"]):
            conteo = group["Nota"].value_counts()
            resumen.append({
                "Área": area, "Cód": "ACTA", "Competencia": comp,
                "AD": conteo.get("AD", 0), "A": conteo.get("A", 0), "B": conteo.get("B", 0),
                "C": conteo.get("C", 0), "EXO": conteo.get("EXO", 0), "Total": conteo.sum()
            })
        df_resumen = pd.DataFrame(resumen)
        
    return df_resumen, pd.DataFrame(alumnos_extraidos)

# ==========================================
# 2. MOTOR INTELIGENTE PARA EXCEL SIAGIE
# ==========================================
def procesar_excel_inteligente(df, nombre_hoja):
    if nombre_hoja.strip().upper() in ["GENERALIDADES", "PARAMETROS", "PARÁMETROS", "MAIN"]: 
        return [], []
    
    mapa_competencias = {}
    for col in df.columns:
        for val in df[col]:
            val_clean = str(val).replace('\n', ' ').strip()
            match = re.search(r'^(?:C|c)?0*(\d+)\s*=\s*(.+)', val_clean)
            if match:
                num = int(match.group(1))
                mapa_competencias[num] = str(match.group(2)).strip()

    df_str = df.astype(str).apply(lambda x: x.str.strip().str.upper())
    
    col_nombres = -1
    for c in df_str.columns:
        if df_str[c].head(25).str.contains(r'\bNOMBRES\b|\bAPELLIDO\b', regex=True).any():
            col_nombres = c
            break

    columnas_notas = []
    for col in df_str.columns:
        if df_str[col].isin(['NL', 'CALIF.', 'NIVEL DE LOGRO', 'CALIF']).any():
            columnas_notas.append(col)
    
    if not columnas_notas: return [], []

    idx_cabecera = -1
    for col in columnas_notas:
        idx = df_str[df_str[col].isin(['NL', 'CALIF.', 'NIVEL DE LOGRO', 'CALIF'])].index[0]
        idx_cabecera = max(idx_cabecera, idx)

    area = "OTRA ÁREA"
    n = nombre_hoja.upper()
    if "COMU" in n: area = "COMUNICACIÓN"
    elif "MATE" in n: area = "MATEMÁTICA"
    elif "CIENC" in n or "CYT" in n: area = "CIENCIA Y TECNOLOGÍA"
    elif "DESARR" in n or "PCC" in n or "DPCC" in n: area = "DESARROLLO PERS., CIUD. Y CÍVICA"
    elif "CCSS" in n or "SOCIAL" in n: area = "CIENCIAS SOCIALES"
    elif "PPSS" in n: area = "PERSONAL SOCIAL" 
    elif "ART" in n: area = "ARTE Y CULTURA"
    elif "EFIS" in n or "FIS" in n: area = "EDUCACIÓN FÍSICA"
    elif "ETRA" in n or "EPT" in n: area = "EDUCACIÓN PARA EL TRABAJO"
    elif "EREL" in n or "RELIG" in n: area = "EDUCACIÓN RELIGIOSA"
    elif "INGL" in n: area = "INGLÉS"
    elif "CAST" in n: area = "CASTELLANO COMO 2DA LENGUA"
    elif "TIC" in n: area = "COMP. TRANSV. TIC"
    elif "GEST" in n or "AUTO" in n: area = "COMP. TRANSV. AUTONOMÍA"

    resultados = []
    alumnos_detalle = []
    
    for i, col_idx in enumerate(columnas_notas):
        num_comp = i + 1
        raw_code = str(df.iloc[idx_cabecera - 1, col_idx]).strip() if idx_cabecera > 0 else ""
            
        match_num = re.search(r'\d+', raw_code)
        if match_num: num_comp = int(match_num.group(0))

        if num_comp in mapa_competencias:
            comp_name = mapa_competencias[num_comp]
        else:
            comp_name = raw_code
            if comp_name.lower() in ['nan', 'nl', 'calif.', ''] and idx_cabecera > 1:
                comp_name = str(df.iloc[idx_cabecera - 2, col_idx]).strip()
            if comp_name.lower() in ['nan', '', 'nl']:
                comp_name = f"Competencia {num_comp}"
        
        comp_final = f"C{num_comp} - {comp_name[:120]}"

        notas_raw = df_str.iloc[idx_cabecera + 1:, col_idx]
        valid_idx = notas_raw[notas_raw.isin(['AD', 'A', 'B', 'C', 'EXO'])].index
        
        for idx in valid_idx:
            nota = notas_raw.loc[idx]
            nombre_est = str(df.iloc[idx, col_nombres]).strip() if col_nombres != -1 else f"Estudiante de la Fila {idx}"
            if nombre_est.upper() not in ['NAN', 'NONE', '']:
                nombre_est = re.sub(r'\s+', ' ', nombre_est)
                alumnos_detalle.append({
                    "Estudiante": nombre_est,
                    "Área": area,
                    "Cód": f"C{num_comp}",
                    "Competencia": comp_final,
                    "Nota": nota,
                    "Nivel de Logro": MAPA_LOGROS.get(nota, "")
                })

        conteo = notas_raw[notas_raw.isin(['AD', 'A', 'B', 'C', 'EXO'])].value_counts()
        if not conteo.empty:
            resultados.append({
                "Área": area,
                "Cód": f"C{num_comp}",
                "Competencia": comp_final,
                "AD": conteo.get("AD", 0), "A": conteo.get("A", 0),
                "B": conteo.get("B", 0), "C": conteo.get("C", 0),
                "EXO": conteo.get("EXO", 0),
                "Total": conteo.get("AD", 0) + conteo.get("A", 0) + conteo.get("B", 0) + conteo.get("C", 0)
            })
            
    return resultados, alumnos_detalle

# ==========================================
# 3. EXPORTADOR EXCEL 3D / 2D
# ==========================================
@st.cache_data(show_spinner=False)
def generar_excel_con_graficos(df_resultados, estilo_3d=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultados.to_excel(writer, sheet_name='Consolidado_Bimestre', index=False)
        worksheet = writer.sheets['Consolidado_Bimestre']
        worksheet.set_column('A:A', 22); worksheet.set_column('B:B', 8)
        worksheet.set_column('C:C', 55); worksheet.set_column('D:H', 10)
        worksheet.set_column('I:I', 12)  
        
        fila_grafico = 1
        for idx, row in df_resultados.iterrows():
            categorias = ['AD', 'A', 'B', 'C']
            valores = [row['AD'], row['A'], row['B'], row['C']]
            colores = ['#2ed573', '#1e90ff', '#ffa502', '#ff4757'] 
            
            if estilo_3d:
                fig = plt.figure(figsize=(6, 3.5))
                ax = fig.add_subplot(111, projection='3d')
                ax.bar3d([0, 1, 2, 3], [0, 0, 0, 0], [0, 0, 0, 0], [0.65]*4, [0.65]*4, valores, color=colores, shade=True, alpha=0.95)
                max_val = max(valores) if max(valores) > 0 else 1
                for i in range(len(valores)):
                    ax.text(i + 0.325, 0.325, valores[i] + (max_val * 0.05), str(valores[i]), color='black', fontsize=10, fontweight='bold', ha='center', va='bottom')
                ax.set_xticks([0.3, 1.3, 2.3, 3.3]); ax.set_xticklabels(categorias, fontsize=10, fontweight='bold')
                ax.set_yticks([]) 
                ax.set_title(f"[MODO 3D] {row['Cód']} - {row['Área']}", fontsize=11, fontweight='bold', pad=15)
                ax.view_init(elev=25, azim=-45); ax.grid(False) 
                ax.xaxis.pane.fill = False; ax.yaxis.pane.fill = False; ax.zaxis.pane.fill = False
                plt.tight_layout()
            else:
                fig, ax = plt.subplots(figsize=(5.5, 3))
                bars = ax.bar(categorias, valores, color=colores, edgecolor='none', width=0.6)
                max_val = max(valores) if max(valores) > 0 else 1
                for i, v in enumerate(valores):
                    ax.text(i, v + (max_val * 0.05), str(v), color='black', fontsize=10, fontweight='bold', ha='center', va='bottom')
                ax.set_title(f"[MODO 2D] {row['Cód']} - {row['Área']}", fontsize=11, fontweight='bold', pad=10)
                for spine in ['top', 'right']: ax.spines[spine].set_visible(False)
                plt.tight_layout()

            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=100)
            img_buffer.seek(0)
            plt.close(fig) 
            
            worksheet.insert_image(f"K{fila_grafico + 1}", f"grafico_{idx}.png", {'image_data': img_buffer})
            fila_grafico += 18 if estilo_3d else 16
    return output.getvalue()

# ==========================================
# 4. ORQUESTADOR PRINCIPAL
# ==========================================
def procesar_todo(files_list):
    todos_resultados = []
    todos_alumnos = []
    nivel_final = "No detectado"
    grado_final = "No detectado"
    seccion_final = "No detectada"
    
    for uploaded_file in files_list:
        name = uploaded_file.name
        if name.lower().endswith('.pdf'):
            try:
                df_res, df_alum = procesar_acta_siagie(uploaded_file)
                if not df_res.empty: 
                    todos_resultados.extend(df_res.to_dict('records'))
                    todos_alumnos.extend(df_alum.to_dict('records'))
            except Exception as e:
                st.error(f"⚠️ Error PDF: {e}")
        elif name.lower().endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, header=None)
                res, alum = procesar_excel_inteligente(df, "CSV")
                if res: 
                    todos_resultados.extend(res)
                    todos_alumnos.extend(alum)
            except Exception as e:
                st.error(f"⚠️ Error CSV: {e}")
        else:
            try:
                xls_dict = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                n, g, s = extraer_metadatos_siagie(xls_dict)
                if n != "No detectado": nivel_final = n
                if g != "No detectado": grado_final = g
                if s != "No detectada": seccion_final = s
                
                for sheet_name, df_sheet in xls_dict.items():
                    res, alum = procesar_excel_inteligente(df_sheet, sheet_name)
                    if res: 
                        todos_resultados.extend(res)
                        todos_alumnos.extend(alum)
            except Exception as e:
                st.error(f"⚠️ Error Excel: {e}\n```\n{traceback.format_exc()}\n```")
                
    return pd.DataFrame(todos_resultados), pd.DataFrame(todos_alumnos), nivel_final, grado_final, seccion_final

# ==========================================
# FUNCIONES DE COLOR PARA LAS TABLAS
# ==========================================
def aplicar_colores_logro(val):
    color = ''
    texto = 'black'
    val_str = str(val).strip().upper()
    
    if val_str in ['AD', 'LOGRO DESTACADO']:
        color = '#1e90ff' # Azul
        texto = 'white'
    elif val_str in ['A', 'LOGRO ESPERADO']:
        color = '#2ed573' # Verde
        texto = 'black'
    elif val_str in ['B', 'EN PROCESO']:
        color = '#ffa502' # Amarillo
        texto = 'black'
    elif val_str in ['C', 'EN INICIO']:
        color = '#ff4757' # Rojo
        texto = 'white'
        
    if color:
        return f'background-color: {color}; color: {texto}; font-weight: bold; text-align: center;'
    return ''

# ==========================================
# 5. INTERFAZ Y RENDERIZADO 
# ==========================================
try:
    if uploaded_files:
        
        # Reseteo de memoria en caliente para evitar errores de transición
        if "app_version" not in st.session_state or st.session_state.app_version != "v5_hyper":
            st.session_state.clear()
            st.session_state.app_version = "v5_hyper"
            
        archivos_actuales = [(f.name, f.size) for f in uploaded_files]
        
        if "archivos_subidos" not in st.session_state or st.session_state.archivos_subidos != archivos_actuales:
            with st.spinner('Procesando documentos a ultra velocidad...'):
                df_resumen, df_alumnos, nivel, grado, seccion = procesar_todo(uploaded_files)
                
                # Seguro contra archivos con lectura parcial
                if not df_alumnos.empty and "Cód" not in df_alumnos.columns:
                    df_alumnos["Cód"] = "C?" 
                    
                st.session_state.df_resumen = df_resumen
                st.session_state.df_alumnos = df_alumnos
                st.session_state.nivel = nivel
                st.session_state.grado = grado
                st.session_state.seccion = seccion
                st.session_state.archivos_subidos = archivos_actuales
        
        df_resumen = st.session_state.df_resumen
        df_alumnos = st.session_state.df_alumnos
        nivel = st.session_state.nivel
        grado = st.session_state.grado
        seccion = st.session_state.seccion
        
        if not df_resumen.empty:
            df_resumen = df_resumen.sort_values(by=["Área", "Competencia"])
            st.success(f"✅ ¡Datos leídos exitosamente! Se procesaron {len(df_alumnos['Estudiante'].unique()) if not df_alumnos.empty else 0} estudiantes.")
            
            col_n, col_g, col_s = st.columns(3)
            with col_n: st.info(f"🏫 **Nivel:** {nivel}")
            with col_g: st.info(f"📖 **Grado:** {grado}")
            with col_s: st.info(f"👥 **Sección:** {seccion}")
            
            # --- TRES PESTAÑAS PRINCIPALES ---
            tab_general, tab_alumnos, tab_panoramica = st.tabs([
                "📈 Análisis General", 
                "🧑‍🎓 Reporte Individual", 
                "👁️ Visión Panorámica"
            ])
            
            # ----------------------------------------
            # PESTAÑA 1: ANÁLISIS GENERAL (Gráficos)
            # ----------------------------------------
            with tab_general:
                st.subheader("📋 Consolidado de Aula")
                st.dataframe(df_resumen, use_container_width=True)
                
                st.divider()
                st.markdown("### ⚙️ Descarga de Reportes")
                opcion_grafico = st.radio(
                    "Estilo de gráficos para el Excel:",
                    ["✨ Gráficos en 3D (Volumen)", "📊 Gráficos en 2D (Planos)"],
                    horizontal=True
                )
                es_estilo_3d = "3D" in opcion_grafico
                
                # Excel 100% aislado en Caché
                excel_data = generar_excel_con_graficos(df_resumen, es_estilo_3d)

                nombre_archivo = f"Reporte_SIAGIE_{nivel.replace(' ', '_')}_{grado.replace(' ', '_')}_{seccion.replace(' ', '_')}.xlsx"
                st.download_button(
                    label="📥 Descargar Excel con Gráficos",
                    data=excel_data,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.divider()
                st.markdown("### 📈 Visualización Web")
                areas = df_resumen["Área"].unique()
                for area in areas:
                    with st.expander(f"📘 {area}", expanded=False):
                        df_area = df_resumen[df_resumen["Área"] == area]
                        for _, row in df_area.iterrows():
                            st.markdown(f"**{row['Competencia']}**")
                            col_torta, col_barra = st.columns([1, 2])
                            
                            valores = [row["AD"], row["A"], row["B"], row["C"]]
                            etiquetas = ["AD", "A", "B", "C"]
                            colores = ["#2ed573", "#1e90ff", "#ffa502", "#ff4757"] 
                            
                            with col_torta:
                                if sum(valores) > 0:
                                    img_torta = obtener_torta_img(tuple(valores), tuple(etiquetas), tuple(colores), es_estilo_3d)
                                    st.image(img_torta)
                            
                            with col_barra:
                                if sum(valores) > 0:
                                    img_barra = obtener_barra_img(tuple(valores), tuple(etiquetas), tuple(colores), es_estilo_3d)
                                    st.image(img_barra)
                            st.divider()

            # ----------------------------------------
            # PESTAÑA 2: REPORTE INDIVIDUAL (Alumnos)
            # ----------------------------------------
            with tab_alumnos:
                if not df_alumnos.empty:
                    df_alumnos_unicos = df_alumnos.drop_duplicates()
                    lista_estudiantes = sorted(df_alumnos_unicos["Estudiante"].unique())
                    
                    st.markdown("### 🔍 Búsqueda de Estudiante")
                    estudiante_seleccionado = st.selectbox("Selecciona un nombre para ver sus calificaciones detalladas:", lista_estudiantes)
                    
                    df_filtrado = df_alumnos_unicos[df_alumnos_unicos["Estudiante"] == estudiante_seleccionado]
                    conteo_notas = df_filtrado["Nota"].value_counts()
                    
                    st.markdown("#### 🎯 Resumen de Logros")
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔵 Logro Destacado (AD)", conteo_notas.get("AD", 0))
                    col2.metric("🟢 Logro Esperado (A)", conteo_notas.get("A", 0))
                    col3.metric("🟠 En Proceso (B)", conteo_notas.get("B", 0))
                    col4.metric("🔴 En Inicio (C)", conteo_notas.get("C", 0))
                    
                    st.markdown("#### 📖 Detalle Curricular")
                    df_mostrar = df_filtrado[["Área", "Competencia", "Nota", "Nivel de Logro"]].sort_values(by=["Área", "Competencia"]).reset_index(drop=True)
                    
                    try:
                        df_coloreado = df_mostrar.style.applymap(aplicar_colores_logro, subset=["Nota", "Nivel de Logro"])
                    except Exception:
                        df_coloreado = df_mostrar.style.map(aplicar_colores_logro, subset=["Nota", "Nivel de Logro"])
                        
                    st.dataframe(df_coloreado, use_container_width=True)
                else:
                    st.info("No se pudieron extraer los nombres de los estudiantes.")

            # ----------------------------------------
# PESTAÑA 3: VISIÓN PANORÁMICA (Matriz)
# ----------------------------------------
# ----------------------------------------
# PESTAÑA 3: VISIÓN PANORÁMICA (Matriz)
# ----------------------------------------
# 5. Contenido Pestaña 3
      # 5. Contenido Pestaña 3
        # 5. Contenido Pestaña 3 (dentro de tab_panoramica)
        with tab_panoramica:
            if 'df_alumnos' in st.session_state and not st.session_state.df_alumnos.empty:
                st.markdown("### 👁️ Matriz de Logros de toda el Aula")
                
                # Generación
                matriz, excel_matriz_bytes, top_c, top_logro = preparar_matriz_panoramica(st.session_state.df_alumnos)
                
                # Visualización de la Matriz
                try:
                    matriz_coloreada = matriz.style.map(aplicar_colores_logro)
                except Exception:
                    matriz_coloreada = matriz.style.applymap(aplicar_colores_logro)
                st.dataframe(matriz_coloreada, use_container_width=True)
                
                st.divider()
                
                # Estadísticas con N° y %
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("#### 🚩 Áreas con más alumnos en Inicio (C)")
                    st.table(top_c.rename("Estadística"))
                with col2:
                    st.markdown("#### 🏆 Áreas con más alumnos en Logro (A/AD)")
                    st.table(top_logro.rename("Estadística"))
               
                          
                st.markdown("#### 📥 Exportar Matriz")
                nombre_matriz = f"Matriz_Panoramica_{str(st.session_state.get('grado', 'Grado')).replace(' ', '_')}.xlsx"
                st.download_button(
                    label="📥 Descargar Matriz Panorámica en Excel",
                    data=excel_matriz_bytes,
                    file_name=nombre_matriz,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Carga documentos para ver la visión panorámica.")
    else:
        st.info("📂 Sube tus archivos (Actas PDF o Registros Excel) para comenzar.")

except Exception as e_fatal:
    st.error(f"⚠️ ERROR FATAL EN LA INTERFAZ. Cópialo y envíalo por el chat:\n\n```\n{traceback.format_exc()}\n```")
