import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
import datetime
import io
import re
import html
import numpy as np
from pathlib import Path

st.set_page_config(page_title="Procesador SOV v2.0", layout="wide")

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

# Plantilla de orden estándar definida de manera global
CONTEO_TEMPLATE = [
    ("Chery 01-18 | |19-31", "Codificación Audiovisuales", "ANCHERY"),
    ("Chery 01-18 | |19-31", "Codificación Impresos", "ANCHERY"),
    ("Chery 01-18 | |19-31", "Notas Audiovisuales", "ANCHERY"),
    ("Chery 01-18 | |19-31", "Notas Impresos", "ANCHERY"),
    ("Chery - Changan, Competencias", "Codificación Audiovisuales", "ANCHERY"),
    ("Chery - Changan, Competencias", "Codificación Impresos", "ANCHERY"),
    ("Chery - Changan, Competencias", "Notas Audiovisuales", "ANCHERY"),
    ("Chery - Changan, Competencias", "Notas Impresos", "ANCHERY"),
    ("Comfenalco Valle", "Codificación Audiovisuales", "ACOMFEVALLE"),
    ("Comfenalco Valle", "Codificación Impresos", "ACOMFEVALLE"),
    ("Comfenalco Valle", "Notas Audiovisuales", "ACOMFEVALLE"),
    ("Comfenalco Valle", "Notas Impresos", "ACOMFEVALLE"),
    ("Federación Nacional de Avicultores de Colombia", "Codificación Audiovisuales", "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Codificación Impresos", "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Notas Audiovisuales", "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Notas Impresos", "ANFENAVI"),
    ("Fundación Santa Fe de Bogotá", "Codificación Audiovisuales", "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá", "Codificación Impresos", "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá", "Notas Audiovisuales", "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá", "Notas Impresos", "FSANTAFE_AN"),
    ("Nissan", "Codificación Audiovisuales", "ANNISSAN"),
    ("Nissan", "Codificación Impresos", "ANNISSAN"),
    ("Nissan", "Notas Audiovisuales", "ANNISSAN"),
    ("Nissan", "Notas Impresos", "ANNISSAN"),
    ("Nissan, Competencia", "Codificación Audiovisuales", "ANNISSAN"),
    ("Nissan, Competencia", "Codificación Impresos", "ANNISSAN"),
    ("Nissan, Competencia", "Notas Audiovisuales", "ANNISSAN"),
    ("Nissan, Competencia", "Notas Impresos", "ANNISSAN"),
    ("Tigo", "Codificación Audiovisuales", "TIGOAN"),
    ("Tigo", "Codificación Impresos", "TIGOAN"),
    ("Tigo", "Notas Audiovisuales", "TIGOAN"),
    ("Tigo", "Notas Impresos", "TIGOAN"),
    ("Universidad Simón Bolívar", "Codificación Audiovisuales", "USIMONAN"),
    ("Universidad Simón Bolívar", "Codificación Impresos", "USIMONAN"),
    ("Universidad Simón Bolívar", "Notas Audiovisuales", "USIMONAN"),
    ("Universidad Simón Bolívar", "Notas Impresos", "USIMONAN"),
    ("Universidad Tecnológica de Bolívar", "Codificación Audiovisuales", "UTB_AN"),
    ("Universidad Tecnológica de Bolívar", "Codificación Impresos", "UTB_AN"),
    ("Universidad Tecnológica de Bolívar", "Notas Audiovisuales", "UTB_AN"),
    ("Universidad Tecnológica de Bolívar", "Notas Impresos", "UTB_AN")
]

# Buscar el archivo de configuración con cualquier variante de nombre/mayúsculas
def find_config_path():
    base = Path(__file__).parent
    for f in base.iterdir():
        if f.suffix.lower() == '.xlsx' and 'config' in f.stem.lower():
            return f
    return None

CONFIG_PATH = find_config_path()

def extract_link_from_cell(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return None

def convert_html_entities(text):
    if not isinstance(text, str):
        return text
    text = html.unescape(text)
    html_entities = {
        '&#xF3;': 'ó', '&#xE1;': 'á', '&#xE9;': 'é', '&#xED;': 'í',
        '&#xFA;': 'ú', '&#xF1;': 'ñ', '&#xDC;': 'Ü', '&#xFC;': 'ü',
        '&#xC1;': 'Á', '&#xC9;': 'É', '&#xCD;': 'Í', '&#xD3;': 'Ó',
        '&#xDA;': 'Ú', '&#xD1;': 'Ñ', '&#xC7;': 'Ç', '&#xE7;': 'ç',
    }
    for entity, char in html_entities.items():
        text = text.replace(entity, char)

    def replace_hex_entity(match):
        try:
            return chr(int(match.group(1), 16))
        except Exception:
            return match.group(0)

    def replace_decimal_entity(match):
        try:
            return chr(int(match.group(1)))
        except Exception:
            return match.group(0)

    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex_entity, text)
    text = re.sub(r'&#(\d+);', replace_decimal_entity, text)

    for bad, good in {'\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
                      'Â': '', 'â': '', '€': '', '™': ''}.items():
        text = text.replace(bad, good)
    return text

def clean_text(text):
    if not isinstance(text, str):
        return text
    return convert_html_entities(text).strip()

def clean_cuerpo(text):
    if not isinstance(text, str) or text.strip() == '':
        return text
    text = convert_html_entities(text)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    return text.strip()

def get_client_category(filename):
    """Determina a qué categoría pertenece el archivo según su nombre y estilo de codificación."""
    fn = filename.lower()
    
    is_competencia = False
    if re.search(r'\b(c|com|comp|competencia|competencias|changan)\b', fn) or "_c" in fn or "-c" in fn:
        is_competencia = True
        
    # 1. Detección de Chery (Marca vs Competencia)
    if "chery" in fn or "anchery" in fn:
        if is_competencia:
            return "Chery - Changan, Competencias"
        else:
            return "Chery 01-18 | |19-31"
            
    # 2. Detección de Nissan (Marca vs Competencia)
    elif "niss" in fn or "nissan" in fn or "annissan" in fn:
        if is_competencia:
            return "Nissan, Competencia"
        else:
            return "Nissan"
            
    # 3. Detección de Comfenalco Valle
    elif "comfe" in fn or "comfenalco" in fn or "acomfevalle" in fn:
        return "Comfenalco Valle"
        
    # 4. Detección de Federación Nacional de Avicultores de Colombia
    elif any(x in fn for x in ["fenavi", "avicultores", "avicola", "anfenavi"]):
        return "Federación Nacional de Avicultores de Colombia"
        
    # 5. Detección de Fundación Santa Fe de Bogotá
    elif any(x in fn for x in ["fsant", "santa", "santafe", "fsantafe_an"]):
        return "Fundación Santa Fe de Bogotá"
        
    # 6. Detección de Tigo
    elif "tigo" in fn or "tigoan" in fn:
        return "Tigo"
        
    # 7. Detección de Universidad Simón Bolívar
    elif any(x in fn for x in ["simon", "usimon", "usim", "usimonan"]):
        return "Universidad Simón Bolívar"
        
    # 8. Detección de Universidad Tecnológica de Bolívar
    elif any(x in fn for x in ["utb", "tecnologica", "utb_an"]):
        return "Universidad Tecnológica de Bolívar"
        
    return None

def build_conteo_df(filename, av_count, grafica_count):
    """Construye un DataFrame con el formato y orden para un archivo aplicando reglas de replicación."""
    matched_client = get_client_category(filename)
    
    rows = []
    for client, tipo, codigo in CONTEO_TEMPLATE:
        val = 0
        is_replicated_client = client in [
            "Chery 01-18 | |19-31",
            "Chery - Changan, Competencias",
            "Nissan",
            "Nissan, Competencia"
        ]
        
        if client == matched_client:
            if tipo == "Notas Audiovisuales":
                val = av_count
            elif tipo == "Notas Impresos":
                val = grafica_count
            elif tipo == "Codificación Audiovisuales":
                val = av_count if is_replicated_client else 0
            elif tipo == "Codificación Impresos":
                val = grafica_count if is_replicated_client else 0
        
        rows.append({
            "Cliente / Categoría": client,
            "Tipo de Conteo": tipo,
            "Código": codigo,
            "Cantidad": val
        })
        
    return pd.DataFrame(rows), matched_client

def build_consolidated_conteo(processed_files):
    """Genera una tabla de conteo acumulando todos los archivos con sus respectivas reglas aplicadas."""
    client_av = {}
    client_graf = {}
    
    for item in processed_files:
        matched_client = get_client_category(item['filename'])
        if matched_client:
            client_av[matched_client] = client_av.get(matched_client, 0) + item['av_count']
            client_graf[matched_client] = client_graf.get(matched_client, 0) + item['grafica_count']
            
    rows = []
    for client, tipo, codigo in CONTEO_TEMPLATE:
        val = 0
        av_count = client_av.get(client, 0)
        graf_count = client_graf.get(client, 0)
        
        is_replicated_client = client in [
            "Chery 01-18 | |19-31",
            "Chery - Changan, Competencias",
            "Nissan",
            "Nissan, Competencia"
        ]
        
        if tipo == "Notas Audiovisuales":
            val = av_count
        elif tipo == "Notas Impresos":
            val = graf_count
        elif tipo == "Codificación Audiovisuales":
            val = av_count if is_replicated_client else 0
        elif tipo == "Codificación Impresos":
            val = graf_count if is_replicated_client else 0
            
        rows.append({
            "Cliente / Categoría": client,
            "Tipo de Conteo": tipo,
            "Código": codigo,
            "Cantidad": val
        })
        
    return pd.DataFrame(rows)

def to_excel_consolidated(df_conteo):
    """Genera un archivo Excel exclusivo para la tabla consolidada de conteos."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Conteo Consolidado"
    
    for col_idx, col_name in enumerate(df_conteo.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        
    for row_idx, row_data in enumerate(df_conteo.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def to_excel_from_df(df, final_order, filename, av_count, grafica_count):
    output = io.BytesIO()
    cols = [c for c in final_order if c in df.columns]
    df_out = df[cols].copy()

    for col in df_out.columns:
        if hasattr(df_out[col].dtype, 'pyarrow_dtype'):
            df_out[col] = df_out[col].astype(object)

    wb = Workbook()
    
    # 1. Hoja "Resultado"
    ws = wb.active
    ws.title = 'Resultado'

    for i, col_name in enumerate(df_out.columns, start=1):
        cell = ws.cell(row=1, column=i, value=col_name)
        cell.font = Font(bold=True)

    link_cols = {'Link Nota', 'Link (Streaming - Imagen)'}

    for row_idx, row_data in enumerate(df_out.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row_data, start=1):
            col_name = df_out.columns[col_idx - 1]
            cell = ws.cell(row=row_idx, column=col_idx)
            if col_name == 'Fecha' and pd.notna(value):
                if isinstance(value, pd.Timestamp):
                    cell.value = value.to_pydatetime()
                    cell.number_format = 'DD/MM/YYYY'
                else:
                    cell.value = value
            elif col_name in link_cols and pd.notna(value) and isinstance(value, str) and value.startswith('http'):
                cell.value = 'Link'
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.value = value if pd.notna(value) else None

    for i, col_name in enumerate(df_out.columns, start=1):
        letter = ws.cell(row=1, column=i).column_letter
        if col_name in ['Título', 'Resumen - Aclaracion']:
            ws.column_dimensions[letter].width = 50
        elif col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 20

    # 2. Hoja "Conteo" (pestaña adicional propia del archivo)
    ws2 = wb.create_sheet(title='Conteo')
    df_conteo, _ = build_conteo_df(filename, av_count, grafica_count)
    
    for col_idx, col_name in enumerate(df_conteo.columns, start=1):
        cell = ws2.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        
    for row_idx, row_data in enumerate(df_conteo.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_data, start=1):
            ws2.cell(row=row_idx, column=col_idx, value=val)
            
    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    output.seek(0)
    return output.getvalue()

def load_config(config_source):
    """Carga region_map e internet_map desde archivo (ruta o file-like)."""
    config_sheets = pd.read_excel(config_source, sheet_name=None, engine='openpyxl')
    region_map = pd.Series(
        config_sheets['Regiones'].iloc[:, 1].values,
        index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
    ).to_dict()
    internet_map = pd.Series(
        config_sheets['Internet'].iloc[:, 1].values,
        index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
    ).to_dict()
    return region_map, internet_map

# ==============================================================================
# PROCESAMIENTO DE UN SOLO DOSSIER
# ==============================================================================

def process_dossier(dossier_file, region_map, internet_map):
    """Procesa un archivo dossier y retorna (df_procesado, av_count, grafica_count)."""
    wb = load_workbook(dossier_file)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1] if cell.value is not None]
    rows = []
    for row in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row):
            continue
        row_data = dict(zip(headers, [c.value for c in row[:len(headers)]]))
        for lc in ['URL Nota AV', 'URL (Streaming - Imagen)', 'URL Nota', 'Link Nota AV', 'Link (Streaming - Imagen)']:
            if lc in headers:
                idx = headers.index(lc)
                if idx < len(row):
                    ext = extract_link_from_cell(row[idx])
                    if ext:
                        row_data[lc] = ext
        rows.append(row_data)

    df = pd.DataFrame(rows)

    tipo_medio_map = {
        'online': 'Internet', 'internet': 'Internet',
        'diario': 'Prensa',
        'am': 'Radio', 'fm': 'Radio',
        'aire': 'Televisión', 'cable': 'Televisión',
        'revista': 'Revistas', 'revistas': 'Revistas',
    }
    df['Tipo de Medio'] = (
        df['Tipo de Medio'].astype(str).str.lower().str.strip()
        .map(tipo_medio_map)
        .fillna(df['Tipo de Medio'].astype(str).str.strip())
    )

    is_av = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
    is_grafica = df['Tipo de Medio'].isin(['Prensa', 'Internet', 'Revistas'])
    is_internet = df['Tipo de Medio'] == 'Internet'

    # Buscarv Internet en columna Medio
    df.loc[is_internet, 'Medio'] = (
        df.loc[is_internet, 'Medio']
        .astype(str).str.lower().str.strip()
        .map(internet_map)
        .fillna(df.loc[is_internet, 'Medio'])
    )

    # Buscarv Región
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)

    df['ID Noticia'] = df.get('NoticiaId', pd.Series(dtype=str))
    df['Fecha'] = pd.to_datetime(df.get('Fecha', pd.Series(dtype=str)), dayfirst=True, errors='coerce').dt.normalize()
    df['Hora'] = df.get('Hora', pd.Series(dtype=str))
    df['Sección - Programa'] = df.get('Sección - Programa', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Título'] = df.get('Título', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Autor - Conductor'] = df.get('Autor - Conductor', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Nro. Pagina'] = df.get('Nro. Pagina', pd.Series(dtype=str))
    df['Dimensión'] = df.get('Dimensioncm2', pd.Series(dtype=str))
    df['Duración - Nro. Caracteres'] = df.get('Duración - Nro. Caracteres', pd.Series(dtype=str))

    # AV: mover Duración a Dimensión y dejar Duración en 0
    df.loc[is_av, 'Dimensión'] = df.loc[is_av, 'Duración - Nro. Caracteres']
    df.loc[is_av, 'Duración - Nro. Caracteres'] = 0

    cpe_av = df.get('CPE', pd.Series([np.nan] * len(df)))
    cpe_grafica = df.get('Valor de Nota', pd.Series([np.nan] * len(df)))
    df['CPE'] = np.where(is_av, cpe_av, np.where(is_grafica, cpe_grafica, np.nan))

    df['Tier'] = df.get('Tier', pd.Series(dtype=str))
    df['Audiencia'] = df.get('Audiencia', pd.Series(dtype=str))
    df['Tono'] = df.get('Tono', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Tema'] = df.get('Tematica', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Temas Generales - Tema'] = df.get('Temas Generales - Tema', pd.Series(dtype=str)).astype(str).apply(clean_text)

    cuerpo_cleaned = df.get('CuerpoEs', pd.Series([''] * len(df))).astype(str).apply(clean_cuerpo)

    def fmt_grafica(text):
        if not isinstance(text, str) or not text.strip():
            return text
        parrafos = [p.strip() for p in text.split('\n') if p.strip()]
        return '\n\n'.join(parrafos) if len(parrafos) > 1 else text

    df['Resumen - Aclaracion'] = np.where(is_av, cuerpo_cleaned, cuerpo_cleaned.apply(fmt_grafica))

    url_nota_av = df.get('URL Nota AV', df.get('Link Nota AV', pd.Series([''] * len(df)))).fillna('').astype(str)
    url_streaming = df.get('URL (Streaming - Imagen)', pd.Series([''] * len(df))).fillna('').astype(str)
    link_nota_av = url_nota_av.str.replace(r'\.com\.ar', '.com.co', regex=True)
    df['Link Nota'] = np.where(is_av, link_nota_av, np.where(is_grafica, url_streaming, ''))
    df['Link Nota'] = df['Link Nota'].replace('', np.nan)

    df['Link (Streaming - Imagen)'] = df.get('URL Nota', pd.Series([''] * len(df))).fillna('').astype(str).replace('', np.nan)

    menciones_av = df.get('Menciones - Empresa', pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    menciones_grafica = df.get('Empresa rel.', pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    df['Menciones - Empresa'] = np.where(is_av, menciones_av, np.where(is_grafica, menciones_grafica, menciones_av))

    # Expansión por ; en Menciones - Empresa
    rows_expanded = []
    for _, row in df.iterrows():
        menciones = [m.strip() for m in str(row['Menciones - Empresa']).split(';') if m.strip()]
        if not menciones:
            rows_expanded.append(row.to_dict())
        else:
            for m in menciones:
                new_row = row.to_dict()
                new_row['Menciones - Empresa'] = m
                rows_expanded.append(new_row)
    df = pd.DataFrame(rows_expanded).reset_index(drop=True)

    is_av_f = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
    is_grafica_f = df['Tipo de Medio'].isin(['Prensa', 'Internet', 'Revistas'])

    return df, int(is_av_f.sum()), int(is_grafica_f.sum())

# ==============================================================================
# INTERFAZ
# ==============================================================================

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

st.title("🚀 Procesador SOV v2.0")
st.markdown("Transforma archivos de entrada al formato estándar de salida con todos los mapeos aplicados.")

# --- Configuración ---
config_source = None
config_label = ""

if CONFIG_PATH is not None:
    config_source = CONFIG_PATH
    config_label = f"✅ Configuración cargada desde el repo: `{CONFIG_PATH.name}`"
    st.success(config_label)
else:
    st.warning("⚠️ No se encontró `Configuracion.xlsx` en el repositorio. Súbelo manualmente:")
    config_upload = st.file_uploader("Subir Configuracion.xlsx", type=["xlsx"], key="config_upload")
    if config_upload:
        config_source = config_upload
        st.success(f"✅ Configuración cargada: {config_upload.name}")

with st.expander("📋 Ver mapeo de columnas aplicado"):
    st.markdown("""
| Columna Salida | Fuente |
|---|---|
| ID Noticia | NoticiaId |
| Fecha | Fecha |
| Hora | Hora |
| Medio | Medio (Internet: reemplazado por hoja Internet) |
| Tipo de Medio | Tipo de Medio (normalizado) |
| Sección - Programa | Sección - Programa |
| Región | Generada desde hoja Regiones en Configuracion.xlsx |
| Título | Título |
| Autor - Conductor | Autor - Conductor |
| Nro. Pagina | Nro. Pagina |
| Dimensión | Dimensioncm2 (Gráfica) / Duración - Nro. Caracteres (AV) |
| Duración - Nro. Caracteres | Duración - Nro. Caracteres (Gráfica) / 0 (AV) |
| CPE | CPE (AV) / Valor de Nota (Gráfica) |
| Tier | Tier |
| Audiencia | Audiencia |
| Tono | Tono |
| Tema | Tematica |
| Temas Generales - Tema | Temas Generales - Tema |
| Resumen - Aclaracion | CuerpoEs |
| Link Nota | URL Nota AV con .ar→.co (AV) / URL (Streaming - Imagen) (Gráfica) |
| Link (Streaming - Imagen) | URL Nota |
| Menciones - Empresa | Menciones - Empresa (AV) / Empresa rel. (Gráfica) — expandido por ; |
""")

st.markdown("---")

# --- Uploader de dossiers (múltiples) ---
st.subheader("📂 Archivos a procesar")
uploaded_dossiers = st.file_uploader(
    "Sube uno o varios archivos Dossier (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"dossiers_{st.session_state['uploader_key']}"
)

if uploaded_dossiers:
    nombres = [f.name for f in uploaded_dossiers]
    st.info(f"{len(uploaded_dossiers)} archivo(s) cargado(s): {', '.join(nombres)}")

# --- Resultados previos persistentes estructurados en Pestañas Estándar ---
if st.session_state.get('resultados'):
    st.markdown("---")
    
    # Creamos las dos pestañas solicitadas en la interfaz de usuario
    tab_archivos, tab_consolidado = st.tabs(["📋 Detalle por Archivo", "📊 Conteo Total Consolidado"])
    
    with tab_archivos:
        col_title, col_clear = st.columns([6, 1])
        col_title.subheader("Detalle de Archivos")
        if col_clear.button("🗑️ Borrar consultas", type="secondary"):
            st.session_state['resultados'] = []
            st.session_state['uploader_key'] += 1
            st.rerun()

        for r in st.session_state['resultados']:
            with st.container():
                col_nombre, col_graf, col_av, col_total, col_check, col_dl = st.columns([3, 1, 1, 1, 1, 1])
                col_nombre.markdown(f"**{r['nombre']}**")
                col_graf.metric("🗞️ Gráficas", r['graficas'])
                col_av.metric("🎬 AV", r['av'])
                col_total.metric("Total", r['total'])
                
                descargar = col_check.checkbox("Descargar", value=False, key=f"chk_{r['nombre']}")
                if descargar:
                    col_dl.download_button(
                        label="📥 Descargar Excel",
                        data=r['excel'],
                        file_name=r['filename'],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{r['nombre']}"
                    )
                
                if r.get('matched_client'):
                    st.caption(f"🎯 Cliente detectado: **{r['matched_client']}**")
                else:
                    st.caption("⚠️ No se pudo auto-detectar el cliente en este archivo.")
                st.markdown("---")
                
    with tab_consolidado:
        st.subheader("📊 Conteo Total de Registros")
        st.markdown("Tabla combinada y consolidada de los archivos procesados en el orden estándar de salida:")
        
        # Reunimos los datos acumulados de cada archivo procesado
        listado_archivos = []
        for r in st.session_state['resultados']:
            listado_archivos.append({
                'filename': r['nombre'],
                'av_count': r['av'],
                'grafica_count': r['graficas']
            })
            
        df_consolidado = build_consolidated_conteo(listado_archivos)
        
        # Mostramos la tabla consolidada directamente
        st.dataframe(df_consolidado, use_container_width=True, hide_index=True)
        
        # Generar botón para descargar el Excel únicamente de la tabla consolidada de conteos
        excel_cons = to_excel_consolidated(df_consolidado)
        st.download_button(
            label="📥 Descargar Tabla Conteo Consolidada (Excel)",
            data=excel_cons,
            file_name=f"Conteo_Consolidado_SOV_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_total_consolidado"
        )

# --- Botón principal ---
can_run = bool(uploaded_dossiers and config_source)

if st.button("▶️ Iniciar Proceso", disabled=not can_run, type="primary"):
    try:
        region_map, internet_map = load_config(config_source)
    except Exception as e:
        st.error(f"Error cargando Configuracion.xlsx: {e}. Debe tener hojas 'Regiones' e 'Internet'.")
        st.stop()

    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio",
        "Sección - Programa", "Región", "Título", "Autor - Conductor",
        "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres",
        "CPE", "Tier", "Audiencia", "Tono", "Tema",
        "Temas Generales - Tema", "Resumen - Aclaracion",
        "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"
    ]

    resultados = []
    progress_bar = st.progress(0)

    for i, dossier_file in enumerate(uploaded_dossiers):
        with st.spinner(f"Procesando {dossier_file.name}..."):
            try:
                df, av_count, grafica_count = process_dossier(dossier_file, region_map, internet_map)
                excel_data = to_excel_from_df(df, final_order, dossier_file.name, av_count, grafica_count)
                df_conteo, matched = build_conteo_df(dossier_file.name, av_count, grafica_count)
                
                resultados.append({
                    'nombre': dossier_file.name,
                    'graficas': grafica_count,
                    'av': av_count,
                    'total': len(df),
                    'excel': excel_data,
                    'conteo_df': df_conteo,
                    'matched_client': matched,
                    'filename': f"SOV_{dossier_file.name.replace('.xlsx','')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                })
            except Exception as e:
                st.error(f"Error procesando {dossier_file.name}: {e}")
        progress_bar.progress((i + 1) / len(uploaded_dossiers))

    st.session_state['resultados'] = resultados
    st.balloons()
    st.rerun()
