import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import datetime
import io
import re
import html
import numpy as np
import json
from pathlib import Path

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA — debe ir primero
# ==============================================================================
st.set_page_config(
    page_title="SOV · Conteo v3",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==============================================================================
# CSS GLOBAL — diseño oscuro, tipografía monoespaciada para datos, acentos cyan
# ==============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* Reset base */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Fondo general */
.stApp {
    background: #0d0f14;
    color: #e2e8f0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #111318 !important;
    border-right: 1px solid #1e2130;
}

/* Header principal */
.sov-header {
    background: linear-gradient(135deg, #0d1117 0%, #131929 50%, #0d1117 100%);
    border: 1px solid #1e3a5f;
    border-radius: 16px;
    padding: 32px 40px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.sov-header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, #00d4ff, #0099cc, transparent);
}
.sov-header h1 {
    font-size: 2rem;
    font-weight: 700;
    color: #f0f6ff;
    margin: 0 0 6px 0;
    letter-spacing: -0.03em;
}
.sov-header p {
    color: #64748b;
    font-size: 0.9rem;
    margin: 0;
    font-weight: 400;
}
.sov-badge {
    display: inline-block;
    background: #00d4ff18;
    border: 1px solid #00d4ff40;
    color: #00d4ff;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 10px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

/* Cards de métricas */
.metric-card {
    background: #111318;
    border: 1px solid #1e2130;
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    transition: border-color 0.2s ease;
}
.metric-card:hover { border-color: #00d4ff40; }
.metric-card .metric-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 2rem;
    font-weight: 700;
    color: #00d4ff;
    display: block;
    line-height: 1;
}
.metric-card .metric-label {
    font-size: 0.75rem;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-top: 6px;
    display: block;
}

/* Archivo card */
.file-card {
    background: #111318;
    border: 1px solid #1e2130;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 10px;
    transition: all 0.2s ease;
}
.file-card:hover { border-color: #00d4ff30; }
.file-card .file-name {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    color: #94a3b8;
    margin-bottom: 8px;
}
.file-card .client-tag {
    background: #00d4ff12;
    border: 1px solid #00d4ff30;
    color: #00d4ff;
    font-size: 0.72rem;
    padding: 2px 10px;
    border-radius: 20px;
    font-weight: 500;
}
.file-card .warn-tag {
    background: #f59e0b12;
    border: 1px solid #f59e0b40;
    color: #f59e0b;
    font-size: 0.72rem;
    padding: 2px 10px;
    border-radius: 20px;
}

/* Tabla de conteo */
.conteo-table {
    background: #111318;
    border: 1px solid #1e2130;
    border-radius: 12px;
    overflow: hidden;
}

/* Panel de números */
.nums-panel {
    background: #080a0f;
    border: 1px solid #1e2130;
    border-radius: 12px;
    padding: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    color: #00d4ff;
    line-height: 2;
    white-space: pre;
    overflow-y: auto;
    max-height: 600px;
}

/* Sección de codificación manual */
.manual-client-block {
    background: #111318;
    border: 1px solid #1e2130;
    border-radius: 12px;
    padding: 18px 22px;
    margin-bottom: 12px;
    transition: border-color 0.2s ease;
}
.manual-client-block:hover { border-color: #1e3a5f; }
.manual-client-block h4 {
    color: #94a3b8;
    font-size: 0.82rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin: 0 0 14px 0;
}

/* Inputs de número — oscurecer el fondo */
div[data-testid="stNumberInput"] input {
    background: #0d0f14 !important;
    border-color: #1e2130 !important;
    color: #e2e8f0 !important;
    font-family: 'JetBrains Mono', monospace !important;
    border-radius: 8px !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #111318;
    border-radius: 10px;
    padding: 4px;
    border: 1px solid #1e2130;
    gap: 2px;
}
.stTabs [data-baseweb="tab"] {
    background: transparent;
    border-radius: 8px;
    color: #64748b;
    font-size: 0.82rem;
    font-weight: 500;
    padding: 8px 18px;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: #1e2a3a !important;
    color: #00d4ff !important;
}

/* Botones primarios */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0066cc, #0099ff) !important;
    border: none !important;
    border-radius: 10px !important;
    color: white !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    padding: 10px 24px !important;
    letter-spacing: 0.02em;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 20px #0099ff40 !important;
}
.stButton > button[kind="secondary"] {
    background: #111318 !important;
    border: 1px solid #1e2130 !important;
    border-radius: 10px !important;
    color: #94a3b8 !important;
    font-weight: 500 !important;
    font-size: 0.82rem !important;
}

/* Download buttons */
.stDownloadButton > button {
    background: #0d1117 !important;
    border: 1px solid #00d4ff30 !important;
    border-radius: 10px !important;
    color: #00d4ff !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    width: 100%;
}
.stDownloadButton > button:hover {
    border-color: #00d4ff80 !important;
    background: #00d4ff08 !important;
}

/* Alertas */
.stSuccess { background: #0d2a1a !important; border-color: #00c853 !important; }
.stWarning { background: #221a00 !important; border-color: #f59e0b !important; }
.stError   { background: #2a0d0d !important; border-color: #ef4444 !important; }

/* Expander */
.streamlit-expanderHeader {
    background: #111318 !important;
    border-radius: 10px !important;
    border: 1px solid #1e2130 !important;
    color: #94a3b8 !important;
    font-size: 0.82rem !important;
}

/* DataFrames */
.stDataFrame { border-radius: 10px; overflow: hidden; }
div[data-testid="stDataFrame"] > div {
    background: #111318;
    border: 1px solid #1e2130;
    border-radius: 10px;
}

/* Separador */
hr { border-color: #1e2130 !important; margin: 24px 0 !important; }

/* Info box custom */
.info-box {
    background: #0d1929;
    border: 1px solid #1e3a5f;
    border-radius: 10px;
    padding: 14px 18px;
    font-size: 0.82rem;
    color: #7cb8d6;
    margin-bottom: 16px;
}

/* Total badge en panel numérico */
.num-row {
    display: flex;
    justify-content: space-between;
    padding: 2px 0;
    border-bottom: 1px solid #1e213008;
}
.num-highlight { color: #00d4ff; font-weight: 700; }
.num-zero { color: #2a3040; }

/* Badge de estado en pestaña de archivos */
.status-ok   { color: #22c55e; font-size: 0.75rem; }
.status-warn { color: #f59e0b; font-size: 0.75rem; }

/* File uploader */
div[data-testid="stFileUploader"] {
    background: #111318 !important;
    border: 1px dashed #1e2130 !important;
    border-radius: 12px !important;
    padding: 8px !important;
}

/* Selectbox, text_area */
textarea {
    background: #0d0f14 !important;
    border-color: #1e2130 !important;
    color: #00d4ff !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.85rem !important;
    line-height: 2 !important;
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONSTANTES Y DATOS ESTRUCTURALES
# ==============================================================================

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
    ("Universidad Tecnológica de Bolívar", "Notas Impresos", "UTB_AN"),
]

UNIQUE_CLIENTS = list(dict.fromkeys(c for c, _, _ in CONTEO_TEMPLATE))

REPLICATED_CLIENTS = {
    "Chery 01-18 | |19-31",
    "Chery - Changan, Competencias",
    "Nissan",
    "Nissan, Competencia",
}

_CODE_TO_CLIENT_SIMPLE = {
    "acomfevalle": "Comfenalco Valle",
    "anfenavi": "Federación Nacional de Avicultores de Colombia",
    "fsantafe_an": "Fundación Santa Fe de Bogotá",
    "tigoan": "Tigo",
    "usimonan": "Universidad Simón Bolívar",
    "utb_an": "Universidad Tecnológica de Bolívar",
}

_LEGACY_KEYWORDS = {
    "Comfenalco Valle": ["comfe", "comfenalco"],
    "Federación Nacional de Avicultores de Colombia": ["fenavi", "avicultores", "avicola"],
    "Fundación Santa Fe de Bogotá": ["fsant", "santa", "santafe"],
    "Tigo": ["tigo"],
    "Universidad Simón Bolívar": ["simon", "usimon", "usim"],
    "Universidad Tecnológica de Bolívar": ["utb", "tecnologica"],
}

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

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
        try: return chr(int(match.group(1), 16))
        except: return match.group(0)

    def replace_decimal_entity(match):
        try: return chr(int(match.group(1)))
        except: return match.group(0)

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
    fn = Path(filename).stem.lower()
    fn_clean = re.sub(r'^[\d\s\-_]+', '', fn).strip()
    tokens = [t for t in re.split(r'[^a-z0-9]', fn_clean) if t]
    comp_keywords = {"c", "com", "comp", "competencia", "competencias", "changan"}
    is_competencia = any(t in comp_keywords for t in tokens)

    if "anchery" in fn_clean:
        return "Chery - Changan, Competencias" if is_competencia else "Chery 01-18 | |19-31"
    if "annissan" in fn_clean:
        return "Nissan, Competencia" if is_competencia else "Nissan"
    for code, client in _CODE_TO_CLIENT_SIMPLE.items():
        if code in fn_clean:
            return client

    if any(kw in t for t in tokens for kw in ["chery"]):
        return "Chery - Changan, Competencias" if is_competencia else "Chery 01-18 | |19-31"
    if any(kw in t for t in tokens for kw in ["niss", "nissan"]):
        return "Nissan, Competencia" if is_competencia else "Nissan"
    for client, kws in _LEGACY_KEYWORDS.items():
        if any(kw in t for t in tokens for kw in kws):
            return client
    return None


def build_consolidated_conteo(processed_files, manual_codif=None):
    """Genera tabla consolidada acumulando todos los archivos + codificación manual."""
    manual_codif = manual_codif or {}
    client_av = {}
    client_graf = {}

    for item in processed_files:
        matched_client = get_client_category(item['filename'])
        if matched_client:
            client_av[matched_client] = client_av.get(matched_client, 0) + item['av_count']
            client_graf[matched_client] = client_graf.get(matched_client, 0) + item['grafica_count']

    rows = []
    for client, tipo, codigo in CONTEO_TEMPLATE:
        av_count  = client_av.get(client, 0)
        graf_count = client_graf.get(client, 0)
        manual_av      = int(manual_codif.get(client, {}).get('av', 0) or 0)
        manual_impresos = int(manual_codif.get(client, {}).get('impresos', 0) or 0)
        is_rep = client in REPLICATED_CLIENTS

        if tipo == "Notas Audiovisuales":
            val = av_count
        elif tipo == "Notas Impresos":
            val = graf_count
        elif tipo == "Codificación Audiovisuales":
            val = (av_count if is_rep else 0) + manual_av
        elif tipo == "Codificación Impresos":
            val = (graf_count if is_rep else 0) + manual_impresos
        else:
            val = 0

        rows.append({
            "Cliente / Categoría": client,
            "Tipo de Conteo": tipo,
            "Código": codigo,
            "Cantidad": val,
        })

    return pd.DataFrame(rows)


def to_excel_consolidated(df_conteo):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Conteo Consolidado"

    header_fill = PatternFill("solid", fgColor="0D1929")
    header_font = Font(bold=True, color="00D4FF", name="Calibri")
    thin = Side(style='thin', color="1E2130")
    border = Border(bottom=Side(style='thin', color="1E2130"))

    for col_idx, col_name in enumerate(df_conteo.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    alt_fill = PatternFill("solid", fgColor="111318")
    for row_idx, row_data in enumerate(df_conteo.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if row_idx % 2 == 0:
                cell.fill = alt_fill
            cell.border = border

    for col in ws.columns:
        max_len = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

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
    ws = wb.active
    ws.title = 'Resultado'

    header_fill = PatternFill("solid", fgColor="0D1929")
    header_font = Font(bold=True, color="00D4FF", name="Calibri")

    for i, col_name in enumerate(df_out.columns, start=1):
        cell = ws.cell(row=1, column=i, value=col_name)
        cell.font = header_font
        cell.fill = header_fill

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

    # Hoja Conteo dentro del archivo individual
    ws2 = wb.create_sheet(title='Conteo')
    # Conteo individual: solo este archivo, sin codificación manual
    listado = [{'filename': filename, 'av_count': av_count, 'grafica_count': grafica_count}]
    df_conteo = build_consolidated_conteo(listado)

    for col_idx, col_name in enumerate(df_conteo.columns, start=1):
        cell = ws2.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="00D4FF")
        cell.fill = PatternFill("solid", fgColor="0D1929")

    for row_idx, row_data in enumerate(df_conteo.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_data, start=1):
            ws2.cell(row=row_idx, column=col_idx, value=val)

    for col in ws2.columns:
        max_len = max(len(str(c.value or '')) for c in col)
        ws2.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

    wb.save(output)
    output.seek(0)
    return output.getvalue()


def load_config(config_source):
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


def process_dossier(dossier_file, region_map, internet_map):
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

    df.loc[is_internet, 'Medio'] = (
        df.loc[is_internet, 'Medio']
        .astype(str).str.lower().str.strip()
        .map(internet_map)
        .fillna(df.loc[is_internet, 'Medio'])
    )

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
# SESSION STATE — inicialización
# ==============================================================================

def _init_state():
    defaults = {
        'uploader_key': 0,
        'resultados': [],
        # Codificación manual: persiste entre reruns de la sesión
        'manual_codif': {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS},
        # Flag para saber si el usuario guardó codificación manual
        'manual_saved': False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ==============================================================================
# HEADER
# ==============================================================================

st.markdown("""
<div class="sov-header">
    <div class="sov-badge">📡 Media Intelligence · SOV</div>
    <h1>Conteo v3</h1>
    <p>Procesa dossiers de monitoreo, calcula conteos y gestiona codificación manual por cliente.</p>
</div>
""", unsafe_allow_html=True)


# ==============================================================================
# SIDEBAR — Configuración y acciones globales
# ==============================================================================

with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    st.markdown("---")

    config_source = None

    if CONFIG_PATH is not None:
        st.success(f"✅ Config: `{CONFIG_PATH.name}`")
        config_source = CONFIG_PATH
    else:
        st.warning("No se encontró `Configuracion.xlsx` en el repo.")
        config_upload = st.file_uploader("Subir Configuracion.xlsx", type=["xlsx"], key="config_upload")
        if config_upload:
            config_source = config_upload
            st.success(f"✅ {config_upload.name}")

    st.markdown("---")
    st.markdown("### 💾 Codificación manual")
    st.caption("Guarda/restaura valores entre sesiones exportando el JSON.")

    # Exportar JSON
    json_bytes = json.dumps(st.session_state['manual_codif'], ensure_ascii=False, indent=2).encode('utf-8')
    st.download_button(
        "⬇️ Exportar JSON",
        data=json_bytes,
        file_name=f"codif_manual_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        key="dl_manual_sidebar",
        use_container_width=True,
    )

    # Importar JSON
    archivo_json = st.file_uploader("⬆️ Importar JSON", type=["json"], key="upload_manual_sidebar")
    if archivo_json is not None:
        try:
            data_cargada = json.load(archivo_json)
            valores_validados = {
                c: data_cargada.get(c, {'av': 0, 'impresos': 0})
                for c in UNIQUE_CLIENTS
            }
            st.session_state['manual_codif'] = valores_validados
            st.session_state['manual_saved'] = True
            st.success("Codificación cargada.")
            st.rerun()
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

    st.markdown("---")
    st.markdown("### 📋 Convención de nombres")
    st.caption(
        "`<fecha> <CÓDIGO> [m|com]`\n\n"
        "Ej: `19 ANCHERY m`, `19 ANNISSAN com`, `19 FSANTAFE_AN`\n\n"
        "Códigos válidos: `ANCHERY`, `ANNISSAN`, `ACOMFEVALLE`, `ANFENAVI`, "
        "`FSANTAFE_AN`, `TIGOAN`, `USIMONAN`, `UTB_AN`"
    )

    st.markdown("---")
    if st.button("🗑️ Limpiar resultados", type="secondary", use_container_width=True):
        st.session_state['resultados'] = []
        st.session_state['uploader_key'] += 1
        st.rerun()


# ==============================================================================
# CUERPO PRINCIPAL
# ==============================================================================

# --- Zona de carga de archivos ---
st.markdown("#### 📂 Dossiers a procesar")
uploaded_dossiers = st.file_uploader(
    "Arrastrá uno o varios archivos `.xlsx`",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"dossiers_{st.session_state['uploader_key']}",
    label_visibility="collapsed",
)

if uploaded_dossiers:
    st.markdown(
        f'<div class="info-box">📎 {len(uploaded_dossiers)} archivo(s) listos: '
        + ', '.join(f"`{f.name}`" for f in uploaded_dossiers)
        + '</div>',
        unsafe_allow_html=True,
    )

can_run = bool(uploaded_dossiers and config_source)

col_btn, col_hint = st.columns([2, 5])
with col_btn:
    run_clicked = st.button(
        "▶ Procesar",
        disabled=not can_run,
        type="primary",
        use_container_width=True,
    )
with col_hint:
    if not config_source:
        st.caption("⚠️ Falta subir `Configuracion.xlsx` en el sidebar.")
    elif not uploaded_dossiers:
        st.caption("⬆️ Subí al menos un dossier para continuar.")

# --- Procesamiento ---
if run_clicked:
    try:
        region_map, internet_map = load_config(config_source)
    except Exception as e:
        st.error(f"Error cargando Configuracion.xlsx: {e}")
        st.stop()

    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio",
        "Sección - Programa", "Región", "Título", "Autor - Conductor",
        "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres",
        "CPE", "Tier", "Audiencia", "Tono", "Tema",
        "Temas Generales - Tema", "Resumen - Aclaracion",
        "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa",
    ]

    nuevos = []
    progress_bar = st.progress(0, text="Iniciando proceso…")

    for i, dossier_file in enumerate(uploaded_dossiers):
        progress_bar.progress((i) / len(uploaded_dossiers), text=f"Procesando {dossier_file.name}…")
        try:
            df, av_count, grafica_count = process_dossier(dossier_file, region_map, internet_map)
            excel_data = to_excel_from_df(df, final_order, dossier_file.name, av_count, grafica_count)
            matched = get_client_category(dossier_file.name)
            nuevos.append({
                'nombre': dossier_file.name,
                'graficas': grafica_count,
                'av': av_count,
                'total': len(df),
                'excel': excel_data,
                'matched_client': matched,
                'filename': f"SOV_{dossier_file.name.replace('.xlsx','')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            })
        except Exception as e:
            st.error(f"❌ Error en {dossier_file.name}: {e}")

    progress_bar.progress(1.0, text="✅ Proceso completado")
    # Acumulamos: no borramos resultados previos, los extendemos
    st.session_state['resultados'].extend(nuevos)
    st.session_state['uploader_key'] += 1
    st.balloons()
    st.rerun()


# ==============================================================================
# RESULTADOS (sólo si hay datos)
# ==============================================================================

if st.session_state.get('resultados'):
    resultados = st.session_state['resultados']

    # KPIs globales rápidos
    total_av   = sum(r['av'] for r in resultados)
    total_graf = sum(r['graficas'] for r in resultados)
    total_notas = sum(r['total'] for r in resultados)
    total_archivos = len(resultados)

    st.markdown("---")
    c1, c2, c3, c4 = st.columns(4)
    for col, label, val in [
        (c1, "Archivos", total_archivos),
        (c2, "Total registros", total_notas),
        (c3, "🎬 Audiovisuales", total_av),
        (c4, "🗞️ Gráficas", total_graf),
    ]:
        col.markdown(
            f'<div class="metric-card">'
            f'<span class="metric-value">{val}</span>'
            f'<span class="metric-label">{label}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("")

    # =========================================================================
    # TABS
    # =========================================================================
    tab_archivos, tab_consolidado, tab_manual = st.tabs([
        "📋  Archivos",
        "📊  Conteo consolidado",
        "✏️  Codificación manual",
    ])

    # -------------------------------------------------------------------------
    # TAB 1 — Detalle por archivo
    # -------------------------------------------------------------------------
    with tab_archivos:
        for r in resultados:
            with st.container():
                st.markdown(
                    f'<div class="file-card">'
                    f'<div class="file-name">📄 {r["nombre"]}</div>'
                    + (
                        f'<span class="client-tag">🎯 {r["matched_client"]}</span>'
                        if r.get("matched_client")
                        else '<span class="warn-tag">⚠️ Sin cliente detectado</span>'
                    )
                    + '</div>',
                    unsafe_allow_html=True,
                )

                col_g, col_a, col_t, col_dl = st.columns([1, 1, 1, 2])
                col_g.metric("Gráficas", r['graficas'])
                col_a.metric("Audiovisuales", r['av'])
                col_t.metric("Total", r['total'])
                col_dl.download_button(
                    "📥 Descargar Excel",
                    data=r['excel'],
                    file_name=r['filename'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_file_{r['nombre']}_{r['filename']}",
                    use_container_width=True,
                )
                st.markdown("")

    # -------------------------------------------------------------------------
    # TAB 2 — Conteo consolidado
    # -------------------------------------------------------------------------
    with tab_consolidado:
        # Recalcular en cada render para reflejar cambios en manual_codif
        listado_archivos = [
            {'filename': r['nombre'], 'av_count': r['av'], 'grafica_count': r['graficas']}
            for r in resultados
        ]
        df_consolidado = build_consolidated_conteo(listado_archivos, st.session_state['manual_codif'])

        col_table, col_numeric = st.columns([3, 1])

        with col_table:
            st.markdown("##### Tabla de conteo")
            st.dataframe(
                df_consolidado.style.apply(
                    lambda col: [
                        'color: #00d4ff; font-weight: 700' if v > 0 else 'color: #2a3040'
                        for v in col
                    ] if col.name == 'Cantidad' else ['' for _ in col],
                    axis=0,
                ),
                use_container_width=True,
                hide_index=True,
                height=600,
            )

            excel_cons = to_excel_consolidated(df_consolidado)
            st.download_button(
                "📥 Descargar Excel consolidado",
                data=excel_cons,
                file_name=f"Conteo_SOV_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_consolidado",
                use_container_width=True,
            )

        with col_numeric:
            st.markdown("##### 📋 Copiar columna")
            st.caption("Seleccioná y copiá directamente en tu plantilla Excel.")

            numeros = df_consolidado["Cantidad"].astype(str).tolist()
            st.text_area(
                "Valores (40 filas, orden estándar):",
                value="\n".join(numeros),
                height=560,
                key="raw_numeric_area",
            )

    # -------------------------------------------------------------------------
    # TAB 3 — Codificación manual
    # -------------------------------------------------------------------------
    with tab_manual:
        st.markdown(
            '<div class="info-box">'
            '✏️ Ingresá aquí las cantidades de <strong>Codificación Audiovisuales</strong> y '
            '<strong>Codificación Impresos</strong> para notas analizadas fuera de los dossiers. '
            'Estos valores <strong>se suman</strong> al conteo automático y se reflejan '
            'inmediatamente en la pestaña <em>Conteo consolidado</em>.'
            '</div>',
            unsafe_allow_html=True,
        )

        # Tip sobre persistencia
        st.info(
            "💡 **Persistencia entre sesiones:** los valores aquí ingresados duran mientras "
            "no cerrés la pestaña del navegador. Para conservarlos entre sesiones, exportá el JSON "
            "desde el sidebar y volvé a importarlo la próxima vez."
        )

        # Formulario de codificación manual
        with st.form("form_codif_manual", clear_on_submit=False):
            nuevos_valores = {}
            for client in UNIQUE_CLIENTS:
                st.markdown(
                    f'<div class="manual-client-block"><h4>{client}</h4></div>',
                    unsafe_allow_html=True,
                )
                # Necesitamos poner los inputs fuera del div para que Streamlit los renderice
                col_av, col_im = st.columns(2)
                actual = st.session_state['manual_codif'].get(client, {'av': 0, 'impresos': 0})
                val_av = col_av.number_input(
                    "🎬 Codif. Audiovisuales",
                    min_value=0, step=1,
                    value=int(actual.get('av', 0) or 0),
                    key=f"mav_{client}",
                )
                val_im = col_im.number_input(
                    "🗞️ Codif. Impresos",
                    min_value=0, step=1,
                    value=int(actual.get('impresos', 0) or 0),
                    key=f"mim_{client}",
                )
                nuevos_valores[client] = {'av': val_av, 'impresos': val_im}
                st.markdown("")

            col_save, col_reset = st.columns([2, 1])
            guardar = col_save.form_submit_button("💾 Guardar valores", type="primary", use_container_width=True)
            reset   = col_reset.form_submit_button("🔄 Resetear todo", use_container_width=True)

            if guardar:
                st.session_state['manual_codif'] = nuevos_valores
                st.session_state['manual_saved'] = True
                st.success("✅ Valores guardados. El conteo consolidado ya refleja estos cambios.")
                st.rerun()

            if reset:
                st.session_state['manual_codif'] = {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS}
                st.session_state['manual_saved'] = False
                st.rerun()

        # Vista previa rápida de lo que se va a sumar
        if st.session_state.get('manual_saved'):
            totales_manual = {
                'av': sum(v['av'] for v in st.session_state['manual_codif'].values()),
                'impresos': sum(v['impresos'] for v in st.session_state['manual_codif'].values()),
            }
            if totales_manual['av'] > 0 or totales_manual['impresos'] > 0:
                st.markdown("---")
                st.markdown("**Totales ingresados manualmente:**")
                ma, mi = st.columns(2)
                ma.metric("🎬 AV manual total", totales_manual['av'])
                mi.metric("🗞️ Impresos manual total", totales_manual['impresos'])

else:
    # Estado vacío — instrucciones
    st.markdown("---")
    st.markdown("""
<div style="text-align:center; padding: 60px 20px; color: #2a3040;">
    <div style="font-size: 3rem; margin-bottom: 16px;">📡</div>
    <div style="font-size: 1rem; color: #3a4555; font-weight: 500;">
        Subí los dossiers y hacé clic en <strong style="color:#00d4ff;">Procesar</strong> para comenzar.
    </div>
    <div style="font-size: 0.8rem; color: #2a3040; margin-top: 10px;">
        Podés gestionar la codificación manual en cualquier momento desde el sidebar.
    </div>
</div>
""", unsafe_allow_html=True)

    # Acceso rápido a codificación manual incluso sin archivos
    with st.expander("✏️ Ingresar codificación manual (sin procesar dossiers)", expanded=False):
        st.markdown(
            '<div class="info-box">'
            'Podés ingresar y guardar tu codificación manual ahora. '
            'Los valores se sumarán al conteo cuando proceses archivos.'
            '</div>',
            unsafe_allow_html=True,
        )
        with st.form("form_codif_manual_empty", clear_on_submit=False):
            nuevos_valores = {}
            for client in UNIQUE_CLIENTS:
                st.markdown(
                    f'<div class="manual-client-block"><h4>{client}</h4></div>',
                    unsafe_allow_html=True,
                )
                col_av, col_im = st.columns(2)
                actual = st.session_state['manual_codif'].get(client, {'av': 0, 'impresos': 0})
                val_av = col_av.number_input(
                    "🎬 Codif. Audiovisuales",
                    min_value=0, step=1,
                    value=int(actual.get('av', 0) or 0),
                    key=f"empty_mav_{client}",
                )
                val_im = col_im.number_input(
                    "🗞️ Codif. Impresos",
                    min_value=0, step=1,
                    value=int(actual.get('impresos', 0) or 0),
                    key=f"empty_mim_{client}",
                )
                nuevos_valores[client] = {'av': val_av, 'impresos': val_im}
                st.markdown("")

            col_s, col_r = st.columns([2, 1])
            if col_s.form_submit_button("💾 Guardar", type="primary", use_container_width=True):
                st.session_state['manual_codif'] = nuevos_valores
                st.session_state['manual_saved'] = True
                st.success("✅ Guardado. Subí dossiers cuando estés listo.")
                st.rerun()
            if col_r.form_submit_button("🔄 Resetear", use_container_width=True):
                st.session_state['manual_codif'] = {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS}
                st.session_state['manual_saved'] = False
                st.rerun()
