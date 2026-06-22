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
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="SOV · Conteo v3",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# CSS — tema claro
# ==============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f8fafc; color: #1e293b; }

section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
    min-width: 320px !important;
}

/* Header */
.sov-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #1a4a7a 60%, #1e3a5f 100%);
    border-radius: 14px;
    padding: 26px 34px;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
}
.sov-header::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, #60a5fa, #3b82f6, transparent);
}
.sov-header h1 { font-size: 1.7rem; font-weight: 700; color: #f0f9ff; margin: 0 0 4px 0; letter-spacing: -0.02em; }
.sov-header p  { color: #93c5fd; font-size: 0.86rem; margin: 0; }
.sov-badge {
    display: inline-block;
    background: #ffffff18;
    border: 1px solid #ffffff30;
    color: #bfdbfe;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.67rem;
    padding: 2px 10px;
    border-radius: 20px;
    margin-bottom: 8px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
}

/* Paso numerado */
.step-block {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 18px 22px;
    margin-bottom: 14px;
    box-shadow: 0 1px 3px #1e293b08;
}
.step-block .step-label {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
}
.step-num {
    background: #1d4ed8;
    color: #ffffff;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    font-weight: 700;
    width: 24px; height: 24px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.step-num.done { background: #16a34a; }
.step-num.warn { background: #d97706; }
.step-title { font-weight: 600; font-size: 0.88rem; color: #1e293b; }
.step-desc  { font-size: 0.78rem; color: #64748b; margin: 0; }

/* Tarjetas de métricas */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 16px;
    text-align: center;
    box-shadow: 0 1px 3px #1e293b08;
}
.metric-card .mv { font-family: 'JetBrains Mono', monospace; font-size: 1.8rem; font-weight: 700; color: #1d4ed8; display: block; line-height: 1; }
.metric-card .ml { font-size: 0.7rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.08em; margin-top: 5px; display: block; }

/* Tarjeta de archivo */
.file-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 8px;
    box-shadow: 0 1px 2px #1e293b06;
}
.file-card .fn { font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #475569; margin-bottom: 5px; }
.client-tag { background: #eff6ff; border: 1px solid #bfdbfe; color: #1d4ed8; font-size: 0.7rem; padding: 2px 10px; border-radius: 20px; font-weight: 500; }
.warn-tag   { background: #fffbeb; border: 1px solid #fde68a; color: #92400e; font-size: 0.7rem; padding: 2px 10px; border-radius: 20px; }

/* Bloque de cliente en codificación manual */
.mc-block {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-left: 3px solid #3b82f6;
    border-radius: 10px;
    padding: 12px 16px 4px 16px;
    margin-bottom: 8px;
}
.mc-block h4 { color: #1e40af; font-size: 0.78rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; margin: 0 0 8px 0; }

/* Info box */
.info-box { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 10px; padding: 11px 15px; font-size: 0.81rem; color: #1e40af; margin-bottom: 12px; }
.warn-box { background: #fffbeb; border: 1px solid #fde68a; border-radius: 10px; padding: 11px 15px; font-size: 0.81rem; color: #92400e; margin-bottom: 12px; }
.ok-box   { background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px; padding: 11px 15px; font-size: 0.81rem; color: #166534; margin-bottom: 12px; }

/* JSON panel en sidebar */
.json-panel {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 10px;
}
.json-panel h4 { font-size: 0.78rem; font-weight: 600; color: #374151; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.05em; }
.json-status { font-size: 0.74rem; padding: 4px 10px; border-radius: 6px; display: inline-block; margin-bottom: 8px; }
.json-status.loaded { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
.json-status.empty  { background: #f8fafc; color: #94a3b8; border: 1px solid #e2e8f0; }

/* Guía de orden */
.order-guide { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; }
.og-row { display: flex; align-items: center; padding: 6px 12px; border-bottom: 1px solid #f1f5f9; gap: 8px; }
.og-row:last-child { border-bottom: none; }
.og-row:nth-child(even) { background: #f8fafc; }
.og-row:nth-child(odd)  { background: #ffffff; }
.og-row.hl { background: #eff6ff !important; }
.og-row.hdr { background: #1e3a5f !important; }
.og-num  { color: #94a3b8; width: 20px; text-align: right; flex-shrink: 0; font-size: 0.68rem; }
.og-tipo { color: #475569; width: 200px; flex-shrink: 0; font-size: 0.73rem; }
.og-cli  { color: #1e293b; flex: 1; font-size: 0.72rem; }
.og-cod  { color: #94a3b8; font-size: 0.67rem; }
.hdr .og-num, .hdr .og-tipo, .hdr .og-cli, .hdr .og-cod { color: #e0f2fe !important; font-weight: 600; }
.hl .og-tipo { color: #1d4ed8 !important; font-weight: 600; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: #f1f5f9; border-radius: 10px; padding: 3px; border: 1px solid #e2e8f0; gap: 2px; }
.stTabs [data-baseweb="tab"] { background: transparent; border-radius: 8px; color: #64748b; font-size: 0.81rem; font-weight: 500; padding: 7px 15px; border: none !important; }
.stTabs [aria-selected="true"] { background: #ffffff !important; color: #1d4ed8 !important; box-shadow: 0 1px 4px #1e293b14; }

/* Botones */
.stButton > button[kind="primary"] { background: #1d4ed8 !important; border: none !important; border-radius: 8px !important; color: white !important; font-weight: 600 !important; font-size: 0.83rem !important; padding: 9px 20px !important; }
.stButton > button[kind="primary"]:hover { background: #1e40af !important; }
.stButton > button[kind="secondary"] { background: #ffffff !important; border: 1px solid #e2e8f0 !important; border-radius: 8px !important; color: #475569 !important; font-weight: 500 !important; font-size: 0.81rem !important; }
.stDownloadButton > button { background: #ffffff !important; border: 1px solid #bfdbfe !important; border-radius: 8px !important; color: #1d4ed8 !important; font-size: 0.79rem !important; font-weight: 500 !important; width: 100%; }
.stDownloadButton > button:hover { background: #eff6ff !important; border-color: #3b82f6 !important; }

/* Inputs */
div[data-testid="stNumberInput"] input { background: #f8fafc !important; border-color: #e2e8f0 !important; color: #1e293b !important; font-family: 'JetBrains Mono', monospace !important; border-radius: 8px !important; }
textarea { background: #f8fafc !important; border-color: #e2e8f0 !important; color: #1d4ed8 !important; font-family: 'JetBrains Mono', monospace !important; font-size: 0.84rem !important; line-height: 2 !important; }
div[data-testid="stFileUploader"] { background: #ffffff !important; border: 1.5px dashed #bfdbfe !important; border-radius: 10px !important; }

hr { border-color: #e2e8f0 !important; margin: 18px 0 !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONSTANTES
# ==============================================================================
CONTEO_TEMPLATE = [
    ("Chery 01-18 | |19-31",                          "Codificación Audiovisuales", "ANCHERY"),
    ("Chery 01-18 | |19-31",                          "Codificación Impresos",      "ANCHERY"),
    ("Chery 01-18 | |19-31",                          "Notas Audiovisuales",        "ANCHERY"),
    ("Chery 01-18 | |19-31",                          "Notas Impresos",             "ANCHERY"),
    ("Chery - Changan, Competencias",                 "Codificación Audiovisuales", "ANCHERY"),
    ("Chery - Changan, Competencias",                 "Codificación Impresos",      "ANCHERY"),
    ("Chery - Changan, Competencias",                 "Notas Audiovisuales",        "ANCHERY"),
    ("Chery - Changan, Competencias",                 "Notas Impresos",             "ANCHERY"),
    ("Comfenalco Valle",                              "Codificación Audiovisuales", "ACOMFEVALLE"),
    ("Comfenalco Valle",                              "Codificación Impresos",      "ACOMFEVALLE"),
    ("Comfenalco Valle",                              "Notas Audiovisuales",        "ACOMFEVALLE"),
    ("Comfenalco Valle",                              "Notas Impresos",             "ACOMFEVALLE"),
    ("Federación Nacional de Avicultores de Colombia","Codificación Audiovisuales", "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia","Codificación Impresos",      "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia","Notas Audiovisuales",        "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia","Notas Impresos",             "ANFENAVI"),
    ("Fundación Santa Fe de Bogotá",                  "Codificación Audiovisuales", "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá",                  "Codificación Impresos",      "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá",                  "Notas Audiovisuales",        "FSANTAFE_AN"),
    ("Fundación Santa Fe de Bogotá",                  "Notas Impresos",             "FSANTAFE_AN"),
    ("Nissan",                                        "Codificación Audiovisuales", "ANNISSAN"),
    ("Nissan",                                        "Codificación Impresos",      "ANNISSAN"),
    ("Nissan",                                        "Notas Audiovisuales",        "ANNISSAN"),
    ("Nissan",                                        "Notas Impresos",             "ANNISSAN"),
    ("Nissan, Competencia",                           "Codificación Audiovisuales", "ANNISSAN"),
    ("Nissan, Competencia",                           "Codificación Impresos",      "ANNISSAN"),
    ("Nissan, Competencia",                           "Notas Audiovisuales",        "ANNISSAN"),
    ("Nissan, Competencia",                           "Notas Impresos",             "ANNISSAN"),
    ("Tigo",                                          "Codificación Audiovisuales", "TIGOAN"),
    ("Tigo",                                          "Codificación Impresos",      "TIGOAN"),
    ("Tigo",                                          "Notas Audiovisuales",        "TIGOAN"),
    ("Tigo",                                          "Notas Impresos",             "TIGOAN"),
    ("Universidad Simón Bolívar",                     "Codificación Audiovisuales", "USIMONAN"),
    ("Universidad Simón Bolívar",                     "Codificación Impresos",      "USIMONAN"),
    ("Universidad Simón Bolívar",                     "Notas Audiovisuales",        "USIMONAN"),
    ("Universidad Simón Bolívar",                     "Notas Impresos",             "USIMONAN"),
    ("Universidad Tecnológica de Bolívar",            "Codificación Audiovisuales", "UTB_AN"),
    ("Universidad Tecnológica de Bolívar",            "Codificación Impresos",      "UTB_AN"),
    ("Universidad Tecnológica de Bolívar",            "Notas Audiovisuales",        "UTB_AN"),
    ("Universidad Tecnológica de Bolívar",            "Notas Impresos",             "UTB_AN"),
]

UNIQUE_CLIENTS   = list(dict.fromkeys(c for c, _, _ in CONTEO_TEMPLATE))
REPLICATED_CLIENTS = {"Chery 01-18 | |19-31", "Chery - Changan, Competencias", "Nissan", "Nissan, Competencia"}

_CODE_MAP = {
    "acomfevalle": "Comfenalco Valle",
    "anfenavi":    "Federación Nacional de Avicultores de Colombia",
    "fsantafe_an": "Fundación Santa Fe de Bogotá",
    "tigoan":      "Tigo",
    "usimonan":    "Universidad Simón Bolívar",
    "utb_an":      "Universidad Tecnológica de Bolívar",
}
_LEGACY = {
    "Comfenalco Valle":                               ["comfe","comfenalco"],
    "Federación Nacional de Avicultores de Colombia": ["fenavi","avicultores","avicola"],
    "Fundación Santa Fe de Bogotá":                   ["fsant","santa","santafe"],
    "Tigo":                                           ["tigo"],
    "Universidad Simón Bolívar":                      ["simon","usimon","usim"],
    "Universidad Tecnológica de Bolívar":             ["utb","tecnologica"],
}

# ==============================================================================
# FUNCIONES
# ==============================================================================

def find_config_path():
    base = Path(__file__).parent
    for f in base.iterdir():
        if f.suffix.lower() == '.xlsx' and 'config' in f.stem.lower():
            return f
    return None

CONFIG_PATH = find_config_path()


def extract_link_from_cell(cell):
    return cell.hyperlink.target if cell.hyperlink and cell.hyperlink.target else None


def convert_html_entities(text):
    if not isinstance(text, str): return text
    text = html.unescape(text)
    for entity, char in {'&#xF3;':'ó','&#xE1;':'á','&#xE9;':'é','&#xED;':'í','&#xFA;':'ú',
                          '&#xF1;':'ñ','&#xDC;':'Ü','&#xFC;':'ü','&#xC1;':'Á','&#xC9;':'É',
                          '&#xCD;':'Í','&#xD3;':'Ó','&#xDA;':'Ú','&#xD1;':'Ñ','&#xC7;':'Ç','&#xE7;':'ç'}.items():
        text = text.replace(entity, char)
    text = re.sub(r'&#x([0-9A-Fa-f]+);', lambda m: chr(int(m.group(1),16)), text)
    text = re.sub(r'&#(\d+);',           lambda m: chr(int(m.group(1))),    text)
    for b,g in {'\u201c':'"','\u201d':'"','\u2018':"'",'\u2019':"'",'Â':'','â':'','€':'','™':''}.items():
        text = text.replace(b, g)
    return text

def clean_text(t):
    return convert_html_entities(t).strip() if isinstance(t, str) else t

def clean_cuerpo(t):
    if not isinstance(t, str) or not t.strip(): return t
    t = convert_html_entities(t)
    t = re.sub(r'<br\s*/?>', '\n', t, flags=re.IGNORECASE)
    t = re.sub(r'<[^>]+>', '', t)
    return t.strip()

def get_client_category(filename):
    fn = re.sub(r'^[\d\s\-_]+', '', Path(filename).stem.lower()).strip()
    tokens = [t for t in re.split(r'[^a-z0-9]', fn) if t]
    comp = any(t in {"c","com","comp","competencia","competencias","changan"} for t in tokens)
    if "anchery"  in fn: return "Chery - Changan, Competencias" if comp else "Chery 01-18 | |19-31"
    if "annissan" in fn: return "Nissan, Competencia" if comp else "Nissan"
    for code, client in _CODE_MAP.items():
        if code in fn: return client
    if any(kw in t for t in tokens for kw in ["chery"]):
        return "Chery - Changan, Competencias" if comp else "Chery 01-18 | |19-31"
    if any(kw in t for t in tokens for kw in ["niss","nissan"]):
        return "Nissan, Competencia" if comp else "Nissan"
    for client, kws in _LEGACY.items():
        if any(kw in t for t in tokens for kw in kws): return client
    return None

def build_consolidated_conteo(processed_files, manual_codif=None):
    mc = manual_codif or {}
    cav, cgraf = {}, {}
    for item in processed_files:
        c = get_client_category(item['filename'])
        if c:
            cav[c]   = cav.get(c, 0)   + item['av_count']
            cgraf[c] = cgraf.get(c, 0) + item['grafica_count']
    rows = []
    for client, tipo, codigo in CONTEO_TEMPLATE:
        av   = cav.get(client, 0); gr = cgraf.get(client, 0)
        m_av = int(mc.get(client, {}).get('av', 0) or 0)
        m_im = int(mc.get(client, {}).get('impresos', 0) or 0)
        rep  = client in REPLICATED_CLIENTS
        if   tipo == "Notas Audiovisuales":        val = av
        elif tipo == "Notas Impresos":             val = gr
        elif tipo == "Codificación Audiovisuales": val = (av if rep else 0) + m_av
        elif tipo == "Codificación Impresos":      val = (gr if rep else 0) + m_im
        else:                                      val = 0
        rows.append({"Cliente / Categoría": client, "Tipo de Conteo": tipo, "Código": codigo, "Cantidad": val})
    return pd.DataFrame(rows)

def to_excel_consolidated(df):
    out = io.BytesIO(); wb = Workbook(); ws = wb.active; ws.title = "Conteo Consolidado"
    hf = PatternFill("solid", fgColor="1E3A5F"); hfont = Font(bold=True, color="FFFFFF", name="Calibri")
    bot = Border(bottom=Side(style='thin', color="E2E8F0")); alt = PatternFill("solid", fgColor="F8FAFC")
    for ci, cn in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=ci, value=cn); c.font = hfont; c.fill = hf; c.border = bot
    for ri, rd in enumerate(df.itertuples(index=False), 2):
        for ci, v in enumerate(rd, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            if ri % 2 == 0: c.fill = alt
            c.border = bot
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(max(len(str(c.value or '')) for c in col)+4, 14)
    wb.save(out); out.seek(0); return out.getvalue()

def to_excel_from_df(df, final_order, filename, av_count, grafica_count):
    out = io.BytesIO(); cols = [c for c in final_order if c in df.columns]; df_out = df[cols].copy()
    for col in df_out.columns:
        if hasattr(df_out[col].dtype, 'pyarrow_dtype'): df_out[col] = df_out[col].astype(object)
    wb = Workbook(); ws = wb.active; ws.title = 'Resultado'
    hf = PatternFill("solid", fgColor="1E3A5F"); hfont = Font(bold=True, color="FFFFFF", name="Calibri")
    for i, cn in enumerate(df_out.columns, 1):
        c = ws.cell(row=1, column=i, value=cn); c.font = hfont; c.fill = hf
    lc = {'Link Nota','Link (Streaming - Imagen)'}
    for ri, rd in enumerate(df_out.itertuples(index=False), 2):
        for ci, value in enumerate(rd, 1):
            cn = df_out.columns[ci-1]; cell = ws.cell(row=ri, column=ci)
            if cn == 'Fecha' and pd.notna(value):
                if isinstance(value, pd.Timestamp): cell.value = value.to_pydatetime(); cell.number_format = 'DD/MM/YYYY'
                else: cell.value = value
            elif cn in lc and pd.notna(value) and isinstance(value,str) and value.startswith('http'):
                cell.value = 'Link'; cell.hyperlink = value
                cell.font = Font(color="1D4ED8", underline="single"); cell.alignment = Alignment(horizontal='left')
            else:
                cell.value = value if pd.notna(value) else None
    for i, cn in enumerate(df_out.columns, 1):
        ltr = ws.cell(row=1,column=i).column_letter
        ws.column_dimensions[ltr].width = 50 if cn in ['Título','Resumen - Aclaracion'] else 15 if cn in lc else 20
    ws2 = wb.create_sheet(title='Conteo')
    df_c = build_consolidated_conteo([{'filename':filename,'av_count':av_count,'grafica_count':grafica_count}])
    for ci,cn in enumerate(df_c.columns,1):
        c=ws2.cell(row=1,column=ci,value=cn); c.font=Font(bold=True,color="FFFFFF"); c.fill=hf
    for ri,rd in enumerate(df_c.itertuples(index=False),2):
        for ci,v in enumerate(rd,1): ws2.cell(row=ri,column=ci,value=v)
    for col in ws2.columns:
        ws2.column_dimensions[col[0].column_letter].width = max(max(len(str(c.value or '')) for c in col)+4, 14)
    wb.save(out); out.seek(0); return out.getvalue()

def load_config(src):
    sheets = pd.read_excel(src, sheet_name=None, engine='openpyxl')
    rmap = pd.Series(sheets['Regiones'].iloc[:,1].values,
                     index=sheets['Regiones'].iloc[:,0].astype(str).str.lower().str.strip()).to_dict()
    imap = pd.Series(sheets['Internet'].iloc[:,1].values,
                     index=sheets['Internet'].iloc[:,0].astype(str).str.lower().str.strip()).to_dict()
    return rmap, imap

def process_dossier(dossier_file, rmap, imap):
    wb = load_workbook(dossier_file); sheet = wb.active
    headers = [c.value for c in sheet[1] if c.value is not None]
    rows = []
    for row in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row): continue
        rd = dict(zip(headers, [c.value for c in row[:len(headers)]]))
        for lc in ['URL Nota AV','URL (Streaming - Imagen)','URL Nota','Link Nota AV','Link (Streaming - Imagen)']:
            if lc in headers:
                idx = headers.index(lc)
                if idx < len(row):
                    ext = extract_link_from_cell(row[idx])
                    if ext: rd[lc] = ext
        rows.append(rd)
    df = pd.DataFrame(rows)
    tmap = {'online':'Internet','internet':'Internet','diario':'Prensa','am':'Radio','fm':'Radio',
            'aire':'Televisión','cable':'Televisión','revista':'Revistas','revistas':'Revistas'}
    df['Tipo de Medio'] = df['Tipo de Medio'].astype(str).str.lower().str.strip().map(tmap).fillna(df['Tipo de Medio'].astype(str).str.strip())
    is_av = df['Tipo de Medio'].isin(['Radio','Televisión'])
    is_gr = df['Tipo de Medio'].isin(['Prensa','Internet','Revistas'])
    is_in = df['Tipo de Medio'] == 'Internet'
    df.loc[is_in,'Medio'] = df.loc[is_in,'Medio'].astype(str).str.lower().str.strip().map(imap).fillna(df.loc[is_in,'Medio'])
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(rmap)
    df['ID Noticia']    = df.get('NoticiaId', pd.Series(dtype=str))
    df['Fecha']         = pd.to_datetime(df.get('Fecha',pd.Series(dtype=str)), dayfirst=True, errors='coerce').dt.normalize()
    df['Hora']          = df.get('Hora', pd.Series(dtype=str))
    for c in ['Sección - Programa','Título','Autor - Conductor']:
        df[c] = df.get(c, pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Nro. Pagina']   = df.get('Nro. Pagina', pd.Series(dtype=str))
    df['Dimensión']     = df.get('Dimensioncm2', pd.Series(dtype=str))
    df['Duración - Nro. Caracteres'] = df.get('Duración - Nro. Caracteres', pd.Series(dtype=str))
    df.loc[is_av,'Dimensión'] = df.loc[is_av,'Duración - Nro. Caracteres']
    df.loc[is_av,'Duración - Nro. Caracteres'] = 0
    df['CPE'] = np.where(is_av, df.get('CPE',pd.Series([np.nan]*len(df))),
                         np.where(is_gr, df.get('Valor de Nota',pd.Series([np.nan]*len(df))), np.nan))
    df['Tier']     = df.get('Tier', pd.Series(dtype=str))
    df['Audiencia'] = df.get('Audiencia', pd.Series(dtype=str))
    df['Tono']     = df.get('Tono', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Tema']     = df.get('Tematica', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Temas Generales - Tema'] = df.get('Temas Generales - Tema', pd.Series(dtype=str)).astype(str).apply(clean_text)
    cuerpo = df.get('CuerpoEs',pd.Series(['']*len(df))).astype(str).apply(clean_cuerpo)
    def fmt(t):
        if not isinstance(t,str) or not t.strip(): return t
        ps = [p.strip() for p in t.split('\n') if p.strip()]
        return '\n\n'.join(ps) if len(ps)>1 else t
    df['Resumen - Aclaracion'] = np.where(is_av, cuerpo, cuerpo.apply(fmt))
    url_av   = df.get('URL Nota AV', df.get('Link Nota AV', pd.Series(['']*len(df)))).fillna('').astype(str)
    url_str  = df.get('URL (Streaming - Imagen)', pd.Series(['']*len(df))).fillna('').astype(str)
    df['Link Nota'] = np.where(is_av, url_av.str.replace(r'\.com\.ar','.com.co',regex=True),
                               np.where(is_gr, url_str, '')).replace('', np.nan)
    df['Link (Streaming - Imagen)'] = df.get('URL Nota',pd.Series(['']*len(df))).fillna('').astype(str).replace('',np.nan)
    m_av  = df.get('Menciones - Empresa',pd.Series(['']*len(df))).fillna('').astype(str).apply(clean_text)
    m_gr  = df.get('Empresa rel.',       pd.Series(['']*len(df))).fillna('').astype(str).apply(clean_text)
    df['Menciones - Empresa'] = np.where(is_av, m_av, np.where(is_gr, m_gr, m_av))
    rows_exp = []
    for _, row in df.iterrows():
        menc = [m.strip() for m in str(row['Menciones - Empresa']).split(';') if m.strip()]
        if not menc: rows_exp.append(row.to_dict())
        else:
            for m in menc: nr=row.to_dict(); nr['Menciones - Empresa']=m; rows_exp.append(nr)
    df = pd.DataFrame(rows_exp).reset_index(drop=True)
    return df, int(df['Tipo de Medio'].isin(['Radio','Televisión']).sum()), int(df['Tipo de Medio'].isin(['Prensa','Internet','Revistas']).sum())

# ==============================================================================
# SESSION STATE
# ==============================================================================
def _init():
    for k, v in {
        'uploader_key': 0,
        'resultados':   [],
        'manual_codif': {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS},
        'manual_saved': False,
        # Bandera para procesar el JSON importado SIN hacer rerun inmediato
        '_json_pending': None,
    }.items():
        if k not in st.session_state: st.session_state[k] = v

_init()

# Procesar JSON pendiente ANTES de renderizar (evita el bug de rerun)
if st.session_state['_json_pending'] is not None:
    try:
        data = json.loads(st.session_state['_json_pending'])
        st.session_state['manual_codif'] = {c: data.get(c, {'av':0,'impresos':0}) for c in UNIQUE_CLIENTS}
        st.session_state['manual_saved']  = True
    except Exception:
        pass
    st.session_state['_json_pending'] = None

# ==============================================================================
# HEADER
# ==============================================================================
st.markdown("""
<div class="sov-header">
    <div class="sov-badge">📡 Media Intelligence · SOV</div>
    <h1>Conteo v3</h1>
    <p>Procesa dossiers de monitoreo, calcula conteos y consolida la codificación manual por cliente.</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# SIDEBAR — Configuración + JSON
# ==============================================================================
with st.sidebar:

    # --- Config ---
    st.markdown("### ⚙️ Configuración")
    config_source = None
    if CONFIG_PATH is not None:
        st.success(f"✅ `{CONFIG_PATH.name}` detectado")
        config_source = CONFIG_PATH
    else:
        st.warning("No se encontró `Configuracion.xlsx` en el repositorio.")
        cfg_up = st.file_uploader("Subir Configuracion.xlsx", type=["xlsx"], key="cfg_up")
        if cfg_up: config_source = cfg_up; st.success(f"✅ {cfg_up.name}")

    st.markdown("---")

    # --- Panel JSON ---
    st.markdown("### 💾 Guardar / restaurar codificación")

    # Estado actual
    tiene_datos = any(
        (v.get('av',0) or 0) + (v.get('impresos',0) or 0) > 0
        for v in st.session_state['manual_codif'].values()
    )
    if tiene_datos:
        tot_av  = sum(v.get('av',0) or 0       for v in st.session_state['manual_codif'].values())
        tot_imp = sum(v.get('impresos',0) or 0 for v in st.session_state['manual_codif'].values())
        st.markdown(
            f'<div class="json-panel">'
            f'<h4>Codificación en memoria</h4>'
            f'<span class="json-status loaded">✅ Cargada — AV: {tot_av} | Impresos: {tot_imp}</span>'
            f'<p style="font-size:0.74rem;color:#374151;margin:4px 0 0 0;">'
            f'Estos valores se están sumando al conteo consolidado.</p>'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="json-panel">'
            '<h4>Codificación en memoria</h4>'
            '<span class="json-status empty">Sin datos ingresados aún</span>'
            '</div>',
            unsafe_allow_html=True,
        )

    # Exportar
    json_bytes = json.dumps(st.session_state['manual_codif'], ensure_ascii=False, indent=2).encode('utf-8')
    st.download_button(
        "⬇️ Exportar JSON (guardar para mañana)",
        data=json_bytes,
        file_name=f"codif_manual_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        key="dl_json_sidebar",
        use_container_width=True,
        help="Descargue este archivo al finalizar su sesión. Mañana impórtelo aquí abajo para restaurar los valores.",
    )

    # Importar — sin rerun inmediato
    st.markdown("**Importar JSON de sesión anterior:**")
    json_up = st.file_uploader(
        "Seleccione el archivo .json guardado",
        type=["json"],
        key="json_up_sidebar",
        label_visibility="collapsed",
        help="Seleccione el archivo .json que exportó en una sesión anterior.",
    )
    if json_up is not None:
        raw = json_up.read().decode('utf-8')
        if st.session_state['_json_pending'] != raw:
            st.session_state['_json_pending'] = raw
            st.rerun()

    if st.session_state.get('manual_saved') and tiene_datos:
        st.success("Codificación importada y activa.")

    st.markdown("---")

    # --- Convención de nombres ---
    with st.expander("📋 Convención de nombres de archivo"):
        st.caption(
            "Formato: `<fecha> <CÓDIGO> [m|com]`\n\n"
            "Ejemplos:\n"
            "- `19 ANCHERY m` → Chery (marca)\n"
            "- `19 ANCHERY com` → Chery (competencia)\n"
            "- `19 ANNISSAN m` → Nissan (marca)\n"
            "- `19 FSANTAFE_AN` → Fundación Santa Fe\n\n"
            "Códigos válidos: `ANCHERY`, `ANNISSAN`, `ACOMFEVALLE`, `ANFENAVI`, "
            "`FSANTAFE_AN`, `TIGOAN`, `USIMONAN`, `UTB_AN`"
        )

    st.markdown("---")
    if st.button("🗑️ Limpiar resultados procesados", type="secondary", use_container_width=True):
        st.session_state['resultados'] = []
        st.session_state['uploader_key'] += 1
        st.rerun()

# ==============================================================================
# CUERPO — Pasos de uso
# ==============================================================================

# ── PASO 1: Codificación manual ────────────────────────────────────────────────
tiene_manual = any(
    (v.get('av',0) or 0) + (v.get('impresos',0) or 0) > 0
    for v in st.session_state['manual_codif'].values()
)
paso1_num_class = "done" if tiene_manual else "warn" if not tiene_manual else ""

st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_manual else ''}">{1}</span>
    <div>
      <div class="step-title">Codificación manual {'✅' if tiene_manual else ''}</div>
      <p class="step-desc">Si ya tiene notas codificadas de períodos anteriores, ingréselas aquí antes de procesar los dossiers.
      Si viene de una sesión anterior, restaure el JSON desde el panel lateral.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

with st.expander("✏️ Ingresar / revisar codificación manual", expanded=not tiene_manual):

    # Guía de orden
    with st.expander("📖 ¿Qué valores van aquí y en qué orden aparecen en el conteo?", expanded=False):
        st.markdown("""
**Regla general:**

| Fila en el conteo | ¿Quién la llena? |
|---|---|
| Codificación Audiovisuales | **Usted manualmente** (notas ya analizadas) |
| Codificación Impresos | **Usted manualmente** (notas ya analizadas) |
| Notas Audiovisuales | La app, automáticamente al procesar el dossier |
| Notas Impresos | La app, automáticamente al procesar el dossier |

**Excepción — Chery y Nissan:** la Codificación se autocompleta igual que las Notas (son "clientes replicados"). Solo ingrese un valor manual si tiene notas extra fuera del dossier.

**Orden de las 40 filas:**
""")
        rows_html = ""
        for i, (client, tipo, codigo) in enumerate(CONTEO_TEMPLATE, 1):
            hl = "og-row hl" if "Codificación" in tipo else "og-row"
            rows_html += (
                f'<div class="{hl}">'
                f'<span class="og-num">{i}</span>'
                f'<span class="og-tipo">{tipo}</span>'
                f'<span class="og-cli">{client}</span>'
                f'<span class="og-cod">{codigo}</span>'
                f'</div>'
            )
        st.markdown(
            '<div class="order-guide">'
            '<div class="og-row hdr"><span class="og-num">#</span>'
            '<span class="og-tipo">Tipo de Conteo</span>'
            '<span class="og-cli">Cliente</span>'
            '<span class="og-cod">Código</span></div>'
            + rows_html + '</div>',
            unsafe_allow_html=True,
        )
        st.caption("Las filas resaltadas en azul son las de Codificación — las que usted ingresa aquí.")

    st.markdown(
        '<div class="info-box">Los valores de Codificación se <strong>suman</strong> al conteo automático. '
        'Ingrese el acumulado total que ya tiene analizado para cada cliente. '
        'Si no tiene nada para un cliente, déjelo en 0.</div>',
        unsafe_allow_html=True,
    )

    with st.form("form_manual", clear_on_submit=False):
        nuevos = {}
        for client in UNIQUE_CLIENTS:
            st.markdown(f'<div class="mc-block"><h4>{client}</h4></div>', unsafe_allow_html=True)
            ca, ci2 = st.columns(2)
            act = st.session_state['manual_codif'].get(client, {'av':0,'impresos':0})
            val_av = ca.number_input("🎬 Codif. Audiovisuales", min_value=0, step=1,
                                     value=int(act.get('av',0) or 0), key=f"mav_{client}")
            val_im = ci2.number_input("🗞️ Codif. Impresos",    min_value=0, step=1,
                                     value=int(act.get('impresos',0) or 0), key=f"mim_{client}")
            nuevos[client] = {'av': val_av, 'impresos': val_im}
            st.markdown("")

        cs, cr = st.columns([3, 1])
        if cs.form_submit_button("💾 Guardar valores de codificación", type="primary", use_container_width=True):
            st.session_state['manual_codif'] = nuevos
            st.session_state['manual_saved']  = True
            st.rerun()
        if cr.form_submit_button("🔄 Restablecer", use_container_width=True):
            st.session_state['manual_codif'] = {c: {'av':0,'impresos':0} for c in UNIQUE_CLIENTS}
            st.session_state['manual_saved']  = False
            st.rerun()

st.markdown("")

# ── PASO 2: Cargar y procesar dossiers ────────────────────────────────────────
tiene_resultados = bool(st.session_state.get('resultados'))

st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_resultados else ''}">{2}</span>
    <div>
      <div class="step-title">Cargar y procesar dossiers {'✅' if tiene_resultados else ''}</div>
      <p class="step-desc">Cargue los archivos .xlsx del dossier. La app detecta el cliente por el nombre del archivo
      y calcula las Notas AV y Notas Impresos automáticamente.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Seleccione uno o varios archivos .xlsx",
    type=["xlsx"], accept_multiple_files=True,
    key=f"dos_{st.session_state['uploader_key']}",
    label_visibility="collapsed",
)

if uploaded:
    st.markdown(
        f'<div class="info-box">📎 {len(uploaded)} archivo(s) listos para procesar: '
        + ', '.join(f'<code>{f.name}</code>' for f in uploaded) + '</div>',
        unsafe_allow_html=True,
    )

can_run = bool(uploaded and config_source)
cb, ch = st.columns([2, 5])
with cb:
    run = st.button("▶ Procesar dossiers", disabled=not can_run, type="primary", use_container_width=True)
with ch:
    if not config_source:
        st.caption("⚠️ Primero cargue `Configuracion.xlsx` en el panel lateral.")
    elif not uploaded:
        st.caption("Seleccione al menos un archivo .xlsx para continuar.")

if run:
    try: rmap, imap = load_config(config_source)
    except Exception as e: st.error(f"Error cargando Configuracion.xlsx: {e}"); st.stop()
    FINAL_ORDER = ["ID Noticia","Fecha","Hora","Medio","Tipo de Medio","Sección - Programa",
                   "Región","Título","Autor - Conductor","Nro. Pagina","Dimensión",
                   "Duración - Nro. Caracteres","CPE","Tier","Audiencia","Tono","Tema",
                   "Temas Generales - Tema","Resumen - Aclaracion",
                   "Link Nota","Link (Streaming - Imagen)","Menciones - Empresa"]
    nuevos = []
    pb = st.progress(0, text="Iniciando...")
    for i, f in enumerate(uploaded):
        pb.progress(i/len(uploaded), text=f"Procesando {f.name}...")
        try:
            df, av, gr = process_dossier(f, rmap, imap)
            exc = to_excel_from_df(df, FINAL_ORDER, f.name, av, gr)
            nuevos.append({'nombre': f.name, 'graficas': gr, 'av': av, 'total': len(df),
                           'excel': exc, 'matched_client': get_client_category(f.name),
                           'filename': f"SOV_{f.name.replace('.xlsx','')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"})
        except Exception as e: st.error(f"Error en {f.name}: {e}")
    pb.progress(1.0, text="✅ Listo")
    st.session_state['resultados'].extend(nuevos)
    st.session_state['uploader_key'] += 1
    st.balloons()
    st.rerun()

st.markdown("")

# ── PASO 3: Resultados ─────────────────────────────────────────────────────────
st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_resultados else ''}">{3}</span>
    <div>
      <div class="step-title">Resultados y conteo consolidado {'✅' if tiene_resultados else ''}</div>
      <p class="step-desc">Aquí verá el conteo final con las Notas calculadas automáticamente
      más la Codificación ingresada en el Paso 1. Copie la columna de números directamente a su plantilla de Excel.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

if tiene_resultados:
    resultados = st.session_state['resultados']
    c1,c2,c3,c4 = st.columns(4)
    for col, lbl, val in [
        (c1, "Archivos",          len(resultados)),
        (c2, "Total registros",   sum(r['total']   for r in resultados)),
        (c3, "🎬 Audiovisuales",  sum(r['av']      for r in resultados)),
        (c4, "🗞️ Gráficas",       sum(r['graficas'] for r in resultados)),
    ]:
        col.markdown(f'<div class="metric-card"><span class="mv">{val}</span><span class="ml">{lbl}</span></div>',
                     unsafe_allow_html=True)
    st.markdown("")

    tab_cons, tab_arch = st.tabs(["📊 Conteo consolidado", "📋 Detalle por archivo"])

    # ── Conteo consolidado ──
    with tab_cons:
        listado = [{'filename':r['nombre'],'av_count':r['av'],'grafica_count':r['graficas']} for r in resultados]
        df_cons = build_consolidated_conteo(listado, st.session_state['manual_codif'])

        if tiene_manual:
            st.markdown(
                '<div class="ok-box">✅ El conteo incluye su codificación manual. '
                'Si modifica los valores en el Paso 1, el conteo se actualizará automáticamente.</div>',
                unsafe_allow_html=True,
            )

        col_t, col_n = st.columns([3, 1])

        with col_t:
            st.markdown("##### Tabla de conteo (40 filas en orden estándar)")
            st.dataframe(
                df_cons.style.apply(
                    lambda col: ['color:#1d4ed8;font-weight:700' if v>0 else 'color:#cbd5e1' for v in col]
                    if col.name == 'Cantidad' else ['' for _ in col], axis=0
                ),
                use_container_width=True, hide_index=True, height=620,
            )
            exc_cons = to_excel_consolidated(df_cons)
            st.download_button(
                "📥 Descargar Excel de conteo",
                data=exc_cons,
                file_name=f"Conteo_SOV_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_cons", use_container_width=True,
            )

        with col_n:
            st.markdown("##### 📋 Columna para copiar")
            st.caption("Seleccione todo el texto y péguelo directamente en su plantilla de Excel.")
            st.text_area(
                "40 valores en orden estándar:",
                value="\n".join(df_cons["Cantidad"].astype(str).tolist()),
                height=580, key="nums_area",
            )

    # ── Detalle por archivo ──
    with tab_arch:
        for r in resultados:
            st.markdown(
                f'<div class="file-card"><div class="fn">📄 {r["nombre"]}</div>'
                + (f'<span class="client-tag">🎯 {r["matched_client"]}</span>'
                   if r.get("matched_client") else '<span class="warn-tag">⚠️ Cliente no detectado</span>')
                + '</div>', unsafe_allow_html=True,
            )
            cg, ca, ct, cdl = st.columns([1,1,1,2])
            cg.metric("Gráficas", r['graficas'])
            ca.metric("Audiovisuales", r['av'])
            ct.metric("Total", r['total'])
            cdl.download_button("📥 Descargar Excel", data=r['excel'], file_name=r['filename'],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{r['nombre']}_{r['filename']}", use_container_width=True)
            st.markdown("")

else:
    st.markdown(
        '<div class="warn-box">⏳ Procese al menos un dossier en el Paso 2 para ver el conteo consolidado aquí.</div>',
        unsafe_allow_html=True,
    )
