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
    page_title="SOV · Conteo y Codificación",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Archivo para persistencia local automática
BACKUP_FILE = "sov_backup_estado.json"

# ==============================================================================
# ESTILOS CSS (Tema Claro Moderno y Profesional)
# ==============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { 
    font-family: 'Plus Jakarta Sans', sans-serif; 
}
.stApp { 
    background: #fbfcfd; 
    color: #1e293b; 
}

/* Contenedor del Sidebar */
section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #f1f5f9 !important;
    min-width: 320px !important;
}

/* Cabezote Principal */
.sov-header {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}
.sov-header h1 { 
    font-size: 1.6rem; 
    font-weight: 700; 
    color: #0f172a; 
    margin: 0 0 4px 0; 
    letter-spacing: -0.02em; 
}
.sov-header p { 
    color: #64748b; 
    font-size: 0.88rem; 
    margin: 0; 
}
.sov-badge {
    display: inline-block;
    background: #eff6ff;
    border: 1px solid #dbeafe;
    color: #2563eb;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    padding: 3px 10px;
    border-radius: 6px;
    margin-bottom: 12px;
    font-weight: 600;
    text-transform: uppercase;
}

/* Tarjetas de Pasos */
.step-block {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.01);
}
.step-label {
    display: flex;
    align-items: flex-start;
    gap: 12px;
}
.step-num {
    background: #2563eb;
    color: #ffffff;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    font-weight: 700;
    width: 24px; 
    height: 24px;
    border-radius: 6px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    margin-top: 2px;
}
.step-num.done { 
    background: #10b981; 
}
.step-title { 
    font-weight: 600; 
    font-size: 0.95rem; 
    color: #0f172a; 
    margin-bottom: 4px;
}
.step-desc { 
    font-size: 0.82rem; 
    color: #64748b; 
    margin: 0; 
}

/* Indicadores Métricos */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    box-shadow: 0 1px 2px rgba(0,0,0,0.01);
}
.metric-card .mv { 
    font-family: 'JetBrains Mono', monospace; 
    font-size: 1.6rem; 
    font-weight: 700; 
    color: #2563eb; 
    display: block; 
    line-height: 1; 
}
.metric-card .ml { 
    font-size: 0.72rem; 
    color: #64748b; 
    text-transform: uppercase; 
    letter-spacing: 0.05em; 
    margin-top: 6px; 
    display: block; 
    font-weight: 500;
}

/* Alertas y Notificaciones */
.info-box { 
    background: #f0fdf4; 
    border: 1px solid #bbf7d0; 
    border-radius: 8px; 
    padding: 12px 16px; 
    font-size: 0.82rem; 
    color: #166534; 
    margin-bottom: 16px; 
}
.warn-box { 
    background: #fffbeb; 
    border: 1px solid #fde68a; 
    border-radius: 8px; 
    padding: 12px 16px; 
    font-size: 0.82rem; 
    color: #92400e; 
    margin-bottom: 16px; 
}

/* Bloques de clientes para codificación */
.mc-block {
    background: #f8fafc;
    border-left: 3px solid #cbd5e1;
    border-radius: 0 8px 8px 0;
    padding: 8px 12px;
    margin-top: 14px;
    margin-bottom: 8px;
}
.mc-block h4 { 
    color: #334155; 
    font-size: 0.8rem; 
    font-weight: 600; 
    margin: 0;
    text-transform: uppercase;
    letter-spacing: 0.02em;
}

/* Guía de Orden Estándar */
.order-guide { 
    background: #ffffff; 
    border: 1px solid #e2e8f0; 
    border-radius: 8px; 
    overflow: hidden; 
    font-family: 'JetBrains Mono', monospace; 
    font-size: 0.75rem; 
}
.og-row { 
    display: flex; 
    align-items: center; 
    padding: 8px 12px; 
    border-bottom: 1px solid #f1f5f9; 
    gap: 8px; 
}
.og-row:last-child { border-bottom: none; }
.og-row:nth-child(even) { background: #f8fafc; }
.og-row.hl { background: #eff6ff !important; }
.og-row.hdr { background: #0f172a !important; color: #ffffff !important; }
.og-num { color: #94a3b8; width: 24px; text-align: right; }
.og-tipo { color: #334155; width: 220px; font-weight: 500; }
.og-cli { color: #0f172a; flex: 1; }
.og-cod { color: #64748b; font-size: 0.7rem; }
.hdr .og-num, .hdr .og-tipo, .hdr .og-cli, .hdr .og-cod { color: #ffffff !important; font-weight: 600; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { 
    background: #f1f5f9; 
    border-radius: 8px; 
    padding: 4px; 
    border: none;
    gap: 4px; 
}
.stTabs [data-baseweb="tab"] { 
    background: transparent; 
    border-radius: 6px; 
    color: #64748b; 
    font-size: 0.82rem; 
    font-weight: 500; 
    padding: 8px 16px; 
    border: none !important; 
}
.stTabs [aria-selected="true"] { 
    background: #ffffff !important; 
    color: #2563eb !important; 
    box-shadow: 0 1px 3px rgba(0,0,0,0.05); 
}

/* Personalización de Inputs */
div[data-testid="stNumberInput"] input { 
    background: #ffffff !important; 
    border-color: #cbd5e1 !important; 
    color: #0f172a !important; 
    border-radius: 6px !important; 
}
textarea { 
    background: #ffffff !important; 
    border-color: #cbd5e1 !important; 
    color: #0f172a !important; 
    font-family: 'JetBrains Mono', monospace !important; 
    font-size: 0.85rem !important; 
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONSTANTES Y ESTRUCTURA DEL CONTEO (40 Filas en orden específico)
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
    ("Federación Nacional de Avicultores de Colombia", "Codificación Audiovisuales", "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Codificación Impresos",      "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Notas Audiovisuales",        "ANFENAVI"),
    ("Federación Nacional de Avicultores de Colombia", "Notas Impresos",             "ANFENAVI"),
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

UNIQUE_CLIENTS = list(dict.fromkeys(c for c, _, _ in CONTEO_TEMPLATE))
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
# PERSISTENCIA LOCAL DE ESTADO (Backup automático de datos manuales)
# ==============================================================================
def save_local_backup(data):
    try:
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_local_backup():
    if Path(BACKUP_FILE).exists():
        try:
            with open(BACKUP_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

# ==============================================================================
# FUNCIONES AUXILIARES DE PROCESAMIENTO
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
    fn = re.sub(r'^[\d\s\-_]+', '', Path(filename).name.lower()).strip()
    tokens = [t for t in re.split(r'[^a-z0-9]', fn) if t]
    comp = any(t in {"c","com","comp","competencia","competencias","changan"} for t in tokens)
    
    if "anchery" in fn: 
        return "Chery - Changan, Competencias" if comp else "Chery 01-18 | |19-31"
    if "annissan" in fn: 
        return "Nissan, Competencia" if comp else "Nissan"
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
    
    # Sumar automáticamente los dossiers procesados
    for item in processed_files:
        c = get_client_category(item['filename'])
        if c:
            cav[c]   = cav.get(c, 0)   + item['av_count']
            cgraf[c] = cgraf.get(c, 0) + item['grafica_count']
            
    rows = []
    for client, tipo, codigo in CONTEO_TEMPLATE:
        av = cav.get(client, 0)
        gr = cgraf.get(client, 0)
        
        m_av = int(mc.get(client, {}).get('av', 0) or 0)
        m_im = int(mc.get(client, {}).get('impresos', 0) or 0)
        rep  = client in REPLICATED_CLIENTS
        
        if tipo == "Notas Audiovisuales":
            val = av
        elif tipo == "Notas Impresos":
            val = gr
        elif tipo == "Codificación Audiovisuales":
            val = (av if rep else 0) + m_av
        elif tipo == "Codificación Impresos":
            val = (gr if rep else 0) + m_im
        else:
            val = 0
            
        rows.append({
            "Cliente / Categoría": client, 
            "Tipo de Conteo": tipo, 
            "Código": codigo, 
            "Cantidad": val
        })
    return pd.DataFrame(rows)

def to_excel_consolidated(df):
    out = io.BytesIO(); wb = Workbook(); ws = wb.active; ws.title = "Conteo Consolidado"
    hf = PatternFill("solid", fgColor="2563EB"); hfont = Font(bold=True, color="FFFFFF", name="Calibri")
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
    hf = PatternFill("solid", fgColor="2563EB"); hfont = Font(bold=True, color="FFFFFF", name="Calibri")
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
                cell.font = Font(color="2563EB", underline="single"); cell.alignment = Alignment(horizontal='left')
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
    
    link_nota_arr = np.where(is_av, url_av.str.replace(r'\.com\.ar','.com.co',regex=True),
                             np.where(is_gr, url_str, ''))
    df['Link Nota'] = pd.Series(link_nota_arr, index=df.index).replace('', np.nan)
    
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
# INICIALIZACIÓN DE ESTADOS
# ==============================================================================
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

if 'resultados' not in st.session_state:
    st.session_state['resultados'] = []

# Intentar recuperar el estado de codificación guardado automáticamente
if 'manual_codif' not in st.session_state:
    saved_data = load_local_backup()
    if saved_data:
        st.session_state['manual_codif'] = {c: saved_data.get(c, {'av': 0, 'impresos': 0}) for c in UNIQUE_CLIENTS}
    else:
        st.session_state['manual_codif'] = {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS}

# Estados para controlar los mensajes de éxito/error del JSON importado sin ciclos de reinicio
if 'json_loaded_success' not in st.session_state:
    st.session_state['json_loaded_success'] = False
if 'json_loaded_error' not in st.session_state:
    st.session_state['json_loaded_error'] = None

# ==============================================================================
# FUNCIÓN CALLBACK PARA CARGA DEL JSON (Previene el error de rerun infinito)
# ==============================================================================
def on_json_uploaded():
    uploaded_file = st.session_state.get('json_uploader_sidebar')
    if uploaded_file is None:
        st.session_state['json_loaded_success'] = False
        st.session_state['json_loaded_error'] = None
        return
    try:
        raw_data = uploaded_file.read().decode('utf-8')
        imported_data = json.loads(raw_data)
        
        # Validar y limpiar la estructura importada
        cleaned_data = {}
        for c in UNIQUE_CLIENTS:
            client_data = imported_data.get(c, {'av': 0, 'impresos': 0})
            cleaned_data[c] = {
                'av': int(client_data.get('av', 0)),
                'impresos': int(client_data.get('impresos', 0))
            }
        
        # Guardar en memoria y persistir en disco
        st.session_state['manual_codif'] = cleaned_data
        save_local_backup(cleaned_data)
        
        st.session_state['json_loaded_success'] = True
        st.session_state['json_loaded_error'] = None
    except Exception as e:
        st.session_state['json_loaded_success'] = False
        st.session_state['json_loaded_error'] = f"Formato inválido: {str(e)}"

# ==============================================================================
# CABECERA PRINCIPAL
# ==============================================================================
st.markdown("""
<div class="sov-header">
    <div class="sov-badge">Sistema Integrado · SOV</div>
    <h1>Conteo y Consolidación de Dossiers</h1>
    <p>Procese los archivos de monitoreo y consolide las métricas con su codificación manual acumulada.</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# SIDEBAR (Control de Configuración y Respaldo Externo)
# ==============================================================================
with st.sidebar:
    st.markdown("### ⚙️ Configuración del Sistema")
    
    # Detección de Configuración
    config_source = None
    if CONFIG_PATH is not None:
        st.success(f"📂 Archivo de marcas detectado:\n`{CONFIG_PATH.name}`")
        config_source = CONFIG_PATH
    else:
        st.warning("No se encontró `Configuracion.xlsx` en la carpeta base.")
        cfg_up = st.file_uploader("Cargue el archivo de configuración (.xlsx)", type=["xlsx"], key="cfg_up")
        if cfg_up:
            config_source = cfg_up
            st.success(f"✅ Cargado: {cfg_up.name}")

    st.markdown("---")
    st.markdown("### 💾 Guardar / Cargar Sesión")
    
    # Indicador visual del estado de memoria
    codigos_activos = sum((v.get('av',0) + v.get('impresos',0) > 0) for v in st.session_state['manual_codif'].values())
    
    if codigos_activos > 0:
        st.info(f"✨ Memoria Activa: {codigos_activos} cliente(s) con valores cargados.")
    else:
        st.caption("No hay valores manuales en memoria actualmente.")

    # Exportación manual de seguridad
    json_bytes = json.dumps(st.session_state['manual_codif'], ensure_ascii=False, indent=2).encode('utf-8')
    st.download_button(
        "⬇️ Descargar Copia de Respaldo (.json)",
        data=json_bytes,
        file_name=f"sov_codif_manual_{datetime.datetime.now().strftime('%Y%m%d')}.json",
        mime="application/json",
        use_container_width=True,
        help="Guarde una copia externa para importarla en cualquier momento o equipo."
    )

    # Importación de archivo externo usando CALLBACK seguro
    st.markdown("**Cargar Respaldo Externo:**")
    st.file_uploader(
        "Subir archivo .json",
        type=["json"],
        key="json_uploader_sidebar",
        label_visibility="collapsed",
        on_change=on_json_uploaded
    )
    
    # Mostrar estado de la carga de forma controlada
    if st.session_state['json_loaded_success']:
        st.success("✅ Copia de seguridad restaurada correctamente.")
    elif st.session_state['json_loaded_error'] is not None:
        st.error(f"❌ {st.session_state['json_loaded_error']}")

    st.markdown("---")
    if st.button("🗑️ Reiniciar Sesión", type="secondary", use_container_width=True):
        st.session_state['resultados'] = []
        st.session_state['manual_codif'] = {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS}
        st.session_state['json_loaded_success'] = False
        st.session_state['json_loaded_error'] = None
        save_local_backup(st.session_state['manual_codif'])
        st.session_state['uploader_key'] += 1
        st.rerun()

# ==============================================================================
# PANEL PRINCIPAL (FLUJO PASO A PASO)
# ==============================================================================

# ──────────────────────────────────────────────────────────────────────────────
# PASO 1: Codificación Manual (Persistente)
# ──────────────────────────────────────────────────────────────────────────────
tiene_manual = any((v.get('av',0) or 0) + (v.get('impresos',0) or 0) > 0 for v in st.session_state['manual_codif'].values())

st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_manual else ''}">1</span>
    <div>
      <div class="step-title">Codificación Manual de Notas {'' if not tiene_manual else '✓'}</div>
      <p class="step-desc">Registre los acumulados ya analizados de periodos previos. Se conservarán automáticamente para sus próximas sesiones.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

with st.expander("✏️ Ver y Modificar Datos Manuales de Codificación", expanded=not tiene_manual):
    st.markdown(
        '<div class="warn-box">⚠️ <strong>Importante:</strong> Para los clientes <strong>Chery</strong> y <strong>Nissan</strong>, '
        'el conteo automático ya incluye su codificación en base al archivo procesado. Registre cantidades manuales adicionales aquí solo si tiene notas '
        'por fuera de los dossiers.</div>',
        unsafe_allow_html=True
    )
    
    # Botón para ver la guía de orden
    with st.expander("📖 Estructura del Conteo (40 Filas en Orden Estándar)", expanded=False):
        rows_html = ""
        for i, (client, tipo, codigo) in enumerate(CONTEO_TEMPLATE, 1):
            hl = "og-row hl" if "Codificación" in tipo else "og-row"
            rows_html += f'<div class="{hl}"><span class="og-num">{i}</span><span class="og-tipo">{tipo}</span><span class="og-cli">{client}</span><span class="og-cod">{codigo}</span></div>'
        st.markdown(f'<div class="order-guide">{rows_html}</div>', unsafe_allow_html=True)

    # Formulario para entrada de datos manuales
    with st.form("form_manual_data"):
        nuevos_valores = {}
        for client in UNIQUE_CLIENTS:
            st.markdown(f'<div class="mc-block"><h4>{client}</h4></div>', unsafe_allow_html=True)
            col_av, col_imp = st.columns(2)
            
            valor_actual = st.session_state['manual_codif'].get(client, {'av': 0, 'impresos': 0})
            
            val_av = col_av.number_input(
                "🎬 Audiovisuales Codificadas", 
                min_value=0, step=1, 
                value=int(valor_actual.get('av', 0)), 
                key=f"input_av_{client}"
            )
            val_im = col_imp.number_input(
                "🗞️ Impresas Codificadas", 
                min_value=0, step=1, 
                value=int(valor_actual.get('impresos', 0)), 
                key=f"input_im_{client}"
            )
            nuevos_valores[client] = {'av': val_av, 'impresos': val_im}

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn_save, col_btn_clear = st.columns([3, 1])
        
        if col_btn_save.form_submit_button("💾 Guardar Cambios y Actualizar Conteo", type="primary", use_container_width=True):
            st.session_state['manual_codif'] = nuevos_valores
            save_local_backup(nuevos_valores)
            st.rerun()
            
        if col_btn_clear.form_submit_button("🔄 Reestablecer a Cero", use_container_width=True):
            vacio = {c: {'av': 0, 'impresos': 0} for c in UNIQUE_CLIENTS}
            st.session_state['manual_codif'] = vacio
            save_local_backup(vacio)
            st.rerun()

# ──────────────────────────────────────────────────────────────────────────────
# PASO 2: Procesar Nuevos Dossiers
# ──────────────────────────────────────────────────────────────────────────────
tiene_resultados = len(st.session_state['resultados']) > 0

st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_resultados else ''}">2</span>
    <div>
      <div class="step-title">Carga y Procesamiento de Dossiers {'' if not tiene_resultados else '✓'}</div>
      <p class="step-desc">Suba uno o varios archivos de monitoreo en formato Excel. El sistema identificará el cliente automáticamente.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Seleccione archivos .xlsx",
    type=["xlsx"], 
    accept_multiple_files=True,
    key=f"uploader_dossiers_{st.session_state['uploader_key']}",
    label_visibility="collapsed"
)

if uploaded_files:
    st.markdown(
        f'<div class="info-box">📎 {len(uploaded_files)} archivo(s) cargado(s) y listos para procesar.</div>',
        unsafe_allow_html=True
    )

# Bloque de ejecución
col_btn_run, col_status_run = st.columns([2, 5])
habilitado = bool(uploaded_files and config_source)

with col_btn_run:
    ejecutar = st.button("▶ Procesar Archivos", disabled=not habilitado, type="primary", use_container_width=True)

with col_status_run:
    if not config_source:
        st.caption("⚠️ Requiere el archivo de configuración en el panel izquierdo para procesar.")
    elif not uploaded_files:
        st.caption("Seleccione al menos un archivo .xlsx para iniciar.")

if ejecutar:
    try:
        rmap, imap = load_config(config_source)
    except Exception as e:
        st.error(f"Error al abrir el archivo de configuración: {e}")
        st.stop()

    FINAL_ORDER = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa",
        "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión",
        "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema",
        "Temas Generales - Tema", "Resumen - Aclaracion",
        "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"
    ]
    
    nuevos_resultados = []
    progreso = st.progress(0, text="Iniciando procesamiento...")
    
    for indice, f in enumerate(uploaded_files):
        progreso.progress(indice / len(uploaded_files), text=f"Procesando: {f.name}...")
        try:
            df, conteo_av, conteo_graf = process_dossier(f, rmap, imap)
            archivo_procesado = to_excel_from_df(df, FINAL_ORDER, f.name, conteo_av, conteo_graf)
            
            nuevos_resultados.append({
                'nombre': f.name, 
                'graficas': conteo_graf, 
                'av': conteo_av, 
                'total': len(df),
                'excel': archivo_procesado, 
                'matched_client': get_client_category(f.name),
                'filename': f"SOV_{f.name.replace('.xlsx','')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            })
        except Exception as e:
            st.error(f"Error procesando {f.name}: {e}")
            
    progreso.progress(1.0, text="Procesamiento completado con éxito.")
    st.session_state['resultados'].extend(nuevos_resultados)
    st.session_state['uploader_key'] += 1
    st.rerun()

# ──────────────────────────────────────────────────────────────────────────────
# PASO 3: Resultados Consolidados
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="step-block">
  <div class="step-label">
    <span class="step-num {'done' if tiene_resultados else ''}">3</span>
    <div>
      <div class="step-title">Resultados y Conteo Consolidado {'' if not tiene_resultados else '✓'}</div>
      <p class="step-desc">Consulte los indicadores unificados y obtenga la columna ordenada de datos para su reporte.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

if tiene_resultados:
    resultados = st.session_state['resultados']
    
    # Métricas Globales
    mc_1, mc_2, mc_3, mc_4 = st.columns(4)
    mc_1.markdown(f'<div class="metric-card"><span class="mv">{len(resultados)}</span><span class="ml">Dossiers Procesados</span></div>', unsafe_allow_html=True)
    mc_2.markdown(f'<div class="metric-card"><span class="mv">{sum(r["total"] for r in resultados)}</span><span class="ml">Total de Registros</span></div>', unsafe_allow_html=True)
    mc_3.markdown(f'<div class="metric-card"><span class="mv">{sum(r["av"] for r in resultados)}</span><span class="ml">Total Audiovisual</span></div>', unsafe_allow_html=True)
    mc_4.markdown(f'<div class="metric-card"><span class="mv">{sum(r["graficas"] for r in resultados)}</span><span class="ml">Total Gráfico</span></div>', unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Pestañas de Presentación de Resultados
    pestaña_conteo, pestaña_archivos = st.tabs(["📊 Conteo Unificado", "📋 Detalles por Archivo"])
    
    with pestaña_conteo:
        lista_dossiers = [{'filename': r['nombre'], 'av_count': r['av'], 'grafica_count': r['graficas']} for r in resultados]
        df_final = build_consolidated_conteo(lista_dossiers, st.session_state['manual_codif'])
        
        st.markdown(
            '<div class="info-box">✅ Conteo actualizado. Los cambios manuales realizados en el paso 1 se ven reflejados de forma inmediata.</div>',
            unsafe_allow_html=True
        )
        
        col_tabla, col_copia = st.columns([3, 1])
        
        with col_tabla:
            st.markdown("##### Estructura Consolidada (40 Registros)")
            
            # Aplicar estilos para destacar filas con valores reales
            st.dataframe(
                df_final.style.apply(
                    lambda col: ['color: #2563eb; font-weight: bold;' if val > 0 else 'color: #94a3b8;' for val in col]
                    if col.name == 'Cantidad' else ['' for _ in col], axis=0
                ),
                use_container_width=True, 
                hide_index=True, 
                height=650
            )
            
            archivo_conteo_excel = to_excel_consolidated(df_final)
            st.download_button(
                "📥 Descargar Conteo en Excel (.xlsx)",
                data=archivo_conteo_excel,
                file_name=f"SOV_Conteo_Final_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col_copia:
            st.markdown("##### 📋 Copiar Datos Numéricos")
            st.caption("Seleccione y copie la columna de valores para pegarla en su plantilla matriz.")
            
            valores_linea = "\n".join(df_final["Cantidad"].astype(str).tolist())
            st.text_area(
                "Valores listos para copiar:",
                value=valores_linea,
                height=600,
                key="conteo_copia_rapida"
            )
            
    with pestaña_archivos:
        for item in resultados:
            st.markdown(f"""
            <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; margin-bottom: 12px;">
                <div style="font-weight: 600; font-size: 0.88rem; color: #0f172a; margin-bottom: 4px;">📄 {item['nombre']}</div>
                <div style="font-size: 0.76rem; color: #64748b; margin-bottom: 10px;">
                    Identificado como: <span style="background: #eff6ff; color: #2563eb; padding: 2px 6px; border-radius: 4px; font-weight: 500;">{item['matched_client'] if item['matched_client'] else 'No Reconocido'}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            col_det_1, col_det_2, col_det_3, col_det_down = st.columns([1, 1, 1, 2])
            col_det_1.metric("Gráficos", item['graficas'])
            col_det_2.metric("Audiovisual", item['av'])
            col_det_3.metric("Total", item['total'])
            
            col_det_down.download_button(
                "📥 Descargar Dossier Filtrado", 
                data=item['excel'], 
                file_name=item['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_file_{item['nombre']}",
                use_container_width=True
            )
            st.markdown("<hr style='margin: 12px 0; border: 0; border-top: 1px solid #f1f5f9;'>", unsafe_allow_html=True)

else:
    st.markdown(
        '<div class="warn-box">⏳ Cargue y procese los dossiers en el Paso 2 para habilitar el reporte final.</div>',
        unsafe_allow_html=True
    )
