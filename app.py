import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
import datetime
import io
import re
import html
import numpy as np

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador SOV v2.0", layout="wide")

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

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
        except:
            return match.group(0)

    def replace_decimal_entity(match):
        try:
            return chr(int(match.group(1)))
        except:
            return match.group(0)

    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex_entity, text)
    text = re.sub(r'&#(\d+);', replace_decimal_entity, text)

    custom_replacements = {
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        'Â': '', 'â': '', '€': '', '™': ''
    }
    for entity, char in custom_replacements.items():
        text = text.replace(entity, char)
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

def to_excel_from_df(df, final_order):
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df].copy()

    for col in df_to_excel.columns:
        if hasattr(df_to_excel[col].dtype, 'pyarrow_dtype'):
            df_to_excel[col] = df_to_excel[col].astype(object)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Resultado'

    for col_idx, col_name in enumerate(df_to_excel.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)

    link_columns = {'Link Nota', 'Link (Streaming - Imagen)'}

    for row_idx, row_data in enumerate(df_to_excel.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row_data, start=1):
            col_name = df_to_excel.columns[col_idx - 1]
            cell = ws.cell(row=row_idx, column=col_idx)

            if col_name == 'Fecha' and pd.notna(value):
                if isinstance(value, pd.Timestamp):
                    cell.value = value.to_pydatetime()
                    cell.number_format = 'DD/MM/YYYY'
                else:
                    cell.value = value
            elif col_name in link_columns and pd.notna(value) and isinstance(value, str) and value.startswith('http'):
                cell.value = 'Link'
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.value = value if pd.notna(value) else None

    for col_idx, col_name in enumerate(df_to_excel.columns, start=1):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        if col_name in ['Título', 'Resumen - Aclaracion']:
            ws.column_dimensions[col_letter].width = 50
        elif col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            ws.column_dimensions[col_letter].width = 15
        else:
            ws.column_dimensions[col_letter].width = 20

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ==============================================================================
# LÓGICA PRINCIPAL DE PROCESAMIENTO
# ==============================================================================

def run_process(dossier_file, config_file):
    st.markdown("---")
    progress_text = st.empty()

    # --- Paso 1: Configuración ---
    progress_text.info("Paso 1/4: Cargando archivo de configuración...")
    try:
        config_sheets = pd.read_excel(config_file, sheet_name=None, engine='openpyxl')
        region_map = pd.Series(
            config_sheets['Regiones'].iloc[:, 1].values,
            index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
    except Exception as e:
        st.error(f"Error al cargar Configuracion.xlsx: {e}. Asegúrate de que tenga la hoja 'Regiones'.")
        st.stop()

    # --- Paso 2: Leer Dossier ---
    progress_text.info("Paso 2/4: Leyendo el archivo de entrada...")
    try:
        wb = load_workbook(dossier_file)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1] if cell.value is not None]

        rows = []
        for row in sheet.iter_rows(min_row=2):
            if all(c.value is None for c in row):
                continue
            row_values = [c.value for c in row[:len(headers)]]
            row_data = dict(zip(headers, row_values))

            for link_col_name in ['URL Nota AV', 'URL (Streaming - Imagen)', 'URL Nota', 'Link Nota AV', 'Link (Streaming - Imagen)']:
                if link_col_name in headers:
                    idx = headers.index(link_col_name)
                    if idx < len(row):
                        extracted = extract_link_from_cell(row[idx])
                        if extracted:
                            row_data[link_col_name] = extracted

            rows.append(row_data)
    except Exception as e:
        st.error(f"Error al leer el Dossier: {e}")
        st.stop()

    df = pd.DataFrame(rows)

    # --- Paso 3: Transformaciones ---
    progress_text.info("Paso 3/4: Aplicando transformaciones y mapeos...")

    tipo_medio_map = {
        'online': 'Internet',
        'diario': 'Diario',
        'am': 'Radio AM',
        'fm': 'Radio FM',
        'aire': 'Televisión Aire',
        'cable': 'Televisión Cable',
        'revista': 'Revistas',
        'revistas': 'Revistas',
    }
    df['Tipo de Medio'] = df['Tipo de Medio'].astype(str).str.lower().str.strip().map(tipo_medio_map).fillna(df['Tipo de Medio'].astype(str).str.strip())

    is_av = df['Tipo de Medio'].isin(['Radio AM', 'Radio FM', 'Televisión Aire', 'Televisión Cable'])
    is_grafica = df['Tipo de Medio'].isin(['Diario', 'Internet', 'Revistas'])

    df['ID Noticia'] = df.get('NoticiaId', pd.Series(dtype=str))
    df['Fecha'] = pd.to_datetime(df.get('Fecha', pd.Series(dtype=str)), dayfirst=True, errors='coerce').dt.normalize()
    df['Hora'] = df.get('Hora', pd.Series(dtype=str))
    df['Medio'] = df.get('Medio', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Sección - Programa'] = df.get('Sección - Programa', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Región'] = df['Medio'].str.lower().str.strip().map(region_map)
    df['Título'] = df.get('Título', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Autor - Conductor'] = df.get('Autor - Conductor', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Nro. Pagina'] = df.get('Nro. Pagina', pd.Series(dtype=str))
    df['Dimensión'] = df.get('Dimensioncm2', pd.Series(dtype=str))
    df['Duración - Nro. Caracteres'] = df.get('Duración - Nro. Caracteres', pd.Series(dtype=str))

    cpe_av = df.get('CPE', pd.Series([np.nan] * len(df)))
    cpe_grafica = df.get('Valor de Nota', pd.Series([np.nan] * len(df)))
    df['CPE'] = np.where(is_av, cpe_av, np.where(is_grafica, cpe_grafica, np.nan))

    df['Tier'] = df.get('Tier', pd.Series(dtype=str))
    df['Audiencia'] = df.get('Audiencia', pd.Series(dtype=str))
    df['Tono'] = df.get('Tono', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Tema'] = df.get('Tematica', pd.Series(dtype=str)).astype(str).apply(clean_text)
    df['Temas Generales - Tema'] = df.get('Temas Generales - Tema', pd.Series(dtype=str)).astype(str).apply(clean_text)

    cuerpo_raw = df.get('CuerpoEs', pd.Series([''] * len(df))).astype(str)
    cuerpo_cleaned = cuerpo_raw.apply(clean_cuerpo)

    def format_resumen_grafica(text):
        if not isinstance(text, str) or text.strip() == '':
            return text
        parrafos = [p.strip() for p in text.split('\n') if p.strip()]
        return '\n\n'.join(parrafos) if len(parrafos) > 1 else text

    resumen_grafica = cuerpo_cleaned.apply(format_resumen_grafica)
    df['Resumen - Aclaracion'] = np.where(is_av, cuerpo_cleaned, resumen_grafica)

    url_nota_av = df.get('URL Nota AV', df.get('Link Nota AV', pd.Series([''] * len(df)))).fillna('').astype(str)
    url_streaming = df.get('URL (Streaming - Imagen)', pd.Series([''] * len(df))).fillna('').astype(str)
    df['Link Nota'] = np.where(is_av, url_nota_av, np.where(is_grafica, url_streaming, ''))
    df['Link Nota'] = df['Link Nota'].replace('', np.nan)

    url_nota = df.get('URL Nota', pd.Series([''] * len(df))).fillna('').astype(str)
    df['Link (Streaming - Imagen)'] = url_nota.replace('', np.nan)

    # Menciones - Empresa: AV → "Menciones - Empresa", Gráfica → "Empresa rel."
    menciones_av = df.get('Menciones - Empresa', pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    menciones_grafica = df.get('Empresa rel.', pd.Series([''] * len(df))).fillna('').astype(str).apply(clean_text)
    df['Menciones - Empresa'] = np.where(is_av, menciones_av, np.where(is_grafica, menciones_grafica, menciones_av))

    # --- Expansión de filas por Menciones - Empresa (separadas por ;) ---
    rows_expanded = []
    for _, row in df.iterrows():
        menciones = [m.strip() for m in str(row['Menciones - Empresa']).split(';') if m.strip()]
        if not menciones:
            rows_expanded.append(row)
        else:
            for mencion in menciones:
                new_row = row.copy()
                new_row['Menciones - Empresa'] = mencion
                rows_expanded.append(new_row)
    df = pd.DataFrame(rows_expanded).reset_index(drop=True)

    # Recalcular is_av / is_grafica sobre el df expandido para los conteos
    is_av_exp = df['Tipo de Medio'].isin(['Radio AM', 'Radio FM', 'Televisión Aire', 'Televisión Cable'])
    is_grafica_exp = df['Tipo de Medio'].isin(['Diario', 'Internet', 'Revistas'])

    # --- Paso 4: Guardar resultados en session_state ---
    progress_text.info("Paso 4/4: Generando resultados...")

    av_count = int(is_av_exp.sum())
    grafica_count = int(is_grafica_exp.sum())

    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio",
        "Sección - Programa", "Región", "Título", "Autor - Conductor",
        "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres",
        "CPE", "Tier", "Audiencia", "Tono", "Tema",
        "Temas Generales - Tema", "Resumen - Aclaracion",
        "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"
    ]

    excel_data = to_excel_from_df(df, final_order)

    # Persistir en session_state para que sobrevivan al clic de descarga
    st.session_state['resultado_listo'] = True
    st.session_state['excel_data'] = excel_data
    st.session_state['filename'] = f"SOV_Procesado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.session_state['total_count'] = len(df)
    st.session_state['av_count'] = av_count
    st.session_state['grafica_count'] = grafica_count

    progress_text.success("¡Proceso completado!")
    st.balloons()

# ==============================================================================
# INTERFAZ STREAMLIT
# ==============================================================================

st.title("🚀 Procesador SOV v2.0")
st.markdown("Transforma el archivo de entrada al formato estándar de salida con todos los mapeos aplicados.")

st.info(
    "**Instrucciones:**\n\n"
    "1. Sube el archivo Dossier de entrada y el archivo `Configuracion.xlsx`.\n"
    "2. El archivo de configuración debe tener una hoja **Regiones** con columna A = Medio y columna B = Región.\n"
    "3. Haz clic en **Iniciar Proceso** y descarga el resultado."
)

with st.expander("📋 Ver mapeo de columnas aplicado"):
    st.markdown("""
| Columna Salida | Fuente |
|---|---|
| ID Noticia | NoticiaId |
| Fecha | Fecha |
| Hora | Hora |
| Medio | Medio |
| Tipo de Medio | Tipo de Medio (normalizado) |
| Sección - Programa | Sección - Programa |
| Región | Generada desde Configuracion.xlsx / hoja Regiones |
| Título | Título |
| Autor - Conductor | Autor - Conductor |
| Nro. Pagina | Nro. Pagina |
| Dimensión | Dimensioncm2 |
| Duración - Nro. Caracteres | Duración - Nro. Caracteres |
| CPE | CPE (AV) / Valor de Nota (Gráfica) |
| Tier | Tier |
| Audiencia | Audiencia |
| Tono | Tono |
| Tema | Tematica |
| Temas Generales - Tema | Temas Generales - Tema |
| Resumen - Aclaracion | CuerpoEs |
| Link Nota | URL Nota AV (AV) / URL (Streaming - Imagen) (Gráfica) |
| Link (Streaming - Imagen) | URL Nota |
| Menciones - Empresa | Menciones - Empresa (AV) / Empresa rel. (Gráfica) |
""")

uploaded_files = st.file_uploader(
    "Sube aquí el Dossier y el archivo Configuracion.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
)

dossier_file, config_file = None, None

if uploaded_files:
    for file in uploaded_files:
        if 'config' in file.name.lower():
            config_file = file
        else:
            dossier_file = file

    if dossier_file:
        st.success(f"✅ Dossier: {dossier_file.name}")
    else:
        st.warning("⚠️ No se detectó el archivo Dossier.")

    if config_file:
        st.success(f"✅ Configuración: {config_file.name}")
    else:
        st.warning("⚠️ No se detectó el archivo Configuracion.xlsx.")

# Mostrar conteos si ya hay resultado (persiste aunque se haga clic en descargar)
if st.session_state.get('resultado_listo'):
    st.markdown("---")
    st.subheader("📊 Resumen")
    col1, col2, col3 = st.columns(3)
    col1.metric("🗞️ Gráficas (Diario, Online, Revistas)", st.session_state['grafica_count'])
    col2.metric("🎬 AV (AM, FM, Aire, Cable)", st.session_state['av_count'])
    col3.metric("Total de filas", st.session_state['total_count'])
    st.download_button(
        label="📥 Descargar Resultado",
        data=st.session_state['excel_data'],
        file_name=st.session_state['filename'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.button(
    "▶️ Iniciar Proceso",
    disabled=not (dossier_file and config_file),
    type="primary"
):
    run_process(dossier_file, config_file)
