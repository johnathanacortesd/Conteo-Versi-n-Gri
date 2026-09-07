"""Lectura y conteo de dossiers / resultados Grill (sin Streamlit)."""
from __future__ import annotations

import html
import re

import numpy as np
import pandas as pd

from excel_io import iter_sheet_records, load_xlsx_workbook, pick_worksheet, read_excel_sheets


def extract_link_from_cell(cell):
    if cell is None:
        return None
    return cell.hyperlink.target if cell.hyperlink and cell.hyperlink.target else None


def _col(df, *names):
    """Primera columna existente; no inventa una serie vacía que pise datos Grill."""
    for name in names:
        if name in df.columns:
            return df[name]
    return pd.Series([np.nan] * len(df), index=df.index)


def convert_html_entities(text):
    if not isinstance(text, str):
        return text
    text = html.unescape(text)
    for entity, char in {
        "&#xF3;": "ó",
        "&#xE1;": "á",
        "&#xE9;": "é",
        "&#xED;": "í",
        "&#xFA;": "ú",
        "&#xF1;": "ñ",
        "&#xDC;": "Ü",
        "&#xFC;": "ü",
        "&#xC1;": "Á",
        "&#xC9;": "É",
        "&#xCD;": "Í",
        "&#xD3;": "Ó",
        "&#xDA;": "Ú",
        "&#xD1;": "Ñ",
        "&#xC7;": "Ç",
        "&#xE7;": "ç",
    }.items():
        text = text.replace(entity, char)
    text = re.sub(r"&#x([0-9A-Fa-f]+);", lambda m: chr(int(m.group(1), 16)), text)
    text = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), text)
    for b, g in {
        "\u201c": '"',
        "\u201d": '"',
        "\u2018": "'",
        "\u2019": "'",
        "Â": "",
        "â": "",
        "€": "",
        "™": "",
    }.items():
        text = text.replace(b, g)
    return text


def clean_text(t):
    return convert_html_entities(t).strip() if isinstance(t, str) else t


def clean_cuerpo(t):
    if not isinstance(t, str) or not t.strip():
        return t
    t = convert_html_entities(t)
    t = re.sub(r"<br\s*/?>", "\n", t, flags=re.IGNORECASE)
    t = re.sub(r"<[^>]+>", "", t)
    return t.strip()


def load_config(src):
    sheets = read_excel_sheets(src)
    if "Regiones" not in sheets or "Internet" not in sheets:
        raise RuntimeError(
            "El archivo de configuración no tiene las hojas «Regiones» e «Internet»."
        )
    rmap = pd.Series(
        sheets["Regiones"].iloc[:, 1].values,
        index=sheets["Regiones"].iloc[:, 0].astype(str).str.lower().str.strip(),
    ).to_dict()
    imap = pd.Series(
        sheets["Internet"].iloc[:, 1].values,
        index=sheets["Internet"].iloc[:, 0].astype(str).str.lower().str.strip(),
    ).to_dict()
    return rmap, imap


def process_dossier(dossier_file, rmap, imap):
    wb = load_xlsx_workbook(dossier_file)
    sheet = pick_worksheet(wb)
    rows = []
    link_cols = [
        "URL Nota AV",
        "URL (Streaming - Imagen)",
        "URL Nota",
        "Link Nota AV",
        "Link (Streaming - Imagen)",
        "Link Nota",
    ]
    for rd, cells in iter_sheet_records(sheet):
        for lc in link_cols:
            if lc in cells:
                ext = extract_link_from_cell(cells[lc])
                if ext:
                    rd[lc] = ext
        rows.append(rd)
    df = pd.DataFrame(rows)
    if df.empty:
        raise RuntimeError(
            "El archivo no tiene filas de datos. Verifique que la hoja tenga encabezados "
            "y notas (no se omiten filas solo porque «Tipo de nota» o «Empresas Consulta» "
            "vengan vacíos)."
        )
    if "Tipo de Medio" not in df.columns:
        raise RuntimeError(
            "No se encontró la columna «Tipo de Medio». "
            "El archivo debe ser un dossier o un resultado Grill válido."
        )
    tmap = {
        "online": "Internet",
        "internet": "Internet",
        "diario": "Prensa",
        "am": "Radio",
        "fm": "Radio",
        "aire": "Televisión",
        "cable": "Televisión",
        "revista": "Revistas",
        "revistas": "Revistas",
    }
    df["Tipo de Medio"] = (
        df["Tipo de Medio"]
        .astype(str)
        .str.lower()
        .str.strip()
        .map(tmap)
        .fillna(df["Tipo de Medio"].astype(str).str.strip())
    )
    is_av = df["Tipo de Medio"].isin(["Radio", "Televisión"])
    is_gr = df["Tipo de Medio"].isin(["Prensa", "Internet", "Revistas"])
    is_in = df["Tipo de Medio"] == "Internet"
    if "Medio" in df.columns:
        df.loc[is_in, "Medio"] = (
            df.loc[is_in, "Medio"]
            .astype(str)
            .str.lower()
            .str.strip()
            .map(imap)
            .fillna(df.loc[is_in, "Medio"])
        )
        df["Región"] = df["Medio"].astype(str).str.lower().str.strip().map(rmap)
    else:
        df["Medio"] = np.nan
        df["Región"] = np.nan
    df["ID Noticia"] = _col(df, "NoticiaId", "ID Noticia")
    df["Fecha"] = pd.to_datetime(_col(df, "Fecha"), dayfirst=True, errors="coerce").dt.normalize()
    df["Hora"] = _col(df, "Hora")
    for c in ["Sección - Programa", "Título", "Autor - Conductor"]:
        df[c] = _col(df, c).astype(str).apply(clean_text)
    df["Nro. Pagina"] = _col(df, "Nro. Pagina")
    df["Dimensión"] = _col(df, "Dimensioncm2", "Dimensión")
    df["Duración - Nro. Caracteres"] = _col(df, "Duración - Nro. Caracteres")
    df.loc[is_av, "Dimensión"] = df.loc[is_av, "Duración - Nro. Caracteres"]
    df.loc[is_av, "Duración - Nro. Caracteres"] = 0
    df["CPE"] = np.where(
        is_av, _col(df, "CPE"), np.where(is_gr, _col(df, "Valor de Nota", "CPE"), np.nan)
    )
    df["Tier"] = _col(df, "Tier")
    df["Audiencia"] = _col(df, "Audiencia")
    df["Tono"] = _col(df, "Tono").astype(str).apply(clean_text)
    df["Tema"] = _col(df, "Tematica", "Tema").astype(str).apply(clean_text)
    df["Temas Generales - Tema"] = _col(df, "Temas Generales - Tema").astype(str).apply(clean_text)
    cuerpo = _col(df, "CuerpoEs", "Resumen - Aclaracion").astype(str).apply(clean_cuerpo)

    def fmt(t):
        if not isinstance(t, str) or not t.strip():
            return t
        ps = [p.strip() for p in t.split("\n") if p.strip()]
        return "\n\n".join(ps) if len(ps) > 1 else t

    df["Resumen - Aclaracion"] = np.where(is_av, cuerpo, cuerpo.apply(fmt))

    url_av = _col(df, "URL Nota AV", "Link Nota AV", "Link Nota").fillna("").astype(str)
    # Dossier crudo: URL (Streaming - Imagen). Resultado Grill: ya trae Link Nota.
    url_str = _col(df, "URL (Streaming - Imagen)", "Link Nota").fillna("").astype(str)

    link_nota_arr = np.where(
        is_av,
        url_av.str.replace(r"\.com\.ar", ".com.co", regex=True),
        np.where(is_gr, url_str, ""),
    )
    df["Link Nota"] = pd.Series(link_nota_arr, index=df.index).replace("", np.nan)

    df["Link (Streaming - Imagen)"] = (
        _col(df, "URL Nota", "Link (Streaming - Imagen)").fillna("").astype(str).replace("", np.nan)
    )
    m_av = _col(df, "Menciones - Empresa").fillna("").astype(str).apply(clean_text)
    m_gr = (
        _col(df, "Empresa rel.", "Menciones - Empresa", "Empresas Consulta")
        .fillna("")
        .astype(str)
        .apply(clean_text)
    )
    df["Menciones - Empresa"] = np.where(is_av, m_av, np.where(is_gr, m_gr, m_av))
    rows_exp = []
    for _, row in df.iterrows():
        menc = [m.strip() for m in str(row["Menciones - Empresa"]).split(";") if m.strip()]
        if not menc:
            rows_exp.append(row.to_dict())
        else:
            for m in menc:
                nr = row.to_dict()
                nr["Menciones - Empresa"] = m
                rows_exp.append(nr)
    df = pd.DataFrame(rows_exp).reset_index(drop=True)
    return (
        df,
        int(df["Tipo de Medio"].isin(["Radio", "Televisión"]).sum()),
        int(df["Tipo de Medio"].isin(["Prensa", "Internet", "Revistas"]).sum()),
    )
