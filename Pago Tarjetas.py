import datetime
import io
import os
import socket
import subprocess
import sys
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Pago Tarjetas (Ganancia USD)",
    page_icon="💳",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------
# ESTILOS CSS PERSONALIZADOS (OPTIMIZADO PARA MÓVIL)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0E1117 !important;
            color: #FAFAFA !important;
        }
        [data-testid="stHeader"] { background-color: rgba(0, 0, 0, 0) !important; }
        .block-container { padding: 1rem 0.5rem 2rem 0.5rem; max-width: 740px; }

        /* Estilos del botón refrescar */
        div[data-testid="stButton"] > button {
            background-color: #1E222B !important;
            border: 1.5px solid #107C41 !important;
            border-radius: 8px !important;
            padding: 0.4rem 0.8rem !important;
            width: 100% !important; 
        }
        div[data-testid="stButton"] > button * {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }
        div[data-testid="stButton"] > button:hover {
            background-color: #107C41 !important;
            border-color: #107C41 !important;
        }

        /* Estilo para el Radio Button de Filtro General */
        div[data-testid="stRadio"] label p {
            color: #FFFFFF !important;
            font-size: 0.92rem !important;
            font-weight: 700 !important;
        }
        div[data-testid="stRadio"] div[role="radiogroup"] {
            gap: 10px !important;
        }

        /* Contenedores de KPI */
        .kpi-container {
            background-color: #1E222B;
            border: 1.5px solid #2D323E;
            border-radius: 12px;
            padding: 14px 12px;
            margin-top: 15px;
            margin-bottom: 15px;
        }
        .kpi-title {
            color: #00D1B2 !important;
            font-size: 1.1rem;
            font-weight: 800;
            margin: 0 0 10px 0;
            border-bottom: 1px solid #2D323E;
            padding-bottom: 6px;
        }
        .kpi-subtitle {
            color: #94A3B8;
            font-size: 0.85rem;
            font-weight: 800;
            text-transform: uppercase;
            margin: 12px 0 8px 0;
            letter-spacing: 0.5px;
        }
        
        .kpi-grid {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 8px;
            text-align: center;
        }

        .kpi-card {
            background-color: #14171E;
            border: 1px solid #252A36;
            padding: 8px 6px;
            border-radius: 8px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            flex: 1 1 calc(33.333% - 10px);
            min-width: 95px;
            max-width: 210px;
            box-sizing: border-box;
        }
        .kpi-card-highlight {
            background: linear-gradient(135deg, #132A24 0%, #14171E 100%);
            border: 1.5px solid #00E676;
            padding: 8px 6px;
            border-radius: 8px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            flex: 1 1 calc(33.333% - 10px);
            min-width: 95px;
            max-width: 210px;
            box-sizing: border-box;
        }

        @media (max-width: 480px) {
            .kpi-card, .kpi-card-highlight {
                flex: 1 1 calc(33.333% - 6px);
                min-width: 88px;
            }
        }

        .kpi-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            color: #94A3B8;
            font-weight: 700;
            margin-bottom: 4px;
            line-height: 1.15;
        }
        
        .kpi-val-corte { font-size: 0.98rem; font-weight: 700; color: #38BDF8; }
        .kpi-val-dias { font-size: 0.98rem; font-weight: 700; color: #FBBF24; }
        .kpi-val-pagos { font-size: 0.98rem; font-weight: 700; color: #C084FC; }
        .kpi-val-menor { font-size: 0.98rem; font-weight: 700; color: #A7F3D0; }
        .kpi-val-mayor { font-size: 0.98rem; font-weight: 700; color: #F43F5E; }
        .kpi-val-total { font-size: 1.02rem; font-weight: 800; color: #00E676; }
        .kpi-val-real { font-size: 0.98rem; font-weight: 700; color: #00E676; }
        .kpi-val-teo { font-size: 0.98rem; font-weight: 700; color: #38BDF8; }
        .kpi-val-dif { font-size: 0.98rem; font-weight: 700; color: #F59E0B; }
        .kpi-val-gasto { font-size: 0.98rem; font-weight: 700; color: #FB7185; }
        .kpi-val-gan-total { font-size: 1.1rem; font-weight: 900; color: #00E676; }

        .kpi-date-sub {
            font-size: 0.76rem;
            color: #CBD5E1;
            margin-top: 3px;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# HELPER: LECTURA DE VALORES
# ---------------------------------------------------------
def get_raw_str(val):
    if pd.isnull(val):
        return "-"
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if s in ["", "-", "None", "NaT"]:
        return "-"
    if " 00:00:00" in s:
        s = s.replace(" 00:00:00", "")
    return s


# ---------------------------------------------------------
# HELPER: TABLA TASAS AUSTRO FUTURO (TÍTULOS EN 2 LÍNEAS)
# ---------------------------------------------------------
def render_excel_table(df):
    html = '<div style="overflow-x: auto; border-radius: 10px; border: 1.5px solid #107C41; margin-top: 10px; margin-bottom: 20px; box-shadow: 0 4px 16px rgba(0, 0, 0, 0.45);">'
    html += '<table style="width:100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif; font-size: 0.92rem; color: #FAFAFA;">'
    html += '<thead><tr style="background: linear-gradient(135deg, #107C41 0%, #0D5C30 100%); color: #FFFFFF;">'

    for col in df.columns:
        align = "right" if col in ["Tasa Vigente", "Tasa"] else "left"
        col_str = str(col).strip()
        col_title = col_str.replace(" ", "<br>", 1) if " " in col_str else col_str
        html += f'<th style="padding: 10px 14px; border-bottom: 2px solid #1B9E52; text-align: {align}; font-weight: 700; line-height: 1.25; vertical-align: middle;">{col_title}</th>'

    html += '</tr></thead><tbody>'

    for idx, row in df.iterrows():
        row_bg = "#1A1D24" if idx % 2 == 0 else "#222733"
        html += f'<tr style="background-color: {row_bg}; color: #FAFAFA;">'
        for col in df.columns:
            align = "right" if col in ["Tasa Vigente", "Tasa"] else "left"
            style_extra = " font-weight: bold; color: #00E676;" if col == "Tasa Vigente" else ""
            html += f'<td style="padding: 10px 14px; border-bottom: 1px solid #2A323D; text-align: {align};{style_extra}">{row[col]}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html


# ---------------------------------------------------------
# HELPER: TABLA RESUMEN DE RESULTADOS (2D STICKY OPTIMIZADA)
# ---------------------------------------------------------
def render_resultados_table(df):
    if df is None or df.empty:
        return ""
    
    css = """<style>
/* Contenedor principal con scroll 2D */
.tbl-res-wrapper {
    max-height: 480px;
    width: 100%;
    overflow-y: auto;
    overflow-x: auto;
    border: 1.8px solid #107C41;
    border-radius: 12px;
    background-color: #14171E;
    box-shadow: 0 6px 20px rgba(0,0,0,0.6);
    margin-top: 10px;
    margin-bottom: 25px;
    display: block;
    position: relative;
}

/* Configuración estricta de la tabla */
.tbl-res-sticky {
    width: 100%;
    border-collapse: separate !important;
    border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    margin: 0;
}

/* Encabezados (TH) */
.tbl-res-sticky th {
    position: sticky !important;
    top: 0 !important;
    z-index: 20 !important;
    background-color: #0D5C30 !important;
    color: #FFFFFF !important;
    border-bottom: 2.5px solid #1B9E52 !important;
    padding: 10px 4px !important;
    text-align: center !important;
    font-weight: 800 !important;
    font-size: 0.82rem !important;
    line-height: 1.25 !important;
    vertical-align: middle !important;
    white-space: nowrap !important;
    min-width: 90px;
}

/* Esquina superior izquierda (TH:first-child) */
.tbl-res-sticky th:first-child {
    position: sticky !important;
    top: 0 !important;
    left: 0 !important;
    z-index: 50 !important;
    text-align: left !important;
    width: 140px !important;
    min-width: 140px !important;
    max-width: 140px !important;
    box-shadow: 2px 0 5px rgba(0,0,0,0.3) !important;
}

/* Celdas de datos normales (TD) */
.tbl-res-sticky td {
    position: static !important;
    z-index: auto !important;
    padding: 9px 5px !important;
    vertical-align: middle !important;
    text-align: center !important;
    white-space: nowrap !important;
    font-family: monospace, sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
}

/* Columna 1 fija (.lbl-sticky-col) */
.tbl-res-sticky .lbl-sticky-col {
    position: sticky !important;
    left: 0 !important;
    z-index: 30 !important;
    font-weight: 700 !important;
    color: #FBBF24 !important;
    font-size: 0.80rem !important;
    line-height: 1.25 !important;
    text-align: left !important;
    width: 140px !important;
    min-width: 140px !important;
    max-width: 140px !important;
    box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important;
    background-clip: padding-box !important;
    white-space: normal !important;
    word-wrap: break-word !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
}

/* Clases CSS de colores por fila (sin inline-styles en fondo) */
.row-even td { background-color: #181C24 !important; border-top: 1px solid #262C38 !important; }
.row-even .lbl-sticky-col { background-color: #181C24 !important; }

.row-odd td { background-color: #1E222B !important; border-top: 1px solid #262C38 !important; }
.row-odd .lbl-sticky-col { background-color: #1E222B !important; }

.row-total-pagado td { background-color: #1E2A38 !important; border-top: 1.5px solid #38BDF8 !important; border-bottom: 1.5px solid #38BDF8 !important; }
.row-total-pagado .lbl-sticky-col { background-color: #1E2A38 !important; }

.row-neta td { background-color: #174D30 !important; border-top: 2px solid #00E676 !important; border-bottom: 2px solid #00E676 !important; }
.row-neta .lbl-sticky-col { background-color: #174D30 !important; }
</style>"""

    html = f'{css}<div class="tbl-res-wrapper"><table class="tbl-res-sticky"><thead><tr>'
    for idx, col in enumerate(df.columns):
        if idx == 0:
            col_title = ""
        else:
            col_str = str(col).strip()
            col_title = col_str.replace(" ", "<br>", 1) if " " in col_str else col_str

        html += f'<th>{col_title}</th>'
    html += '</tr></thead><tbody>'

    first_col = df.columns[0]

    for idx, row in df.iterrows():
        row_title_raw = str(row[first_col]).strip()
        is_neta = "ganancia total neta" in row_title_raw.lower()
        is_total_pagado = "valor total pagado" in row_title_raw.lower()

        if is_neta:
            row_class = "row-neta"
        elif is_total_pagado:
            row_class = "row-total-pagado"
        else:
            row_class = "row-even" if idx % 2 == 0 else "row-odd"

        html += f'<tr class="{row_class}">'
        
        for c_idx, col in enumerate(df.columns):
            val = row[col]

            if c_idx == 0:
                html += f'<td class="lbl-sticky-col">{val}</td>'
            else:
                if pd.isnull(val) or str(val).strip() in ["", "-", "None", "nan"]:
                    val_str = "-"
                    color_style = "color: #94A3B8;"
                elif isinstance(val, (int, float)):
                    if "años" in row_title_raw.lower() or "meses" in row_title_raw.lower():
                        val_str = f"{val:.1f}"
                        color_style = "color: #E2E8F0;"
                    elif "días" in row_title_raw.lower() or "numero" in row_title_raw.lower() or "número" in row_title_raw.lower():
                        val_str = f"{int(val)}"
                        color_style = "color: #E2E8F0;"
                    else:
                        val_str = f"${val:.2f}"
                        if val < 0:
                            color_style = "color: #FF5252; font-weight: 800;"
                        elif val > 0:
                            color_style = "color: #00E676; font-weight: 700;"
                        else:
                            color_style = "color: #94A3B8;"
                else:
                    val_str = str(val)
                    color_style = "color: #FAFAFA;"

                if is_neta:
                    color_style = "color: #00E676; font-weight: 900; font-size: 1.05rem;"
                elif is_total_pagado:
                    color_style = "color: #38BDF8; font-weight: 800; font-size: 1.00rem;"

                html += f'<td style="{color_style}">{val_str}</td>'
        
        html += '</tr>'

    html += '</tbody></table></div>'
    return html


# ---------------------------------------------------------
# HELPER: TABLA DETALLE DE PAGOS POR TARJETA (TÍTULOS EN 2 LÍNEAS)
# ---------------------------------------------------------
def render_payment_table(df, title="📋 Detalle de Pagos Realizados"):
    if df is None or df.empty:
        return ""
    
    real_vals = pd.to_numeric(df["Ganancia Real"], errors='coerce')
    max_real = real_vals.max() if not real_vals.empty else None
    min_real = real_vals.min() if not real_vals.empty else None
    has_highlights = (max_real is not None and min_real is not None and max_real != min_real)

    html = f'<div style="margin-top: 5px; margin-bottom: 25px;">'
    html += f'<h5 style="color: #00D1B2 !important; margin-bottom: 8px; font-size: 0.95rem; font-weight: 700;">{title}</h5>'
    
    html += '<div style="overflow: auto; max-height: 480px; border-radius: 10px; border: 1.5px solid #107C41; box-shadow: 0 4px 16px rgba(0, 0, 0, 0.45);">'
    html += '<table style="width:100%; border-collapse: separate; border-spacing: 0; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif; font-size: 0.85rem; color: #FAFAFA;">'
    
    html += '<thead><tr>'
    
    header_map = {
        "Fecha de Pago": "Fecha<br>de Pago",
        "Valor Pagado": "Valor<br>Pagado",
        "Días Financ.": "Días<br>Financ.",
        "Tasa Aplicada": "Tasa<br>Aplicada",
        "Ganancia Real": "Ganancia<br>Real",
        "Ganancia Teórica": "Ganancia<br>Teórica",
        "Maxidólares": "Maxi<br>dólares"
    }

    for col in df.columns:
        col_str = str(col).strip()
        col_title = header_map.get(col_str, col_str.replace(" ", "<br>", 1) if " " in col_str else col_str)

        html += f'<th style="position: sticky; top: 0; z-index: 20; background: #0D5C30; padding: 10px 4px; border-bottom: 2px solid #1B9E52; text-align: center; font-weight: 800; font-size: 0.80rem; line-height: 1.25; vertical-align: middle; color: #FFFFFF;">{col_title}</th>'
    html += '</tr></thead><tbody>'

    totals = {col: 0.0 for col in df.columns}
    count = 0

    curr_cols = ["Valor Pagado", "Ganancia Real", "Ganancia Teórica", "Maxidólares"]

    for idx, row in df.iterrows():
        v_real = pd.to_numeric(row["Ganancia Real"], errors='coerce')
        is_max = has_highlights and pd.notnull(v_real) and (v_real == max_real)
        is_min = has_highlights and pd.notnull(v_real) and (v_real == min_real)

        if is_max:
            row_bg = "#524300"
            row_color = "#FFF176"
            font_w = "bold"
            border_left = "4px solid #FFEE58"
        elif is_min:
            row_bg = "#4A151B"
            row_color = "#FF8A80"
            font_w = "bold"
            border_left = "4px solid #EF5350"
        else:
            row_bg = "#1A1D24" if idx % 2 == 0 else "#222733"
            row_color = "#FAFAFA"
            font_w = "normal"
            border_left = "none"

        html += f'<tr style="background-color: {row_bg}; color: {row_color}; font-weight: {font_w}; border-left: {border_left};">'
        
        for col in df.columns:
            val = row[col]
            
            if col == "Fecha de Pago":
                val_str = get_raw_str(val)
            elif col in curr_cols:
                v_num = pd.to_numeric(val, errors='coerce')
                if pd.notnull(v_num):
                    val_str = f"${v_num:.2f}"
                    totals[col] += v_num
                else:
                    val_str = str(val)
            elif col == "Días Financ.":
                v_num = pd.to_numeric(val, errors='coerce')
                val_str = f"{int(v_num)}" if pd.notnull(v_num) else str(val)
            elif col == "Tasa Aplicada":
                v_num = pd.to_numeric(val, errors='coerce')
                if pd.notnull(v_num):
                    rate = v_num * 100 if v_num <= 1.0 else v_num
                    val_str = f"{rate:.2f}%"
                else:
                    val_str = str(val)
            else:
                val_str = str(val)

            if is_max:
                style_extra = " color: #FFF176; font-weight: 700;"
            elif is_min:
                style_extra = " color: #FF8A80; font-weight: 700;"
            else:
                style_extra = ""
                if col in ["Ganancia Real", "Ganancia Teórica", "Maxidólares"]:
                    style_extra = " color: #00E676; font-weight: 600;"
                elif col == "Valor Pagado":
                    style_extra = " font-weight: 600;"

            html += f'<td style="padding: 8px 4px; border-bottom: 1px solid #2A323D; text-align: center; vertical-align: middle;{style_extra}">{val_str}</td>'
        
        html += '</tr>'
        count += 1

    if count > 0:
        html += '<tr style="background-color: #133322; font-weight: bold; border-top: 2px solid #107C41; color: #FFFFFF;">'
        for col in df.columns:
            if col == "Fecha de Pago":
                val_str = "TOTAL"
            elif col in curr_cols:
                val_str = f"${totals[col]:.2f}"
            else:
                val_str = "-"
            
            style_extra = " color: #00E676;" if col in curr_cols else ""
            html += f'<td style="padding: 10px 4px; text-align: center; vertical-align: middle;{style_extra}">{val_str}</td>'
        html += '</tr>'

    html += '</tbody></table></div></div>'
    return html


# ---------------------------------------------------------
# HELPER: EXTRAER KPIs DE CADA HOJA
# ---------------------------------------------------------
def extract_sheet_kpis(excel_bytes, sheet_name):
    try:
        df_k = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None)
        s_upper = sheet_name.upper().strip()

        if "PRODUBANCO" in s_upper:
            val_col = 11            # Columna L
            date_col = 12           # Columna M
        else:
            val_col = 10            # Columna K
            date_col = 11           # Columna L

        def get_val(r, c=val_col, default=0.0):
            if df_k.shape[1] > c and df_k.shape[0] > r:
                v = df_k.iloc[r, c]
                return v if pd.notnull(v) else default
            return default

        corte_val = get_val(0, val_col, "-")
        dias_val = get_val(3, val_col, "-")
        total_pagos_val = get_val(4, val_col, 0)

        menor_val = get_val(5, val_col, 0.0)
        menor_fecha = get_val(5, date_col, "-")

        mayor_val = get_val(6, val_col, 0.0)
        mayor_fecha = get_val(6, date_col, "-")

        total_val = get_val(8, val_col, 0.0)

        ganancia_real_val = get_val(10, val_col, 0.0)
        ganancia_teorica_val = get_val(11, val_col, 0.0)
        dif_teorica_real_val = get_val(12, val_col, 0.0)

        gastos_impresion_val = get_val(14, val_col, 0.0)
        gastos_transferencias_val = get_val(16, val_col, 0.0)

        maxidolares_val = None
        millas_val = None
        anualidad_val = None

        if "PRODUBANCO" in s_upper:
            maxidolares_val = get_val(20, val_col, 0.0)
            anualidad_val = get_val(21, val_col, 0.0)
            ganancia_total_val = get_val(23, val_col, 0.0)
        elif "DINNER" in s_upper or "DINER" in s_upper:
            millas_val = get_val(18, val_col, 0.0)
            anualidad_val = get_val(19, val_col, 0.0)
            
            g_tot = None
            for r in range(20, min(30, df_k.shape[0])):
                row_txt = " ".join([str(df_k.iloc[r, c]) for c in range(df_k.shape[1]) if pd.notnull(df_k.iloc[r, c])]).upper()
                if "GANANCIA TOTAL" in row_txt or "GANACIA TOTAL" in row_txt:
                    g_tot = df_k.iloc[r, val_col]
                    break
            if g_tot is None or not pd.notnull(g_tot):
                g_tot = get_val(23, val_col, get_val(21, val_col, 0.0))
            ganancia_total_val = g_tot
        else: # American Express
            ganancia_total_val = get_val(21, val_col, 0.0)

        def fmt_curr(val):
            if isinstance(val, (int, float)) and pd.notnull(val):
                return f"${float(val):.2f}"
            return str(val)

        def fmt_num(val):
            if isinstance(val, (int, float)) and pd.notnull(val):
                return f"{int(val)}"
            return str(val)

        def fmt_miles(val):
            if isinstance(val, (int, float)) and pd.notnull(val):
                return f"{val:.0f}"
            return str(val)

        corte_str = str(corte_val).strip()
        if "FIN DE MES" in corte_str.upper():
            corte_display = "Fin de Mes"
        elif isinstance(corte_val, (int, float)) and pd.notnull(corte_val):
            try:
                corte_display = str(int(corte_val))
            except Exception:
                corte_display = corte_str
        else:
            corte_display = corte_str if corte_str else "-"

        return {
            "corte": corte_display,
            "dias": fmt_num(dias_val),
            "total_pagos": fmt_num(total_pagos_val),
            "menor": fmt_curr(menor_val),
            "menor_fecha": get_raw_str(menor_fecha),
            "mayor": fmt_curr(mayor_val),
            "mayor_fecha": get_raw_str(mayor_fecha),
            "total": fmt_curr(total_val),
            "ganancia_real": fmt_curr(ganancia_real_val),
            "ganancia_teorica": fmt_curr(ganancia_teorica_val),
            "dif_teorica_real": fmt_curr(dif_teorica_real_val),
            "gastos_impresion": fmt_curr(gastos_impresion_val),
            "gastos_transferencias": fmt_curr(gastos_transferencias_val),
            "maxidolares": fmt_curr(maxidolares_val) if maxidolares_val is not None else None,
            "millas": fmt_miles(millas_val) if millas_val is not None else None,
            "anualidad": fmt_curr(anualidad_val) if anualidad_val is not None else None,
            "ganancia_total": fmt_curr(ganancia_total_val)
        }
    except Exception:
        return {
            "corte": "-", "dias": "-", "total_pagos": "-",
            "menor": "$0.00", "menor_fecha": "-", "mayor": "$0.00", "mayor_fecha": "-",
            "total": "$0.00", "ganancia_real": "$0.00", "ganancia_teorica": "$0.00",
            "dif_teorica_real": "$0.00", "gastos_impresion": "$0.00",
            "gastos_transferencias": "$0.00", "maxidolares": None, "millas": None, "anualidad": None,
            "ganancia_total": "$0.00"
        }


def extract_sheet_table(excel_bytes, sheet_name, col_indices, col_names):
    try:
        df_raw = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None)
        
        valid_cols = [c for c in col_indices if c < df_raw.shape[1]]
        if len(valid_cols) < len(col_indices):
            return pd.DataFrame()
            
        df_sub = df_raw.iloc[:, valid_cols].copy()
        df_sub.columns = col_names
        
        rows = []
        for idx, row in df_sub.iterrows():
            val_date = row[col_names[0]]
            if pd.notnull(val_date):
                val_str = get_raw_str(val_date)
                if val_str not in ["-", ""] and not val_str.lower().startswith("fecha") and not val_str.lower().startswith("total"):
                    rows.append(row)
                    
        if not rows:
            return pd.DataFrame()
            
        return pd.DataFrame(rows).reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


def extract_resultados_sheet(excel_bytes):
    """Extrae datos vivos de la hoja Resultados y formatea los nombres en 2 líneas."""
    try:
        excel_file = pd.ExcelFile(excel_bytes)
        res_sheet = next((s for s in excel_file.sheet_names if "resultados" in s.strip().lower()), None)
        
        if not res_sheet:
            return pd.DataFrame()

        df_res = pd.read_excel(excel_bytes, sheet_name=res_sheet)
        
        cols = []
        for idx, c in enumerate(df_res.columns):
            if idx == 0 or str(c).startswith("Unnamed"):
                cols.append("")
            else:
                cols.append(str(c).strip())
        df_res.columns = cols
        
        col_first = df_res.columns[0]
        
        # Eliminar la antigua fila A13
        df_res = df_res[~df_res[col_first].astype(str).str.contains("Ganancia Real por dinero|Ganacia Real Neta", case=False, na=False)].copy()
        
        # MAPEO CON FORMATO EN 2 LÍNEAS DE TEXTO (<br>)
        rename_map = {
            "Anos Pagando tarjetas": "Años Pagando<br>Tarjetas",
            "Meses Pagando tarjetas": "Meses Pagando<br>Tarjetas",
            "Dias Pagando tarjetas": "Días Pagando<br>Tarjetas",
            "Numero Total de Pagos": "Número Total<br>de Pagos",
            "Valor Promedio Pagado por mes": "Valor Promedio<br>Pagado por Mes",
            "Valor Total Pagado en todos los periodos": "Valor Total Pagado<br>Todos los Periodos",
            "Ganancia dinero mantenido en cuenta de ahorros hasta pago tarjeta (Ganancia por dinero en cuenta)": "Ganancia Real<br>en Cuenta",
            "Ganancia dinero si hubiera dejado el dinero en cuenta de ahorros los 50 dias antes de pagar tarjeta (Ganacia Teorica)": "Ganancia<br>Teórica",
            "Diferencia dinero en cuenta 50 dias - dinero en cuenta los dias que en verdad estuvo dinero en la cuenta ahorros": "Diferencia<br>(Real - Teórica)",
            "Gastos Tarjetas por impresion tarjeta o sacar nueva tarjeta por robo (Costo Tarjetas)": "Impresión<br>Tarjetas",
            "Gastos Transferencias Bancarias para pagar tarjeta cada periodo (Costos Transferencias)": "Gastos<br>Transferencias",
            "Ganancia Maxidolares Millas": "Maxidólares /<br>Millas",
            "Costo Anualidad Tarjeta": "Costo Anualidad<br>Tarjeta",
            "Ganancia Total Neta": "GANANCIA TOTAL<br>NETA"
        }
        
        df_res[col_first] = df_res[col_first].astype(str).str.strip().map(lambda x: rename_map.get(x, x))
        return df_res
    except Exception:
        return pd.DataFrame()


# ---------------------------------------------------------
# CONEXIÓN GOOGLE DRIVE & CARGA DE DATOS
# ---------------------------------------------------------
EXCEL_URL = "https://drive.google.com/uc?export=download&id=1FaRzJGrucpF0GuGuPfs-K9NVA-PIXhZt"

@st.cache_data(ttl=300)
def load_all_excel_sheets():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    response = requests.get(EXCEL_URL, headers=headers, timeout=30)
    response.raise_for_status()
    excel_bytes = io.BytesIO(response.content)

    # 1. Cargar Hoja 'Cuenta Austro Futuro'
    df_tasas = pd.read_excel(excel_bytes, sheet_name="Cuenta Austro Futuro")
    df_tasas.columns = [str(c).strip() for c in df_tasas.columns]
    df_tasas["Fecha Vigencia"] = df_tasas["Fecha"].apply(get_raw_str)
    df_tasas["Tasa Vigente"] = df_tasas["Tasa"].apply(lambda x: f"{x * 100:.2f}%" if pd.notnull(x) and isinstance(x, (int, float)) else str(x))
    df_tasas_clean = df_tasas[["Fecha Vigencia", "Tasa Vigente"]].copy()

    # 2. Cargar Hoja 'Resultados'
    df_resultados = extract_resultados_sheet(excel_bytes)

    # 3. Hojas de Tarjetas
    card_sheets = {
        "American Express": "American Express",
        "Produbanco Mastercard": "Produbanco Mastercard",
        "Dinners Visa": "Dinners Visa ",
        "Dinners Mastercard": "Dinners Mastercard"
    }

    dict_kpis = {}
    dict_tables = {}

    for card_label, sheet_name in card_sheets.items():
        dict_kpis[card_label] = extract_sheet_kpis(excel_bytes, sheet_name)
        
        if card_label in ["American Express", "Dinners Visa", "Dinners Mastercard"]:
            dict_tables[card_label] = extract_sheet_table(
                excel_bytes,
                sheet_name,
                col_indices=[0, 1, 5, 6, 7, 8],
                col_names=["Fecha de Pago", "Valor Pagado", "Días Financ.", "Tasa Aplicada", "Ganancia Real", "Ganancia Teórica"]
            )
        elif card_label == "Produbanco Mastercard":
            dict_tables[card_label] = extract_sheet_table(
                excel_bytes,
                sheet_name,
                col_indices=[0, 1, 5, 6, 7, 8, 9],
                col_names=["Fecha de Pago", "Valor Pagado", "Días Financ.", "Tasa Aplicada", "Ganancia Real", "Ganancia Teórica", "Maxidólares"]
            )

    return df_tasas_clean, df_resultados, dict_kpis, dict_tables

try:
    with st.spinner("Conectando con Google Drive..."):
        df_tasas, df_resultados, dict_kpis, dict_tables = load_all_excel_sheets()
except Exception as e:
    st.error(f"Error al cargar la información desde Google Drive: {e}")
    st.stop()

# ---------------------------------------------------------
# CABECERA & BOTÓN REFRESCAR
# ---------------------------------------------------------
col_title, col_btn = st.columns([2.8, 1.2], vertical_alignment="center")

with col_title:
    st.markdown("<h3 style='margin: 0; color: #FFFFFF !important;'>💳 Pago Tarjetas (Ganancia USD)</h3>", unsafe_allow_html=True)

with col_btn:
    if st.button("🔄 Refrescar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ---------------------------------------------------------
# 1. TABLA DINÁMICA: TASAS AUSTRO FUTURO
# ---------------------------------------------------------
st.markdown("<h4 style='color: #00D1B2 !important; margin-top: 12px; margin-bottom: 2px;'>📈 Tasas Austro Futuro</h4>", unsafe_allow_html=True)

if not df_tasas.empty:
    st.markdown(render_excel_table(df_tasas), unsafe_allow_html=True)
else:
    st.info("No se encontraron registros en Austro Futuro.")

# ---------------------------------------------------------
# 2. GRÁFICOS DINÁMICOS TIPO PASTEL
# ---------------------------------------------------------
if not df_resultados.empty:
    col_first = df_resultados.columns[0]
    card_cols = [c for c in df_resultados.columns[1:] if "total" not in str(c).lower()]
    
    row_pagado = df_resultados[df_resultados[col_first].astype(str).str.contains("Valor Total Pagado", case=False, na=False)]
    row_ganancia = df_resultados[df_resultados[col_first].astype(str).str.contains("Ganancia Total", case=False, na=False)]

    if not row_pagado.empty and not row_ganancia.empty and len(card_cols) > 0:
        short_names = {
            "American Express": "Amex",
            "American<br>Express": "Amex",
            "Produbanco Mastercard": "Produbanco MC",
            "Produbanco<br>Mastercard": "Produbanco MC",
            "Dinners Visa": "Diners Visa",
            "Dinners<br>Visa": "Diners Visa",
            "Dinners Mastercard": "Diners MC",
            "Dinners<br>Mastercard": "Diners MC"
        }

        color_map = {
            "Amex": "#1E40AF",          # Azul Cobalto Intenso
            "Produbanco MC": "#059669", # Verde Esmeralda
            "Diners Visa": "#6B21A8",    # Púrpura Profundo
            "Diners MC": "#A855F7"       # Morado Suave
        }

        data_pagado = []
        for card in card_cols:
            clean_card = card.replace("<br>", " ").strip()
            short_label = short_names.get(clean_card, clean_card)
            val = pd.to_numeric(row_pagado[card].values[0], errors='coerce')
            data_pagado.append({"Tarjeta": short_label, "Valor Pagado": val if pd.notnull(val) else 0.0})
        df_pie_pagado = pd.DataFrame(data_pagado)

        data_ganancia = []
        for card in card_cols:
            clean_card = card.replace("<br>", " ").strip()
            short_label = short_names.get(clean_card, clean_card)
            val = pd.to_numeric(row_ganancia[card].values[0], errors='coerce')
            data_ganancia.append({"Tarjeta": short_label, "Ganancia Neta": val if pd.notnull(val) else 0.0})
        df_pie_ganancia = pd.DataFrame(data_ganancia)

        col_g1, col_g2 = st.columns(2)

        clean_font = dict(color='#FFFFFF', size=11, family='Segoe UI, Roboto, Helvetica, Arial, sans-serif')

        with col_g1:
            fig_pagado = px.pie(
                df_pie_pagado, 
                values='Valor Pagado', 
                names='Tarjeta', 
                title='💳 % Valor Pagado por Tarjeta',
                hole=0.38,
                color='Tarjeta',
                color_discrete_map=color_map
            )
            fig_pagado.update_traces(
                textinfo='label+percent',
                textposition='inside',
                insidetextfont=clean_font,
                hovertemplate='<b>%{label}</b><br>Valor: $%{value:.2f}<br>Porcentaje: %{percent}'
            )
            fig_pagado.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#FFFFFF", size=12),
                showlegend=False,
                margin=dict(l=5, r=5, t=35, b=5),
                height=290
            )
            st.plotly_chart(fig_pagado, use_container_width=True, config={'displayModeBar': False})

        with col_g2:
            df_pie_ganancia['Ganancia_Plot'] = df_pie_ganancia['Ganancia Neta'].apply(lambda x: max(0.0, x))
            
            fig_ganancia = px.pie(
                df_pie_ganancia, 
                values='Ganancia_Plot', 
                names='Tarjeta', 
                title='💰 % Ganancia Neta por Tarjeta',
                hole=0.38,
                color='Tarjeta',
                color_discrete_map=color_map
            )
            fig_ganancia.update_traces(
                textinfo='label+percent',
                textposition='inside',
                insidetextfont=clean_font,
                hovertemplate='<b>%{label}</b><br>Ganancia: $%{value:.2f}<br>Porcentaje: %{percent}'
            )
            fig_ganancia.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#FFFFFF", size=12),
                showlegend=False,
                margin=dict(l=5, r=5, t=35, b=5),
                height=290
            )
            st.plotly_chart(fig_ganancia, use_container_width=True, config={'displayModeBar': False})

# ---------------------------------------------------------
# 3. TABLA RESUMEN: RESULTADOS PAGOS TODAS LAS TARJETAS
# ---------------------------------------------------------
st.markdown("<h4 style='color: #00D1B2 !important; margin-top: 10px; margin-bottom: 2px;'>📊 Resultados Pagos Todas Las Tarjetas</h4>", unsafe_allow_html=True)

if not df_resultados.empty:
    st.markdown(render_resultados_table(df_resultados), unsafe_allow_html=True)
else:
    st.info("No se encontraron resultados consolidados en la hoja Resultados.")

st.markdown("---")

# ---------------------------------------------------------
# 4. FILTRO GENERAL POR TARJETA
# ---------------------------------------------------------
st.markdown("<h4 style='color: #FFFFFF !important; margin-bottom: 6px;'>🔍 Filtro General</h4>", unsafe_allow_html=True)

opciones_tarjetas = ["Todas", "American Express", "Produbanco Mastercard", "Dinners Visa", "Dinners Mastercard"]

tarjeta_seleccionada = st.radio(
    label="Seleccionar Tarjeta a Analizar:",
    options=opciones_tarjetas,
    index=0,
    horizontal=True
)

# ---------------------------------------------------------
# 5. DESPLIEGUE DE KPIs Y TABLAS INDIVIDUALES (ETIQUETAS EN 2 LÍNEAS)
# ---------------------------------------------------------
def render_kpi_block(card_label, kpi_data, table_data=None):
    rewards_html = ""
    if kpi_data.get("maxidolares") is not None:
        rewards_html += f'<div class="kpi-card"><div class="kpi-label">Maxi<br>dólares</div><div class="kpi-val-real">{kpi_data["maxidolares"]}</div></div>'
    if kpi_data.get("millas") is not None:
        rewards_html += f'<div class="kpi-card"><div class="kpi-label">Millas<br>Acumuladas</div><div class="kpi-val-real">{kpi_data["millas"]}</div></div>'
    if kpi_data.get("anualidad") is not None:
        rewards_html += f'<div class="kpi-card"><div class="kpi-label">Anualidad<br>Tarjeta</div><div class="kpi-val-gasto">{kpi_data["anualidad"]}</div></div>'

    html_str = f"""<div class="kpi-container">
<div class="kpi-title">📊 Resumen de Métricas - {card_label}</div>
<div class="kpi-subtitle">💳 Consumos y Operativa</div>
<div class="kpi-grid">
<div class="kpi-card"><div class="kpi-label">Día de<br>Corte</div><div class="kpi-val-corte">{kpi_data['corte']}</div></div>
<div class="kpi-card"><div class="kpi-label">Días<br>Financ.</div><div class="kpi-val-dias">{kpi_data['dias']}</div></div>
<div class="kpi-card"><div class="kpi-label">Total<br>Pagos</div><div class="kpi-val-pagos">{kpi_data['total_pagos']}</div></div>
<div class="kpi-card"><div class="kpi-label">Menor<br>Pago</div><div class="kpi-val-menor">{kpi_data['menor']}</div><div class="kpi-date-sub">📅 {kpi_data['menor_fecha']}</div></div>
<div class="kpi-card"><div class="kpi-label">Mayor<br>Pago</div><div class="kpi-val-mayor">{kpi_data['mayor']}</div><div class="kpi-date-sub">📅 {kpi_data['mayor_fecha']}</div></div>
<div class="kpi-card"><div class="kpi-label">Total<br>Pagado</div><div class="kpi-val-total">{kpi_data['total']}</div></div>
</div>
<div class="kpi-subtitle">💰 Rendimientos y Gastos</div>
<div class="kpi-grid">
<div class="kpi-card"><div class="kpi-label">Ganancia Real<br>en Cuenta</div><div class="kpi-val-real">{kpi_data['ganancia_real']}</div></div>
<div class="kpi-card"><div class="kpi-label">Ganancia<br>Teórica</div><div class="kpi-val-teo">{kpi_data['ganancia_teorica']}</div></div>
<div class="kpi-card"><div class="kpi-label">Diferencia<br>Real - Teórica</div><div class="kpi-val-dif">{kpi_data['dif_teorica_real']}</div></div>
<div class="kpi-card"><div class="kpi-label">Gastos<br>Impresión</div><div class="kpi-val-gasto">{kpi_data['gastos_impresion']}</div></div>
<div class="kpi-card"><div class="kpi-label">Gastos<br>Transfer.</div><div class="kpi-val-gasto">{kpi_data['gastos_transferencias']}</div></div>
{rewards_html}
<div class="kpi-card-highlight"><div class="kpi-label" style="color: #00E676;">Ganancia<br>Total</div><div class="kpi-val-gan-total">{kpi_data['ganancia_total']}</div></div>
</div>
</div>"""
    st.markdown(html_str, unsafe_allow_html=True)

    if table_data is not None and not table_data.empty:
        st.markdown(render_payment_table(table_data, f"📋 Pagos Realizados ({card_label})"), unsafe_allow_html=True)


if tarjeta_seleccionada == "Todas":
    for card_label in opciones_tarjetas[1:]:
        if card_label in dict_kpis:
            render_kpi_block(card_label, dict_kpis[card_label], dict_tables.get(card_label))
else:
    if tarjeta_seleccionada in dict_kpis:
        render_kpi_block(tarjeta_seleccionada, dict_kpis[tarjeta_seleccionada], dict_tables.get(tarjeta_seleccionada))

# ---------------------------------------------------------
# EJECUCIÓN LOCAL OPCIONAL
# ---------------------------------------------------------
if __name__ == "__main__" and not os.environ.get("STREAMLIT_RUNNING"):
    os.environ["STREAMLIT_RUNNING"] = "true"
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    subprocess.run([sys.executable, "-m", "streamlit", "run", sys.argv[0], "--server.address=0.0.0.0"])