"""Tema visual de la app: sidebar oscuro con la marca, lienzo claro con paneles flotantes,
tarjetas de indicadores, pestañas tipo píldora y tablas con formato contable.

Paleta corporativa en azul rey sobre azul profundo, con un patrón de líneas y nodos muy
sutil (motivo "circuito") que aporta el aire tecnológico sin restarle sobriedad.
"""
import base64
import os

import pandas as pd
import streamlit as st
from pandas.api.types import is_numeric_dtype

from excel_export import LOGO_PATH

# ---------------------------------------------------------------------------
# Paleta corporativa: azul rey como color de marca, con acentos semánticos para
# que un dato en verde o en rojo signifique siempre lo mismo en toda la app.
# ---------------------------------------------------------------------------
AZUL = "#1D4ED8"          # azul rey — color principal
AZUL_OSC = "#16307A"      # azul profundo — fondos y botones
AZUL_CLARO = "#3B82F6"
CIAN = "#22D3EE"          # acento tecnológico del patrón de fondo

VERDE = "#0EA36B"         # positivo: ingresos, conciliados
VERDE_OSC = AZUL          # acento principal (se mantiene el nombre por compatibilidad)
NARANJA = "#E08700"       # ámbar: requiere revisión
ROJO = "#DC2743"          # pendiente / salida
GRIS = "#64748B"          # neutro

# Dorado: se usa SOLO en la pantalla de acceso. En finanzas el oro comunica valor y
# categoría, pero dentro de las tablas competiría con el ámbar de "requiere revisión",
# así que se mantiene fuera del área de trabajo.
ORO = "#D4AF37"
ORO_CLARO = "#F0D98B"
ORO_OSCURO = "#A87F22"

_SIDEBAR_TOP = "#14245C"
_SIDEBAR_BOT = "#0A1233"
_TEXTO = "#152036"
_BORDE = "#E2E7F0"


@st.cache_data(show_spinner=False)
def _logo_b64():
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


# Patrón de líneas + nodos (motivo "circuito") como SVG en línea, muy tenue.
_PATRON_SVG = """
<svg xmlns='http://www.w3.org/2000/svg' width='220' height='220'>
  <g fill='none' stroke='%23ffffff' stroke-opacity='0.05' stroke-width='1.4'>
    <path d='M10 40 L60 40 L80 15'/>
    <path d='M0 120 L50 120 L75 150 L140 150'/>
    <path d='M120 0 L120 40 L170 60 L170 110'/>
    <path d='M180 130 L210 160 L210 210'/>
    <path d='M30 180 L70 180 L95 210'/>
  </g>
  <g fill='%2360A5FA' fill-opacity='0.26'>
    <circle cx='10' cy='40' r='3'/><circle cx='80' cy='15' r='3'/>
    <circle cx='140' cy='150' r='3'/><circle cx='170' cy='110' r='3'/>
  </g>
  <g fill='%2322D3EE' fill-opacity='0.22'>
    <circle cx='0' cy='120' r='3'/><circle cx='120' cy='0' r='3'/>
    <circle cx='210' cy='210' r='3'/><circle cx='95' cy='210' r='3'/>
  </g>
</svg>
""".replace("\n", "").replace("  ", "")


def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    h1, h2, h3, h4 {{ font-family: 'Sora', sans-serif; letter-spacing: -0.015em; }}

    /* Contenedor ancho y compacto: las tablas contables necesitan todo el espacio posible */
    .block-container {{
        padding-top: 2.6rem; padding-bottom: 1.5rem;
        padding-left: 1.6rem; padding-right: 1.6rem;
        max-width: 1680px;
    }}
    /* OJO: el encabezado de Streamlit es un <header>, no un <div>; con el selector
       div[data-testid=...] la regla no aplicaba y la franja superior seguía tapando la
       barra de navegación, dejándola sin poder pulsar. Se selecciona por atributo y se
       le quitan los eventos de puntero, devolviéndoselos solo a sus propios botones. */
    [data-testid="stHeader"] {{ background: transparent; pointer-events: none; }}

    /* Se ocultan los accesos de desarrollo: Fork, GitHub, "Manage app", el menú de
       opciones y la insignia de Streamlit. Quien entra solo debe ver la conciliación,
       sin puertas al código ni a la configuración del servidor. */
    [data-testid="stToolbar"],
    [data-testid="manage-app-button"],
    [data-testid="stStatusWidget"],
    [data-testid="stDecoration"],
    #MainMenu, footer,
    .viewerBadge_container__1QSob,
    a[href*="streamlit.io/cloud"],
    a[href*="share.streamlit.io"] {{ display: none !important; }}
    div[data-testid="stElementContainer"]:has(> div:empty) {{ display: none; }}

    /* ==================== SIDEBAR OSCURO ==================== */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {_SIDEBAR_TOP} 0%, {_SIDEBAR_BOT} 100%);
        border-right: none;
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] h4,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] label p,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {{
        color: #DCE4F5 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
    section[data-testid="stSidebar"] small {{ color: #9AAAD0 !important; }}
    section[data-testid="stSidebar"] svg {{ fill: #8C9CC4; color: #8C9CC4; }}
    section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.10); }}

    /* Marca en el tope del sidebar */
    .istho-brand {{
        display: flex; align-items: center; gap: 0.7rem;
        padding: 0.2rem 0 1rem; margin-bottom: 0.6rem;
        border-bottom: 1px solid rgba(255,255,255,0.10);
    }}
    .istho-brand img {{
        height: 34px; background: #fff; padding: 5px 8px; border-radius: 9px;
    }}
    .istho-brand .istho-brand-txt {{ line-height: 1.2; }}
    .istho-brand .istho-brand-txt b {{
        display: block; color: #fff; font-family: 'Sora', sans-serif;
        font-size: 0.92rem; font-weight: 700;
    }}
    .istho-brand .istho-brand-txt span {{ color: #8C9CC4; font-size: 0.72rem; }}

    /* Encabezado de paso numerado */
    .istho-step {{
        display: flex; align-items: center; gap: 0.55rem;
        margin: 1.1rem 0 0.55rem;
    }}
    .istho-step i {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 22px; height: 22px; border-radius: 7px; flex-shrink: 0;
        background: rgba(96,165,250,0.20); color: #93C5FD;
        font-style: normal; font-size: 0.74rem; font-weight: 700;
    }}
    .istho-step span {{
        color: #fff; font-family: 'Sora', sans-serif; font-size: 0.85rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.045em;
    }}

    /* Zona de carga de archivos dentro del sidebar oscuro */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
        background: rgba(255,255,255,0.045);
        border: 1px dashed rgba(255,255,255,0.20);
        border-radius: 10px;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: {AZUL_CLARO};
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{
        background: rgba(255,255,255,0.10); color: #DCE4F5; border: 1px solid rgba(255,255,255,0.18);
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button p {{
        color: #DCE4F5 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {{
        color: #8C9CC4 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * {{ color: #DCE4F5 !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] small {{ color: #8C9CC4 !important; }}

    section[data-testid="stSidebar"] [data-testid="stNumberInputContainer"] {{
        background: rgba(255,255,255,0.06); border-color: rgba(255,255,255,0.16);
    }}
    section[data-testid="stSidebar"] [data-testid="stNumberInputContainer"] input {{
        color: #fff;
    }}

    section[data-testid="stSidebar"] div[data-testid="stButton"] button {{
        background: {AZUL}; color: #fff; border: none; font-weight: 700;
        border-radius: 10px; padding: 0.7rem 1rem; font-size: 0.95rem;
        box-shadow: 0 4px 14px rgba(29,78,216,0.42);
    }}
    section[data-testid="stSidebar"] div[data-testid="stButton"] button:hover {{
        background: {AZUL_CLARO}; color: #fff;
    }}

    /* Selector "Retomar conciliación guardada": cada opción incluye la fecha y hora del
       último guardado (para distinguir versiones si dos personas trabajaron el mismo mes)
       y con la letra normal del sidebar el texto no cabía y se cortaba. */
    div[class*="st-key-periodo_a_retomar"] [data-baseweb="select"] * {{
        font-size: 0.72rem !important;
    }}

    /* ==================== BARRA DE NAVEGACIÓN SUPERIOR ==================== */
    /* El contenedor de Streamlit es un flex en COLUMNA, así que lo que centra
       horizontalmente es align-items (justify-content solo movería en vertical). */
    div[class*="st-key-nav_hojas"] {{ margin: 0 0 0.8rem; align-items: center; }}
    /* `width: fit-content` + `margin auto` centra el grupo completo; con solo
       justify-content se centrarían los botones dentro de un grupo ya alineado a la
       izquierda, que era lo que estaba pasando. */
    div[class*="st-key-nav_hojas"] [role="radiogroup"],
    div[class*="st-key-nav_hojas"] [data-baseweb="button-group"] {{
        display: inline-flex; justify-content: center; gap: 5px;
        width: fit-content; margin-left: auto; margin-right: auto;
        background: #E6EBF5; padding: 5px; border-radius: 12px;
    }}
    div[class*="st-key-nav_hojas"] [data-testid="stElementContainer"],
    div[class*="st-key-nav_hojas"] [data-testid="stSegmentedControl"] {{ text-align: center; }}
    div[class*="st-key-nav_hojas"] button {{
        border-radius: 9px !important; border: none !important;
        background: transparent !important; color: #5C6C64 !important;
        font-weight: 600 !important; font-size: 0.86rem !important;
        padding: 0.42rem 1.05rem !important;
    }}
    div[class*="st-key-nav_hojas"] button:hover {{ background: rgba(255,255,255,0.65) !important; }}
    div[class*="st-key-nav_hojas"] button[aria-checked="true"],
    div[class*="st-key-nav_hojas"] button[aria-pressed="true"] {{
        background: #fff !important; color: {VERDE_OSC} !important;
        box-shadow: 0 1px 3px rgba(16,24,20,0.16);
    }}

    /* ==================== HERO ==================== */
    .istho-hero {{
        position: relative; display: flex; align-items: center; gap: 1.2rem;
        padding: 1.1rem 1.7rem; margin-bottom: 0.6rem;
        border-radius: 15px; overflow: hidden;
        background: linear-gradient(120deg, #0C1740 0%, #1B3A9E 52%, #10245E 100%);
        box-shadow: 0 12px 30px rgba(9, 20, 55, 0.32);
    }}
    .istho-hero::before {{
        content: ""; position: absolute; inset: 0;
        background-image: url("data:image/svg+xml,{_PATRON_SVG}");
        background-repeat: repeat; opacity: 0.9;
    }}
    .istho-hero::after {{
        content: ""; position: absolute; inset: 0;
        background:
            radial-gradient(circle at 8% 25%, rgba(59,130,246,0.30), transparent 42%),
            radial-gradient(circle at 92% 85%, rgba(34,211,238,0.20), transparent 44%);
    }}
    .istho-hero-logo {{
        position: relative; z-index: 1; height: 42px; flex-shrink: 0;
        background: #fff; padding: 6px 10px; border-radius: 10px;
        filter: drop-shadow(0 3px 8px rgba(0,0,0,0.40));
    }}
    .istho-hero-text {{ position: relative; z-index: 1; }}
    .istho-hero-text h1 {{ color: #fff; font-size: 1.45rem; font-weight: 800; margin: 0; line-height: 1.15; }}
    .istho-hero-text p {{
        color: rgba(255,255,255,0.70); margin: 0.2rem 0 0; font-size: 0.85rem; font-weight: 500;
    }}

    /* Variante delgada: en las hojas de tablas el encabezado se reduce a una sola línea
       para que los registros empiecen lo más arriba posible. */
    .istho-hero.compacto {{ padding: 0.6rem 1.4rem; margin-bottom: 0.45rem; border-radius: 12px; }}
    .istho-hero.compacto .istho-hero-logo {{ height: 30px; padding: 4px 8px; border-radius: 8px; }}
    .istho-hero.compacto .istho-hero-text h1 {{ font-size: 1.05rem; }}
    .istho-hero.compacto .istho-hero-text p {{ display: none; }}
    .istho-hero.compacto .istho-hero-badge b {{ font-size: 0.75rem; }}
    .istho-hero.compacto .istho-hero-badge span {{ font-size: 0.68rem; }}
    .istho-hero-badge {{
        position: relative; z-index: 1; margin-left: auto; text-align: right;
        white-space: nowrap; flex-shrink: 0; padding-left: 1rem;
    }}
    .istho-hero-badge b {{
        display: block; color: rgba(255,255,255,0.85); font-size: 0.82rem; font-weight: 700;
    }}
    .istho-hero-badge span {{ color: rgba(255,255,255,0.45); font-size: 0.72rem; }}

    /* ==================== SECCIONES ==================== */
    .istho-section {{
        display: flex; align-items: baseline; gap: 0.55rem;
        margin: 0.95rem 0 0.5rem;
    }}
    .istho-section h4 {{
        margin: 0; font-size: 0.95rem; font-weight: 700; color: {_TEXTO};
    }}
    .istho-section em {{ font-style: normal; font-size: 0.78rem; color: #8A9791; }}

    /* ==================== TARJETAS DE INDICADORES ==================== */
    .istho-stats {{
        display: grid; grid-template-columns: repeat(6, minmax(0, 1fr));
        gap: 0.6rem; margin-bottom: 0.3rem;
    }}
    @media (max-width: 1350px) {{
        .istho-stats {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    }}
    @media (max-width: 900px) {{
        .istho-stats {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 560px) {{
        .istho-stats {{ grid-template-columns: 1fr; }}
    }}
    .istho-card {{
        position: relative; background: #fff;
        border: 1px solid {_BORDE}; border-radius: 12px;
        padding: 0.7rem 0.85rem 0.8rem;
        box-shadow: 0 1px 2px rgba(16,24,20,0.04), 0 6px 16px rgba(16,24,20,0.035);
        transition: transform .14s ease, box-shadow .14s ease;
    }}
    .istho-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 2px 4px rgba(16,24,20,0.05), 0 10px 24px rgba(16,24,20,0.08);
    }}
    .istho-card-top {{
        display: flex; align-items: center; gap: 0.4rem; margin-bottom: 0.35rem;
    }}
    .istho-chip {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 22px; height: 22px; border-radius: 7px; font-size: 0.72rem; flex-shrink: 0;
        background: color-mix(in srgb, var(--accent) 14%, transparent);
    }}
    .istho-card .istho-label {{
        font-size: 0.64rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em;
        color: #7A8A82; line-height: 1.2;
    }}
    .istho-card .istho-value {{
        font-family: 'Sora', sans-serif; font-size: 1.3rem; font-weight: 700;
        color: {_TEXTO}; line-height: 1.05;
    }}
    .istho-card .istho-sub {{ font-size: 0.67rem; color: #97A39C; margin-top: 0.2rem; }}
    .istho-card::after {{
        content: ""; position: absolute; left: 0.85rem; right: 0.85rem; bottom: 0;
        height: 3px; border-radius: 3px 3px 0 0; background: var(--accent, {VERDE});
        opacity: 0.85;
    }}

    /* Explicación al pasar el cursor (solo CSS, sin JavaScript) */
    .istho-card .istho-tip {{
        position: absolute; z-index: 999;
        left: 50%; top: calc(100% + 11px);
        transform: translateX(-50%) translateY(-5px);
        width: max-content; max-width: 265px;
        background: #142244; color: #E8EDF8;
        font-family: 'Inter', sans-serif; font-size: 0.745rem; font-weight: 500;
        line-height: 1.4; text-transform: none; letter-spacing: 0; text-align: left;
        padding: 0.55rem 0.72rem; border-radius: 9px;
        box-shadow: 0 8px 22px rgba(10,20,14,0.30);
        opacity: 0; visibility: hidden; pointer-events: none;
        transition: opacity .16s ease, transform .16s ease, visibility .16s;
    }}
    .istho-card .istho-tip::after {{
        content: ""; position: absolute; bottom: 100%; left: 50%;
        transform: translateX(-50%);
        border: 6px solid transparent; border-bottom-color: #142244;
    }}
    .istho-card:hover .istho-tip {{
        opacity: 1; visibility: visible; transform: translateX(-50%) translateY(0);
    }}
    /* En las tarjetas de los extremos el globo se ancla hacia adentro para que no
       quede cortado por el borde de la pantalla. */
    .istho-stats .istho-card:nth-last-child(-n+2) .istho-tip {{
        left: auto; right: 0; transform: translateX(0) translateY(-5px);
    }}
    .istho-stats .istho-card:nth-last-child(-n+2):hover .istho-tip {{
        transform: translateX(0) translateY(0);
    }}
    .istho-stats .istho-card:nth-last-child(-n+2) .istho-tip::after {{
        left: auto; right: 26px; transform: none;
    }}
    .istho-stats .istho-card:first-child .istho-tip {{
        left: 0; right: auto; transform: translateX(0) translateY(-5px);
    }}
    .istho-stats .istho-card:first-child:hover .istho-tip {{
        transform: translateX(0) translateY(0);
    }}
    .istho-stats .istho-card:first-child .istho-tip::after {{
        left: 26px; right: auto; transform: none;
    }}
    /* --- Que el globo quede SIEMPRE por encima de todo ---
       No basta con un z-index alto en el globo: cada bloque de Streamlit se pinta según su
       orden en el documento, así que el segundo grupo de tarjetas (y la tabla que sigue)
       tapaban el globo del primero. Hay que elevar también la tarjeta y todos los
       contenedores que la envuelven, y dejarlos con overflow visible para que no lo
       recorten. */
    div[data-testid="stMarkdownContainer"]:has(.istho-stats),
    div[data-testid="stMarkdown"]:has(.istho-stats),
    div[data-testid="stElementContainer"]:has(.istho-stats),
    div[data-testid="stVerticalBlock"]:has(.istho-stats),
    .istho-stats {{ overflow: visible !important; }}

    .istho-card:hover {{ z-index: 900; }}
    .istho-stats:has(.istho-card:hover),
    div[data-testid="stMarkdownContainer"]:has(.istho-card:hover),
    div[data-testid="stMarkdown"]:has(.istho-card:hover),
    div[data-testid="stElementContainer"]:has(.istho-card:hover) {{
        position: relative; z-index: 900;
    }}

    /* ==================== PANEL (contenedor con borde) ==================== */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 16px !important;
        border-color: {_BORDE} !important;
        background: #fff;
        box-shadow: 0 1px 2px rgba(16,24,20,0.04), 0 8px 22px rgba(16,24,20,0.045);
    }}

    /* ==================== PESTAÑAS TIPO PÍLDORA ====================
       Esta versión de Streamlit monta las tabs con react-aria (no baseweb):
       el contenedor es [role="tablist"] y cada tab es [data-testid="stTab"]. */
    .stTabs [role="tablist"] {{
        gap: 4px; background: #E6EBF5; padding: 4px; border-radius: 11px;
        border-bottom: none !important; display: inline-flex; flex-wrap: wrap;
        margin-bottom: 0.5rem;
    }}
    .stTabs .react-aria-SelectionIndicator {{ display: none !important; }}
    .stTabs [data-testid="stTab"] {{
        border-radius: 9px; padding: 0.38rem 0.8rem;
        background: transparent; border: none; transition: background .13s ease;
    }}
    .stTabs [data-testid="stTab"] p {{
        font-size: 0.79rem !important; font-weight: 600 !important;
        color: #63736B !important; margin: 0;
    }}
    .stTabs [data-testid="stTab"]:hover {{ background: rgba(255,255,255,0.75); }}
    .stTabs [data-testid="stTab"][aria-selected="true"] {{
        background: #fff; box-shadow: 0 1px 3px rgba(16,24,20,0.16);
    }}
    .stTabs [data-testid="stTab"][aria-selected="true"] p {{ color: {VERDE_OSC} !important; }}

    /* ==================== TABLAS ==================== */
    div[data-testid="stDataFrame"] {{
        border-radius: 13px; overflow: hidden; border: 1px solid {_BORDE};
        box-shadow: 0 1px 2px rgba(16,24,20,0.04), 0 6px 18px rgba(16,24,20,0.04);
    }}

    /* ==================== BARRA DE FILTROS ==================== */
    div[class*="st-key-toolbar_"] {{
        background: #FFFFFF !important;
        border: 1px solid {_BORDE} !important;
        border-radius: 11px !important;
        padding: 0.6rem 0.9rem 0.1rem !important;
        margin-bottom: 0.45rem !important;
        box-shadow: 0 1px 2px rgba(16,24,20,0.035);
    }}
    div[class*="st-key-toolbar_"] label p {{
        font-size: 0.68rem !important; font-weight: 700 !important;
        color: #63736B !important; text-transform: uppercase; letter-spacing: 0.04em;
    }}
    div[class*="st-key-toolbar_"] div[data-testid="stVerticalBlock"] {{ gap: 0.2rem; }}

    /* ==================== ETIQUETA DE CONTEO ==================== */
    .istho-badge {{
        display: inline-flex; align-items: center; gap: 0.35rem;
        background: #E8EEFC; color: {AZUL_OSC};
        font-size: 0.74rem; font-weight: 600;
        padding: 0.2rem 0.65rem; border-radius: 999px;
        margin: 0 0 0.4rem;
    }}

    /* ==================== PANTALLA DE CARGA ==================== */
    .istho-loader {{
        position: relative; overflow: hidden;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        gap: 1.15rem; padding: 3.2rem 2rem;
        border-radius: 18px; margin: 0.5rem 0 1rem;
        background: linear-gradient(120deg, #0C1740 0%, #1B3A9E 52%, #10245E 100%);
        box-shadow: 0 12px 30px rgba(9, 20, 55, 0.32);
    }}
    .istho-loader::before {{
        content: ""; position: absolute; inset: 0;
        background-image: url("data:image/svg+xml,{_PATRON_SVG}");
        background-repeat: repeat; opacity: 0.85;
    }}
    /* Barrido de luz que recorre la tarjeta */
    .istho-loader::after {{
        content: ""; position: absolute; top: 0; bottom: 0; width: 45%;
        background: linear-gradient(90deg, transparent, rgba(96,165,250,0.16), transparent);
        animation: istho-sweep 2.1s linear infinite;
    }}
    @keyframes istho-sweep {{
        0%   {{ left: -45%; }}
        100% {{ left: 100%; }}
    }}

    .istho-loader-logo {{
        position: relative; z-index: 1; height: 46px;
        background: #fff; padding: 7px 11px; border-radius: 11px;
        animation: istho-latido 1.9s ease-in-out infinite;
    }}
    @keyframes istho-latido {{
        0%, 100% {{ transform: scale(1);    box-shadow: 0 0 0 0 rgba(59,130,246,0.45); }}
        50%      {{ transform: scale(1.04); box-shadow: 0 0 0 14px rgba(59,130,246,0); }}
    }}

    .istho-loader-txt {{
        position: relative; z-index: 1; text-align: center;
        color: #fff; font-family: 'Sora', sans-serif;
        font-size: 1.12rem; font-weight: 700; letter-spacing: 0.01em;
    }}
    .istho-loader-sub {{
        position: relative; z-index: 1;
        color: rgba(255,255,255,0.60); font-size: 0.85rem; margin-top: -0.55rem;
    }}

    /* Bolitas titilando */
    .istho-dots {{ position: relative; z-index: 1; display: flex; gap: 0.55rem; }}
    .istho-dots i {{
        width: 12px; height: 12px; border-radius: 50%; display: block;
        animation: istho-bounce 1.25s ease-in-out infinite both;
    }}
    .istho-dots i:nth-child(1) {{ background: {AZUL}; animation-delay: -0.34s; }}
    .istho-dots i:nth-child(2) {{ background: {AZUL_CLARO}; animation-delay: -0.17s; }}
    .istho-dots i:nth-child(3) {{ background: {CIAN}; animation-delay: 0s; }}
    @keyframes istho-bounce {{
        0%, 75%, 100% {{ transform: translateY(0) scale(0.72); opacity: 0.40; }}
        35%           {{ transform: translateY(-11px) scale(1); opacity: 1; }}
    }}

    /* Barra de progreso indeterminada */
    .istho-bar {{
        position: relative; z-index: 1; width: min(340px, 70%); height: 4px;
        background: rgba(255,255,255,0.12); border-radius: 99px; overflow: hidden;
    }}
    .istho-bar span {{
        position: absolute; height: 100%; width: 38%; border-radius: 99px;
        background: linear-gradient(90deg, {AZUL_CLARO}, {CIAN});
        animation: istho-slide 1.5s ease-in-out infinite;
    }}
    @keyframes istho-slide {{
        0%   {{ left: -38%; }}
        100% {{ left: 100%; }}
    }}

    /* ==================== BOTÓN DE DESCARGA POR TABLA ==================== */
    div[class*="st-key-btn_dl_"] button {{
        background: linear-gradient(135deg, {AZUL_OSC}, {AZUL}); color: #fff;
        font-size: 0.82rem; font-weight: 700; padding: 0.42rem 0.8rem; border-radius: 9px;
        border: none; letter-spacing: 0.01em;
        box-shadow: 0 2px 8px rgba(27,94,32,0.28);
        transition: transform .12s ease, box-shadow .12s ease;
    }}
    div[class*="st-key-btn_dl_"] button:hover {{
        background: linear-gradient(135deg, #112663, #1A45C0); color: #fff;
        transform: translateY(-1px); box-shadow: 0 4px 12px rgba(27,94,32,0.34);
    }}
    div[class*="st-key-btn_dl_"] button:disabled {{
        background: #D8DEDA; color: #9AA69F; box-shadow: none;
    }}
    div[class*="st-key-btn_dl_"] button p {{ font-size: 0.82rem !important; font-weight: 700 !important; }}
    </style>
    """, unsafe_allow_html=True)


def _flat(html):
    """Streamlit interpreta la indentación de 4+ espacios como bloque de código,
    incluso con unsafe_allow_html. Se aplana a una sola línea para evitarlo."""
    return "".join(line.strip() for line in html.strip().splitlines())


def sidebar_brand(nombre="ISTHO S.A.S.", sub="Conciliación bancaria"):
    logo_b64 = _logo_b64()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" />' if logo_b64 else ""
    st.sidebar.markdown(_flat(f"""
    <div class="istho-brand">
        {logo_html}
        <div class="istho-brand-txt"><b>{nombre}</b><span>{sub}</span></div>
    </div>
    """), unsafe_allow_html=True)


def sidebar_step(numero, titulo):
    st.sidebar.markdown(_flat(f"""
    <div class="istho-step"><i>{numero}</i><span>{titulo}</span></div>
    """), unsafe_allow_html=True)


def hero(title, subtitle, empresa=None, periodo=None, compacto=False):
    logo_b64 = _logo_b64()
    logo_html = (f'<img src="data:image/png;base64,{logo_b64}" class="istho-hero-logo" />'
                 if logo_b64 else "")
    badge_html = ""
    if empresa:
        badge_html = _flat(f"""
        <div class="istho-hero-badge">
            <b>{empresa}</b>
            <span>{periodo or ""}</span>
        </div>
        """)
    clase = "istho-hero compacto" if compacto else "istho-hero"
    st.markdown(_flat(f"""
    <div class="{clase}">
        {logo_html}
        <div class="istho-hero-text">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        {badge_html}
    </div>
    """), unsafe_allow_html=True)


def section(titulo, nota=None):
    nota_html = f"<em>{nota}</em>" if nota else ""
    st.markdown(_flat(f"""
    <div class="istho-section"><h4>{titulo}</h4>{nota_html}</div>
    """), unsafe_allow_html=True)


def stat_cards(items):
    """items: dicts con {label, value, icon, sub (opcional), accent (opcional hex),
    tip (opcional: explicación que aparece al pasar el cursor)}."""
    tarjetas = "".join(_flat(f"""
        <div class="istho-card" style="--accent:{it.get('accent', VERDE)}">
            <div class="istho-card-top">
                <span class="istho-chip">{it.get('icon', '•')}</span>
                <span class="istho-label">{it['label']}</span>
            </div>
            <div class="istho-value">{it['value']}</div>
            {f'<div class="istho-sub">{it["sub"]}</div>' if it.get('sub') else ''}
            {f'<span class="istho-tip">{it["tip"]}</span>' if it.get('tip') else ''}
        </div>
    """) for it in items)
    st.markdown(f'<div class="istho-stats">{tarjetas}</div>', unsafe_allow_html=True)


def badge(texto):
    st.markdown(f'<div class="istho-badge">🔎 {texto}</div>', unsafe_allow_html=True)


def loader(mensaje, detalle=""):
    """HTML de la pantalla de carga. Las animaciones son CSS puro, así que siguen
    corriendo en el navegador mientras Python está ocupado procesando."""
    logo_b64 = _logo_b64()
    logo_html = (f'<img src="data:image/png;base64,{logo_b64}" class="istho-loader-logo" />'
                 if logo_b64 else "")
    detalle_html = f'<div class="istho-loader-sub">{detalle}</div>' if detalle else ""
    return _flat(f"""
    <div class="istho-loader">
        {logo_html}
        <div class="istho-loader-txt">{mensaje}</div>
        {detalle_html}
        <div class="istho-dots"><i></i><i></i><i></i></div>
        <div class="istho-bar"><span></span></div>
    </div>
    """)


_COLS_FECHA = {"fecha", "fecha banco", "fecha contabilidad"}
_COLS_MONEDA = {"valor", "valor banco", "valor contabilidad", "diferencia"}


# Anchos en píxeles. Las descripciones y los valores mandan: por el nombre y el monto es
# que se decide si un movimiento cruza o no, así que nunca deben quedar recortados.
_ANCHOS = {
    "id": 62, "origen": 85, "tipo": 76, "dif. días": 62, "grupo": 58,
    "fecha": 96, "fecha banco": 96, "fecha contabilidad": 104,
    "valor": 145, "valor banco": 145, "valor contabilidad": 155, "diferencia": 135,
    # El extracto trunca sus textos a ~28 caracteres, así que su columna no necesita
    # tanto ancho; el espacio sobrante se le da a la descripción contable, que es larga.
    "descripción banco": 250, "descripción contabilidad": 560, "descripción": 620,
    "comprobante": 110, "documento": 120,
    "motivo": 200, "conciliado el": 130, "confianza": 200,
}

# Columnas que se ocultan en pantalla para dejarle todo el espacio a descripciones y
# valores. Siguen yendo completas en el Excel que se descarga.
OCULTAR_EN_PANTALLA = ("Comprobante", "Documento", "Conciliado el")


def tabla(df, height=780, row_height=34, ocultar=OCULTAR_EN_PANTALLA):
    """st.dataframe con formato contable de moneda/fecha, columnas dimensionadas según su
    contenido y alto generoso, para ver la mayor cantidad de movimientos de una sola vez.
    `ocultar` quita de la vista las columnas accesorias (siguen en el Excel)."""
    df = df.drop(columns=[c for c in ocultar if c in df.columns], errors="ignore").copy()
    # Las celdas vacías llegan como NaN y Streamlit las dibuja como "None", que en un informe
    # contable se lee como un dato. Se convierten a texto vacío. Ojo: en pandas 3 `astype(str)`
    # deja los NaN intactos, por eso la conversión se hace valor por valor.
    for col in df.columns:
        if col.lower() in _COLS_FECHA or is_numeric_dtype(df[col]):
            continue
        df[col] = df[col].map(lambda v: "" if pd.isna(v) else str(v))

    column_config = {}
    for col in df.columns:
        key = col.lower()
        ancho = _ANCHOS.get(key)
        if key in _COLS_FECHA:
            column_config[col] = st.column_config.DateColumn(col, format="DD/MM/YYYY", width=ancho)
        elif key in _COLS_MONEDA:
            column_config[col] = st.column_config.NumberColumn(col, format="accounting", width=ancho)
        elif key == "dif. días":
            column_config[col] = st.column_config.NumberColumn("Dif. días", format="plain", width=ancho)
        elif ancho:
            column_config[col] = st.column_config.TextColumn(col, width=ancho)
    st.dataframe(df, use_container_width=True, hide_index=True,
                 height=min(height, 60 + len(df) * row_height),
                 row_height=row_height, column_config=column_config)


# ===========================================================================
# PANTALLA DE ACCESO
#
# Azul marino profundo con acentos dorados: en interfaces financieras el azul
# comunica confianza y el oro señala categoría. El dorado se usa SOLO aquí; en
# las tablas competiría con el ámbar de "requiere revisión".
# ===========================================================================

# Retícula de nodos y trazos: alude a la trazabilidad del dinero sin caer en
# iconos literales de billetes, que restarían seriedad.
_PATRON_LOGIN = """
<svg xmlns='http://www.w3.org/2000/svg' width='260' height='260'>
  <g fill='none' stroke='%23D4AF37' stroke-opacity='0.10' stroke-width='1'>
    <path d='M20 60 L90 60 L120 30 L210 30'/>
    <path d='M0 150 L60 150 L95 185 L180 185 L215 150 L260 150'/>
    <path d='M140 0 L140 55 L200 90 L200 140'/>
    <path d='M40 230 L100 230 L130 200'/>
  </g>
  <g fill='none' stroke='%233B82F6' stroke-opacity='0.12' stroke-width='1'>
    <path d='M0 95 L45 95 L75 125 L160 125'/>
    <path d='M225 200 L250 225 L250 260'/>
  </g>
  <g fill='%23D4AF37' fill-opacity='0.22'>
    <circle cx='90' cy='60' r='2.6'/><circle cx='180' cy='185' r='2.6'/>
    <circle cx='200' cy='90' r='2.6'/><circle cx='100' cy='230' r='2.6'/>
  </g>
  <g fill='%233B82F6' fill-opacity='0.26'>
    <circle cx='45' cy='95' r='2.6'/><circle cx='160' cy='125' r='2.6'/>
    <circle cx='140' cy='55' r='2.6'/>
  </g>
</svg>
""".replace("\n", "").replace("  ", "")


def login_css():
    """Estilos exclusivos de la pantalla de acceso. Se inyectan solo cuando se
    muestra el login, para no teñir de azul oscuro el resto de la aplicación."""
    st.markdown(f"""
    <style>
    /* Fondo a pantalla completa: degradado marino con halos dorado y azul */
    .stApp {{
        background:
            radial-gradient(circle at 12% 16%, rgba(212,175,55,0.18), transparent 45%),
            radial-gradient(circle at 88% 84%, rgba(34,211,238,0.18), transparent 46%),
            linear-gradient(135deg, #0C1740 0%, #1B3A9E 52%, #10245E 100%) !important;
    }}
    .stApp::before {{
        content: ""; position: fixed; inset: 0; pointer-events: none;
        background-image: url("data:image/svg+xml,{_PATRON_LOGIN}");
        background-repeat: repeat; opacity: 0.75;
    }}

    /* En el acceso no hay nada que configurar: se oculta el panel lateral */
    section[data-testid="stSidebar"] {{ display: none !important; }}
    .block-container {{ padding-top: 5vh !important; max-width: 1100px !important; }}

    /* ---------------- Tarjeta ---------------- */
    /* Sobre el azul vivo del encabezado la tarjeta va en azul profundo translúcido: si
       fuera blanca translúcida se fundiría con el fondo y el texto perdería contraste. */
    div[class*="st-key-login_card"] {{
        position: relative;
        background: rgba(8, 18, 52, 0.62) !important;
        border: 1px solid rgba(212,175,55,0.34) !important;
        border-radius: 20px !important;
        padding: 2.3rem 2.2rem 1.6rem !important;
        box-shadow: 0 30px 70px rgba(4, 10, 34, 0.55), inset 0 1px 0 rgba(255,255,255,0.09);
        backdrop-filter: blur(16px);
    }}
    /* Filo dorado superior, como el canto de un documento de valor */
    div[class*="st-key-login_card"]::before {{
        content: ""; position: absolute; top: 0; left: 14%; right: 14%; height: 2px;
        background: linear-gradient(90deg, transparent, {ORO}, {ORO_CLARO}, {ORO}, transparent);
        border-radius: 2px;
    }}

    .istho-login-cab {{ text-align: center; margin-bottom: 1.5rem; }}
    .istho-login-logo {{
        height: 54px; background: #fff; padding: 9px 14px; border-radius: 13px;
        box-shadow: 0 10px 26px rgba(0,0,0,0.42), 0 0 0 1px rgba(212,175,55,0.30);
        margin-bottom: 1.15rem;
    }}
    .istho-login-eyebrow {{
        color: {ORO}; font-size: 0.68rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.22em; margin-bottom: 0.5rem;
    }}
    .istho-login-cab h1 {{
        color: #FFFFFF; font-family: 'Sora', sans-serif;
        font-size: 1.72rem; font-weight: 800; margin: 0; letter-spacing: -0.02em;
    }}
    .istho-login-cab p {{
        color: rgba(226,236,255,0.62); font-size: 0.86rem; margin: 0.45rem 0 0;
    }}
    .istho-login-sep {{
        width: 54px; height: 2px; margin: 1.2rem auto 0; border-radius: 2px;
        background: linear-gradient(90deg, transparent, {ORO}, transparent);
    }}

    /* ---------------- Campo de clave ---------------- */
    div[class*="st-key-login_card"] label p {{
        color: rgba(226,236,255,0.78) !important;
        font-size: 0.76rem !important; font-weight: 600 !important;
        text-transform: uppercase; letter-spacing: 0.09em;
    }}
    /* OJO: en esta versión el fondo blanco lo pone stTextInputRootElement; el atributo
       data-baseweb ya no existe, así que hay que apuntar a ese testid o el campo queda
       blanco y el texto se vuelve ilegible sobre la tarjeta oscura. */
    div[class*="st-key-login_card"] [data-testid="stTextInputRootElement"] {{
        background: rgba(255,255,255,0.07) !important;
        border: 1px solid rgba(255,255,255,0.16) !important;
        border-radius: 12px !important;
        transition: border-color .16s ease, box-shadow .16s ease;
    }}
    div[class*="st-key-login_card"] [data-testid="stTextInputRootElement"]:focus-within {{
        border-color: {ORO} !important;
        box-shadow: 0 0 0 3px rgba(212,175,55,0.16) !important;
    }}
    /* El botón de "ver clave" vive dentro del campo: se le quita el dorado del botón
       principal para que no compita con la acción de ingresar. */
    div[class*="st-key-login_card"] [data-testid="stTextInputRootElement"] button {{
        background: transparent !important; box-shadow: none !important;
        border: none !important; padding: 0 0.6rem !important;
    }}
    div[class*="st-key-login_card"] [data-testid="stTextInputRootElement"] svg {{
        fill: rgba(226,236,255,0.50);
    }}
    div[class*="st-key-login_card"] input {{
        background: transparent !important; color: #F3F7FF !important;
        padding: 0.72rem 0.3rem !important; font-size: 0.95rem !important;
    }}
    div[class*="st-key-login_card"] input::placeholder {{ color: rgba(226,236,255,0.34) !important; }}
    div[class*="st-key-login_card"] [data-testid="stWidgetLabel"] {{ margin-bottom: 0.4rem; }}
    div[class*="st-key-login_card"] svg {{ fill: rgba(226,236,255,0.55); }}

    /* ---------------- Botón dorado ---------------- */
    div[class*="st-key-login_card"] button {{
        background: linear-gradient(135deg, {ORO_OSCURO} 0%, {ORO} 42%, {ORO_CLARO} 100%) !important;
        color: #0E1B33 !important; border: none !important; border-radius: 12px !important;
        font-weight: 800 !important; font-size: 0.88rem !important;
        letter-spacing: 0.1em; text-transform: uppercase;
        padding: 0.78rem 1rem !important; margin-top: 0.35rem;
        box-shadow: 0 10px 26px rgba(212,175,55,0.30);
        transition: transform .13s ease, box-shadow .13s ease, filter .13s ease;
    }}
    div[class*="st-key-login_card"] button p {{
        color: #0E1B33 !important; font-weight: 800 !important; letter-spacing: 0.1em;
    }}
    div[class*="st-key-login_card"] button:hover {{
        filter: brightness(1.07); transform: translateY(-1px);
        box-shadow: 0 14px 32px rgba(212,175,55,0.40);
    }}

    /* Aviso de clave incorrecta, en tono sobrio */
    div[class*="st-key-login_card"] div[data-testid="stAlertContainer"] {{
        background: rgba(220,39,67,0.13) !important;
        border: 1px solid rgba(220,39,67,0.36) !important;
        border-radius: 11px !important;
    }}
    div[class*="st-key-login_card"] div[data-testid="stAlertContainer"] p {{
        color: #FFC9D2 !important; font-size: 0.84rem !important;
    }}

    .istho-login-pie {{
        text-align: center; margin-top: 1.15rem;
        color: rgba(226,236,255,0.62); font-size: 0.75rem; line-height: 1.7;
        text-shadow: 0 1px 3px rgba(4,10,34,0.5);
    }}
    .istho-login-pie b {{ color: {ORO_CLARO}; font-weight: 700; }}
    </style>
    """, unsafe_allow_html=True)


def login_encabezado(titulo, subtitulo, eyebrow="Área financiera"):
    logo_b64 = _logo_b64()
    logo = (f'<img src="data:image/png;base64,{logo_b64}" class="istho-login-logo" />'
            if logo_b64 else "")
    st.markdown(_flat(f"""
    <div class="istho-login-cab">
        {logo}
        <div class="istho-login-eyebrow">{eyebrow}</div>
        <h1>{titulo}</h1>
        <p>{subtitulo}</p>
        <div class="istho-login-sep"></div>
    </div>
    """), unsafe_allow_html=True)


def login_pie(empresa):
    st.markdown(_flat(f"""
    <div class="istho-login-pie">
        <b>{empresa}</b><br>
        Acceso restringido · Si no tienes la clave, solicítala al área financiera.
    </div>
    """), unsafe_allow_html=True)
