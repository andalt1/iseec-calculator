# -*- coding: utf-8 -*-
"""Оформление калькулятора ОСЭЭК: палитра, стили, графики.

Основной цвет NAVY #012169 (фирменный синий факультета государственного
управления МГУ), дополнительный BLUE #5E8FEF. Уровни шкалы кодируются
дивергентной лестницей с нейтральной серединой; принадлежность уровня
всегда дублируется текстовой подписью.
"""

import plotly.graph_objects as go
import streamlit as st

from oseec_core import LEVEL_NAMES, fmt

NAVY = "#012169"
INK = "#232830"
GRAY = "#6B7686"
LINE = "#C9D7EA"
PANEL = "#F4F8FC"
PANEL2 = "#EDF2FA"
BLUE = "#5E8FEF"

# Уровни шкалы: акцентные цвета и светлые заливки полос
LEVEL_ACCENT = ("#A6261C", "#CE6E43", "#8792A3", "#4A77D4", "#123A8C")
LEVEL_TINT = ("#F6E4E2", "#F9EBE3", "#EFF1F5", "#E3EBFA", "#DFE5F3")
# Насыщеннее тинтов — для дуги спидометра, чтобы зоны читались с проектора
GAUGE_ZONES = ("#E4BCB5", "#EDCDB4", "#CDD5E2", "#B7C9EE", "#A5B9E7")
LEVEL_EDGES = (0.0, 25.0, 50.0, 75.0, 100.0, 116.0)

FONT_BODY = "Inter, 'PT Sans', 'Segoe UI', sans-serif"
FONT_SERIF = "'PT Serif', Georgia, 'Times New Roman', serif"

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=PT+Serif:ital,wght@0,400;0,700;1,400&family=Inter:wght@400;500;600;700&display=swap');

html, body,
[data-testid="stAppViewContainer"] *:not([data-testid="stIconMaterial"]):not([class*="material-symbols"]),
[data-testid="stSidebar"] *:not([data-testid="stIconMaterial"]):not([class*="material-symbols"]) {
    font-family: Inter, 'PT Sans', 'Segoe UI', sans-serif;
}
[data-testid="stIconMaterial"], [class*="material-symbols"] {
    font-family: 'Material Symbols Rounded' !important;
}
[data-testid="stAppViewContainer"] {
    background: #FFFFFF;
}
[data-testid="stHeader"] {
    background: transparent;
}
[data-testid="stMainBlockContainer"], .block-container {
    max-width: 1180px;
    padding-top: 1.3rem;
    padding-bottom: 3rem;
}

h1, h2, h3 {
    font-family: 'PT Serif', Georgia, 'Times New Roman', serif !important;
    color: #012169 !important;
    letter-spacing: 0;
}
h1 { font-size: 1.9rem !important; font-weight: 700 !important; }
h2 { font-size: 1.35rem !important; font-weight: 700 !important; }
h3 { font-size: 1.12rem !important; font-weight: 700 !important; }

[data-testid="stSidebar"] {
    background: #F4F8FC;
    border-right: 1px solid #C9D7EA;
}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span {
    font-size: 0.94rem;
}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {
    background: #E3EBFA;
    border-radius: 8px;
}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] span {
    color: #012169;
    font-weight: 600;
}

.oseec-brand {
    padding: 0.2rem 0.2rem 0.6rem 0.2rem;
    border-bottom: 1px solid #C9D7EA;
    margin-bottom: 0.4rem;
}
.oseec-brand .t1 {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 1.22rem;
    color: #012169;
    line-height: 1.25;
}
.oseec-brand .t2 {
    font-size: 0.78rem;
    color: #6B7686;
    margin-top: 0.35rem;
    line-height: 1.45;
}
.oseec-sidefoot {
    font-size: 0.74rem;
    color: #6B7686;
    line-height: 1.5;
    border-top: 1px solid #C9D7EA;
    padding-top: 0.7rem;
    margin-top: 0.6rem;
}

/* --- Плашка-заголовок карточки: залитая полоса во всю ширину --------------- */
.oseec-plate {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    flex-wrap: wrap;
    background: linear-gradient(120deg, #012169 0%, #123278 62%, #1D4396 100%);
    margin: -1.15rem -1.45rem 0.15rem -1.45rem;
    padding: 0.66rem 1.35rem 0.64rem 1.35rem;
    border-radius: 13px 13px 0 0;
}
.oseec-plate .num {
    background: rgba(255, 255, 255, 0.14);
    border: 1.5px solid rgba(255, 255, 255, 0.72);
    color: #FFFFFF;
    font-weight: 600;
    font-size: 0.92rem;
    width: 1.8rem;
    height: 1.8rem;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 auto;
}
.oseec-plate .txt {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 1.13rem;
    color: #FFFFFF;
    line-height: 1.3;
}
.oseec-plate .sub {
    margin-left: auto;
    font-size: 0.8rem;
    color: #C3D2F0;
    text-align: right;
    line-height: 1.35;
    max-width: 46%;
}
.oseec-plate.light {
    background: #F1DFB6;
}
.oseec-plate.light .txt { color: #012169; }
.oseec-plate.light .num {
    background: #FFFFFF;
    border-color: #B08D3E;
    color: #7A5A1A;
}
.oseec-plate.light .sub { color: #8A6B3F; }

/* --- Шапка зоны результатов: заголовок с золотистым подчеркиванием --------- */
.oseec-band {
    display: flex;
    align-items: baseline;
    gap: 0.9rem;
    flex-wrap: wrap;
    padding: 0.15rem 0.3rem 0.45rem 0.3rem;
    border-bottom: 2px solid #D9BC7C;
    margin: 0 0 0.9rem 0;
}
.oseec-band .txt {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 1.42rem;
    color: #012169;
}
.oseec-band .sub {
    font-size: 0.85rem;
    color: #8A6B3F;
    margin-left: auto;
}

/* --- Внутренние подзаголовки и подписи ------------------------------------- */
.oseec-sub {
    display: flex;
    align-items: center;
    gap: 0.55rem;
    font-weight: 600;
    font-size: 1.0rem;
    color: #012169;
    margin: 0.55rem 0 0.15rem 0;
}
.oseec-sub::before {
    content: "";
    width: 5px;
    height: 1.05rem;
    background: #5E8FEF;
    border-radius: 3px;
    flex: 0 0 auto;
}
.oseec-eyebrow {
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    color: #6B7686;
    margin: 0.25rem 0 0.05rem 0;
}
.oseec-lead {
    color: #4A5568;
    font-size: 0.97rem;
    line-height: 1.62;
    margin: -0.35rem 0 0.9rem 0;
}

.oseec-chip {
    display: inline-block;
    padding: 0.28rem 0.85rem;
    border-radius: 999px;
    font-size: 0.95rem;
    font-weight: 600;
    color: #232830;
    border: 1.5px solid;
    white-space: nowrap;
}
.oseec-note {
    border: 1px solid #D9E4F2;
    border-left: 3px solid #5E8FEF;
    border-radius: 8px;
    background: #F7FAFD;
    color: #232830;
    font-size: 0.87rem;
    line-height: 1.55;
    padding: 0.6rem 0.95rem;
    margin: 0.35rem 0;
}
.oseec-crit {
    border: 1px solid #E4C7C3;
    border-left: 4px solid #A6261C;
    border-radius: 10px;
    background: #FBF3F2;
    color: #232830;
    font-size: 0.9rem;
    line-height: 1.55;
    padding: 0.7rem 1rem;
    margin: 0.35rem 0;
}
.oseec-form {
    border: 1px solid #C9D7EA;
    border-radius: 10px;
    background: #FFFFFF;
    padding: 0.55rem 1rem;
    margin: 0.25rem 0;
    font-size: 0.92rem;
}

[data-testid="stMetricLabel"] {
    color: #6B7686 !important;
    font-size: 0.82rem !important;
    height: auto !important;
}
[data-testid="stMetricLabel"] p {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    line-height: 1.35;
}
[data-testid="stMetricValue"] {
    font-family: 'PT Serif', Georgia, serif !important;
    color: #012169 !important;
    font-size: 1.75rem !important;
}

/* Карточка = вертикальный блок, первым элементом которого идет плашка */
div[data-testid="stLayoutWrapper"]:has(> div[data-testid="stVerticalBlock"] > div[data-testid="stElementContainer"] .oseec-plate) {
    margin: 0.3rem 0 1.05rem 0;
}
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-plate) {
    border: 1px solid #C9D7EA !important;
    border-radius: 14px !important;
    background: #FFFFFF;
    box-shadow: 0 2px 10px rgba(1, 33, 105, 0.07);
    padding: 1.15rem 1.45rem 1.35rem 1.45rem !important;
}
/* Тонированный вариант карточки (чередование фонов) */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-plate.z) {
    background: #F0F5FC;
}
/* Панель итога: заливка и усиленная рамка в теплой гамме зоны */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-plate.result) {
    background: linear-gradient(180deg, #FFFFFF 0%, #FDF7EA 100%) !important;
    border: 1px solid #D9BC7C !important;
    box-shadow: 0 4px 14px rgba(146, 106, 26, 0.18);
}
/* Зона результатов: теплое кремовое полотно — температурный контраст к
   сине-белой зоне ввода. Зоной считается контейнер, первым элементом
   которого идет .oseec-band */
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-band) {
    background: #F7EDD8 !important;
    border: 1px solid #DFC58F !important;
    border-radius: 18px !important;
    padding: 1.0rem 1.1rem 1.2rem 1.1rem !important;
    box-shadow: 0 6px 20px rgba(146, 106, 26, 0.16);
}
div[data-testid="stLayoutWrapper"]:has(> div[data-testid="stVerticalBlock"] > div[data-testid="stElementContainer"] .oseec-band) {
    margin: 1.6rem 0 1.05rem 0;
}
/* Английские подсказки Press Enter to apply у полей скрываем */
[data-testid="InputInstructions"] { display: none; }
.oseec-hlabel {
    font-size: 0.85rem;
    color: #6B7686;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 0.15rem;
}
.oseec-hnum {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 3.15rem;
    line-height: 1.02;
    color: #012169;
    margin-bottom: 0.35rem;
}
.oseec-hnum span {
    font-size: 1.02rem;
    color: #6B7686;
    font-weight: 400;
}
.oseec-econtour {
    margin-top: 0.7rem;
    font-size: 0.95rem;
    color: #232830;
    background: #F7EFDC;
    border-radius: 8px;
    padding: 0.5rem 0.8rem;
}

div[data-testid="stExpander"] details {
    border: 1px solid #C9D7EA;
    border-radius: 10px;
    background: #FFFFFF;
}
div[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #012169;
}

.stButton button, .stDownloadButton button,
[data-testid="stFormSubmitButton"] button {
    border-radius: 8px;
    font-weight: 600;
}
.stButton button[kind="primary"], .stDownloadButton button[kind="primary"],
[data-testid="stFormSubmitButton"] button[kind="primary"] {
    background: #012169;
    border: 1px solid #012169;
}
.stButton button[kind="primary"]:hover, .stDownloadButton button[kind="primary"]:hover,
[data-testid="stFormSubmitButton"] button[kind="primary"]:hover {
    background: #0A2E86;
    border-color: #0A2E86;
}

thead tr th {
    background: #EDF2FA !important;
    color: #232830 !important;
}

@media (max-width: 700px) {
    div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-plate) {
        padding: 0.95rem 0.95rem 1.1rem 0.95rem !important;
    }
    .oseec-plate {
        margin: -0.95rem -0.95rem 0.15rem -0.95rem;
        padding: 0.58rem 0.95rem;
    }
    .oseec-plate .sub {
        margin-left: 0;
        max-width: 100%;
        text-align: left;
        flex-basis: 100%;
    }
    div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .oseec-band) {
        padding: 0.8rem 0.75rem 0.9rem 0.75rem !important;
    }
    .oseec-band {
        padding: 0.1rem 0.15rem 0.35rem 0.15rem;
        margin: 0 0 0.7rem 0;
    }
    .oseec-band .sub { margin-left: 0; flex-basis: 100%; }
}

[data-testid="stToolbar"] { visibility: hidden; }
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
</style>
"""


def inject_css() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def sidebar_brand() -> None:
    st.markdown(
        """
        <div class="oseec-brand">
          <div class="t1">Калькулятор ОСЭЭК</div>
          <div class="t2">Интегральный индекс социально-экономической
          эффективности коммуникаций компаний с государственным участием</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def sidebar_footer() -> None:
    st.markdown(
        """
        <div class="oseec-sidefoot">
          Авторы: Алтухов А.С., Бобылева А.З.<br>
          МГУ имени М.В. Ломоносова, факультет государственного управления<br><br>
          Свидетельство о государственной регистрации программы для ЭВМ
          № 2026663079 от 04.05.2026<br><br>
          Версия 2.1 · август 2026
        </div>
        """,
        unsafe_allow_html=True,
    )


def _plate_html(title: str, sub: str = "", num=None, light: bool = False,
                tint: bool = False, result: bool = False) -> str:
    num_html = f'<div class="num">{num}</div>' if num is not None else ""
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    cls = "oseec-plate light" if light else "oseec-plate"
    if tint:
        cls += " z"
    if result:
        cls += " result"
    return (f'<div class="{cls}">{num_html}'
            f'<div class="txt">{title}</div>{sub_html}</div>')


def step_header(num, title: str, sub: str = "", tint: bool = False) -> None:
    """Навесная плашка-заголовок шага с номером."""
    st.markdown(_plate_html(title, sub, num=num, tint=tint),
                unsafe_allow_html=True)


def card_title(title: str, sub: str = "", light: bool = False,
               tint: bool = False, result: bool = False) -> None:
    """Плашка-заголовок карточки без номера (light — светлый вариант)."""
    st.markdown(_plate_html(title, sub, light=light, tint=tint, result=result),
                unsafe_allow_html=True)


def section_band(title: str, sub: str = "") -> None:
    """Полоса-разделитель раздела страницы."""
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    st.markdown(
        f'<div class="oseec-band"><div class="txt">{title}</div>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def lead(text: str) -> None:
    """Вводный абзац страницы приглушенным кеглем."""
    st.markdown(f'<div class="oseec-lead">{text}</div>',
                unsafe_allow_html=True)


def sub_label(text: str) -> None:
    """Внутренний подзаголовок блока с синей меткой."""
    st.markdown(f'<div class="oseec-sub">{text}</div>', unsafe_allow_html=True)


def eyebrow(text: str) -> None:
    """Мелкая надпись-ярлык над группой индикаторов."""
    st.markdown(f'<div class="oseec-eyebrow">{text}</div>',
                unsafe_allow_html=True)


def level_chip(level_idx: int, text: str = None) -> str:
    accent = LEVEL_ACCENT[level_idx]
    tint = LEVEL_TINT[level_idx]
    label = text or LEVEL_NAMES[level_idx]
    return (f'<span class="oseec-chip" style="background:{tint};'
            f'border-color:{accent};">{label}</span>')


# ---------------------------------------------------------------------------
# Графики (Plotly)
# ---------------------------------------------------------------------------

_PLOT_FONT = dict(family=FONT_BODY, size=13, color=INK)


def _base_layout(fig: go.Figure, height: int) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=_PLOT_FONT,
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=LINE,
                        font=dict(family=FONT_BODY, size=12, color=INK)),
    )
    return fig


def _level_bands(fig: go.Figure, y0: float, y1: float,
                 labels: bool = True, yref: str = None) -> None:
    names_short = ("критически низкий", "низкий", "средний", "высокий", "очень высокий")
    kw = {"yref": yref} if yref else {}
    for i in range(5):
        x0, x1 = LEVEL_EDGES[i], LEVEL_EDGES[i + 1]
        fig.add_shape(type="rect", x0=x0, x1=x1, y0=y0, y1=y1,
                      fillcolor=LEVEL_TINT[i], line=dict(width=0),
                      layer="below", **kw)
        if i < 4:
            fig.add_shape(type="line", x0=x1, x1=x1, y0=y0, y1=y1,
                          line=dict(color="#FFFFFF", width=2), layer="below",
                          **kw)
        if labels:
            fig.add_annotation(x=(x0 + x1) / 2, y=y1, yshift=10,
                               text=names_short[i], showarrow=False,
                               font=dict(size=11, color=GRAY), **kw)


def fig_gauge(value: float, level_idx: int) -> go.Figure:
    """Спидометр итогового значения на интерпретационной шкале.

    Дуга разбита на пять зон качественных уровней, стрелка отмечает
    итоговое значение. Наглядный ориентир для демонстрации результата.
    """
    v = min(max(value, 0.0), 116.0)
    fig = go.Figure(go.Indicator(
        mode="gauge",
        value=v,
        gauge=dict(
            axis=dict(range=[0, 116], tickvals=[0, 25, 50, 75, 100],
                      tickwidth=1, tickcolor=LINE, ticklen=6,
                      tickfont=dict(size=11, color=GRAY)),
            bar=dict(color="rgba(0,0,0,0)"),
            bgcolor="rgba(0,0,0,0)",
            borderwidth=0,
            steps=[dict(range=[LEVEL_EDGES[i], LEVEL_EDGES[i + 1]],
                        color=GAUGE_ZONES[i]) for i in range(5)],
            threshold=dict(line=dict(color=NAVY, width=6), thickness=0.92,
                           value=v),
        ),
        domain=dict(x=[0, 1], y=[0, 1]),
    ))
    fig.add_annotation(x=0.5, y=0.30, xref="paper", yref="paper",
                       text=f"<b>{fmt(value)}</b>", showarrow=False,
                       font=dict(family=FONT_SERIF, size=33, color=NAVY))
    fig.add_annotation(x=0.5, y=0.11, xref="paper", yref="paper",
                       text=LEVEL_NAMES[level_idx] + " уровень",
                       showarrow=False, font=dict(size=12.5, color=GRAY))
    fig.update_layout(height=248, margin=dict(l=22, r=22, t=14, b=4),
                      paper_bgcolor="rgba(0,0,0,0)", font=_PLOT_FONT)
    return fig


def fig_scale(value: float, value_e: float = None,
              scenario_points: dict = None) -> go.Figure:
    """Линейка шкалы интерпретации с маркером итогового значения."""
    fig = go.Figure()
    _level_bands(fig, 0.0, 1.0)
    fig.add_shape(type="line", x0=value, x1=value, y0=-0.06, y1=1.06,
                  line=dict(color=NAVY, width=3))
    fig.add_trace(go.Scatter(
        x=[value], y=[0.5], mode="markers",
        marker=dict(symbol="diamond", size=15, color=NAVY,
                    line=dict(color="#FFFFFF", width=2)),
        hovertemplate="ОСЭЭК<sub>B</sub> = " + fmt(value) + "<extra></extra>",
        showlegend=False))
    fig.add_annotation(x=value, y=-0.06, yshift=-14, text=f"<b>{fmt(value)}</b>",
                       showarrow=False, font=dict(size=13, color=NAVY))
    if value_e is not None:
        fig.add_trace(go.Scatter(
            x=[value_e], y=[0.5], mode="markers",
            marker=dict(symbol="diamond-open", size=15, color=BLUE,
                        line=dict(width=2.5)),
            hovertemplate="ОСЭЭК<sub>E</sub> = " + fmt(value_e) + "<extra></extra>",
            showlegend=False))
        fig.add_annotation(x=value_e, y=1.06, yshift=13,
                           text=f"E: {fmt(value_e)}", showarrow=False,
                           font=dict(size=11.5, color=BLUE))
    if scenario_points:
        for lbl, x in scenario_points.items():
            fig.add_trace(go.Scatter(
                x=[x], y=[0.18], mode="markers",
                marker=dict(symbol="line-ns", size=11, color=GRAY,
                            line=dict(color=GRAY, width=2)),
                hovertemplate=f"{lbl}: {fmt(x)}<extra></extra>",
                showlegend=False))
    fig.update_xaxes(range=[0, 116], showgrid=False, zeroline=False,
                     tickvals=[0, 25, 50, 75, 100], ticks="outside",
                     tickcolor=LINE, tickfont=dict(size=11, color=GRAY))
    fig.update_yaxes(range=[-0.35, 1.45], visible=False, fixedrange=True)
    return _base_layout(fig, 150)


def fig_components(res: dict) -> go.Figure:
    """Декомпозиция: компоненты и субиндексы по стобалльной шкале."""
    rows = [
        ("Медийная устойчивость M<sub>stab</sub>", res["m_stab"], NAVY),
        ("Социальная репутация S<sub>rep</sub>", res["s_rep"], NAVY),
        ("Верификация HR-бренда V<sub>hr</sub>", res["v_hr"], BLUE),
        ("Транспарентность R<sub>transp</sub>", res["r_transp"], BLUE),
        ("Институциональная зрелость R<sub>inst</sub>", res["r_inst"], BLUE),
    ]
    rows = rows[::-1]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=[r[0] for r in rows],
        x=[r[1] for r in rows],
        orientation="h",
        marker=dict(color=[r[2] for r in rows],
                    cornerradius=4,
                    line=dict(color="#FFFFFF", width=1)),
        width=0.55,
        text=[f"<b>{fmt(r[1])}</b>" for r in rows],
        textposition="outside",
        textfont=dict(size=12.5, color=INK),
        cliponaxis=False,
        hovertemplate="%{y}: %{text}<extra></extra>",
        showlegend=False,
    ))
    fig.update_xaxes(range=[0, 119], showgrid=True, gridcolor="#E9EFF8",
                     zeroline=False, tickvals=[0, 25, 50, 75, 100],
                     tickfont=dict(size=11, color=GRAY))
    fig.update_yaxes(showgrid=False, tickfont=dict(size=12.5, color=INK))
    fig.add_shape(type="line", x0=0, x1=0, y0=-0.5, y1=4.5,
                  line=dict(color=LINE, width=1))
    return _base_layout(fig, 260)


def fig_core_structure(res: dict) -> go.Figure:
    """Структура ядра: вклады взвешенных субиндексов."""
    c_m = 0.6 * res["m_stab"]
    c_s = 0.4 * res["s_rep"]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=[""], x=[c_m], orientation="h", name="0,6 × M<sub>stab</sub>",
        marker=dict(color=NAVY, cornerradius=3),
        width=0.5,
        text=f"{fmt(c_m)}" if c_m >= 9 else "", textposition="inside",
        textangle=0, insidetextanchor="middle",
        insidetextfont=dict(color="#FFFFFF", size=11.5),
        hovertemplate="Вклад медийной устойчивости: " + fmt(c_m) + "<extra></extra>"))
    fig.add_trace(go.Bar(
        y=[""], x=[c_s], orientation="h", name="0,4 × S<sub>rep</sub>",
        marker=dict(color=BLUE, cornerradius=3),
        width=0.5,
        text=f"{fmt(c_s)}" if c_s >= 9 else "", textposition="inside",
        textangle=0, insidetextanchor="middle",
        insidetextfont=dict(color="#FFFFFF", size=11.5),
        hovertemplate="Вклад социальной репутации: " + fmt(c_s) + "<extra></extra>"))
    fig.add_annotation(x=res["i_core"], y=0, xshift=8, xanchor="left",
                       text=f"I<sub>Core</sub> = <b>{fmt(res['i_core'])}</b>",
                       showarrow=False, font=dict(size=13, color=INK))
    fig.update_layout(barmode="stack", bargap=0.1,
                      legend=dict(orientation="h", y=-0.5, x=0,
                                  traceorder="normal",
                                  font=dict(size=12, color=INK)))
    fig.update_xaxes(range=[0, 112], showgrid=False, zeroline=False,
                     tickvals=[0, 25, 50, 75, 100],
                     tickfont=dict(size=11, color=GRAY))
    fig.update_yaxes(visible=False, fixedrange=True)
    return _base_layout(fig, 130)


def fig_mc_hist(values, base_val: float, median_val: float) -> go.Figure:
    """Распределение значений имитационного прогона."""
    fig = go.Figure()
    _level_bands(fig, 0, 1, labels=True, yref="paper")
    fig.add_trace(go.Histogram(
        x=values, nbinsx=60,
        marker=dict(color=BLUE, opacity=0.8,
                    line=dict(color="#FFFFFF", width=1)),
        histnorm="probability",
        hovertemplate="Диапазон: %{x}<br>Доля итераций: %{y:.3f}<extra></extra>",
        showlegend=False))
    fig.add_shape(type="line", x0=base_val, x1=base_val, y0=0, y1=0.97,
                  yref="paper", line=dict(color=NAVY, width=2.5))
    fig.add_annotation(x=base_val, y=0.97, yref="paper",
                       text=f"базовый расчет {fmt(base_val)}", showarrow=False,
                       font=dict(size=11.5, color=NAVY), xanchor="left", xshift=5)
    fig.add_shape(type="line", x0=median_val, x1=median_val, y0=0, y1=0.85,
                  yref="paper", line=dict(color=GRAY, width=2, dash="dot"))
    fig.add_annotation(x=median_val, y=0.85, yref="paper",
                       text=f"медиана {fmt(median_val)}", showarrow=False,
                       font=dict(size=11.5, color=GRAY), xanchor="left", xshift=5)
    fig.update_xaxes(range=[0, 116], showgrid=False, zeroline=False,
                     tickvals=[0, 25, 50, 75, 100],
                     tickfont=dict(size=11, color=GRAY))
    fig.update_yaxes(showgrid=True, gridcolor="#E9EFF8", zeroline=False,
                     tickfont=dict(size=11, color=GRAY),
                     title=dict(text="доля итераций",
                                font=dict(size=11.5, color=GRAY)))
    fig.update_layout(bargap=0.03)
    return _base_layout(fig, 320)


PLOTLY_CONFIG = {"displayModeBar": False, "locale": "ru"}
