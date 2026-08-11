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
.block-container {
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

.oseec-step {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    margin: 1.6rem 0 0.6rem 0;
}
.oseec-step .num {
    background: #012169;
    color: #FFFFFF;
    font-weight: 600;
    font-size: 0.95rem;
    width: 1.9rem;
    height: 1.9rem;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 auto;
}
.oseec-step .txt {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 1.18rem;
    color: #012169;
}
.oseec-step .sub {
    font-size: 0.86rem;
    color: #6B7686;
    margin-left: 0.2rem;
}

.oseec-hero {
    border: 1px solid #C9D7EA;
    border-left: 5px solid #012169;
    border-radius: 14px;
    background: linear-gradient(180deg, #FFFFFF 0%, #F4F8FC 100%);
    padding: 1.35rem 1.7rem 1.25rem 1.7rem;
    margin: 0.4rem 0 0.9rem 0;
}
.oseec-hero .lbl {
    font-size: 0.86rem;
    color: #6B7686;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
.oseec-hero .val {
    font-family: 'PT Serif', Georgia, serif;
    font-weight: 700;
    font-size: 3.3rem;
    line-height: 1.05;
    color: #012169;
}
.oseec-hero .unit {
    font-size: 1.05rem;
    color: #6B7686;
    font-weight: 400;
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
    border: 1px solid #C9D7EA;
    border-radius: 10px;
    background: #F4F8FC;
    color: #232830;
    font-size: 0.88rem;
    line-height: 1.55;
    padding: 0.7rem 1rem;
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

[data-testid="stVerticalBlockBorderWrapper"] > div {
    border-color: #C9D7EA !important;
    border-radius: 12px !important;
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

.stButton button, .stDownloadButton button {
    border-radius: 8px;
    font-weight: 600;
}
.stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {
    background: #012169;
    border: 1px solid #012169;
}
.stButton button[kind="primary"]:hover, .stDownloadButton button[kind="primary"]:hover {
    background: #0A2E86;
    border-color: #0A2E86;
}

thead tr th {
    background: #EDF2FA !important;
    color: #232830 !important;
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
          Версия 2.0 · август 2026
        </div>
        """,
        unsafe_allow_html=True,
    )


def step_header(num, title: str, sub: str = "") -> None:
    sub_html = f'<span class="sub">{sub}</span>' if sub else ""
    st.markdown(
        f'<div class="oseec-step"><div class="num">{num}</div>'
        f'<div class="txt">{title}</div>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def level_chip(level_idx: int, text: str = None) -> str:
    accent = LEVEL_ACCENT[level_idx]
    tint = LEVEL_TINT[level_idx]
    label = text or LEVEL_NAMES[level_idx]
    return (f'<span class="oseec-chip" style="background:{tint};'
            f'border-color:{accent};">{label}</span>')


def hero_result(value: float, level_idx: int, subtitle: str,
                extra_html: str = "") -> None:
    st.markdown(
        f"""
        <div class="oseec-hero">
          <div class="lbl">{subtitle}</div>
          <div style="display:flex;align-items:baseline;gap:1.1rem;flex-wrap:wrap;">
            <div class="val">{fmt(value)}<span class="unit"> балла</span></div>
            <div>{level_chip(level_idx, LEVEL_NAMES[level_idx] + " уровень")}</div>
          </div>
          {extra_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


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
