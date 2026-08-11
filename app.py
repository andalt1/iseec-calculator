# -*- coding: utf-8 -*-
"""Калькулятор ОСЭЭК — интегральный индекс социально-экономической
эффективности коммуникаций компаний с государственным участием.

Авторы методики: Алтухов А.С., Бобылева А.З.
МГУ имени М.В. Ломоносова, факультет государственного управления.
Свидетельство о государственной регистрации программы для ЭВМ
№ 2026663079 от 04.05.2026.
"""

import streamlit as st

import inputs_state as inp
import ui_theme as ui

st.set_page_config(
    page_title="Калькулятор ОСЭЭК",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="auto",
    menu_items={"about": "Калькулятор ОСЭЭК · свидетельство о государственной "
                         "регистрации программы для ЭВМ № 2026663079"},
)

ui.inject_css()
inp.ensure_defaults()

pages = [
    st.Page("page_calc.py", title="Калькулятор",
            icon=":material/calculate:", default=True),
    st.Page("page_mc.py", title="Проверка устойчивости",
            icon=":material/query_stats:"),
    st.Page("page_method.py", title="Методика",
            icon=":material/menu_book:"),
    st.Page("page_about.py", title="О программе",
            icon=":material/workspace_premium:"),
]

with st.sidebar:
    ui.sidebar_brand()

nav = st.navigation(pages)

with st.sidebar:
    ui.sidebar_footer()

nav.run()
