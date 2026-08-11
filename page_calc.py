# -*- coding: utf-8 -*-
"""Страница «Калькулятор» — расчет индекса ОСЭЭК по данным пользователя."""

import streamlit as st

import dadata_api
import inputs_state as inp
import oseec_core as oc
import ui_theme as ui
from oseec_core import fmt
from report_docx import build_report

inp.ensure_defaults()

st.title("Расчет индекса ОСЭЭК")
st.markdown(
    "Калькулятор применим к любой организации: введите значения показателей "
    "по блокам расчетной модели — компоненты, субиндексы, поправочные "
    "коэффициенты и итоговое значение индекса рассчитываются автоматически. "
    "Для случаев неполноты открытых данных предусмотрены сценарии, заданные "
    "методикой."
)

def _do_lookup() -> None:
    info = dadata_api.lookup_inn(st.session_state.get("c_inn", ""))
    st.session_state["c_info"] = info
    if "error" in info:
        return
    st.session_state["c_name"] = info["name"]
    emp = info.get("employee_count")
    if isinstance(emp, int) and emp > 0:
        if emp > 100_000:
            st.session_state["kscale"] = list(inp.SCALE_OPTIONS)[0]
        elif emp >= 10_000:
            st.session_state["kscale"] = list(inp.SCALE_OPTIONS)[1]
        else:
            st.session_state["kscale"] = list(inp.SCALE_OPTIONS)[2]


nc1, nc2, nc3 = st.columns([2.3, 1.1, 0.85], vertical_alignment="bottom")
with nc1:
    st.text_input(
        "Наименование организации",
        key="c_name",
        placeholder="Введите название или найдите компанию по ИНН",
    )
with nc2:
    st.text_input("ИНН", key="c_inn", placeholder="10 или 12 цифр")
with nc3:
    st.button("Найти по ИНН", on_click=_do_lookup, use_container_width=True,
              help="Название организации и справочные сведения подставятся "
                   "автоматически")

info = st.session_state.get("c_info")
if info:
    if "error" in info:
        st.markdown(f"<div class='oseec-note'>{info['error']}</div>",
                    unsafe_allow_html=True)
    else:
        parts = [f"<b>{info['full_name']}</b>", f"ИНН {info['inn']}"]
        if info.get("okved"):
            sec = f" — {info['okved_section']}" if info.get("okved_section") else ""
            parts.append(f"основной ОКВЭД {info['okved']}{sec}")
        if info.get("region"):
            parts.append(info["region"])
        if isinstance(info.get("employee_count"), int) and info["employee_count"] > 0:
            parts.append(f"численность по справочнику: "
                         f"{info['employee_count']:,} чел.".replace(",", " ")
                         + " — ступень K_scale на шаге 3 выбрана "
                           "автоматически, проверьте ее")
        status_note = ""
        if info.get("status") and info["status"] != "ACTIVE":
            status_note = ("<br><b>Внимание:</b> по данным справочника "
                           "организация не имеет статуса действующей.")
        st.markdown(
            f"<div class='oseec-note'>{' · '.join(parts)}{status_note}<br>"
            "<span style='color:#6B7686'>Сведения справочные и в расчет "
            "индекса не входят.</span></div>",
            unsafe_allow_html=True)

bc1, bc2, _ = st.columns([1.35, 0.75, 2.1])
with bc1:
    st.button("Демонстрационный пример", use_container_width=True,
              on_click=inp.load_demo_into_state,
              help="Заполняет форму условными данными вымышленной компании, "
                   "чтобы показать порядок работы")
with bc2:
    st.button("Очистить", use_container_width=True,
              on_click=inp.clear_all_inputs)

# ---------------------------------------------------------------------------
ui.step_header(1, "Медийная устойчивость",
               "субиндекс M_stab — текущее информационное присутствие")

st.radio(
    "Источник медиаданных",
    options=[inp.TRACK_MANUAL, inp.TRACK_MONITOR, inp.TRACK_NONE],
    key="c_track",
    captions=[
        "Формула (2): кодирование до 100 уникальных публикаций за отчетный "
        "год по тональности; порог репрезентативности — 12 публикаций",
        "Формула (1): годовое значение медиапоказателя сопоставляется с "
        "отраслевым эталоном — средним пиковых значений трех крупнейших "
        "компаний отрасли за три года",
        "Правило нулевого значения: субиндекс принимается равным 0,00 и "
        "получает статус критического ограничения",
    ],
)

track = st.session_state["c_track"]

if track == inp.TRACK_MANUAL:
    st.markdown("**Помесячное распределение публикаций верифицированного корпуса**")
    st.caption(
        "По каждому месяцу укажите общее число публикаций и число негативных; "
        "позитивные и нейтральные материалы учитываются суммарно."
    )
    for half in (0, 1):
        cols = st.columns(6)
        for j in range(6):
            i = half * 6 + j
            with cols[j]:
                st.number_input(f"{oc.MONTHS[i]} — всего", min_value=0,
                                max_value=1000, step=1, key=f"mt{i}")
                st.number_input("из них негативных", min_value=0,
                                max_value=1000, step=1, key=f"mn{i}",
                                label_visibility="visible")
elif track == inp.TRACK_MONITOR:
    c1, c2 = st.columns(2)
    with c1:
        st.number_input(
            "Годовое значение медиапоказателя компании (X_fact)",
            min_value=0.0, step=100.0, key="xfact")
    with c2:
        st.number_input(
            "Отраслевой эталон (X_ref)",
            min_value=0.0, step=100.0, key="xref",
            help="Среднее арифметическое пиковых значений трех крупнейших "
                 "компаний отрасли за последние три года по данным той же "
                 "системы мониторинга.")
    st.markdown("**Помесячные значения медиапоказателя** — для коэффициента "
                "волатильности")
    for half in (0, 1):
        cols = st.columns(6)
        for j in range(6):
            i = half * 6 + j
            with cols[j]:
                st.number_input(oc.MONTHS[i], min_value=0.0, step=10.0,
                                key=f"mm{i}")
else:
    st.markdown(
        '<div class="oseec-note">Открытые источники не содержат публикаций, '
        "атрибутированных непосредственно к анализируемому юридическому лицу, "
        "либо объем корпуса ниже порога репрезентативности. Субиндекс "
        "медийной устойчивости принимается равным 0,00 — нижняя граница "
        "внешней наблюдаемости без присвоения отрицательной оценки.</div>",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
ui.step_header(2, "Социальная репутация",
               "субиндекс S_rep — накопленный репутационный капитал")

st.markdown("**Верификация HR-бренда (V_hr)**")
st.radio("Положение компании в открытых рейтингах работодателей",
         options=[inp.HR_IN, inp.HR_SCEN], key="hr_mode",
         label_visibility="collapsed")

if st.session_state["hr_mode"] == inp.HR_IN:
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        st.number_input("Число участников рейтинга (N)", min_value=2,
                        max_value=100000, step=1, key="hr_total")
    with c2:
        st.number_input("Позиция компании (Rank)", min_value=1,
                        max_value=int(st.session_state["hr_total"]),
                        step=1, key="hr_rank")
    with c3:
        vhr_prev = oc.v_hr(int(st.session_state["hr_rank"]),
                           int(st.session_state["hr_total"]))
        st.metric("V_hr по формуле (5)", fmt(vhr_prev))
else:
    st.radio(
        "Сценарное значение компонента",
        options=list(inp.SCEN_VALUES),
        key="hr_scen",
        captions=[
            "Консервативная граница диапазона",
            "Базовый сценарий, применяемый в расчетах диссертационного исследования",
            "Верхняя граница диапазона",
        ],
    )
    st.caption(
        "Отсутствие самостоятельной позиции юридического лица в рейтингах "
        "работодателей фиксируется как ограничение источниковой базы. Итог "
        "рассчитывается по выбранному сценарию, влияние крайних значений "
        "диапазона показывается в результатах."
    )

st.markdown("**Транспарентность (R_transp)** — чек-лист из десяти индикаторов")
tc1, tc2 = st.columns(2)
with tc1:
    st.markdown("Блок 1. Корпоративная открытость")
    for i, label in enumerate(oc.TRANSP_BLOCK1):
        st.checkbox(label, key=f"t1_{i}")
with tc2:
    st.markdown("Блок 2. Компенсаторное раскрытие социально значимой информации")
    for i, label in enumerate(oc.TRANSP_BLOCK2):
        st.checkbox(label, key=f"t2_{i}")
st.caption(
    "Индикатор получает значение 1 при подтверждении раскрытия официальными "
    "источниками на уровне самой компании. Второй блок операционализирует "
    "понятие компенсаторной транспарентности: раскрытие социально значимых "
    "последствий деятельности в сферах, где публичное информирование допустимо."
)

st.markdown("**Институциональная зрелость (R_inst)** — чек-лист из шести признаков")
ic1, ic2 = st.columns(2)
with ic1:
    st.markdown("Блок 1. Каналы и процедуры обратной связи")
    for i, label in enumerate(oc.INST_BLOCK1):
        st.checkbox(label, key=f"i1_{i}")
with ic2:
    st.markdown("Блок 2. Институциональное закрепление коммуникационной функции")
    for i, label in enumerate(oc.INST_BLOCK2):
        st.checkbox(label, key=f"i2_{i}")
st.caption(
    "Оба чек-листа агрегируются в две ступени: среднее значение индикаторов "
    "внутри блока, затем среднее двух блоков по шкале от 0 до 100. Такой "
    "порядок снижает риск двойного учета близких по смыслу признаков."
)

# ---------------------------------------------------------------------------
ui.step_header(3, "Поправочные коэффициенты",
               "институциональные условия деятельности")

cc1, cc2 = st.columns(2)
with cc1:
    st.markdown("**Риск коммуникационной среды (K_risk)** — каскадное правило")
    st.radio("K_risk", options=list(inp.RISK_OPTIONS), key="krisk",
             label_visibility="collapsed",
             captions=[
                 "Ограничения, прямо затрагивающие раскрытие сведений о "
                 "технологиях, производственных процессах, контрактах или "
                 "продукции",
                 "Прямых ограничений нет, однако компания включена в перечень "
                 "стратегических организаций либо действует на территории с "
                 "особым правовым режимом",
                 "Специальные правовые, отраслевые или территориальные "
                 "условия, заметно влияющие на раскрытие информации, не "
                 "выявлены",
             ])
    st.caption("Категория присваивается по наиболее жесткому из применимых "
               "условий, одновременно действующие режимы не суммируются.")
with cc2:
    st.markdown("**Масштаб организации (K_scale)** — по численности персонала")
    st.radio("K_scale", options=list(inp.SCALE_OPTIONS), key="kscale",
             label_visibility="collapsed",
             captions=[
                 "Крупная организация с разветвленной системой внутренних и "
                 "внешних коммуникаций",
                 "Организация среднего или крупного масштаба со стандартной "
                 "поправкой",
                 "Организация меньшего масштаба",
                 "Нижняя граница внешней оценки масштаба: более высокая "
                 "поправка без подтвержденных данных не применяется",
             ])

# ---------------------------------------------------------------------------
ui.step_header(4, "Расширенный контур",
               "при наличии данных управленческого учета")

st.toggle("Рассчитать расширенный контур ОСЭЭК_E", key="ext_on")
if st.session_state["ext_on"]:
    st.caption(
        "Поправки задаются по ступеням таблицы условий методики: методика "
        "фиксирует, подтверждает ли управленческая отчетность экономическую "
        "и социальную отдачу, а также качество бюджетного исполнения."
    )
    ec1, ec2, ec3 = st.columns(3)
    with ec1:
        st.markdown("**Экономическая отдача (k_roi)**")
        st.radio("k_roi", options=list(inp.ROI_OPTIONS), key="kroi",
                 label_visibility="collapsed")
    with ec2:
        st.markdown("**Социальная отдача (k_sroi)**")
        st.radio("k_sroi", options=list(inp.SROI_OPTIONS), key="ksroi",
                 label_visibility="collapsed")
    with ec3:
        st.markdown("**Бюджетная дисциплина (k_budget)**")
        st.radio("k_budget", options=list(inp.BUDGET_OPTIONS), key="kbud",
                 label_visibility="collapsed")

# ---------------------------------------------------------------------------
# Расчет и результаты
# ---------------------------------------------------------------------------

base = inp.collect_base_inputs()
ext = inp.collect_ext_inputs()
res = oc.compute(base, ext)

st.divider()
st.header("Результаты расчета")

company = st.session_state["c_name"].strip()
subtitle = ("Базовый контур ОСЭЭК · " + company) if company else \
    "Базовый контур ОСЭЭК — внешняя оценка по открытым источникам"

extra = ""
if "oseec_e" in res:
    extra = (f'<div style="margin-top:0.55rem;font-size:0.95rem;color:{ui.INK};">'
             f'Расширенный контур: ОСЭЭК<sub>E</sub> = ОСЭЭК<sub>B</sub> × '
             f'{fmt(res["k_eff"])} = <b>{fmt(res["oseec_e"])}</b> балла — '
             f'{res["level_e_name"].lower()} уровень</div>')

ui.hero_result(res["oseec_b"], res["level"], subtitle, extra)

scen_points = None
if res.get("hr_scenarios"):
    scen_points = {f"V_hr = {int(k)}": v for k, v in res["hr_scenarios"].items()}
st.plotly_chart(
    ui.fig_scale(res["oseec_b"], res.get("oseec_e"), scen_points),
    use_container_width=True, config=ui.PLOTLY_CONFIG)

# Критические ограничения и примечания
if res["critical"]:
    items = "".join(
        f"<div class='oseec-crit'><b>Критическое ограничение:</b> {kind} "
        f"«{name}» — {fmt(val)} балла при пороговом уровне "
        f"{'40' if kind == 'субиндекс' else '30'} баллов.</div>"
        for kind, name, val in res["critical"])
    st.markdown(items, unsafe_allow_html=True)
    st.caption(
        "Расчетная величина индекса при присвоении статуса не изменяется: "
        "слабый результат уже отражен в формуле. Статус дополняет "
        "качественную характеристику оценки указанием на проблемную зону."
    )
for note in res["notes"]:
    st.markdown(f"<div class='oseec-note'>{note}</div>", unsafe_allow_html=True)

# Сценарный диапазон V_hr
if res.get("hr_scenarios"):
    s0, s50, s100 = (res["hr_scenarios"][0.0], res["hr_scenarios"][50.0],
                     res["hr_scenarios"][100.0])
    st.markdown("**Чувствительность к сценарию V_hr**")
    m1, m2, m3 = st.columns(3)
    m1.metric("ОСЭЭК_B при V_hr = 0", fmt(s0),
              delta=fmt(s0 - res["oseec_b"]), delta_color="normal")
    m2.metric("ОСЭЭК_B при V_hr = 50", fmt(s50))
    m3.metric("ОСЭЭК_B при V_hr = 100", fmt(s100),
              delta=fmt(s100 - res["oseec_b"]), delta_color="normal")

st.subheader("Декомпозиция оценки")
d1, d2 = st.columns([3, 2])
with d1:
    st.markdown("Компоненты и субиндексы, баллы")
    st.plotly_chart(ui.fig_components(res), use_container_width=True,
                    config=ui.PLOTLY_CONFIG)
with d2:
    st.markdown("Структура ядра индекса")
    st.plotly_chart(ui.fig_core_structure(res), use_container_width=True,
                    config=ui.PLOTLY_CONFIG)
    st.markdown(
        f"""
        <div class="oseec-form">I<sub>Core</sub> = 0,6 × {fmt(res['m_stab'])} +
        0,4 × {fmt(res['s_rep'])} = <b>{fmt(res['i_core'])}</b></div>
        <div class="oseec-form">ОСЭЭК<sub>B</sub> = {fmt(res['i_core'])} ×
        {fmt(res['k_risk'])} × {fmt(res['k_scale'])} =
        <b>{fmt(res['oseec_b'])}</b></div>
        """,
        unsafe_allow_html=True)
    if "oseec_e" in res:
        st.markdown(
            f"""
            <div class="oseec-form">K<sub>eff</sub> = 1 + {fmt(ext.k_roi)} +
            {fmt(ext.k_sroi)} + {fmt(ext.k_budget)} =
            <b>{fmt(res['k_eff'])}</b></div>
            <div class="oseec-form">ОСЭЭК<sub>E</sub> = {fmt(res['oseec_b'])} ×
            {fmt(res['k_eff'])} = <b>{fmt(res['oseec_e'])}</b></div>
            """,
            unsafe_allow_html=True)

if res.get("i_media") is not None:
    st.caption(
        f"Медийный блок: I_media = {fmt(res['i_media'])}, V_vol = "
        f"{fmt(res['v_vol'])}, M_stab = I_media / (1 + V_vol) = "
        f"{fmt(res['m_stab'])}."
        + (f" Корпус: {res['n_total']} публикаций, из них негативных — "
           f"{res['n_neg']}." if res.get("n_total") else ""))

# Выгрузка и переходы
st.divider()
b1, b2, _ = st.columns([1.2, 1.2, 2])
with b1:
    st.download_button(
        "Скачать протокол расчета (Word)",
        data=build_report(res, company, base),
        file_name="Протокол_расчета_ОСЭЭК.docx",
        mime=("application/vnd.openxmlformats-officedocument"
              ".wordprocessingml.document"),
        type="primary",
        use_container_width=True)
with b2:
    if st.button("Проверить устойчивость результата", use_container_width=True):
        st.switch_page("page_mc.py")
