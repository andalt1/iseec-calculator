# -*- coding: utf-8 -*-
"""Страница «Калькулятор» — расчет индекса ОСЭЭК по данным пользователя."""

import html
import time

import streamlit as st

import dadata_api
import inputs_state as inp
import oseec_core as oc
import ui_theme as ui
from oseec_core import LEVEL_NAMES, fmt
from report_docx import build_report

inp.ensure_defaults()

st.title("Расчет индекса ОСЭЭК")
ui.lead(
    "Укажите объект оценки и заполните четыре шага — субиндексы, поправочные "
    "коэффициенты и итоговое значение индекса рассчитываются автоматически. "
    "Для случаев неполноты открытых данных предусмотрены сценарии, заданные "
    "методикой."
)


def _run_calc() -> None:
    st.session_state["show_results"] = True


def _do_lookup() -> None:
    if not st.session_state.get("c_inn", "").strip():
        st.session_state["c_info"] = None
        return
    now = time.time()
    if now - st.session_state.get("_last_lookup_ts", 0.0) < 2.0:
        return
    st.session_state["_last_lookup_ts"] = now
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


# --- Карточка «Объект оценки» ----------------------------------------------
with st.container(border=True):
    ui.card_title("Объект оценки",
                  "введите ИНН и нажмите Enter — название подставится")
    nc1, nc2 = st.columns([2.3, 1.95], vertical_alignment="bottom")
    with nc1:
        st.text_input(
            "Наименование организации",
            key="c_name",
            placeholder="Введите название или найдите компанию по ИНН",
        )
    with nc2:
        with st.form("inn_form", border=False):
            fc1, fc2 = st.columns([1.1, 0.85], vertical_alignment="bottom")
            with fc1:
                st.text_input("ИНН", key="c_inn",
                              placeholder="10 или 12 цифр")
            with fc2:
                st.form_submit_button(
                    "Найти по ИНН", on_click=_do_lookup,
                    use_container_width=True)

    info = st.session_state.get("c_info")
    if info:
        if "error" in info:
            st.markdown(
                f"<div class='oseec-note'>{html.escape(str(info['error']))}"
                "</div>",
                unsafe_allow_html=True)
        else:
            parts = [f"<b>{html.escape(str(info['full_name']))}</b>",
                     f"ИНН {html.escape(str(info['inn']))}"]
            if info.get("okved"):
                sec = (f" — {html.escape(str(info['okved_section']))}"
                       if info.get("okved_section") else "")
                parts.append(
                    f"основной ОКВЭД {html.escape(str(info['okved']))}{sec}")
            if info.get("region"):
                parts.append(html.escape(str(info["region"])))
            if (isinstance(info.get("employee_count"), int)
                    and info["employee_count"] > 0):
                parts.append(
                    f"численность по справочнику: "
                    f"{info['employee_count']:,} чел.".replace(",", " ")
                    + " — ступень K_scale на шаге 3 выбрана автоматически, "
                      "проверьте ее")
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
                  on_click=inp.load_demo_into_state)
    with bc2:
        st.button("Очистить", use_container_width=True,
                  on_click=inp.clear_all_inputs)


# --- Шаг 1. Медийная устойчивость ------------------------------------------
with st.container(border=True):
    ui.step_header(1, "Медийная устойчивость",
                   "субиндекс M_stab — текущее информационное присутствие",
                   tint=True)

    st.radio(
        "Источник медиаданных",
        options=[inp.TRACK_MANUAL, inp.TRACK_MONITOR, inp.TRACK_NONE],
        key="c_track",
        help="Ручной протокол: кодирование до 100 уникальных публикаций за "
             "отчетный год по тональности, порог репрезентативности — "
             "12 публикаций. Система мониторинга: отраслевой эталон — среднее "
             "пиковых значений трех крупнейших компаний отрасли за три года "
             "по данным той же системы.",
        captions=[
            "Кодирование корпуса публикаций по тональности — формула (2)",
            "Сопоставление с отраслевым эталоном — формула (1)",
            "Субиндекс 0,00 со статусом критического ограничения",
        ],
    )

    track = st.session_state["c_track"]

    if track == inp.TRACK_MANUAL:
        ui.sub_label("Помесячное распределение публикаций верифицированного "
                     "корпуса")
        st.caption(
            "По каждому месяцу — публикаций всего и из них негативных; "
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
        _tot_all = sum(int(st.session_state[f"mt{i}"]) for i in range(12))
        _over = [oc.MONTHS[i] for i in range(12)
                 if int(st.session_state[f"mn{i}"])
                 > int(st.session_state[f"mt{i}"])]
        if _tot_all > 100:
            st.markdown(
                f"<div class='oseec-note'><b>Проверьте объем корпуса.</b> "
                f"Введено {_tot_all} публикаций, тогда как методика "
                "предусматривает верифицированный корпус объемом до 100 "
                "уникальных материалов за отчетный год; более крупные массивы "
                "сокращаются до этого объема по протоколу кодирования.</div>",
                unsafe_allow_html=True)
        if _over:
            st.markdown(
                "<div class='oseec-note'><b>Негативных больше, чем всего "
                "публикаций</b> — в расчет негативные приняты равными общему "
                "числу месяца: " + ", ".join(_over) + ".</div>",
                unsafe_allow_html=True)
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
        ui.sub_label("Помесячные значения медиапоказателя — для коэффициента "
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
            '<div class="oseec-note">Открытые источники не содержат '
            "публикаций, атрибутированных непосредственно к анализируемому "
            "юридическому лицу, либо объем корпуса ниже порога "
            "репрезентативности. Субиндекс медийной устойчивости принимается "
            "равным 0,00 — нижняя граница внешней наблюдаемости без "
            "присвоения отрицательной оценки.</div>",
            unsafe_allow_html=True,
        )


# --- Шаг 2. Социальная репутация -------------------------------------------
with st.container(border=True):
    ui.step_header(2, "Социальная репутация",
                   "субиндекс S_rep — накопленный репутационный капитал")

    ui.sub_label("Верификация HR-бренда (V_hr)")
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
                "Базовый сценарий, применяемый в расчетах диссертационного "
                "исследования",
                "Верхняя граница диапазона",
            ],
        )
        st.caption(
            "Итог рассчитывается по выбранному сценарию, влияние крайних "
            "значений диапазона показывается в результатах."
        )

    ui.sub_label("Транспарентность (R_transp) — чек-лист из десяти "
                 "индикаторов")
    tc1, tc2 = st.columns(2)
    with tc1:
        ui.eyebrow("Блок 1 · Корпоративная открытость")
        for i, label in enumerate(oc.TRANSP_BLOCK1):
            st.checkbox(label, key=f"t1_{i}")
    with tc2:
        ui.eyebrow("Блок 2 · Компенсаторное раскрытие социально значимой "
                   "информации")
        for i, label in enumerate(oc.TRANSP_BLOCK2):
            st.checkbox(label, key=f"t2_{i}")

    ui.sub_label("Институциональная зрелость (R_inst) — чек-лист из шести "
                 "признаков")
    ic1, ic2 = st.columns(2)
    with ic1:
        ui.eyebrow("Блок 1 · Каналы и процедуры обратной связи")
        for i, label in enumerate(oc.INST_BLOCK1):
            st.checkbox(label, key=f"i1_{i}")
    with ic2:
        ui.eyebrow("Блок 2 · Институциональное закрепление коммуникационной "
                   "функции")
        for i, label in enumerate(oc.INST_BLOCK2):
            st.checkbox(label, key=f"i2_{i}")

    with st.expander("Правила оценивания индикаторов"):
        st.markdown(
            "Индикатор получает значение 1 при подтверждении раскрытия "
            "официальными источниками на уровне самой компании. Блок "
            "компенсаторного раскрытия операционализирует понятие "
            "компенсаторной транспарентности — раскрытие социально значимых "
            "последствий деятельности в сферах, где публичное информирование "
            "допустимо. Оба чек-листа агрегируются в две ступени: среднее "
            "значение индикаторов внутри блока, затем среднее двух блоков по "
            "шкале от 0 до 100. Такой порядок снижает риск двойного учета "
            "близких по смыслу признаков."
        )


# --- Шаг 3. Поправочные коэффициенты ---------------------------------------
with st.container(border=True):
    ui.step_header(3, "Поправочные коэффициенты",
                   "институциональные условия деятельности",
                   tint=True)

    cc1, cc2 = st.columns(2)
    with cc1:
        ui.sub_label("Риск коммуникационной среды (K_risk)")
        st.radio("K_risk", options=list(inp.RISK_OPTIONS), key="krisk",
                 label_visibility="collapsed",
                 captions=[
                     "Ограничения, прямо затрагивающие раскрытие сведений о "
                     "технологиях, производственных процессах, контрактах или "
                     "продукции",
                     "Прямых ограничений нет, однако компания включена в "
                     "перечень стратегических организаций либо действует на "
                     "территории с особым правовым режимом",
                     "Специальные правовые, отраслевые или территориальные "
                     "условия, заметно влияющие на раскрытие информации, не "
                     "выявлены",
                 ])
        st.caption("Категория присваивается по наиболее жесткому из применимых "
                   "условий, одновременно действующие режимы не суммируются.")
    with cc2:
        ui.sub_label("Масштаб организации (K_scale) — по численности "
                     "персонала")
        st.radio("K_scale", options=list(inp.SCALE_OPTIONS), key="kscale",
                 label_visibility="collapsed",
                 captions=[
                     "Крупная организация с разветвленной системой внутренних "
                     "и внешних коммуникаций",
                     "Организация среднего или крупного масштаба со "
                     "стандартной поправкой",
                     "Организация меньшего масштаба",
                     "Нижняя граница внешней оценки масштаба: более высокая "
                     "поправка без подтвержденных данных не применяется",
                 ])


# --- Шаг 4. Расширенный контур ---------------------------------------------
with st.container(border=True):
    ui.step_header(4, "Расширенный контур",
                   "при наличии данных управленческого учета")

    st.toggle("Рассчитать расширенный контур ОСЭЭК_E", key="ext_on",
              help="Поправки задаются по ступеням таблицы условий методики: "
                   "фиксируется, подтверждает ли управленческая отчетность "
                   "экономическую и социальную отдачу, а также качество "
                   "бюджетного исполнения.")
    if st.session_state["ext_on"]:
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            ui.sub_label("Экономическая отдача (k_roi)")
            st.radio("k_roi", options=list(inp.ROI_OPTIONS), key="kroi",
                     label_visibility="collapsed")
        with ec2:
            ui.sub_label("Социальная отдача (k_sroi)")
            st.radio("k_sroi", options=list(inp.SROI_OPTIONS), key="ksroi",
                     label_visibility="collapsed")
        with ec3:
            ui.sub_label("Бюджетная дисциплина (k_budget)")
            st.radio("k_budget", options=list(inp.BUDGET_OPTIONS), key="kbud",
                     label_visibility="collapsed")
    else:
        st.caption(
            "Расширенный контур доступен при наличии данных управленческого "
            "учета. Без них расчет ведется в рамках базового контура по "
            "открытым источникам."
        )


# ---------------------------------------------------------------------------
# Расчет и результаты
# ---------------------------------------------------------------------------

base = inp.collect_base_inputs()
ext = inp.collect_ext_inputs()
res = oc.compute(base, ext)

company = st.session_state["c_name"].strip()
company_html = html.escape(company)
subtitle = ("базовый контур · " + company_html) if company else \
    "базовый контур — внешняя оценка по открытым источникам"

# --- До ввода данных результаты не показываются -----------------------------
if not inp.has_input_data(base):
    with st.container(border=True):
        ui.card_title("Результаты расчета",
                      "появятся после ввода данных и запуска расчета")
        st.markdown(
            "<div class='oseec-note'>Заполните шаги 1–4 вручную или нажмите "
            "«Демонстрационный пример» в карточке «Объект оценки» — форма "
            "заполнится данными условной компании. После этого здесь "
            "появится кнопка «Рассчитать индекс».</div>",
            unsafe_allow_html=True)
    st.stop()

# --- Для мониторинга обязателен отраслевой эталон ----------------------------
if base.media_track == "monitoring" and base.x_ref <= 0:
    with st.container(border=True):
        ui.card_title("Результаты расчета",
                      "не хватает отраслевого эталона")
        st.markdown(
            "<div class='oseec-note'>Для расчета по данным системы "
            "мониторинга укажите отраслевой эталон X_ref — без него формула "
            "(1) неприменима. Эталон определяется как среднее арифметическое "
            "пиковых значений трех крупнейших компаний отрасли за последние "
            "три года по данным той же системы мониторинга.</div>",
            unsafe_allow_html=True)
    st.stop()

# --- Результат раскрывается по нажатию кнопки --------------------------------
if not st.session_state["show_results"]:
    with st.container(border=True):
        ui.card_title("Результаты расчета",
                      "данные заполнены — можно выполнять расчет")
        st.markdown(
            "<div class='oseec-note'>Данные шагов введены. По нажатию кнопки "
            "калькулятор рассчитает субиндексы, поправочные коэффициенты и "
            "итоговое значение ОСЭЭК, построит декомпозицию оценки и "
            "подготовит протокол расчета. При последующем изменении данных "
            "результат пересчитывается сразу.</div>",
            unsafe_allow_html=True)
        st.button("Рассчитать индекс", type="primary", on_click=_run_calc)
    st.stop()

# --- Зона результатов: теплое полотно ----------------------------------------
with st.container(border=True):
    ui.section_band("Результаты расчета",
                    "обновляются автоматически при изменении данных")

    # --- Панель итога: спидометр + значение ------------------------------------
    with st.container(border=True):
        ui.card_title("Итоговое значение индекса", subtitle, light=True,
                      result=True)
        gcol, tcol = st.columns([1, 1.15], vertical_alignment="center")
        with gcol:
            st.plotly_chart(ui.fig_gauge(res["oseec_b"], res["level"]),
                            use_container_width=True, config=ui.PLOTLY_CONFIG)
        with tcol:
            econtour = ""
            if "oseec_e" in res:
                econtour = (
                    f"<div class='oseec-econtour'>Расширенный контур: "
                    f"ОСЭЭК<sub>E</sub> = ОСЭЭК<sub>B</sub> × {fmt(res['k_eff'])} "
                    f"= <b>{fmt(res['oseec_e'])}</b> балла — "
                    f"{res['level_e_name'].lower()} уровень</div>")
            st.markdown(
                f"""
                <div class="oseec-hlabel">ОСЭЭК базового контура</div>
                <div class="oseec-hnum">{fmt(res['oseec_b'])}<span> балла</span></div>
                <div>{ui.level_chip(res['level'],
                                    LEVEL_NAMES[res['level']] + ' уровень')}</div>
                <div class="oseec-form" style="margin-top:0.8rem;">
                  ОСЭЭК<sub>B</sub> = I<sub>Core</sub> × K<sub>risk</sub> ×
                  K<sub>scale</sub> = {fmt(res['i_core'])} × {fmt(res['k_risk'])} ×
                  {fmt(res['k_scale'])} = <b>{fmt(res['oseec_b'])}</b>
                </div>
                {econtour}
                """,
                unsafe_allow_html=True)

        scen_points = None
        if res.get("hr_scenarios"):
            scen_points = {f"V_hr = {int(k)}": v
                           for k, v in res["hr_scenarios"].items()}
        st.plotly_chart(
            ui.fig_scale(res["oseec_b"], res.get("oseec_e"), scen_points),
            use_container_width=True, config=ui.PLOTLY_CONFIG)

    # --- Критические ограничения и примечания ----------------------------------
    if res["critical"] or res["notes"]:
        with st.container(border=True):
            ui.card_title("Диагностические статусы и примечания", light=True)
            if res["critical"]:
                items = "".join(
                    f"<div class='oseec-crit'><b>Критическое ограничение:</b> "
                    f"{kind} «{name}» — {fmt(val)} балла при пороговом уровне "
                    f"{'40' if kind == 'субиндекс' else '30'} баллов.</div>"
                    for kind, name, val in res["critical"])
                st.markdown(items, unsafe_allow_html=True)
                st.caption(
                    "Расчетная величина индекса при присвоении статуса не "
                    "изменяется: слабый результат уже отражен в формуле. Статус "
                    "дополняет качественную характеристику оценки указанием на "
                    "проблемную зону."
                )
            for note in res["notes"]:
                st.markdown(f"<div class='oseec-note'>{note}</div>",
                            unsafe_allow_html=True)

    # --- Сценарный диапазон V_hr -----------------------------------------------
    if res.get("hr_scenarios"):
        with st.container(border=True):
            s0, s50, s100 = (res["hr_scenarios"][0.0], res["hr_scenarios"][50.0],
                             res["hr_scenarios"][100.0])
            ui.card_title("Чувствительность к сценарию V_hr", light=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("ОСЭЭК_B при V_hr = 0", fmt(s0),
                      delta=fmt(s0 - res["oseec_b"]), delta_color="normal")
            m2.metric("ОСЭЭК_B при V_hr = 50", fmt(s50))
            m3.metric("ОСЭЭК_B при V_hr = 100", fmt(s100),
                      delta=fmt(s100 - res["oseec_b"]), delta_color="normal")
            st.caption(
                "Компания не представлена в рейтингах работодателей как "
                "самостоятельное юридическое лицо: показан разброс итога при "
                "крайних и центральном значениях сценарного диапазона."
            )

    # --- Декомпозиция ----------------------------------------------------------
    with st.container(border=True):
        ui.card_title("Декомпозиция оценки", light=True)
        d1, d2 = st.columns([3, 2])
        with d1:
            st.caption("Компоненты и субиндексы, баллы")
            st.plotly_chart(ui.fig_components(res), use_container_width=True,
                            config=ui.PLOTLY_CONFIG)
        with d2:
            st.caption("Структура ядра индекса")
            st.plotly_chart(ui.fig_core_structure(res), use_container_width=True,
                            config=ui.PLOTLY_CONFIG)
            st.markdown(
                f"""
                <div class="oseec-form">I<sub>Core</sub> = 0,6 × {fmt(res['m_stab'])}
                + 0,4 × {fmt(res['s_rep'])} = <b>{fmt(res['i_core'])}</b></div>
                """,
                unsafe_allow_html=True)
            if "oseec_e" in res:
                st.markdown(
                    f"""
                    <div class="oseec-form">K<sub>eff</sub> = 1 + {fmt(ext.k_roi)} +
                    {fmt(ext.k_sroi)} + {fmt(ext.k_budget)} =
                    <b>{fmt(res['k_eff'])}</b></div>
                    """,
                    unsafe_allow_html=True)

        if res.get("i_media") is not None:
            st.caption(
                f"Медийный блок: I_media = {fmt(res['i_media'])}, V_vol = "
                f"{fmt(res['v_vol'])}, M_stab = I_media / (1 + V_vol) = "
                f"{fmt(res['m_stab'])}."
                + (f" Корпус: {res['n_total']} публикаций, из них негативных — "
                   f"{res['n_neg']}." if res.get("n_total") else ""))

    # --- Выгрузка и переходы ----------------------------------------------------
    b1, b2, _ = st.columns([1.2, 1.2, 2])
    with b1:
        st.download_button(
            "Скачать протокол расчета (Word)",
            data=build_report(res, company, base, ext),
            file_name="Протокол_расчета_ОСЭЭК.docx",
            mime=("application/vnd.openxmlformats-officedocument"
                  ".wordprocessingml.document"),
            type="primary",
            use_container_width=True)
    with b2:
        if st.button("Проверить устойчивость результата", use_container_width=True):
            st.switch_page("page_mc.py")
