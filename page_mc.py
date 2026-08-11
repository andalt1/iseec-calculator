# -*- coding: utf-8 -*-
"""Страница «Проверка устойчивости» — имитационное моделирование Монте-Карло."""

import streamlit as st

import inputs_state as inp
import oseec_core as oc
import ui_theme as ui
from oseec_core import fmt

inp.ensure_defaults()

st.title("Проверка устойчивости результата")
ui.lead(
    "Наличие в архитектуре ОСЭЭК теоретически заданных весов, порогов и "
    "поправочных коэффициентов предполагает проверку итога на устойчивость. "
    "Модуль выполняет 10 000 итераций имитационного моделирования "
    "Монте-Карло с фиксированным начальным значением генератора случайных "
    "чисел, что обеспечивает воспроизводимость результатов."
)

with st.expander("Протокол имитационного варьирования"):
    st.markdown(
        """
        В каждой итерации одновременно изменяются параметры расчетной модели:

        - вес субиндекса медийной устойчивости — в диапазоне от 0,4 до 0,7
          при сохранении суммы весов ядра, равной единице;
        - границы качественных уровней и порог критических ограничений —
          случайный сдвиг в пределах 2,5 балла, десятой части ширины уровня;
        - поправочные коэффициенты — в границах соседней ступени шкалы;
        - отмеченные спорные индикаторы чек-листов — включение в расчет
          с вероятностью 0,5 независимо друг от друга.

        Качественный уровень в каждой итерации определяется заново по
        сдвинутой шкале этой итерации. Результат признается устойчивым при
        сохранении качественного уровня и выводов о слабых компонентах.
        """
    )

base = inp.collect_base_inputs()
res = oc.compute(base)
company = st.session_state["c_name"].strip() or "текущий расчет"

# --- Объект проверки и запуск прогона --------------------------------------
with st.container(border=True):
    ui.card_title("Объект проверки и параметры прогона")
    st.markdown(
        f"<div class='oseec-note'>Проверяется расчет «<b>{company}</b>»: "
        f"ОСЭЭК<sub>B</sub> = <b>{fmt(res['oseec_b'])}</b> балла, уровень — "
        f"{res['level_name'].lower()}. Исходные данные берутся из раздела "
        "«Калькулятор» — для проверки другого объекта измените их там.</div>",
        unsafe_allow_html=True)

    all_indicators = (
        [("transp_b1", i, f"Транспарентность · {t}")
         for i, t in enumerate(oc.TRANSP_BLOCK1)]
        + [("transp_b2", i, f"Транспарентность · {t}")
           for i, t in enumerate(oc.TRANSP_BLOCK2)]
        + [("inst_b1", i, f"Институциональная зрелость · {t}")
           for i, t in enumerate(oc.INST_BLOCK1)]
        + [("inst_b2", i, f"Институциональная зрелость · {t}")
           for i, t in enumerate(oc.INST_BLOCK2)]
    )
    labels = {f"{blk}:{i}": lbl for blk, i, lbl in all_indicators}
    chosen = st.multiselect(
        "Спорные индикаторы чек-листов (включаются с вероятностью 0,5)",
        options=list(labels),
        format_func=lambda k: labels[k],
        placeholder="Не отмечены — варьируются только параметры модели",
        help="Отметьте индикаторы, оценка которых допускает альтернативную "
             "трактовку источников. В каждой итерации значение такого "
             "индикатора меняется на противоположное с вероятностью 0,5.")
    disputed = [(k.split(":")[0], int(k.split(":")[1])) for k in chosen]

    if st.button("Выполнить имитационный прогон", type="primary"):
        with st.spinner("Выполняется 10 000 итераций..."):
            mc = oc.run_mc_custom(
                m_stab_v=res["m_stab"], vhr_v=res["v_hr"],
                transp_b1=base.transp_b1, transp_b2=base.transp_b2,
                inst_b1=base.inst_b1, inst_b2=base.inst_b2,
                k_risk_v=res["k_risk"], k_scale_v=res["k_scale"],
                disputed=disputed)
        st.session_state["mc_custom_result"] = mc
        st.session_state["mc_custom_meta"] = {
            "company": company, "base_val": res["oseec_b"],
            "level": res["level"], "disputed_n": len(disputed)}

mc = st.session_state.get("mc_custom_result")
meta = st.session_state.get("mc_custom_meta")
if mc and meta:
  with st.container(border=True):
    ui.section_band("Итоги имитационного прогона", meta["company"])
    if abs(meta["base_val"] - res["oseec_b"]) > 1e-9:
        st.markdown(
            "<div class='oseec-note'>Данные в калькуляторе изменились после "
            "последнего прогона — выполните прогон заново.</div>",
            unsafe_allow_html=True)
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Сохранение базового уровня", fmt(mc["keep"] * 100, 1) + " %")
    r2.metric("Размах значений", f"{fmt(mc['min'])}–{fmt(mc['max'])}")
    r3.metric("Медиана", fmt(mc["median"]))
    r4.metric("Наблюдавшиеся уровни", str(len(mc["levels_seen"])))

    chips = " ".join(ui.level_chip(i) for i in mc["levels_seen"])
    st.markdown(
        f"<div style='margin:0.25rem 0 0.6rem 0'>Уровни в итерациях "
        f"прогона: {chips}</div>", unsafe_allow_html=True)

    st.plotly_chart(
        ui.fig_mc_hist(mc["values"], meta["base_val"], mc["median"]),
        use_container_width=True, config=ui.PLOTLY_CONFIG)

    keep_pct = mc["keep"] * 100
    if keep_pct >= 90:
        verdict = ("Качественный уровень сохраняется в подавляющем "
                   "большинстве итераций: вывод о положении объекта на "
                   "шкале слабо чувствителен к вариации параметров модели "
                   "в заданных диапазонах.")
    elif keep_pct >= 66:
        verdict = ("Качественный уровень сохраняется в большинстве "
                   "итераций, однако доля переходов заметна: оценка "
                   "находится вблизи границы уровней, и ее интерпретацию "
                   "уместно сопровождать указанием на пограничное "
                   "положение.")
    else:
        verdict = ("Существенная доля итераций дает смену качественного "
                   "уровня: итоговый вывод чувствителен к выбранным "
                   "параметрам модели, и дальнейшее сопоставление уместно "
                   "сопровождать указанием на это методологическое "
                   "ограничение.")
    st.markdown(f"<div class='oseec-note'>{verdict}</div>",
                unsafe_allow_html=True)
    st.caption(
        f"Диапазоны варьирования: вес ядра 0,4–0,7; границы уровней и "
        f"порог ограничений ± 2,5 балла; K_risk "
        f"{fmt(mc['k_risk_range'][0])}–{fmt(mc['k_risk_range'][1])}; "
        f"K_scale {fmt(mc['k_scale_range'][0])}–"
        f"{fmt(mc['k_scale_range'][1])}; спорных индикаторов — "
        f"{meta['disputed_n']}. Начальное значение генератора фиксировано "
        f"({oc.MC_SEED}), результаты воспроизводимы."
    )
