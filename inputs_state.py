# -*- coding: utf-8 -*-
"""Связь виджетов калькулятора с расчетным ядром через session_state."""

import streamlit as st

from oseec_core import BaseInputs, DEMO_EXAMPLE, ExtInputs

TRACK_MANUAL = "Ручной протокол: верифицированный медиакорпус"
TRACK_MONITOR = "Данные системы медиамониторинга"
TRACK_NONE = "Публикации отсутствуют или корпус не сформирован"

HR_IN = "Компания представлена в рейтинге работодателей"
HR_SCEN = "Самостоятельная позиция в рейтингах не выявлена (сценарный подход)"

RISK_OPTIONS = {
    "Повышенный риск — K_risk = 1,10": "elevated",
    "Умеренный риск — K_risk = 1,00": "moderate",
    "Низкий риск — K_risk = 0,90": "low",
}

SCALE_OPTIONS = {
    "Свыше 100 тыс. человек — K_scale = 1,05": "large",
    "От 10 до 100 тыс. человек — K_scale = 1,00": "standard",
    "Менее 10 тыс. человек — K_scale = 0,95": "small",
    "Численность не подтверждена открытыми источниками — K_scale = 0,95": "unknown",
}

ROI_OPTIONS = {
    "ROI отсутствует, отрицателен или не подтвержден внутренними данными — k_roi = 0": 0.0,
    "ROI положителен, но не превышает барьер (WACC или ключевую ставку) — k_roi = 0,05": 0.05,
    "ROI превышает барьерное значение — k_roi = 0,10": 0.10,
}

SROI_OPTIONS = {
    "SROI ниже 1 или расчет не подтвержден — k_sroi = 0": 0.0,
    "SROI в диапазоне от 1 до 2 — k_sroi = 0,05": 0.05,
    "SROI выше 2 при подтвержденном качестве расчета — k_sroi = 0,10": 0.10,
}

BUDGET_OPTIONS = {
    "Перерасход сверх лимита без управленческой санкции — k_budget = −0,05": -0.05,
    "Исполнение бюджета в пределах установленного допуска — k_budget = 0": 0.0,
    "Экономия при выполнении плана и без ухудшения субиндексов — k_budget = 0,05": 0.05,
}

_STEP_TO_RISK = {v: k for k, v in RISK_OPTIONS.items()}
_STEP_TO_SCALE = {v: k for k, v in SCALE_OPTIONS.items()}

DEFAULTS: dict = {
    "c_name": "",
    "c_inn": "",
    "c_track": TRACK_MANUAL,
    "hr_mode": HR_SCEN,
    "hr_rank": 30,
    "hr_total": 100,
    "hr_scen": "Среднее значение диапазона — 50",
    "xfact": 0.0,
    "xref": 0.0,
    "krisk": list(RISK_OPTIONS)[1],
    "kscale": list(SCALE_OPTIONS)[3],
    "ext_on": False,
    "kroi": list(ROI_OPTIONS)[0],
    "ksroi": list(SROI_OPTIONS)[0],
    "kbud": list(BUDGET_OPTIONS)[1],
    "show_results": False,
}
for _i in range(12):
    DEFAULTS[f"mt{_i}"] = 0
    DEFAULTS[f"mn{_i}"] = 0
    DEFAULTS[f"mm{_i}"] = 0.0
for _i in range(5):
    DEFAULTS[f"t1_{_i}"] = False
    DEFAULTS[f"t2_{_i}"] = False
for _i in range(3):
    DEFAULTS[f"i1_{_i}"] = False
    DEFAULTS[f"i2_{_i}"] = False

SCEN_VALUES = {
    "Минимальное значение диапазона — 0": 0.0,
    "Среднее значение диапазона — 50": 50.0,
    "Максимальное значение диапазона — 100": 100.0,
}


def ensure_defaults() -> None:
    """Инициализация и удержание состояния виджетов.

    Повторное присваивание значения переводит ключ в разряд программно
    управляемого состояния: Streamlit не очищает его при переходе на
    страницу, где соответствующий виджет не отрисован. Благодаря этому
    введенные в калькуляторе данные сохраняются при переключении разделов.
    """
    for k, v in DEFAULTS.items():
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]
        else:
            st.session_state[k] = v


def collect_base_inputs() -> BaseInputs:
    ss = st.session_state
    track = {TRACK_MANUAL: "manual", TRACK_MONITOR: "monitoring",
             TRACK_NONE: "none"}[ss["c_track"]]
    return BaseInputs(
        media_track=track,
        x_fact=float(ss["xfact"]),
        x_ref=float(ss["xref"]),
        monthly_metric=[float(ss[f"mm{i}"]) for i in range(12)],
        monthly_total=[int(ss[f"mt{i}"]) for i in range(12)],
        monthly_neg=[min(int(ss[f"mn{i}"]), int(ss[f"mt{i}"])) for i in range(12)],
        hr_in_rating=ss["hr_mode"] == HR_IN,
        hr_rank=int(ss["hr_rank"]),
        hr_total=int(ss["hr_total"]),
        hr_scenario=SCEN_VALUES[ss["hr_scen"]],
        transp_b1=[int(ss[f"t1_{i}"]) for i in range(5)],
        transp_b2=[int(ss[f"t2_{i}"]) for i in range(5)],
        inst_b1=[int(ss[f"i1_{i}"]) for i in range(3)],
        inst_b2=[int(ss[f"i2_{i}"]) for i in range(3)],
        k_risk_step=RISK_OPTIONS[ss["krisk"]],
        k_scale_step=SCALE_OPTIONS[ss["kscale"]],
    )


def has_input_data(base: BaseInputs) -> bool:
    """Проверяет, введены ли содержательные данные для расчета.

    Пустая форма не образует осмысленной оценки, вследствие чего блок
    результатов до появления данных заменяется приглашением к вводу.
    """
    if base.media_track == "none":
        return True
    if base.media_track == "manual" and sum(base.monthly_total) > 0:
        return True
    if base.media_track == "monitoring" and (base.x_fact > 0
                                             or base.x_ref > 0):
        return True
    checklist = (sum(base.transp_b1) + sum(base.transp_b2)
                 + sum(base.inst_b1) + sum(base.inst_b2))
    return checklist > 0


def collect_ext_inputs() -> ExtInputs:
    ss = st.session_state
    return ExtInputs(
        enabled=bool(ss["ext_on"]),
        k_roi=ROI_OPTIONS[ss["kroi"]],
        k_sroi=SROI_OPTIONS[ss["ksroi"]],
        k_budget=BUDGET_OPTIONS[ss["kbud"]],
    )


def load_demo_into_state() -> None:
    """Заполняет калькулятор демонстрационным примером с условными данными."""
    p = DEMO_EXAMPLE
    ss = st.session_state
    ss["c_name"] = "ПАО «Прогресс» (условная компания)"
    ss["c_inn"] = ""
    ss["c_info"] = None
    ss["show_results"] = False
    ss["c_track"] = {"manual": TRACK_MANUAL, "monitoring": TRACK_MONITOR,
                     "none": TRACK_NONE}[p["media_track"]]
    for i in range(12):
        ss[f"mt{i}"] = int(p["monthly_total"][i])
        ss[f"mn{i}"] = int(p["monthly_neg"][i])
        ss[f"mm{i}"] = 0.0
    if p.get("hr_in_rating"):
        ss["hr_mode"] = HR_IN
        ss["hr_rank"] = int(p["hr_rank"])
        ss["hr_total"] = int(p["hr_total"])
    else:
        ss["hr_mode"] = HR_SCEN
        ss["hr_scen"] = "Среднее значение диапазона — 50"
    for i in range(5):
        ss[f"t1_{i}"] = bool(p["transp_b1"][i])
        ss[f"t2_{i}"] = bool(p["transp_b2"][i])
    for i in range(3):
        ss[f"i1_{i}"] = bool(p["inst_b1"][i])
        ss[f"i2_{i}"] = bool(p["inst_b2"][i])
    ss["krisk"] = _STEP_TO_RISK[p["k_risk_step"]]
    ss["kscale"] = _STEP_TO_SCALE[p["k_scale_step"]]
    ss["ext_on"] = False


def clear_all_inputs() -> None:
    """Возвращает калькулятор к пустому состоянию."""
    for k, v in DEFAULTS.items():
        st.session_state[k] = v
    st.session_state["c_info"] = None
    st.session_state.pop("mc_custom_result", None)
    st.session_state.pop("mc_custom_meta", None)
