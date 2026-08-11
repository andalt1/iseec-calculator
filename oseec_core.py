# -*- coding: utf-8 -*-
"""Расчетное ядро методики ОСЭЭК.

Модуль реализует формулы (1)-(11) раздела 3.2 диссертационного исследования,
правила работы с неполными данными, интерпретационную шкалу, статусы
критических ограничений и имитационную проверку устойчивости (Монте-Карло)
по протоколу подраздела Б.12 Приложения Б.

Порядок округления соответствует подразделу Б.1: значения компонентов
фиксируются с точностью до сотых долей и в таком виде применяются в
последующих расчетах; производные величины вычисляются без промежуточного
округления и отображаются с точностью до сотых долей.

Авторы методики: Алтухов А.С., Бобылева А.З.
"""

from dataclasses import dataclass, field
from typing import Optional

import numpy as np

# ---------------------------------------------------------------------------
# Константы методики
# ---------------------------------------------------------------------------

W_MEDIA = 0.6          # вес субиндекса медийной устойчивости в ядре
W_REP = 0.4            # вес субиндекса социальной репутации в ядре
V_VOL_CAP = 1.0        # верхняя граница коэффициента волатильности
MIN_CORPUS = 12        # порог репрезентативности медиакорпуса, публикаций в год
THR_SUBINDEX = 40.0    # порог критического ограничения по субиндексу
THR_COMPONENT = 30.0   # порог критического ограничения по компоненту
LEVEL_BOUNDS = (25.0, 50.0, 75.0)

K_RISK_STEPS = {
    "elevated": 1.10,   # повышенный риск
    "moderate": 1.00,   # умеренный риск
    "low": 0.90,        # низкий риск
}

K_SCALE_STEPS = {
    "large": 1.05,      # свыше 100 тыс. человек
    "standard": 1.00,   # от 10 до 100 тыс. человек
    "small": 0.95,      # менее 10 тыс. человек
    "unknown": 0.95,    # численность не подтверждена открытыми источниками
}

LEVEL_NAMES = (
    "Критически низкий",
    "Низкий",
    "Средний",
    "Высокий",
    "Очень высокий",
)

MONTHS = (
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
)

TRANSP_BLOCK1 = (
    "Годовой отчет",
    "Аудированная финансовая отчетность",
    "Структура собственности",
    "Органы управления",
    "Существенные факты",
)

TRANSP_BLOCK2 = (
    "Устойчивое развитие",
    "Занятость и кадровая политика",
    "Территории присутствия",
    "Экологическая безопасность",
    "Социальные программы",
)

INST_BLOCK1 = (
    "Публичные каналы обратной связи",
    "Регулярные форматы взаимодействия со стейкхолдерами",
    "Процедуры рассмотрения обращений",
)

INST_BLOCK2 = (
    "Матрица существенности или сопоставимый механизм",
    "Коммуникационная политика или регламент",
    "Публичное участие первого лица или уполномоченных",
)


def fmt(x: float, dp: int = 2) -> str:
    """Отображение числа в русской записи с запятой."""
    if x is None:
        return "–"
    return f"{x:.{dp}f}".replace(".", ",")


# ---------------------------------------------------------------------------
# Формулы базового контура
# ---------------------------------------------------------------------------

def i_media_monitoring(x_fact: float, x_ref: float) -> tuple[float, bool]:
    """Формула (1): индекс медийной результативности по данным мониторинга.

    Возвращает значение и признак приведения к границам стобалльной шкалы ядра.
    """
    if x_ref <= 0:
        return 0.0, False
    raw = x_fact / x_ref * 100.0
    capped = min(max(raw, 0.0), 100.0)
    return capped, capped != raw


def i_media_manual(n_pos_neu: int, n_neg: int, n_total: int) -> float:
    """Формула (2): индекс медийной результативности по ручному протоколу."""
    if n_total <= 0:
        return 0.0
    return (1.0 + (n_pos_neu - n_neg) / n_total) * 50.0


def v_vol(monthly: list[float]) -> Optional[float]:
    """Формула (3): коэффициент волатильности как отношение стандартного
    отклонения помесячных значений к их среднему; расчет по генеральной
    совокупности (двенадцать месяцев образуют полный цикл наблюдения),
    верхняя граница 1,0."""
    arr = np.asarray(monthly, dtype=float)
    mu = arr.mean()
    if mu <= 0:
        return None
    sigma = arr.std(ddof=0)
    return min(sigma / mu, V_VOL_CAP)


def m_stab(i_media: float, vol: float) -> float:
    """Формула (4): субиндекс медийной устойчивости с гиперболическим
    демпфированием."""
    return i_media / (1.0 + vol)


def v_hr(rank: int, total: int) -> float:
    """Формула (5): нормирование позиции компании в рейтинге работодателей."""
    if total <= 1:
        return 100.0 if rank == 1 else 0.0
    value = (1.0 - (rank - 1) / (total - 1)) * 100.0
    return min(max(value, 0.0), 100.0)


def checklist_component(block1: list[int], block2: list[int]) -> float:
    """Двухступенчатая агрегация чек-листа: среднее значение индикаторов
    внутри блока, затем среднее двух блоков по шкале от 0 до 100."""
    b1 = float(np.mean(block1)) * 100.0
    b2 = float(np.mean(block2)) * 100.0
    return (b1 + b2) / 2.0


def s_rep(vhr: float, r_transp: float, r_inst: float) -> float:
    """Формула (6): субиндекс социальной репутации как среднее трех
    компонентов."""
    return (vhr + r_transp + r_inst) / 3.0


def i_core(m: float, s: float, w_media: float = W_MEDIA) -> float:
    """Формула (7): ядро индекса как взвешенная сумма субиндексов."""
    return w_media * m + (1.0 - w_media) * s


def oseec_b(core: float, k_risk: float, k_scale: float) -> float:
    """Формулы (8)-(9): приведенная оценка и базовый контур."""
    return core * k_risk * k_scale


def k_eff(k_roi: float, k_sroi: float, k_budget: float) -> float:
    """Формула (10): коэффициент управленческой эффективности."""
    return round(1.0 + k_roi + k_sroi + k_budget, 10)


def oseec_e(b_value: float, keff: float) -> float:
    """Формула (11): расширенный контур."""
    return b_value * keff


def level_of(x: float, bounds: tuple[float, float, float] = LEVEL_BOUNDS) -> int:
    """Качественный уровень по шкале таблицы 29 (индекс 0-4)."""
    if x <= bounds[0]:
        return 0
    if x <= bounds[1]:
        return 1
    if x <= bounds[2]:
        return 2
    if x <= 100.0:
        return 3
    return 4


# ---------------------------------------------------------------------------
# Сквозной расчет
# ---------------------------------------------------------------------------

@dataclass
class BaseInputs:
    """Исходные данные базового контура."""

    # Медийный блок: track = "monitoring" | "manual" | "none"
    media_track: str = "manual"
    x_fact: float = 0.0
    x_ref: float = 0.0
    monthly_metric: list[float] = field(default_factory=lambda: [0.0] * 12)
    monthly_total: list[int] = field(default_factory=lambda: [0] * 12)
    monthly_neg: list[int] = field(default_factory=lambda: [0] * 12)

    # HR-блок: сценарный режим при отсутствии в рейтингах
    hr_in_rating: bool = False
    hr_rank: int = 1
    hr_total: int = 100
    hr_scenario: float = 50.0  # значение сценарного диапазона (0, 50, 100)

    # Чек-листы (значения 0/1)
    transp_b1: list[int] = field(default_factory=lambda: [0] * 5)
    transp_b2: list[int] = field(default_factory=lambda: [0] * 5)
    inst_b1: list[int] = field(default_factory=lambda: [0] * 3)
    inst_b2: list[int] = field(default_factory=lambda: [0] * 3)

    # Поправочные коэффициенты
    k_risk_step: str = "elevated"
    k_scale_step: str = "unknown"


@dataclass
class ExtInputs:
    """Данные расширенного контура (управленческий учет)."""

    enabled: bool = False
    k_roi: float = 0.0
    k_sroi: float = 0.0
    k_budget: float = 0.0


def compute(base: BaseInputs, ext: Optional[ExtInputs] = None) -> dict:
    """Полный расчет ОСЭЭК по правилам раздела 3.2 и Приложения Б."""
    r: dict = {"notes": [], "critical": []}

    # --- Медийная устойчивость -------------------------------------------
    corpus_total = None
    if base.media_track == "monitoring":
        im, capped = i_media_monitoring(base.x_fact, base.x_ref)
        vol = v_vol(base.monthly_metric)
        if vol is None:
            r["i_media"], r["v_vol"], ms = None, None, 0.0
            r["notes"].append(
                "Помесячные значения медиапоказателя не заполнены, субиндекс "
                "медийной устойчивости принят равным 0,00."
            )
        else:
            if capped:
                r["notes"].append(
                    "Расчетное значение индекса медийной результативности "
                    "приведено к границе стобалльной шкалы ядра."
                )
            r["i_media"], r["v_vol"] = im, vol
            ms = m_stab(im, vol)
    elif base.media_track == "manual":
        corpus_total = int(sum(base.monthly_total))
        n_neg = int(sum(base.monthly_neg))
        n_pos_neu = corpus_total - n_neg
        r["n_total"], r["n_neg"], r["n_pos_neu"] = corpus_total, n_neg, n_pos_neu
        if corpus_total < MIN_CORPUS:
            r["i_media"], r["v_vol"], ms = None, None, 0.0
            r["notes"].append(
                f"Объем верифицированного корпуса ({corpus_total}) ниже порога "
                f"репрезентативности ({MIN_CORPUS} публикаций за отчетный год): "
                "субиндекс медийной устойчивости принят равным 0,00 – нижняя "
                "граница внешней наблюдаемости."
            )
        else:
            im = i_media_manual(n_pos_neu, n_neg, corpus_total)
            vol = v_vol([float(x) for x in base.monthly_total])
            r["i_media"], r["v_vol"] = im, vol
            ms = m_stab(im, vol)
    else:  # none — корпус не сформирован
        r["i_media"], r["v_vol"], ms = None, None, 0.0
        r["notes"].append(
            "Публикации, атрибутированные к анализируемому юридическому лицу, "
            "отсутствуют: субиндекс медийной устойчивости принят равным 0,00."
        )

    # --- Компоненты социальной репутации ---------------------------------
    if base.hr_in_rating:
        vhr = v_hr(base.hr_rank, base.hr_total)
        r["hr_scenario_mode"] = False
    else:
        vhr = float(base.hr_scenario)
        r["hr_scenario_mode"] = True
        r["notes"].append(
            "Компания не представлена в рейтингах работодателей как "
            "самостоятельное юридическое лицо: по компоненту верификации "
            "HR-бренда применен сценарный подход."
        )

    rtr = checklist_component(base.transp_b1, base.transp_b2)
    rin = checklist_component(base.inst_b1, base.inst_b2)

    # --- Фиксация компонентов до сотых (правило Б.1) ----------------------
    ms_r, vhr_r = round(ms, 2), round(vhr, 2)
    rtr_r, rin_r = round(rtr, 2), round(rin, 2)
    r.update(m_stab=ms_r, v_hr=vhr_r, r_transp=rtr_r, r_inst=rin_r)
    r["transp_b1_score"] = float(np.mean(base.transp_b1)) * 100.0
    r["transp_b2_score"] = float(np.mean(base.transp_b2)) * 100.0
    r["inst_b1_score"] = float(np.mean(base.inst_b1)) * 100.0
    r["inst_b2_score"] = float(np.mean(base.inst_b2)) * 100.0

    # --- Производные величины без промежуточного округления ---------------
    srep = s_rep(vhr_r, rtr_r, rin_r)
    core = i_core(ms_r, srep)
    kr = K_RISK_STEPS[base.k_risk_step]
    ksc = K_SCALE_STEPS[base.k_scale_step]
    b_val = oseec_b(core, kr, ksc)
    r.update(s_rep=srep, i_core=core, k_risk=kr, k_scale=ksc, oseec_b=b_val)
    r["level"] = level_of(b_val)
    r["level_name"] = LEVEL_NAMES[r["level"]]

    if base.k_scale_step == "unknown":
        r["notes"].append(
            "Численность персонала юридического лица не подтверждена открытыми "
            "источниками: коэффициент масштаба принят по нижней границе 0,95."
        )

    # --- Сценарный диапазон V_hr (правило Б.9) ----------------------------
    if r["hr_scenario_mode"]:
        scen = {}
        for v in (0.0, 50.0, 100.0):
            s_alt = s_rep(v, rtr_r, rin_r)
            scen[v] = oseec_b(i_core(ms_r, s_alt), kr, ksc)
        r["hr_scenarios"] = scen

    # --- Критические ограничения ------------------------------------------
    if ms_r < THR_SUBINDEX:
        r["critical"].append(("субиндекс", "медийная устойчивость", ms_r))
    if srep < THR_SUBINDEX:
        r["critical"].append(("субиндекс", "социальная репутация", srep))
    if vhr_r < THR_COMPONENT:
        r["critical"].append(("компонент", "верификация HR-бренда", vhr_r))
    if rtr_r < THR_COMPONENT:
        r["critical"].append(("компонент", "транспарентность", rtr_r))
    if rin_r < THR_COMPONENT:
        r["critical"].append(("компонент", "институциональная зрелость", rin_r))

    # --- Расширенный контур ------------------------------------------------
    if ext is not None and ext.enabled:
        keff = k_eff(ext.k_roi, ext.k_sroi, ext.k_budget)
        r["k_eff"] = keff
        r["oseec_e"] = oseec_e(b_val, keff)
        r["level_e"] = level_of(r["oseec_e"])
        r["level_e_name"] = LEVEL_NAMES[r["level_e"]]

    return r


# ---------------------------------------------------------------------------
# Демонстрационный пример заполнения (условные данные)
# ---------------------------------------------------------------------------

DEMO_EXAMPLE: dict = {
    "media_track": "manual",
    "monthly_total": [3, 2, 4, 3, 2, 3, 4, 2, 3, 3, 2, 3],
    "monthly_neg": [1, 0, 1, 0, 0, 0, 1, 0, 0, 1, 0, 0],
    "hr_in_rating": True, "hr_rank": 25, "hr_total": 140,
    "transp_b1": [1, 1, 1, 1, 1], "transp_b2": [1, 1, 0, 1, 1],
    "inst_b1": [1, 1, 1], "inst_b2": [0, 1, 1],
    "k_risk_step": "moderate", "k_scale_step": "standard",
}

# ---------------------------------------------------------------------------
# Имитационная проверка устойчивости (Монте-Карло) по протоколу раздела 3.2
# ---------------------------------------------------------------------------

MC_SEED = 20260702
MC_ITER = 10_000
MC_DELTA = 2.5

# ---------------------------------------------------------------------------
# Монте-Карло для произвольных данных пользователя
# ---------------------------------------------------------------------------

def _neighbor_range(value: float, step: float, lo: float, hi: float) -> tuple:
    """Диапазон варьирования коэффициента в границах соседней ступени."""
    return (max(lo, value - step), min(hi, value + step))


def run_mc_custom(m_stab_v: float, vhr_v: float,
                  transp_b1: list[int], transp_b2: list[int],
                  inst_b1: list[int], inst_b2: list[int],
                  k_risk_v: float, k_scale_v: float,
                  disputed: Optional[list[tuple[str, int]]] = None,
                  n_iter: int = MC_ITER, seed: int = MC_SEED) -> dict:
    """Имитационная проверка устойчивости для данных пользователя.

    Протокол повторяет подраздел Б.12: вес ядра варьируется в диапазоне
    0,4-0,7; границы уровней и пороги критических ограничений сдвигаются
    в пределах 2,5 балла; поправочные коэффициенты меняются в границах
    соседней ступени; отмеченные пользователем спорные индикаторы
    чек-листов включаются с вероятностью 0,5. Начальное значение
    генератора фиксировано, результаты воспроизводимы.
    """
    disputed = disputed or []
    rng = np.random.default_rng(seed)

    kr_lo, kr_hi = _neighbor_range(k_risk_v, 0.10, 0.90, 1.10)
    ks_lo, ks_hi = _neighbor_range(k_scale_v, 0.05, 0.95, 1.05)

    def component_values(flips: dict) -> tuple[float, float]:
        tb1 = [(1 - v if ("transp_b1", i) in flips else v)
               for i, v in enumerate(transp_b1)]
        tb2 = [(1 - v if ("transp_b2", i) in flips else v)
               for i, v in enumerate(transp_b2)]
        ib1 = [(1 - v if ("inst_b1", i) in flips else v)
               for i, v in enumerate(inst_b1)]
        ib2 = [(1 - v if ("inst_b2", i) in flips else v)
               for i, v in enumerate(inst_b2)]
        return checklist_component(tb1, tb2), checklist_component(ib1, ib2)

    rtr_base, rin_base = component_values(set())
    s_base = s_rep(vhr_v, round(rtr_base, 2), round(rin_base, 2))
    base_val = oseec_b(i_core(m_stab_v, s_base), k_risk_v, k_scale_v)
    base_lvl = level_of(base_val)
    base_crit_sub = m_stab_v < THR_SUBINDEX or s_base < THR_SUBINDEX

    keep = 0
    levels_seen = set()
    vals = np.empty(n_iter)
    crit_sub_share = 0
    w_draws = np.empty(n_iter)

    for i in range(n_iter):
        w = rng.uniform(0.4, 0.7)
        b = tuple(x + rng.uniform(-MC_DELTA, MC_DELTA) for x in LEVEL_BOUNDS)
        th_sub = THR_SUBINDEX + rng.uniform(-MC_DELTA, MC_DELTA)
        flips = {key for key in disputed if rng.integers(0, 2)}
        kr = rng.uniform(kr_lo, kr_hi)
        ks = rng.uniform(ks_lo, ks_hi)
        rtr_i, rin_i = component_values(flips)
        s_i = s_rep(vhr_v, round(rtr_i, 2), round(rin_i, 2))
        v = oseec_b(i_core(m_stab_v, s_i, w), kr, ks)
        vals[i] = v
        w_draws[i] = w
        lv = level_of(v, b)
        levels_seen.add(lv)
        if lv == base_lvl:
            keep += 1
        if m_stab_v < th_sub or s_i < th_sub:
            crit_sub_share += 1

    return {
        "base_val": base_val,
        "base_level": base_lvl,
        "keep": keep / n_iter,
        "levels_seen": sorted(levels_seen),
        "min": float(vals.min()),
        "max": float(vals.max()),
        "median": float(np.median(vals)),
        "values": vals,
        "w_draws": w_draws,
        "crit_sub_share": crit_sub_share / n_iter,
        "base_crit_sub": base_crit_sub,
        "k_risk_range": (kr_lo, kr_hi),
        "k_scale_range": (ks_lo, ks_hi),
    }
