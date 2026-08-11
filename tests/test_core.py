# -*- coding: utf-8 -*-
"""Контрольные тесты расчетного ядра ОСЭЭК.

Каждое ожидаемое значение взято из печатного текста диссертационного
исследования: таблицы 30-31 (раздел 3.3) и таблицы Б.2-Б.12 (Приложение Б).
Исходные данные аналитической выборки используются только для верификации
расчетного ядра и в интерфейсе приложения не отображаются.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import numpy as np

import oseec_core as oc


def r2(x):
    return round(x, 2)


# ---------------------------------------------------------------------------
# Верификационные данные аналитической выборки (2024 год, Приложение Б)
# ---------------------------------------------------------------------------

COMPANIES: dict = {
    "КАМАЗ": {
        "inputs": dict(
            media_track="manual",
            monthly_total=[7, 7, 2, 3, 0, 4, 3, 4, 2, 2, 4, 3],
            monthly_neg=[3, 1, 0, 0, 0, 0, 0, 2, 0, 0, 4, 0],
            hr_in_rating=True, hr_rank=30, hr_total=152,
            transp_b1=[1, 1, 1, 1, 1], transp_b2=[1, 1, 1, 1, 1],
            inst_b1=[1, 1, 1], inst_b2=[0, 1, 1],
            k_risk_step="elevated", k_scale_step="standard"),
        "expected": dict(m_stab=48.27, v_hr=80.79, r_transp=100.00,
                         r_inst=83.33, s_rep=88.04, i_core=64.18,
                         oseec_b=70.60),
    },
    "Росэнергоатом": {
        "inputs": dict(
            media_track="manual",
            monthly_total=[1, 0, 3, 2, 1, 1, 1, 3, 1, 1, 0, 1],
            monthly_neg=[0, 0, 1, 1, 1, 0, 0, 1, 1, 0, 0, 0],
            hr_in_rating=False, hr_scenario=50.0,
            transp_b1=[0, 1, 1, 1, 1], transp_b2=[0, 0, 1, 0, 1],
            inst_b1=[1, 0, 1], inst_b2=[0, 0, 1],
            k_risk_step="elevated", k_scale_step="standard"),
        "expected": dict(m_stab=38.33, v_hr=50.00, r_transp=60.00,
                         r_inst=50.00, s_rep=53.33, i_core=44.33,
                         oseec_b=48.76),
    },
    "ТВЭЛ": {
        "inputs": dict(
            media_track="manual",
            monthly_total=[0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
            monthly_neg=[0] * 12,
            hr_in_rating=False, hr_scenario=50.0,
            transp_b1=[0, 1, 1, 1, 1], transp_b2=[0, 0, 0, 0, 0],
            inst_b1=[1, 0, 0], inst_b2=[0, 0, 1],
            k_risk_step="elevated", k_scale_step="unknown"),
        "expected": dict(m_stab=0.00, v_hr=50.00, r_transp=40.00,
                         r_inst=33.33, s_rep=41.11, i_core=16.44,
                         oseec_b=17.18),
    },
    "Росатом Недра": {
        "inputs": dict(
            media_track="manual",
            monthly_total=[0, 0, 1, 0, 0, 1, 0, 0, 0, 1, 0, 0],
            monthly_neg=[0] * 12,
            hr_in_rating=False, hr_scenario=50.0,
            transp_b1=[0, 0, 1, 1, 0], transp_b2=[0, 0, 0, 0, 0],
            inst_b1=[1, 0, 0], inst_b2=[0, 0, 0],
            k_risk_step="elevated", k_scale_step="small"),
        "expected": dict(m_stab=0.00, v_hr=50.00, r_transp=20.00,
                         r_inst=16.67, s_rep=28.89, i_core=11.56,
                         oseec_b=12.08),
    },
    "Швабе": {
        "inputs": dict(
            media_track="none",
            hr_in_rating=False, hr_scenario=50.0,
            transp_b1=[0, 0, 1, 1, 1], transp_b2=[0, 0, 0, 0, 0],
            inst_b1=[0, 0, 0], inst_b2=[0, 0, 0],
            k_risk_step="elevated", k_scale_step="unknown"),
        "expected": dict(m_stab=0.00, v_hr=50.00, r_transp=30.00,
                         r_inst=0.00, s_rep=26.67, i_core=10.67,
                         oseec_b=11.15),
    },
    "Вертолеты России": {
        "inputs": dict(
            media_track="none",
            hr_in_rating=False, hr_scenario=50.0,
            transp_b1=[0, 0, 0, 1, 1], transp_b2=[0, 0, 0, 0, 0],
            inst_b1=[0, 0, 0], inst_b2=[0, 0, 0],
            k_risk_step="elevated", k_scale_step="unknown"),
        "expected": dict(m_stab=0.00, v_hr=50.00, r_transp=20.00,
                         r_inst=0.00, s_rep=23.33, i_core=9.33,
                         oseec_b=9.75),
    },
}


def make_inputs(name: str, **overrides) -> oc.BaseInputs:
    kw = dict(COMPANIES[name]["inputs"])
    kw.update(overrides)
    return oc.BaseInputs(**kw)


def test_formula_examples_b4():
    """Подраздел Б.4: контрольные расчеты медийного блока."""
    im = oc.i_media_manual(31, 10, 41)
    assert abs(im - 75.6098) < 5e-5, im
    kamaz = [7, 7, 2, 3, 0, 4, 3, 4, 2, 2, 4, 3]
    assert sum(kamaz) == 41
    assert abs(np.mean(kamaz) - 3.4167) < 5e-5
    assert abs(np.std(kamaz, ddof=0) - 1.9347) < 5e-5
    vol = oc.v_vol(kamaz)
    assert abs(vol - 0.5663) < 5e-5, vol
    assert r2(oc.m_stab(im, vol)) == 48.27

    im2 = oc.i_media_manual(10, 5, 15)
    assert abs(im2 - 66.6667) < 5e-5
    rea = [1, 0, 3, 2, 1, 1, 1, 3, 1, 1, 0, 1]
    assert sum(rea) == 15
    assert abs(np.std(rea, ddof=0) - 0.9242) < 5e-5
    vol2 = oc.v_vol(rea)
    assert abs(vol2 - 0.7394) < 5e-5
    assert r2(oc.m_stab(im2, vol2)) == 38.33


def test_vhr_b7():
    """Подраздел Б.7: V_hr = (1 - 29/151) x 100 = 80,79."""
    assert r2(oc.v_hr(30, 152)) == 80.79
    assert oc.v_hr(1, 152) == 100.0
    assert oc.v_hr(152, 152) == 0.0


def test_checklists_b5_b6():
    """Подразделы Б.5-Б.6: двухступенчатая агрегация чек-листов."""
    assert r2(oc.checklist_component([1, 1, 1, 1, 1], [1, 1, 1, 1, 1])) == 100.00
    assert r2(oc.checklist_component([0, 1, 1, 1, 1], [0, 0, 1, 0, 1])) == 60.00
    assert r2(oc.checklist_component([0, 1, 1, 1, 1], [0, 0, 0, 0, 0])) == 40.00
    assert r2(oc.checklist_component([0, 0, 1, 1, 0], [0, 0, 0, 0, 0])) == 20.00
    assert r2(oc.checklist_component([0, 0, 1, 1, 1], [0, 0, 0, 0, 0])) == 30.00
    assert r2(oc.checklist_component([1, 1, 1], [0, 1, 1])) == 83.33
    assert r2(oc.checklist_component([1, 0, 1], [0, 0, 1])) == 50.00
    assert r2(oc.checklist_component([1, 0, 0], [0, 0, 1])) == 33.33
    assert r2(oc.checklist_component([1, 0, 0], [0, 0, 0])) == 16.67
    assert r2(oc.checklist_component([0, 0, 0], [0, 0, 0])) == 0.00


def test_full_matrix_b6():
    """Таблица Б.6: полная расчетная матрица по шести компаниям выборки."""
    for name, spec in COMPANIES.items():
        res = oc.compute(make_inputs(name))
        exp = spec["expected"]
        for key in ("m_stab", "v_hr", "r_transp", "r_inst"):
            assert r2(res[key]) == exp[key], (name, key, res[key], exp[key])
        assert r2(res["s_rep"]) == exp["s_rep"], (name, res["s_rep"])
        assert r2(res["i_core"]) == exp["i_core"], (name, res["i_core"])
        assert r2(res["oseec_b"]) == exp["oseec_b"], (name, res["oseec_b"])


def test_levels_table29():
    """Таблица 29: шкала интерпретации."""
    assert oc.LEVEL_NAMES[oc.level_of(70.60)] == "Средний"
    assert oc.LEVEL_NAMES[oc.level_of(48.76)] == "Низкий"
    assert oc.LEVEL_NAMES[oc.level_of(17.18)] == "Критически низкий"
    assert oc.LEVEL_NAMES[oc.level_of(25.0)] == "Критически низкий"
    assert oc.LEVEL_NAMES[oc.level_of(25.01)] == "Низкий"
    assert oc.LEVEL_NAMES[oc.level_of(76.0)] == "Высокий"
    assert oc.LEVEL_NAMES[oc.level_of(100.0)] == "Высокий"
    assert oc.LEVEL_NAMES[oc.level_of(100.01)] == "Очень высокий"
    assert oc.LEVEL_NAMES[oc.level_of(115.5)] == "Очень высокий"


def test_critical_statuses():
    """Критические ограничения: субиндекс < 40, компонент < 30."""
    res = oc.compute(make_inputs("Росэнергоатом"))
    kinds = {(k, n) for k, n, _ in res["critical"]}
    assert ("субиндекс", "медийная устойчивость") in kinds
    assert ("субиндекс", "социальная репутация") not in kinds
    res2 = oc.compute(make_inputs("Швабе"))
    kinds2 = {(k, n) for k, n, _ in res2["critical"]}
    assert ("компонент", "транспарентность") not in kinds2  # ровно 30,00
    assert ("компонент", "институциональная зрелость") in kinds2


def test_zero_rule_and_threshold():
    """Правило нулевого значения при корпусе ниже порога 12 публикаций."""
    res = oc.compute(make_inputs("ТВЭЛ"))
    assert res["m_stab"] == 0.00
    assert res["i_media"] is None
    base = oc.BaseInputs(media_track="manual",
                         monthly_total=[1] * 12, monthly_neg=[0] * 12)
    res12 = oc.compute(base)
    assert res12["m_stab"] > 0
    assert res12["v_vol"] == 0.0
    assert r2(res12["i_media"]) == 100.00


def test_vhr_scenarios_b9():
    """Таблица Б.9: ОСЭЭК_B при крайних и центральном сценариях V_hr."""
    expected = {
        "Росэнергоатом": (41.43, 48.76, 56.10),
        "ТВЭЛ": (10.22, 17.18, 24.15),
        "Росатом Недра": (5.11, 12.08, 19.04),
        "Швабе": (4.18, 11.15, 18.11),
        "Вертолеты России": (2.79, 9.75, 16.72),
    }
    for name, (lo, mid, hi) in expected.items():
        res = oc.compute(make_inputs(name))
        scen = res["hr_scenarios"]
        assert r2(scen[0.0]) == lo, (name, scen[0.0])
        assert r2(scen[50.0]) == mid, (name, scen[50.0])
        assert r2(scen[100.0]) == hi, (name, scen[100.0])


def test_kscale_alternative_b10():
    """Таблица Б.10: значения при альтернативной ступени K_scale."""
    alt = {
        "КАМАЗ": ("small", 67.07),
        "Росэнергоатом": ("small", 46.33),
        "ТВЭЛ": ("standard", 18.09),
        "Росатом Недра": ("standard", 12.71),
        "Швабе": ("standard", 11.73),
        "Вертолеты России": ("standard", 10.27),
    }
    for name, (step, expected) in alt.items():
        res = oc.compute(make_inputs(name, k_scale_step=step))
        assert r2(res["oseec_b"]) == expected, (name, res["oseec_b"])


def _score(name: str, w: float, k_risk: float, k_scale: float,
           alt: bool = False) -> float:
    """Оценка по печатным компонентам Б.6 для проверок Б.11-Б.12."""
    comp = {
        "КАМАЗ": (48.27, 80.79, 100.00, 83.33),
        "Росэнергоатом": (38.33, 50.00, 60.00, 50.00),
        "ТВЭЛ": (0.00, 50.00, 40.00, 33.33),
        "Росатом Недра": (0.00, 50.00, 20.00, 16.67),
        "Швабе": (0.00, 50.00, 30.00, 0.00),
        "Вертолеты России": (0.00, 50.00, 20.00, 0.00),
    }
    m, v, t, i = comp[name]
    if alt:
        if name == "КАМАЗ":
            t = 90.00
        elif name == "Росэнергоатом":
            i = 200 / 3
        elif name == "Швабе":
            i = 100 / 6
    s = (v + t + i) / 3
    return (w * m + (1 - w) * s) * k_risk * k_scale


def test_oat_weight_b11():
    """Таблица Б.11: изолированное варьирование веса ядра 0,4-0,7."""
    exp = {
        "КАМАЗ": (66.22, 79.35, 1.00),
        "Росэнергоатом": (47.11, 52.07, 1.00),
        "ТВЭЛ": (12.89, 25.78, 0.95),
        "Росатом Недра": (9.06, 18.11, 0.95),
        "Швабе": (8.36, 16.72, 0.95),
        "Вертолеты России": (7.32, 14.63, 0.95),
    }
    for name, (lo, hi, ks) in exp.items():
        a = _score(name, 0.7, 1.10, ks)
        b = _score(name, 0.4, 1.10, ks)
        assert r2(min(a, b)) == lo and r2(max(a, b)) == hi, (name, a, b)


def test_oat_disputed_b11():
    """Таблица Б.11: изолированное включение спорных индикаторов."""
    assert r2(_score("КАМАЗ", 0.6, 1.10, 1.00, alt=True)) == 69.13
    assert r2(_score("Росэнергоатом", 0.6, 1.10, 1.00, alt=True)) == 51.21
    assert r2(_score("Швабе", 0.6, 1.10, 0.95, alt=True)) == 13.47


def test_mc_reference_b12():
    """Подраздел Б.12: референсный имитационный прогон (seed 20260702).

    Самостоятельная реализация протокола Б.12 с сохранением порядка
    розыгрышей; ожидаемые значения соответствуют прогону с уточненным
    компонентом V_hr ПАО «КАМАЗ», равным 80,79.
    """
    names = ("КАМАЗ", "Росэнергоатом", "ТВЭЛ", "Росатом Недра", "Швабе",
             "Вертолеты России")
    ks_base = dict(zip(names, (1.00, 1.00, 0.95, 0.95, 0.95, 0.95)))
    disputed_names = ("КАМАЗ", "Росэнергоатом", "Швабе")
    rng = np.random.default_rng(20260702)
    base = {c: _score(c, 0.6, 1.10, ks_base[c]) for c in names}
    base_level = {c: oc.level_of(base[c]) for c in names}
    keep = {c: 0 for c in names}
    vals = {c: [] for c in names}
    keep_groups = keep_rea_crit = 0
    w_draws, rea_scores = [], []
    for _ in range(10_000):
        w = rng.uniform(0.4, 0.7)
        b = tuple(x + rng.uniform(-2.5, 2.5) for x in (25.0, 50.0, 75.0))
        th_sub = 40.0 + rng.uniform(-2.5, 2.5)
        disputed = {c: bool(rng.integers(0, 2)) for c in disputed_names}
        k_risk = {c: rng.uniform(1.00, 1.10) for c in names}
        k_scale = {c: (rng.uniform(0.95, 1.05)
                       if c in ("КАМАЗ", "Росэнергоатом")
                       else rng.uniform(0.95, 1.00)) for c in names}
        it = {c: _score(c, w, k_risk[c], k_scale[c], alt=disputed.get(c, False))
              for c in names}
        for c in names:
            vals[c].append(it[c])
            if oc.level_of(it[c], b) == base_level[c]:
                keep[c] += 1
        zeros_max = max(it[c] for c in names[2:])
        if it["КАМАЗ"] > it["Росэнергоатом"] > zeros_max:
            keep_groups += 1
        if 38.33 < th_sub:
            keep_rea_crit += 1
        w_draws.append(w)
        rea_scores.append(it["Росэнергоатом"])

    assert keep_groups == 10_000
    assert keep_rea_crit == 8265
    exp = {
        "КАМАЗ": (8980, 56.73, 82.31, 68.65),
        "Росэнергоатом": (6690, 41.26, 58.05, 48.51),
        "ТВЭЛ": (9360, 11.89, 26.86, 18.94),
        "Росатом Недра": (10_000, 8.31, 18.85, 13.31),
        "Швабе": (10_000, 7.70, 21.10, 13.43),
        "Вертолеты России": (10_000, 6.72, 15.32, 10.76),
    }
    for c, (k, lo, hi, med) in exp.items():
        a = np.array(vals[c])
        assert keep[c] == k, (c, keep[c])
        assert r2(a.min()) == lo and r2(a.max()) == hi, (c, a.min(), a.max())
        assert r2(float(np.median(a))) == med, (c, np.median(a))
    w_arr, r_arr = np.array(w_draws), np.array(rea_scores)
    cross = r_arr > 50
    assert int(cross.sum()) == 3121
    assert round(float(np.median(w_arr[cross])), 3) == 0.479
    assert round(float(np.median(w_arr[~cross])), 3) == 0.586


def test_extended_contour():
    """Формулы (10)-(11): границы коэффициента управленческой эффективности."""
    assert oc.k_eff(0, 0, -0.05) == 0.95
    assert oc.k_eff(0.10, 0.10, 0.05) == 1.25
    base = make_inputs("КАМАЗ")
    ext = oc.ExtInputs(enabled=True, k_roi=0.10, k_sroi=0.05, k_budget=0.0)
    res = oc.compute(base, ext)
    assert r2(res["k_eff"]) == 1.15
    assert r2(res["oseec_e"]) == r2(res["oseec_b"] * 1.15)


def test_monitoring_track():
    """Формула (1) с приведением к границам шкалы."""
    im, capped = oc.i_media_monitoring(45000, 60000)
    assert r2(im) == 75.00 and not capped
    im2, capped2 = oc.i_media_monitoring(70000, 60000)
    assert im2 == 100.0 and capped2
    im3, _ = oc.i_media_monitoring(0, 60000)
    assert im3 == 0.0


def test_demo_example():
    """Демонстрационный пример: условные данные, высокий уровень."""
    p = oc.DEMO_EXAMPLE
    base = oc.BaseInputs(
        media_track=p["media_track"],
        monthly_total=list(p["monthly_total"]),
        monthly_neg=list(p["monthly_neg"]),
        hr_in_rating=p["hr_in_rating"], hr_rank=p["hr_rank"],
        hr_total=p["hr_total"],
        transp_b1=list(p["transp_b1"]), transp_b2=list(p["transp_b2"]),
        inst_b1=list(p["inst_b1"]), inst_b2=list(p["inst_b2"]),
        k_risk_step=p["k_risk_step"], k_scale_step=p["k_scale_step"])
    res = oc.compute(base)
    assert r2(res["oseec_b"]) == 76.75
    assert res["level_name"] == "Высокий"
    assert not res["critical"]


def test_custom_mc_reproducible():
    """Пользовательский Монте-Карло: воспроизводимость и диапазоны ступеней."""
    kw = dict(m_stab_v=48.27, vhr_v=80.79,
              transp_b1=[1] * 5, transp_b2=[1] * 5,
              inst_b1=[1, 1, 1], inst_b2=[0, 1, 1],
              k_risk_v=1.10, k_scale_v=1.00,
              disputed=[("transp_b2", 2)], n_iter=2000)
    a = oc.run_mc_custom(**kw)
    b = oc.run_mc_custom(**kw)
    assert a["min"] == b["min"] and a["median"] == b["median"]
    assert a["k_risk_range"] == (1.00, 1.10)
    assert a["k_scale_range"] == (0.95, 1.05)
    c = oc.run_mc_custom(m_stab_v=0, vhr_v=50, transp_b1=[0] * 5,
                         transp_b2=[0] * 5, inst_b1=[0] * 3, inst_b2=[0] * 3,
                         k_risk_v=0.90, k_scale_v=0.95, n_iter=500)
    assert c["k_risk_range"] == (0.90, 1.00)
    assert c["k_scale_range"] == (0.95, 1.00)


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"OK  {fn.__name__}")
    print(f"\nВсе {len(fns)} контрольных тестов пройдены.")
