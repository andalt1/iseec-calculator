# -*- coding: utf-8 -*-
"""Поиск организации по ИНН через справочник DaData.

Ключ доступа хранится исключительно в защищенном хранилище Streamlit
(st.secrets, переменная DADATA_API_KEY) и в код не включается. При
отсутствии ключа функция поиска недоступна, а калькулятор продолжает
работать с ручным вводом наименования. Справочные сведения (наименование,
отрасль по ОКВЭД, регион, численность) в расчет индекса не входят.
"""

from typing import Optional

import requests
import streamlit as st

_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"

# Укрупненные разделы ОКВЭД 2 для справочного отображения отрасли
_OKVED_SECTIONS = (
    ((1, 3), "сельское, лесное хозяйство, рыболовство"),
    ((5, 9), "добыча полезных ископаемых"),
    ((10, 33), "обрабатывающие производства"),
    ((35, 35), "энергетика и газоснабжение"),
    ((36, 39), "водоснабжение и обращение с отходами"),
    ((41, 43), "строительство"),
    ((45, 47), "торговля"),
    ((49, 53), "транспортировка и хранение"),
    ((55, 56), "гостиницы и общественное питание"),
    ((58, 63), "информация и связь"),
    ((64, 66), "финансы и страхование"),
    ((68, 68), "операции с недвижимостью"),
    ((69, 75), "профессиональная и научно-техническая деятельность"),
    ((77, 82), "административная деятельность"),
    ((84, 84), "государственное управление"),
    ((85, 85), "образование"),
    ((86, 88), "здравоохранение и социальные услуги"),
    ((90, 93), "культура и спорт"),
    ((94, 96), "прочие услуги"),
)


def okved_section_name(okved: str) -> Optional[str]:
    try:
        code = int(str(okved).split(".")[0])
    except (ValueError, AttributeError):
        return None
    for (lo, hi), name in _OKVED_SECTIONS:
        if lo <= code <= hi:
            return name
    return None


def get_api_key() -> Optional[str]:
    try:
        return st.secrets.get("DADATA_API_KEY", None)
    except Exception:
        return None


def lookup_inn(inn: str) -> dict:
    """Возвращает сведения об организации либо словарь с полем error."""
    key = get_api_key()
    if not key:
        return {"error": "Поиск по ИНН сейчас недоступен: ключ доступа к "
                         "справочнику не настроен. Введите наименование "
                         "организации вручную."}
    inn = (inn or "").strip().replace(" ", "")
    if not inn.isdigit() or len(inn) not in (10, 12):
        return {"error": "ИНН состоит из 10 цифр для юридического лица или "
                         "12 цифр для индивидуального предпринимателя."}
    try:
        resp = requests.post(
            _URL,
            json={"query": inn, "branch_type": "MAIN"},
            headers={"Content-Type": "application/json",
                     "Accept": "application/json",
                     "Authorization": f"Token {key}"},
            timeout=8,
        )
    except requests.RequestException:
        return {"error": "Справочник не отвечает. Попробуйте позже или "
                         "введите наименование вручную."}
    if resp.status_code in (401, 403):
        return {"error": "Справочник отклонил ключ доступа. Введите "
                         "наименование вручную."}
    if resp.status_code != 200:
        return {"error": "Справочник временно недоступен. Введите "
                         "наименование вручную."}
    suggestions = resp.json().get("suggestions") or []
    if not suggestions:
        return {"error": "Организация с этим ИНН в справочнике не найдена. "
                         "Проверьте номер или введите наименование вручную."}
    data = suggestions[0].get("data", {}) or {}
    name_block = data.get("name", {}) or {}
    name = (name_block.get("short_with_opf")
            or suggestions[0].get("value") or "")
    address = ((data.get("address", {}) or {}).get("data", {}) or {})
    state = (data.get("state", {}) or {}).get("status")
    okved = data.get("okved")
    return {
        "name": name,
        "full_name": name_block.get("full_with_opf") or name,
        "inn": data.get("inn", inn),
        "okved": okved,
        "okved_section": okved_section_name(okved) if okved else None,
        "region": address.get("region_with_type"),
        "status": state,
        "employee_count": data.get("employee_count"),
    }
