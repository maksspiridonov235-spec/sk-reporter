"""ТЛ пилота SUP-PDR — эталон в репо (слово в слово), без парсера docx."""

from __future__ import annotations

from typing import Any

_CIPHER = "SUP-PDR-ENC-001-DD-ST01-EV"
_OBJECT = (
    "Обустройство Верхнесалымского месторождения. "
    "ВЛ 35 кВ Энергоцентр – т.вр. ВЛ 35 кВ Промысловая 1(2). "
    "ВЛ 35 кВ Энергоцентр – т.вр. ВЛ 35 кВ К-54 1(2). "
    "Этап строительства №1\n"
    "ВЛ 35 кВ «отпайка от ВЛ 35 кВ Промысловая 1,2 – энергоцентр в районе УПСВ»."
)
_RD = "Воздушные линии электроснабжения."
_TL_FILE = "SUP-PDR-ENC-001-DD-ST01-EV.ТЛ_00.1.docx"


def sup_pdr_tl_content() -> dict[str, Any]:
    return {
        "source": _TL_FILE,
        "rows": [
            {"label": "Шифр проекта", "value": _CIPHER},
            {"label": "Объект", "value": _OBJECT},
            {"label": "РД", "value": _RD},
        ],
    }


def sup_pdr_tl_card_fields() -> dict[str, str]:
    return {
        "cipher": _CIPHER,
        "title": _RD.rstrip("."),
        "object_name": _OBJECT.replace("\n", " "),
        "tl_file": _TL_FILE,
    }
