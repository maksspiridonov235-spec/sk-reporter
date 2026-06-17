"""ТЛ проекта SUP-IPL-S101-016-DD-00-RL — данные в репо, без парсера."""

from __future__ import annotations

from typing import Any

_PROJECT_ID = "SUP-IPL-S101-016-DD-00-RL"
_CIPHER = "SUP-IPL-S101-016-DD-00-RL"
_OBJECT = (
    "Обустройство Верхнесалымского месторождения.\n"
    "Нефтегазосборный трубопровод от узла УН41 до узла УН236"
)
_RD = "Рубка леса."


def ipl_s101_tl_content() -> dict[str, Any]:
    return {
        "source": "",
        "rows": [
            {"label": "Шифр проекта", "value": _CIPHER},
            {"label": "Объект", "value": _OBJECT},
            {"label": "РД", "value": _RD},
        ],
    }


def ipl_s101_tl_card_fields() -> dict[str, str]:
    return {
        "id": _PROJECT_ID,
        "cipher": _CIPHER,
        "title": _RD.rstrip("."),
        "object_name": _OBJECT.replace("\n", " "),
        "tl_file": "",
        "vor_file": "",
    }
