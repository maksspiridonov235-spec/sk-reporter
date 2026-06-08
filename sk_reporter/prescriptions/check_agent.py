"""
Агент проверки предписания (Excel).

Читает лист «Форма заполнения предписания»: B18 (содержание замечания),
B19 (нормативный документ). Сверка с Техэксперт, при недоступности — с открытым интернетом.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from sk_reporter.prescriptions.techexpert_client import lookup_normative, parse_normative_reference

MODEL = "gemma4:31b-cloud"
FORM_SHEET = "Форма заполнения предписания"
ROW_CONTENT = 18
ROW_NORMATIVE = 19
COL_VALUE = 2
COL_HINT = 3


SYSTEM_PROMPT = """Ты — ведущий инженер строительного контроля (СК). Проверяешь предписание по нормативной базе.

РЕЖИМ: ПРОФЕССИОНАЛЬНАЯ РЕДАКТУРА B18 + ПОЛНАЯ ССЫЛКА B19. Перепиши содержание замечания техническим языком СК: конкретно, без воды, деловой документальный стиль. Все объективные факты сохрани; оценочные и эмоциональные формулировки убери или замени фактами.

На входе:
- B18 — текст инженера (исходник).
- B19 — ссылка инженера на НД.
- Фрагмент нормативки (если есть).

ЖЁСТКИЕ ПРАВИЛА B18 — ЧТО СОХРАНЯТЬ (объективные факты):
- объект, трубопровод, диаметры (Ø530×10 и т.п.);
- номера стыков/соединений (№ 2/2, 4, 8…);
- методы контроля (радиография, УЗК…), сроки («до начала гидроиспытаний»);
- требуемые действия (ремонт, повторный контроль);
- клеймо сварщика/бригады — только как идентификатор исполнителя («соединения, выполненные бригадой с клеймом «Л»»), если это следует из исходника.

ЖЁСТКИЕ ПРАВИЛА B18 — СТИЛЬ И ТОН:
1. Стиль предписания СК: безлично, нейтрально, по факту нарушения. Без обвинений, без «обиды», без оценок личных качеств бригады.
2. ЗАПРЕЩЕНО оставлять без переработки:
   - оценочные суждения без цифр: «большое количество брака», «систематически допускает», «постоянно нарушает», «работает плохо»;
   - эмоциональные обобщения про бригаду/подрядчика вместо описания нарушения на конкретных стыках.
3. КАК ПЕРЕРАБОТАТЬ такие фразы:
   - «бригада с клеймом «Л» допускает большое количество брака» →
     «на сварных соединениях …, выполненных бригадой с клеймом «Л», выявлены недопустимые дефекты»
     ИЛИ убрать предложение о бригаде, если связь клейма с перечисленными стыками в исходнике не ясна (и задать вопрос инженеру).
   - «брак» без перечня → заменить на «недопустимые дефекты» + перечень стыков; количество — только если есть в исходнике.
4. ЗАПРЕЩЕНО сжимать факты в пустые обобщения («выявлены дефекты» без стыков, диаметров, метода).
5. Можно полностью перестроить фразы, убрать повторы и «воду», исправить орфографию — при сохранении фактов.
6. Не вставляй в B18 номера приказов, ГОСТ, пункты НД — только в B19.
7. Структура B18 (рекомендуемая): [факт контроля и объект] → [что выявлено, где] → [что выполнить и до какого момента].

ЖЁСТКИЕ ПРАВИЛА B19 (нормативный документ):
1. Полное наименование документа бери ТОЛЬКО из блока «Найденный документ» / фрагмента нормативки. НЕ выдумывай название, дату и номер — система подставит официальное наименование из источника.
2. Пункты/подпункты из B19 инженера НЕ МЕНЯЙ. Если инженер указал п. 44 — оставь п. 44.
3. Заменяй сам документ (другой ГОСТ/приказ/номер) только если инженер явно ошибся — обязательно объясни в отчёте почему.
4. В «ИСПРАВЛЕННЫЕ ПОЛЯ» для B19 достаточно краткой ссылки инженера + пункты; полное название добавит система.

АЛГОРИТМ:
1. Выпиши объективные факты из B18.
2. Сверь с фрагментом НД по пунктам B19.
3. Перепиши B18 профессиональным языком СК; убери оценочные формулировки.
4. B19 — разверни полное наименование НД, сохрани пункты инженера.
5. КОНКРЕТИЗАЦИЯ: если не хватает деталей (какие дефекты, размеры, акты НК) — вопросы инженеру.

Если фрагмент нормативки не получен — не меняй пункты B19. В отчёте: ⚠ «нормативка не сверена онлайн».

ФОРМАТ ОТВЕТА (строго):

## РЕЗЮМЕ ПРОВЕРКИ
- [✓ или ⚠] Сверка с нормативной базой: [Техэксперт / интернет / не получено] — укажи источник явно
- [✓ или ⚠] Содержание замечания: [переработано профессионально / что изменено в стиле / недостаточно конкретики]
- [✓ или ⚠] Нормативный документ: [развёрнуто наименование / что и почему изменено]
- Решение: [переработка B18 / правка B19 / оба / без изменений]

## ВОПРОСЫ ИНЖЕНЕРУ
(Что уточнить у автора предписания. Если всё конкретно — одна строка: «Дополнительных вопросов нет.»)
1. [Вопрос] — [на какой пункт НД и какой пробел в B18 опирается]
2. …

## ЧЕРНОВИК ПИСЬМА ИНЖЕНЕРУ
(2–5 предложений для руководителя СК: вежливо, по делу, перечисли что уточнить. Если вопросов нет — «Уточнения не требуются.»)

## ОТЧЁТ О ПРАВКАХ
### Содержание замечания (B18)
- Статус: [без изменений / переработано]
- Стиль и тон: [что убрано/перефразировано — оценочные фразы, «брак», обвинения бригаде]
- Факты: [что сохранено — стыки, диаметры, клеймо, сроки]
- Причина правок: [кратко]

### Нормативный документ (B19)
- Статус: [без изменений / изменено]
- Что сделано: [развёрнуто полное наименование / заменён документ / исправлены пункты / без изменений]
- Причина: [обязательно: почему развернули название, сменили документ или пункт]

## ИСПРАВЛЕННЫЕ ПОЛЯ
Содержание замечания:
[полный профессиональный текст B18 — все объективные факты, без воды и оценочных суждений]

Нормативный документ:
[полный текст B19 — полное наименование НД и пункты инженера]"""


def _cell_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _open_form_sheet(path: Path):
    suffix = path.suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd  # type: ignore[import-untyped]
        except ImportError as e:
            raise ValueError(
                "Формат .xls требует xlrd — сохраните как .xlsx или установите xlrd"
            ) from e
        book = xlrd.open_workbook(str(path))
        try:
            sheet = book.sheet_by_name(FORM_SHEET)
        except xlrd.XLRDError as e:
            raise ValueError(f"Лист «{FORM_SHEET}» не найден") from e
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = FORM_SHEET
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                ws.cell(r + 1, c + 1, value=sheet.cell_value(r, c))
        return wb, ws

    wb = load_workbook(path, data_only=True, read_only=False)
    if FORM_SHEET not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Лист «{FORM_SHEET}» не найден")
    return wb, wb[FORM_SHEET]


def extract_form_fields(filepath: str | Path) -> dict[str, str]:
    """B18, B19 и подсказки из C18, C19."""
    path = Path(filepath)
    wb, ws = _open_form_sheet(path)
    try:
        fields = {
            "content": _cell_str(ws.cell(ROW_CONTENT, COL_VALUE).value),
            "normative": _cell_str(ws.cell(ROW_NORMATIVE, COL_VALUE).value),
            "content_hint": _cell_str(ws.cell(ROW_CONTENT, COL_HINT).value),
            "normative_hint": _cell_str(ws.cell(ROW_NORMATIVE, COL_HINT).value),
        }
    finally:
        wb.close()
    return fields


def _parse_corrected_fields(report_text: str) -> dict[str, str]:
    cleaned = re.sub(r"\*\*([^*]+)\*\*", r"\1", report_text)
    section = re.search(
        r"##\s*ИСПРАВЛЕННЫЕ\s*ПОЛЯ\s*(.*)$",
        cleaned,
        re.DOTALL | re.IGNORECASE,
    )
    body = section.group(1).strip() if section else ""

    content_m = re.search(
        r"Содержание\s+замечания\s*:\s*(.*?)(?=Нормативный\s+документ\s*:|$)",
        body,
        re.DOTALL | re.IGNORECASE,
    )
    norm_m = re.search(
        r"Нормативный\s+документ\s*:\s*(.*?)$",
        body,
        re.DOTALL | re.IGNORECASE,
    )
    out: dict[str, str] = {}
    if content_m:
        out["content"] = content_m.group(1).strip()
    if norm_m:
        out["normative"] = norm_m.group(1).strip()
    return out


def _number_in_text(num: str, text: str) -> bool:
    """Число или дробь (2/2) есть в тексте — не привязано к формату «№»."""
    if not num or not text:
        return False
    if num in text:
        return True
    if "/" in num:
        a, b = num.split("/", 1)
        return a in text and b in text
    return bool(re.search(rf"\b{re.escape(num)}\b", text))


def _missing_b18_facts(original: str, rewritten: str) -> list[str]:
    """
    Какие объективные факты из B18 инженера пропали в переработке.
    Сравнение по смыслу (числа, стыки, клеймо), а не по точной строке regex.
    """
    if not original or not rewritten:
        return []
    missing: list[str] = []

    for m in re.finditer(r"Ø\s*(\d+)\s*[xх×]\s*(\d+)", original, flags=re.IGNORECASE):
        d1, d2 = m.group(1), m.group(2)
        if not (_number_in_text(d1, rewritten) and _number_in_text(d2, rewritten)):
            missing.append(f"диаметр Ø{d1}×{d2}")

    seen_joints: set[str] = set()
    for m in re.finditer(
        r"(?:№№?|стык\w*|соединен\w*)\s*([\d\s.,/]+)",
        original,
        flags=re.IGNORECASE,
    ):
        for j in re.findall(r"\d+(?:/\d+)?", m.group(1)):
            if j in seen_joints:
                continue
            seen_joints.add(j)
            if not _number_in_text(j, rewritten):
                missing.append(f"стык/соединение №{j}")

    # Клеймо не обязательно: при снятии оценочных фраз о бригаде допустимо не повторять клеймо.

    if re.search(r"радиограф", original, flags=re.IGNORECASE):
        if not re.search(r"радиограф", rewritten, flags=re.IGNORECASE):
            missing.append("радиографический контроль")

    if re.search(r"неразрушающ", original, flags=re.IGNORECASE):
        if not re.search(r"неразрушающ", rewritten, flags=re.IGNORECASE):
            missing.append("неразрушающий контроль (НК)")

    if re.search(r"гидравлич|гидроиспыт", original, flags=re.IGNORECASE):
        if not re.search(r"гидравлич|гидроиспыт", rewritten, flags=re.IGNORECASE):
            missing.append("условие/срок до гидроиспытаний")

    return missing


def _usable_doc_title(title: str, engineer_norm: str) -> bool:
    """Наименование из Техэксперт/интернета пригодно для B19 (не эхо запроса)."""
    t = (title or "").strip()
    if len(t) < 35:
        return False
    ref = parse_normative_reference(engineer_norm)
    if ref.number and ref.number not in t:
        return False
    low = t.lower()
    if ref.search_query and low == ref.search_query.lower():
        return False
    if len(t.split()) <= 5 and ref.number in t and "приказ" in low:
        return False
    return bool(
        re.search(
            r"(?:приказ|постановлен|гост|сп|снип|правил|федеральн)",
            t,
            flags=re.IGNORECASE,
        )
    )


def _compose_b19(
    engineer_norm: str,
    normative_lookup: dict[str, Any],
    model_norm: str,
) -> str:
    """B19: официальное наименование из источника + пункты инженера."""
    pts = _normative_points(engineer_norm)
    pts_tail = ", ".join(f"п. {p}" for p in pts) if pts else ""
    doc_title = (normative_lookup.get("doc_title") or "").strip()

    if normative_lookup.get("ok") and _usable_doc_title(doc_title, engineer_norm):
        base = doc_title.rstrip(" .,")
        return f"{base}, {pts_tail}." if pts_tail else f"{base}."

    return (model_norm or engineer_norm).strip()


def _normative_source_info(lookup: dict[str, Any]) -> dict[str, str]:
    """Метаданные источника нормативки для UI и отчёта."""
    src = (lookup.get("source") or "unknown").lower()
    labels = {
        "techexpert": "Техэксперт",
        "internet": "Интернет (открытые источники)",
    }
    label = labels.get(src, src)
    if lookup.get("ok"):
        status = "получен"
    elif lookup.get("error"):
        status = "не получен"
    else:
        status = "не выполнялся"
    return {
        "source": src,
        "label": label,
        "status": status,
        "doc_title": lookup.get("doc_title") or "",
        "source_url": lookup.get("source_url") or "",
        "error": lookup.get("error") or "",
        "te_fallback_error": lookup.get("te_fallback_error") or "",
    }


def _parse_numbered_list(section: str) -> list[str]:
    if not section:
        return []
    items: list[str] = []
    for line in section.splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^\d+[\.)]\s*(.+)$", line)
        if m:
            items.append(m.group(1).strip())
        elif line.startswith("- ") and not items:
            items.append(line[2:].strip())
    return items


def _parse_engineer_questions(report_text: str) -> list[str]:
    section = _extract_report_section(report_text, "ВОПРОСЫ ИНЖЕНЕРУ")
    if not section:
        return []
    if re.search(r"дополнительных\s+вопросов\s+нет", section, re.I):
        return []
    return _parse_numbered_list(section)


def _parse_draft_letter(report_text: str) -> str:
    section = _extract_report_section(report_text, "ЧЕРНОВИК ПИСЬМА ИНЖЕНЕРУ")
    if not section:
        return ""
    if re.search(r"уточнения\s+не\s+требуются", section, re.I):
        return ""
    return section.strip()


_SUBJECTIVE_RE = re.compile(
    r"больш(ое|ая|ой)\s+количеств|"
    r"систематическ|"
    r"постоянн\w*\s+(?:допуск|наруш)|"
    r"работает\s+плохо|"
    r"допускает\s+брак|"
    r"много\s+брака|"
    r"недобросовест",
    re.I,
)


_DEFECT_TYPE_RE = re.compile(
    r"трещин|пор[ыа]|шлак|подрез|несплавлен|включен|вклинен|"
    r"смещен|непровар|расслоен|карман|выгоран|окисн|"
    r"подрезан|вогнут|выпукл|отклонен",
    re.I,
)


def _rule_based_questions(
    fields: dict[str, str], normative_lookup: dict[str, Any]
) -> list[str]:
    """Эвристики, если модель не задала вопросы по размытым формулировкам."""
    content = fields.get("content") or ""
    if not content:
        return []
    questions: list[str] = []
    pts = _normative_points(fields.get("normative") or "")
    pts_hint = f" (п. {', '.join(pts)})" if pts else ""

    if re.search(r"дефект", content, re.I) and not _DEFECT_TYPE_RE.search(content):
        questions.append(
            "Какие именно недопустимые дефекты выявлены (тип, размер, расположение по каждому стыку/соединению)?"
            f"{pts_hint} — в B18 указано только «дефекты» без конкретики."
        )

    if re.search(r"недопустим", content, re.I) and not _DEFECT_TYPE_RE.search(content):
        if not questions:
            questions.append(
                "Чем обоснована оценка «недопустимо» — какие параметры дефекта превышают норму по НД"
                f"{pts_hint}?"
            )

    if _SUBJECTIVE_RE.search(content):
        questions.append(
            "В исходнике есть оценочные формулировки («большое количество брака», "
            "«допускает брак» и т.п.) без цифр. Подтвердите количество дефектных стыков "
            "или согласны убрать оценочный тон из предписания?"
        )

    if re.search(r"брак", content, re.I) and not _SUBJECTIVE_RE.search(content) and not re.search(
        r"количеств|процент|стык|соединен|шов", content, re.I
    ):
        questions.append(
            "В чём проявляется «брак»: перечень дефектных стыков, виды дефектов, статистика?"
        )

    if re.search(r"ремонт", content, re.I) and not re.search(
        r"выруб|шлиф|перевар|зачист|способ", content, re.I
    ):
        questions.append(
            "Какой способ ремонта сварных соединений предусмотрен (вырубка, шлифовка, переварка и т.п.)?"
        )

    if normative_lookup.get("ok") and pts and not questions:
        excerpt = (normative_lookup.get("excerpt") or "")[:800]
        if excerpt and re.search(r"дефект|контрол|качеств", excerpt, re.I):
            if not re.search(r"акт|заключен|протокол|снимок|пленк", content, re.I):
                questions.append(
                    f"Приложены ли акты/заключения НК по стыкам, на которые ссылается предписание{pts_hint}?"
                )

    return questions


def _merge_questions(model_q: list[str], rule_q: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for q in model_q + rule_q:
        key = q.lower()[:80]
        if key in seen:
            continue
        seen.add(key)
        out.append(q)
    return out


def _normative_points(text: str) -> list[str]:
    if not text:
        return []
    return parse_normative_reference(text).points


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _text_changed(before: str, after: str) -> bool:
    return _normalize_text(before) != _normalize_text(after)


def _extract_report_section(report_text: str, heading: str) -> str:
    pattern = rf"##\s*{re.escape(heading)}\s*(.*?)(?=##\s*|\Z)"
    m = re.search(pattern, report_text, re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _guard_engineer_text(
    fields: dict[str, str],
    corrected: dict[str, str],
    normative_lookup: dict[str, Any] | None = None,
) -> tuple[dict[str, str], list[dict[str, str]]]:
    """
    B18: не терять объективные факты при переработке.
    B19: наименование из источника (lookup), пункты — у инженера.
    """
    out = dict(corrected)
    events: list[dict[str, str]] = []
    lookup = normative_lookup or {}
    orig_content = fields.get("content") or ""
    orig_norm = fields.get("normative") or ""
    new_content = out.get("content") or ""
    model_norm = out.get("normative") or ""

    if orig_content and new_content and _text_changed(orig_content, new_content):
        missing = _missing_b18_facts(orig_content, new_content)
        if missing:
            lost_str = "; ".join(missing)
            print(f"[PRESCRIPTION_CHECK] B18 guard: revert — {lost_str}")
            events.append(
                {
                    "field": "content",
                    "level": "warn",
                    "code": "guard_content_revert",
                    "message": (
                        "Переработка B18 не записана в файл: в тексте модели не найдены "
                        f"факты из исходника ({lost_str}). См. блок «Предложение модели»."
                    ),
                    "model_preview": new_content,
                    "lost_facts": lost_str,
                }
            )
            out["content"] = orig_content

    if orig_norm:
        composed = _compose_b19(orig_norm, lookup, model_norm)
        src_label = {
            "techexpert": "Техэксперт",
            "internet": "интернет",
        }.get(lookup.get("source") or "", "источник")
        if _usable_doc_title(lookup.get("doc_title") or "", orig_norm):
            if _text_changed(orig_norm, composed):
                events.append(
                    {
                        "field": "normative",
                        "level": "info",
                        "code": "normative_title_from_source",
                        "message": (
                            f"B19: полное наименование взято из {src_label} "
                            f"({(lookup.get('doc_title') or '')[:80]}…); "
                            f"пункты инженера сохранены."
                        ),
                    }
                )
            out["normative"] = composed
        elif model_norm:
            out["normative"] = model_norm
        else:
            out["normative"] = orig_norm

        new_norm = out.get("normative") or ""
        orig_pts = _normative_points(orig_norm)
        new_pts = _normative_points(new_norm)
        if orig_pts and new_pts and orig_pts != new_pts:
            print(
                f"[PRESCRIPTION_CHECK] B19 guard: points {orig_pts} -> {new_pts}, "
                "restore engineer points"
            )
            base = re.sub(
                r",?\s*(?:[PpПп]\.?\s*|п\.?\s*|пункт\s+)\d+(?:\.\d+)*"
                r"(?:\s*,\s*(?:[PpПп]\.?\s*|п\.?\s*|пункт\s+)\d+(?:\.\d+)*)*",
                "",
                new_norm,
                flags=re.IGNORECASE,
            ).rstrip(" .,")
            pts_tail = ", ".join(f"п. {p}" for p in orig_pts)
            out["normative"] = f"{base}, {pts_tail}." if base else orig_norm
            events.append(
                {
                    "field": "normative",
                    "level": "warn",
                    "code": "guard_points_revert",
                    "message": (
                        f"Модель заменила пункты НД ({', '.join('п. '+p for p in orig_pts)} "
                        f"→ {', '.join('п. '+p for p in new_pts)}). "
                        "В файл записаны пункты инженера."
                    ),
                }
            )

    return out, events


def _report_has_issues(report_text: str) -> bool:
    resume = _extract_report_section(report_text, "РЕЗЮМЕ ПРОВЕРКИ")
    block = resume or report_text
    return "⚠" in block


def _build_review_display(
    report_text: str,
    fields: dict[str, str],
    model_corrected: dict[str, str],
    final_corrected: dict[str, str],
    guard_events: list[dict[str, str]],
    normative_lookup: dict[str, Any],
    engineer_questions: list[str],
    draft_letter: str,
) -> str:
    """Текст отчёта для панели UI."""
    parts: list[str] = []
    src = _normative_source_info(normative_lookup)

    source_lines = [f"**Источник нормативки:** {src['label']} — {src['status']}"]
    if src["doc_title"]:
        source_lines.append(f"**Документ:** {src['doc_title']}")
    if src["source_url"]:
        source_lines.append(f"**URL:** {src['source_url']}")
    if src.get("te_fallback_error"):
        source_lines.append(
            f"**Техэксперт (не использован):** {src['te_fallback_error']}"
        )
    if src["error"] and not normative_lookup.get("ok"):
        source_lines.append(f"**Ошибка поиска:** {src['error']}")
    parts.append("## ИСТОЧНИК НОРМАТИВКИ\n" + "\n".join(f"- {l}" for l in source_lines))

    if engineer_questions:
        q_lines = "\n".join(f"{i}. {q}" for i, q in enumerate(engineer_questions, 1))
        parts.append("## ВОПРОСЫ ИНЖЕНЕРУ\n" + q_lines)
    else:
        parts.append("## ВОПРОСЫ ИНЖЕНЕРУ\nДополнительных вопросов нет — предписание достаточно конкретно.")

    if draft_letter:
        parts.append("## ЧЕРНОВИК ПИСЬМА ИНЖЕНЕРУ\n" + draft_letter)

    resume = _extract_report_section(report_text, "РЕЗЮМЕ ПРОВЕРКИ")
    changes = _extract_report_section(report_text, "ОТЧЁТ О ПРАВКАХ")
    if resume:
        parts.append("## РЕЗЮМЕ ПРОВЕРКИ\n" + resume)
    if changes:
        parts.append("## ОТЧЁТ О ПРАВКАХ\n" + changes)

    model_blocks: list[str] = []
    if model_corrected.get("content") and _text_changed(
        fields.get("content", ""), model_corrected.get("content", "")
    ):
        model_blocks.append(
            "### B18 — как предложила модель\n" + model_corrected["content"]
        )
    if model_corrected.get("normative") and _text_changed(
        fields.get("normative", ""), model_corrected.get("normative", "")
    ):
        model_blocks.append(
            "### B19 — как предложила модель\n" + model_corrected["normative"]
        )
    if model_blocks and (
        _text_changed(model_corrected.get("content", ""), final_corrected.get("content", ""))
        or _text_changed(model_corrected.get("normative", ""), final_corrected.get("normative", ""))
    ):
        parts.append("## ПРЕДЛОЖЕНИЕ МОДЕЛИ\n" + "\n\n".join(model_blocks))

    applied: list[str] = []
    b18_guard = next((e for e in guard_events if e.get("code") == "guard_content_revert"), None)
    if _text_changed(fields.get("content", ""), final_corrected.get("content", "")):
        applied.append("B18 — в файл записана переработка модели")
    elif b18_guard:
        applied.append(
            "B18 — переработка модели не записана: "
            + (b18_guard.get("lost_facts") or "потеряны факты из исходника")
        )
    elif _text_changed(fields.get("content", ""), model_corrected.get("content", "")):
        applied.append("B18 — модель предлагала правки, в файл записан исходник")
    else:
        applied.append("B18 — без изменений")

    b19_from_src = any(e.get("code") == "normative_title_from_source" for e in guard_events)
    if _text_changed(fields.get("normative", ""), final_corrected.get("normative", "")):
        if b19_from_src:
            applied.append(
                "B19 — полное наименование из источника нормативки + пункты инженера"
            )
        else:
            applied.append("B19 — нормативная ссылка изменена в проверенном файле")
    else:
        applied.append("B19 — без изменений")

    parts.append("## ИТОГ ЗАПИСИ В ФАЙЛ\n" + "\n".join(f"- {line}" for line in applied))

    if guard_events:
        parts.append(
            "## ЗАМЕЧАНИЯ СИСТЕМЫ\n"
            + "\n".join(f"- ⚠ {e['message']}" for e in guard_events)
        )

    return "\n\n".join(parts).strip()


def _collect_issues(
    report_text: str,
    fields: dict[str, str],
    model_corrected: dict[str, str],
    final_corrected: dict[str, str],
    guard_events: list[dict[str, str]],
    normative_lookup: dict[str, Any],
    engineer_questions: list[str],
) -> list[dict[str, str]]:
    issues: list[dict[str, str]] = []
    src = _normative_source_info(normative_lookup)

    if fields.get("normative"):
        if normative_lookup.get("ok"):
            msg = f"Нормативка: {src['label']}"
            if src["doc_title"]:
                msg += f" — {src['doc_title'][:80]}"
            if src.get("te_fallback_error"):
                msg += f" (Техэксперт: {src['te_fallback_error'][:100]})"
            issues.append(
                {
                    "level": "info",
                    "code": "normative_source",
                    "message": msg,
                }
            )
        else:
            issues.append(
                {
                    "level": "warn",
                    "code": "normative_lookup",
                    "message": normative_lookup.get("error")
                    or "Не удалось получить текст нормативного документа для сверки",
                }
            )

    if engineer_questions:
        issues.append(
            {
                "level": "warn",
                "code": "clarification_needed",
                "message": (
                    f"Нужно уточнение у инженера: {len(engineer_questions)} "
                    f"вопрос(ов) — см. блок «Вопросы инженеру»"
                ),
            }
        )

    for ev in guard_events:
        issues.append(
            {
                "level": ev.get("level") or "warn",
                "code": ev.get("code") or "guard",
                "message": ev.get("message") or "Система отклонила правку модели",
            }
        )

    if _text_changed(fields.get("content", ""), final_corrected.get("content", "")):
        reason = _extract_field_reason(report_text, "B18") or "см. отчёт о правках"
        issues.append(
            {
                "level": "warn",
                "code": "content_changed",
                "message": f"Изменено содержание замечания (B18): {reason}",
            }
        )

    if _text_changed(fields.get("normative", ""), final_corrected.get("normative", "")):
        reason = _extract_field_reason(report_text, "B19") or "см. отчёт о правках"
        issues.append(
            {
                "level": "warn",
                "code": "normative_changed",
                "message": f"Изменена нормативная ссылка (B19): {reason}",
            }
        )

    if _report_has_issues(report_text) and not issues:
        issues.append(
            {
                "level": "warn",
                "code": "model_flags",
                "message": "Модель отметила замечания в резюме проверки",
            }
        )

    return issues


def _extract_field_reason(report_text: str, field: str) -> str:
    block = _extract_report_section(report_text, "ОТЧЁТ О ПРАВКАХ")
    if not block:
        return ""
    if field.upper() == "B18":
        section = re.search(
            r"###\s*Содержание\s+замечания.*?-\s*Причина:\s*(.+?)(?=###|\Z)",
            block,
            re.DOTALL | re.IGNORECASE,
        )
    else:
        section = re.search(
            r"###\s*Нормативный\s+документ.*?-\s*Причина:\s*(.+?)(?=###|\Z)",
            block,
            re.DOTALL | re.IGNORECASE,
        )
    if not section:
        return ""
    reason = section.group(1).strip()
    reason = re.sub(r"\s*-\s*Статус:.*", "", reason, flags=re.DOTALL)
    return reason.split("\n")[0].strip()[:300]


def _format_normative_lookup(lookup: dict[str, Any]) -> str:
    ref = lookup.get("reference") or {}
    src = lookup.get("source") or "techexpert"
    src_label = {"techexpert": "Техэксперт", "internet": "интернет"}.get(src, src)
    lines = [
        f"Источник нормативки: {src_label}",
        f"Поисковый запрос: {ref.get('search_query') or '—'}",
    ]
    if lookup.get("doc_title"):
        lines.append(f"Найденный документ: {lookup['doc_title']}")
    if lookup.get("source_url"):
        lines.append(f"URL: {lookup['source_url']}")
    if lookup.get("te_fallback_error"):
        lines.append(f"Техэксперт недоступен: {lookup['te_fallback_error']}")
    if lookup.get("ok") and lookup.get("excerpt"):
        lines.append("")
        lines.append(f"--- Фрагмент из {src_label} ---")
        lines.append(str(lookup["excerpt"]))
        lines.append("--- конец фрагмента ---")
    elif lookup.get("error"):
        lines.append(f"Ошибка: {lookup['error']}")
    return "\n".join(lines)


def check_prescription(filepath: str | Path) -> dict:
    """Проверить предписание: Техэксперт + Ollama."""
    path = Path(filepath)
    filename = path.name

    try:
        fields = extract_form_fields(path)
    except Exception as e:
        return {
            "ok": False,
            "report": f"Ошибка чтения формы: {e}",
            "_source_file": filename,
            "issues": [{"level": "error", "message": str(e)}],
        }

    if not fields["content"] and not fields["normative"]:
        return {
            "ok": False,
            "report": "Ячейки B18 и B19 пустые — нечего проверять",
            "_source_file": filename,
            "issues": [{"level": "error", "message": "Пустые B18 и B19"}],
        }

    normative_lookup: dict[str, Any] = {"ok": False, "error": "B19 пуст — поиск не выполнялся"}
    if fields["normative"]:
        try:
            normative_lookup = lookup_normative(fields["normative"])
        except Exception as e:
            print(f"[PRESCRIPTION_CHECK] techexpert error: {e}")
            normative_lookup = {"ok": False, "error": str(e)}

    user_prompt = f"""Проверь предписание (лист «{FORM_SHEET}») по нормативной базе.

Подсказки шаблона:
- Содержание замечания: {fields["content_hint"] or "—"}
- Нормативный документ: {fields["normative_hint"] or "—"}

--- Содержание замечания (B18) ---
{fields["content"] or "(пусто)"}
--- конец ---

--- Нормативный документ (B19) ---
{fields["normative"] or "(пусто)"}
--- конец ---

{_format_normative_lookup(normative_lookup)}

Пункты НД, указанные инженером в B19: {", ".join(_normative_points(fields["normative"])) or "—"}

ВАЖНО:
- B18: полностью перепиши профессиональным языком СК — конкретно, без воды; сохрани объективные факты (стыки, диаметры, методы, сроки, клеймо).
- Убери оценочный и эмоциональный тон («большое количество брака», обвинения бригаде); клеймо «Л» — только как идентификатор исполнителя на конкретных стыках.
- B19: пункты инженера не меняй; полное наименование подставит система из найденного документа.
- Заполни «ВОПРОСЫ ИНЖЕНЕРУ», «ЧЕРНОВИК ПИСЬМА» и «ОТЧЁТ О ПРАВКАХ» (в т.ч. блок «Стиль и тон» для B18).

Ответ строго в формате из инструкции."""

    try:
        import ollama

        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            stream=False,
        )
        report_text = response.get("message", {}).get("content", "").strip()
    except Exception as e:
        print(f"[PRESCRIPTION_CHECK] ollama error: {e}")
        return {
            "ok": False,
            "report": f"Ошибка проверки: {e}",
            "_source_file": filename,
            "normative_lookup": normative_lookup,
            "issues": [{"level": "error", "message": str(e)}],
        }

    if not report_text:
        return {
            "ok": False,
            "report": "Ошибка: пустой ответ модели",
            "_source_file": filename,
            "normative_lookup": normative_lookup,
            "issues": [{"level": "error", "message": "пустой ответ модели"}],
        }

    model_corrected = _parse_corrected_fields(report_text)
    corrected, guard_events = _guard_engineer_text(
        fields, model_corrected, normative_lookup
    )
    model_questions = _parse_engineer_questions(report_text)
    rule_questions = _rule_based_questions(fields, normative_lookup)
    engineer_questions = _merge_questions(model_questions, rule_questions)
    draft_letter = _parse_draft_letter(report_text)
    normative_source = _normative_source_info(normative_lookup)
    review_display = _build_review_display(
        report_text,
        fields,
        model_corrected,
        corrected,
        guard_events,
        normative_lookup,
        engineer_questions,
        draft_letter,
    )
    issues = _collect_issues(
        report_text,
        fields,
        model_corrected,
        corrected,
        guard_events,
        normative_lookup,
        engineer_questions,
    )
    has_errors = any(i.get("level") == "error" for i in issues)
    has_warnings = any(i.get("level") == "warn" for i in issues)

    result = {
        "ok": not has_errors and not has_warnings,
        "report": report_text,
        "review_display": review_display,
        "_source_file": filename,
        "fields": fields,
        "corrected": corrected,
        "model_corrected": model_corrected,
        "guard_events": guard_events,
        "normative_lookup": normative_lookup,
        "normative_source": normative_source,
        "engineer_questions": engineer_questions,
        "draft_letter": draft_letter,
        "issues": issues,
        "has_errors": has_errors,
        "has_warnings": has_warnings,
    }
    print(
        f"[PRESCRIPTION_CHECK] "
        f"{'OK' if result['ok'] else ('ERR' if has_errors else 'WARN')}: {filename} "
        f"(te={'ok' if normative_lookup.get('ok') else 'fail'}, "
        f"issues={len(issues)})"
    )
    return result


def write_checked_copy(src: str | Path, dest: str | Path, check_result: dict) -> None:
    """Копия файла с исправленными B18/B19, если модель их вернула."""
    src_path = Path(src)
    dest_path = Path(dest)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src_path, dest_path)

    corrected = check_result.get("corrected") or {}
    if not corrected.get("content") and not corrected.get("normative"):
        return
    if dest_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return

    wb = load_workbook(dest_path)
    if FORM_SHEET not in wb.sheetnames:
        wb.close()
        return
    ws = wb[FORM_SHEET]
    if corrected.get("content"):
        ws.cell(ROW_CONTENT, COL_VALUE, value=corrected["content"])
    if corrected.get("normative"):
        ws.cell(ROW_NORMATIVE, COL_VALUE, value=corrected["normative"])
    wb.save(dest_path)
    wb.close()
