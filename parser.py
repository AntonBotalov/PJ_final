import re
import logging
from typing import List, Dict, Any, Tuple
from pathlib import Path
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.parts.image import ImagePart
from docx.oxml.ns import qn
from PIL import Image
from io import BytesIO
from regex_patterns import match_element_type

logger = logging.getLogger(__name__)

# Константы
_BULLETS = r"[\-–—•·→■▪□➤*]"

def parse_marker(text: str) -> tuple[str | None, str, int | None]:
    """
    Извлекает маркер из текста, возвращает тип элемента, очищенный текст и уровень (для списков).
    """
    match = re.match(r'^\[(H1|H2|H3|H4|OL|OL:\d+|UL|UL:\d+|TABLE|FIGURE|FORMULA)\]\s*(.*)$', text, re.IGNORECASE)
    if not match:
        return None, text, None
    marker, clean_text = match.groups()
    level = None
    if ':' in marker:
        marker, level = marker.split(':')
        level = int(level)
    elem_type = {
        'H1': 'heading_1', 'H2': 'heading_2', 'H3': 'heading_3', 'H4': 'heading_4',
        'OL': 'list_ordered', 'UL': 'list_bullet',
        'TABLE': 'table_caption', 'FIGURE': 'figure_caption', 'FORMULA': 'formula'
    }.get(marker.upper())
    if not elem_type:
        logger.warning(f"Неверный маркер: {marker}")
        return None, text, None
    return elem_type, clean_text, level

def _list_level(p: Paragraph) -> int:
    ind = p.paragraph_format.left_indent
    cm = ind.cm if ind else 0.0
    return max(0, int(round(cm / 1.25)))

def _is_marker_only(p: Paragraph) -> bool:
    txt = p.text.strip()
    return bool(
        re.match(
            rf'^\s*(?:{_BULLETS}|\d+[.)]\s|[a-zа-я]\)\s|[a-zа-я]\.\s)\s*$', txt
        )
    )

def _detect_word_list(p: Paragraph, raw: str, elements: List[Dict[str, Any]]) -> tuple[bool, bool, int]:
    """
    (is_list, ordered, level)
    ordered True — десятичная нумерация, False — маркированный.
    """
    style = (p.style.name or "").lower()
    pPr = p._p.pPr
    level = _list_level(p)

    # Контекст: библиографический список
    is_bibliography = False
    if elements and elements[-1].get("type") == "heading_1" and \
       "список использованных источников" in elements[-1].get("text", "").lower():
        is_bibliography = True

    # Контекст: нумерованные списки
    is_ordered_context = False
    if elements and elements[-1].get("type") == "paragraph":
        prev_text = elements[-1].get("text", "").lower()
        if any(keyword in prev_text for keyword in ["следующие задачи", "структура документа", "обязательные элементы"]):
            is_ordered_context = True

    # Если стиль явно указывает на заголовок, возвращаем False
    if style.startswith("heading"):
        return False, False, 0

    # Проверка регулярных выражений для списков
    is_ordered = bool(re.match(r'^\s*(?:\d+\)|[a-zа-я]\)|[a-zа-я]\.|[IVXLCDM]+\.)\s+\S', raw)) or \
                 is_bibliography or is_ordered_context
    is_bullet = bool(re.match(rf'^\s*(?:{_BULLETS})\s+\S', raw))
    is_list = style in {"list bullet", "list number", "list paragraph"} or is_ordered or is_bullet

    if is_list:
        return True, is_ordered, level

    return False, False, 0

def parse_document(
    docx_path: str | Path,
    use_markers: bool = False,
    use_regex: bool = True
) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
    doc = Document(docx_path)
    elements: List[Dict[str, Any]] = []
    stats: Dict[str, int] = dict.fromkeys(
        ["headings", "tables", "figures", "formulas", "paragraphs", "list_items"], 0
    )

    para_idx = table_idx = figure_idx = formula_idx = 0
    body = list(doc.element.body.iterchildren())
    i = 0
    while i < len(body):
        node = body[i]

        # ───── ПАРАГРАФ ────────────────────────────────────
        if isinstance(node, CT_P):
            para = Paragraph(node, doc)
            raw = para.text.strip()

            # Пропускаем пустые параграфы
            if not raw and not node.xpath('.//pic:pic'):
                para_idx += 1
                i += 1
                continue

            # 1) Встроенный рисунок
            if node.xpath('.//pic:pic'):
                logger.debug(f"Обнаружено изображение в параграфе {para_idx}")
                img = None
                try:
                    for rId in node.xpath('.//a:blip/@r:embed'):
                        part = doc.part.related_parts.get(rId)
                        if part:
                            blob = getattr(part, "blob", None) or getattr(part, "_blob", None)
                            if blob:
                                img = Image.open(BytesIO(blob))
                                logger.debug(f"Изображение успешно извлечено для параграфа {para_idx}")
                                break
                    if not img:
                        logger.warning(f"Не удалось извлечь изображение для параграфа {para_idx}: нет blob")
                except Exception as e:
                    logger.error(f"Ошибка при извлечении изображения в параграфе {para_idx}: {str(e)}")

                figure_idx += 1
                fig = {"type": "figure", "image": img, "para_idx": para_idx, "number": figure_idx}
                if i + 1 < len(body) and isinstance(body[i + 1], CT_P):
                    nxt = Paragraph(body[i + 1], doc)
                    cap = nxt.text.strip()
                    if cap:
                        cap_type, cap_clean, _ = parse_marker(cap) if use_markers else (None, cap, None)
                        if cap_type == "figure_caption" or (use_regex and match_element_type(cap) == "figure_caption") or cap_type is None:
                            fig["caption"] = cap_clean
                            logger.debug(f"Подпись для изображения №{figure_idx}: '{cap_clean}'")
                            i += 1
                            para_idx += 1
                elements.append(fig)
                stats["figures"] += 1
                para_idx += 1
                i += 1
                continue

            # 2) Строка‑маркер (например, "1)", "–")
            if _is_marker_only(para):
                para_idx += 1
                i += 1
                continue

            # 3) Классификация параграфа
            elem_type = None
            ordered = False
            level = 0
            clean_text = raw

            # ─── Проверка 1: Маркеры (высший приоритет, если use_markers=True) ───
            if use_markers:
                elem_type, clean_text, marker_level = parse_marker(raw)
                if elem_type:
                    logger.debug(f"Маркер классифицировал '{raw[:50]}' как {elem_type}")
                    if elem_type.startswith("list_"):
                        ordered = elem_type == "list_ordered"
                        level = marker_level if marker_level is not None else 0
                        # Создаём элемент списка
                        el = {
                            "type": elem_type,
                            "text": clean_text,
                            "para_idx": para_idx,
                            "ordered": ordered,
                            "list_level": level
                        }
                        elements.append(el)
                        stats["list_items"] += 1
                        para_idx += 1
                        i += 1
                        continue
                    elif elem_type.startswith("heading"):
                        level = int(elem_type.split('_')[1])
                    elif elem_type == "figure_caption":
                        # Если предыдущий элемент — не figure, создаём figure без изображения
                        if not elements or elements[-1].get("type") != "figure":
                            figure_idx += 1
                            elements.append({
                                "type": "figure",
                                "image": None,
                                "para_idx": para_idx,
                                "number": figure_idx,
                                "caption": clean_text
                            })
                            stats["figures"] += 1
                            logger.debug(f"Добавлен figure №{figure_idx} без изображения с подписью '{clean_text}'")
                        else:
                            elements[-1]["caption"] = clean_text
                            logger.debug(f"Добавлена подпись '{clean_text}' к figure №{elements[-1]['number']}")
                        para_idx += 1
                        i += 1
                        continue
                    elif elem_type == "table_caption":
                        # Обрабатываем table_caption позже, при обработке таблицы
                        el = {
                            "type": elem_type,
                            "text": clean_text,
                            "para_idx": para_idx
                        }
                        elements.append(el)
                        para_idx += 1
                        i += 1
                        continue
                    elif elem_type == "formula":
                        formula_idx += 1
                        el = {
                            "type": elem_type,
                            "text": clean_text,
                            "para_idx": para_idx,
                            "number": formula_idx
                        }
                        elements.append(el)
                        stats["formulas"] += 1
                        para_idx += 1
                        i += 1
                        continue

            # ─── Проверка 2: Регулярные выражения (только если маркер не найден) ───
            if not elem_type and use_regex:
                elem_type = match_element_type(clean_text)
                if elem_type in ["heading_1", "heading_2", "heading_3", "heading_4"]:
                    level = int(elem_type.split('_')[1])
                    logger.debug(f"Регулярное выражение классифицировало '{clean_text[:50]}' как {elem_type}")
                elif elem_type in ["list_ordered", "list_bullet"]:
                    is_list, ord_f, lvl = _detect_word_list(para, clean_text, elements)
                    elem_type = "list_ordered" if ord_f else "list_bullet"
                    ordered = ord_f
                    level = lvl
                    logger.debug(f"Регулярное выражение и _detect_word_list классифицировали '{clean_text[:50]}' как {elem_type}")
                elif elem_type == "figure_caption":
                    if not elements or elements[-1].get("type") != "figure":
                        figure_idx += 1
                        elements.append({
                            "type": "figure",
                            "image": None,
                            "para_idx": para_idx,
                            "number": figure_idx,
                            "caption": clean_text
                        })
                        stats["figures"] += 1
                        logger.debug(f"Добавлен figure №{figure_idx} без изображения с подписью '{clean_text}'")
                    else:
                        elements[-1]["caption"] = clean_text
                        logger.debug(f"Добавлена подпись '{clean_text}' к figure №{elements[-1]['number']}")
                    para_idx += 1
                    i += 1
                    continue

            # ─── Проверка стилей Word для заголовков (дополнительная) ───
            if not elem_type and para.style and para.style.name.lower().startswith("heading"):
                try:
                    lvl = int(re.search(r'heading\s*(\d)', para.style.name.lower()).group(1))
                    elem_type = f"heading_{lvl}"
                    level = lvl
                    logger.debug(f"Стиль Word классифицировал '{clean_text[:50]}' как heading_{lvl}")
                except Exception:
                    pass

            # ─── Контекст: библиография ───
            is_bibliography = False
            if elements and elements[-1].get("type") == "heading_1" and \
               "список использованных источников" in elements[-1].get("text", "").lower():
                is_bibliography = True
                if not elem_type or elem_type == "paragraph":
                    elem_type = "list_ordered"
                    ordered = True
                    level = 0
                    logger.debug(f"Контекст библиографии классифицировал '{clean_text[:50]}' как list_ordered")

            # ─── По умолчанию: paragraph ───
            if not elem_type:
                elem_type = "paragraph"
                ordered = False
                level = 0
                logger.debug(f"По умолчанию классифицировано '{clean_text[:50]}' как paragraph")

            # Формируем элемент
            el: Dict[str, Any] = {
                "type": elem_type,
                "text": clean_text,
                "para_idx": para_idx,
            }

            if elem_type.startswith("heading"):
                el["level"] = level
                stats["headings"] += 1
            elif elem_type.startswith("list_"):
                el["ordered"] = ordered
                el["list_level"] = level
                stats["list_items"] += 1
            elif elem_type == "formula":
                formula_idx += 1
                el["number"] = formula_idx
                stats["formulas"] += 1
            elif elem_type == "table_caption":
                para_idx += 1
                i += 1
                continue
            else:
                stats["paragraphs"] += 1

            elements.append(el)
            para_idx += 1

        # ───── ТАБЛИЦА ────────────────────────────────────
        elif isinstance(node, CT_Tbl):
            table_idx += 1
            tbl = {
                "type": "table",
                "content": Table(node, doc),
                "number": table_idx,
                "rows": [[cell.text for cell in row.cells] for row in Table(node, doc).rows]
            }
            stats["tables"] += 1
            if elements and elements[-1].get("type") == "table_caption":
                tbl["caption"] = elements[-1]["text"]
                elements[-1] = tbl  # Заменяем table_caption на table
            else:
                elements.append(tbl)

        i += 1

    return elements, stats