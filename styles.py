from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from lxml.etree import SubElement
import logging

logger = logging.getLogger(__name__)

def set_gost_styles(document: Document) -> None:
    """
    Настраивает стили документа по требованиям ГОСТ 7.32–2017.
    Создаёт или обновляет стили: Normal, Heading 1–4, Caption, List Bullet,
    List Number, Formula, Table Grid, и задаёт outlineLvl для заголовков.
    """
    # ─── Вспомогательные функции ───────────────────────────────────────────
    def _configure_font(
        font,
        name: str = "Times New Roman",
        size: int = 14,
        bold: bool = False,
        italic: bool = False,
        all_caps: bool = False
    ) -> None:
        """Настраивает параметры шрифта."""
        font.name = name
        font.size = Pt(size)
        font.color.rgb = RGBColor(0, 0, 0)
        font.bold = bold
        font.italic = italic
        font.all_caps = all_caps

    def define_heading_style(
        level: int,
        size: int,
        bold: bool = False,
        italic: bool = False,
        all_caps: bool = False,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_before: Pt = Pt(0),
        space_after: Pt = Pt(0),
        first_line_indent: Cm | None = None,
        left_indent: Cm | None = None,
        page_break_before: bool = False,
        line_spacing: float = 1.0
    ) -> None:
        """Определяет стиль для заголовков с указанным уровнем outlineLvl."""
        style_name = f"Heading {level}"
        if style_name in document.styles:
            style = document.styles[style_name]
            logger.debug(f"Модифицируется стиль {style_name}")
        else:
            style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            logger.debug(f"Создан стиль {style_name}")

        _configure_font(
            style.font, size=size, bold=bold, italic=italic, all_caps=all_caps
        )

        pformat = style.paragraph_format
        pformat.alignment = alignment
        pformat.space_before = space_before
        pformat.space_after = space_after
        pformat.page_break_before = page_break_before
        pformat.line_spacing = line_spacing
        if first_line_indent is not None:
            pformat.first_line_indent = first_line_indent
        if left_indent is not None:
            pformat.left_indent = left_indent

        # Задаём outlineLvl для стиля
        ppr = style._element.xpath('w:pPr')[0]
        outline_lvl = SubElement(ppr, qn('w:outlineLvl'))
        outline_lvl.set(qn('w:val'), str(level - 1))  # outlineLvl: 0 для Heading 1, 1 для Heading 2 и т.д.

    def define_list_style(
        style_name: str,
        is_ordered: bool,
        base_indent: float = 1.25,
        hanging_indent: float = -0.63
    ) -> None:
        """Определяет стиль для списков (нумерованных или маркированных)."""
        if style_name in document.styles:
            style = document.styles[style_name]
            logger.debug(f"Модифицируется стиль {style_name}")
        else:
            style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            logger.debug(f"Создан стиль {style_name}")

        _configure_font(style.font)

        pformat = style.paragraph_format
        pformat.left_indent = Cm(base_indent)
        pformat.first_line_indent = Cm(hanging_indent)
        pformat.space_before = Pt(0)
        pformat.space_after = Pt(0)
        pformat.line_spacing = 1.5
        pformat.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # ─── Стиль Normal ─────────────────────────────────────────────────────
    if "Normal" in document.styles:
        style = document.styles["Normal"]
        logger.debug("Модифицируется стиль Normal")
    else:
        style = document.styles.add_style("Normal", WD_STYLE_TYPE.PARAGRAPH)
        logger.debug("Создан стиль Normal")

    _configure_font(style.font, size=14)
    pformat = style.paragraph_format
    pformat.first_line_indent = Cm(1.25)
    pformat.line_spacing = 1.5
    pformat.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pformat.space_before = Pt(0)
    pformat.space_after = Pt(0)

    # ─── Стили заголовков ─────────────────────────────────────────────────
    define_heading_style(
        level=1,
        size=14,
        bold=True,
        all_caps=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=Pt(0),
        space_after=Pt(0),
        page_break_before=True,
        line_spacing=1.0,
        first_line_indent = Cm(0)
    )
    define_heading_style(
        level=2,
        size=14,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_before=Pt(0),
        space_after=Pt(8),
        line_spacing=1.0
    )
    define_heading_style(
        level=3,
        size=14,
        bold=True,
        italic=False,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Cm(1.25),
        space_before=Pt(0),
        space_after=Pt(8),
        line_spacing=1.0
    )
    define_heading_style(
        level=4,
        size=14,
        italic=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Cm(1.25),
        space_before=Pt(0),
        space_after=Pt(8),
        line_spacing=1.0
    )

    # ─── Стиль Caption ────────────────────────────────────────────────────
    if "Caption" in document.styles:
        style = document.styles["Caption"]
        logger.debug("Модифицируется стиль Caption")
    else:
        style = document.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
        logger.debug("Создан стиль Caption")

    _configure_font(style.font, size=14)
    pformat = style.paragraph_format
    pformat.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Для рисунков; для таблиц будет переопределено в formatter.py
    pformat.first_line_indent = Cm(0)
    pformat.space_before = Pt(0)
    pformat.space_after = Pt(8)
    pformat.line_spacing = 1.0

    # ─── Стили списков ────────────────────────────────────────────────────
    define_list_style("List Bullet", is_ordered=False, base_indent=1.25, hanging_indent=-0.63)
    define_list_style("List Number", is_ordered=True, base_indent=1.25, hanging_indent=-0.63)

    # ─── Стиль Formula ────────────────────────────────────────────────────
    if "Formula" in document.styles:
        style = document.styles["Formula"]
        logger.debug("Модифицируется стиль Formula")
    else:
        style = document.styles.add_style("Formula", WD_STYLE_TYPE.PARAGRAPH)
        logger.debug("Создан стиль Formula")

    _configure_font(style.font, size=14)
    pformat = style.paragraph_format
    pformat.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pformat.space_before = Pt(0)
    pformat.space_after = Pt(0)
    pformat.line_spacing = 1.5

    # ─── Стиль Table Grid ─────────────────────────────────────────────────
    if "Table Grid" in document.styles:
        style = document.styles["Table Grid"]
        logger.debug("Модифицируется стиль Table Grid")
    else:
        style = document.styles.add_style("Table Grid", WD_STYLE_TYPE.TABLE)
        logger.debug("Создан стиль Table Grid")

    _configure_font(style.font, size=12)
    pformat = style.paragraph_format
    pformat.space_before = Pt(0)
    pformat.space_after = Pt(0)
    pformat.line_spacing = 1.5
    pformat.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Настройка границ таблицы
    tbl_pr = style._element.xpath("//w:tblPr")[0]
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = SubElement(tbl_pr, qn("w:tblBorders"))
        logger.debug("Создан элемент w:tblBorders для Table Grid")
    else:
        borders.clear()
        logger.debug("Очищены существующие границы Table Grid")

    for border in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border_el = SubElement(borders, qn(f"w:{border}"))
        border_el.set(qn("w:val"), "single")
        border_el.set(qn("w:sz"), "4")  # 0.5 pt (1 pt = 8 sz units, 0.5 pt = 4 sz)
        border_el.set(qn("w:space"), "0")
        border_el.set(qn("w:color"), "auto")

    logger.debug("Все стили ГОСТ 7.32–2017 настроены")