from pathlib import Path
from typing import List, Dict, Any, Optional

from docx import Document
from docx.shared import Cm

import styles
import formatter

__all__ = ["build_document"]

def _apply_page_margins(doc: Document, *, top: float = 2.0, bottom: float = 2.0, left: float = 3.0, right: float = 1.5) -> None:
    """Устанавливает одинаковые поля (см) для всех секций документа."""
    for section in doc.sections:
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)

def build_document(
    elements: List[Dict[str, Any]],
    doc: Document,
    output_path: Optional[str | Path] = None,
    *,
    margins_cm: Optional[Dict[str, float]] = None,
    config: Optional[Dict[str, Any]] = None,
) -> Document:
    """Форматирует существующий документ Word по структуре *elements* согласно ГОСТ 7.32–2017."""
    styles.set_gost_styles(doc)

    margins = margins_cm or {"top": 2, "bottom": 2, "left": 3, "right": 1.5}
    _apply_page_margins(doc, **margins)

    formatter.format_document(elements, doc, config=config)

    if output_path:
        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(out_path)

    return doc