#!/usr/bin/env python3
"""Add the reimbursement signature and archive label to electronic invoices.

The script keeps the source PDF unchanged.  It reads the first page's text
layer, finds the invoice fields, detects the remarks box, and writes a new PDF
named ``company-amount-dzfpNUMBER.pdf``.

Only PyMuPDF (fitz) and Pillow are required.  Both are already available in
the workspace runtime used for this project.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

import fitz
from PIL import Image, ImageDraw, ImageFont, ImageOps


DEFAULT_SIGNATURE = "signature.jpg"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parents[1] / "output" / "pdf"


@dataclass(frozen=True)
class InvoiceFields:
    buyer: str
    seller: str
    invoice_number: str
    total: Decimal
    subtotal: Optional[Decimal]
    buyer_tax_id: str = ""
    seller_tax_id: str = ""
    invoice_date: str = ""
    items: List[dict] = field(default_factory=list)


@dataclass(frozen=True)
class RemarkBox:
    x0: float
    y0: float
    x1: float
    y1: float
    content_x0: float


def _word_rows(page: fitz.Page) -> List[Tuple[float, float, float, float, str]]:
    rows: List[Tuple[float, float, float, float, str]] = []
    for item in page.get_text("words"):
        if len(item) >= 5:
            x0, y0, x1, y1, text = item[:5]
            rows.append((float(x0), float(y0), float(x1), float(y1), str(text)))
    return rows


def _is_numeric_token(text: str) -> bool:
    return bool(re.fullmatch(r"[0-9]{8,20}", text.strip()))


def _parse_money(text: str) -> Optional[Decimal]:
    match = re.search(r"[0-9][0-9,]*\.[0-9]{2}", text)
    if not match:
        return None
    try:
        return Decimal(match.group(0).replace(",", "")).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except InvalidOperation:
        return None


def _money_text(value: Decimal) -> str:
    return format(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), "f")


def _fields_dict(fields: InvoiceFields) -> dict:
    return {
        "buyer": fields.buyer,
        "seller": fields.seller,
        "invoice_number": fields.invoice_number,
        "total": _money_text(fields.total),
        "subtotal": _money_text(fields.subtotal) if fields.subtotal is not None else None,
        "buyer_tax_id": fields.buyer_tax_id,
        "seller_tax_id": fields.seller_tax_id,
        "invoice_date": fields.invoice_date,
        "items": fields.items,
    }


def _name_from_label(
    words: Sequence[Tuple[float, float, float, float, str]], side: str, page_width: float
) -> Optional[str]:
    """Find a name token next to a visible buyer/seller ``名称`` label."""

    if side == "buyer":
        labels = [w for w in words if w[0] < page_width * 0.50 and "名称" in w[4]]
        x_limit = page_width * 0.50
    else:
        labels = [w for w in words if w[0] >= page_width * 0.50 and "名称" in w[4]]
        x_limit = page_width
    if not labels:
        return None

    label = min(labels, key=lambda w: (w[1], w[0]))
    # Flattened text layers sometimes keep the label and value in one token.
    inline_value = re.sub(r"^名称\s*[:：]\s*", "", label[4]).strip()
    if inline_value and inline_value != label[4].strip():
        return inline_value
    candidates: List[Tuple[float, float, float, float, str]] = []
    for word in words:
        x0, y0, x1, y1, text = word
        # Some invoice templates place the first character of the value a few
        # points before the label's reported right edge.  Use the label's
        # left edge plus a small cushion instead of a hard x1 boundary.
        label_start = label[0] + max(10.0, (label[2] - label[0]) * 0.5)
        if x0 < label_start or x0 >= x_limit or abs(y0 - label[1]) > 5:
            continue
        if "名称" in text or text in {"购", "销", "买", "售", "方"}:
            continue
        if _parse_money(text) is not None or _is_numeric_token(text):
            continue
        if not re.search(r"[\u3400-\u9fffA-Za-z]", text):
            continue
        candidates.append(word)
    if not candidates:
        return None
    # A company name is normally one word token.  Pick the widest/longest
    # candidate while keeping the line nearest to the label.
    chosen = max(candidates, key=lambda w: (len(w[4]), w[2] - w[0]))
    return chosen[4].strip()


def _tax_id_from_label(
    words: Sequence[Tuple[float, float, float, float, str]], side: str, page_width: float
) -> str:
    """Find the tax id next to the buyer/seller 统一社会信用代码/纳税人识别号 label."""

    if side == "buyer":
        labels = [w for w in words if w[0] < page_width * 0.50 and "识别号" in w[4]]
        x_limit = page_width * 0.50
    else:
        labels = [w for w in words if w[0] >= page_width * 0.50 and "识别号" in w[4]]
        x_limit = page_width
    if not labels:
        return ""

    label = min(labels, key=lambda w: (w[1], w[0]))
    # 扁平文本层可能把标签和值合成一个 token
    inline_value = re.sub(r"^.*识别号\s*[:：]?\s*", "", label[4]).strip()
    if re.fullmatch(r"[0-9A-Za-z]{15,20}", inline_value):
        return inline_value

    label_start = label[0] + max(10.0, (label[2] - label[0]) * 0.5)
    for word in words:
        x0, y0, x1, y1, text = word
        if x0 < label_start or x0 >= x_limit or abs(y0 - label[1]) > 6:
            continue
        candidate = text.strip()
        if re.fullmatch(r"[0-9A-Za-z]{15,20}", candidate):
            return candidate
    return ""


def _find_invoice_date(page: fitz.Page) -> str:
    """Extract the invoice date in ISO format (YYYY-MM-DD)."""

    flat = re.sub(r"\s+", " ", page.get_text("text"))
    match = re.search(
        r"开票日期\s*[:：]?\s*([0-9]{4})[年\-/]([0-9]{1,2})[月\-/]([0-9]{1,2})日?",
        flat,
    )
    if not match:
        # 兜底：页面上任意 年月日 token
        match = re.search(r"([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日", flat)
    if match:
        year, month, day = match.groups()
        return f"{year}-{int(month):02d}-{int(day):02d}"
    return ""


def _find_invoice_number(page: fitz.Page, words: Sequence[Tuple[float, float, float, float, str]]) -> str:
    text = page.get_text("text")
    flat = re.sub(r"\s+", " ", text)
    match = re.search(r"发票号码\s*[:：]?\s*([0-9]{8,20})", flat)
    if match:
        return match.group(1)
    candidates = [w for w in words if _is_numeric_token(w[4]) and w[1] < page.rect.height * 0.25]
    if candidates:
        return max(candidates, key=lambda w: (w[1] < page.rect.height * 0.18, len(w[4])))[4]
    raise ValueError("Could not identify the invoice number from the PDF text layer")


def _find_amounts(
    page: fitz.Page, words: Sequence[Tuple[float, float, float, float, str]]
) -> Tuple[Decimal, Optional[Decimal]]:
    currency = "\u00a5"
    money_words: List[Tuple[Decimal, Tuple[float, float, float, float, str]]] = []
    for word in words:
        value = _parse_money(word[4])
        if value is not None:
            money_words.append((value, word))

    # The total is the currency token in the lower ``价税合计（小写）`` row.
    total_candidates = [
        (value, word)
        for value, word in money_words
        if currency in word[4] and word[1] > page.rect.height * 0.65
    ]
    if not total_candidates:
        flat = re.sub(r"\s+", " ", page.get_text("text"))
        match = re.search(r"小写[^0-9]{0,20}(?:\u00a5)?\s*([0-9][0-9,]*\.[0-9]{2})", flat)
        if match:
            total = _parse_money(match.group(1))
        else:
            # Last-resort fallback: the largest currency amount on the page.
            currency_candidates = [value for value, word in money_words if currency in word[4]]
            if not currency_candidates:
                raise ValueError("Could not identify the tax-inclusive total from the PDF")
            total = max(currency_candidates)
    else:
        total = max(total_candidates, key=lambda pair: pair[1][1])[0]

    subtotal_candidates = [
        value
        for value, word in money_words
        if word[1] > page.rect.height * 0.58
        and word[1] < page.rect.height * 0.75
        and word[0] < page.rect.width * 0.85
        and value != total
    ]
    subtotal = None
    if subtotal_candidates:
        # In standard invoices the pre-tax total is the largest amount in the
        # ``合计`` row after excluding the tax-inclusive total.
        subtotal = max(subtotal_candidates)
    return total, subtotal


def _number_token(text: str) -> Optional[float]:
    match = re.search(r"-?[0-9][0-9,]*(\.[0-9]+)?", text.replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _find_items(
    page: fitz.Page, words: Sequence[Tuple[float, float, float, float, str]]
) -> List[dict]:
    """解析发票货物/服务明细行（项目名称/规格/单位/数量/单价/金额/税率/税额）。

    列边界由表头 token 的 x 坐标确定；合计行（"合 计"）以下的内容排除。
    """

    header_name = next((w for w in words if "项目名称" in w[4]), None)
    if header_name is None:
        return []
    header_y = header_name[1]
    header_words = [w for w in words if abs(w[1] - header_y) <= 4]
    header_words.sort(key=lambda w: w[0])

    def find_x(pred, start=0):
        for i, w in enumerate(header_words[start:], start=start):
            if pred(w[4]):
                return w[0], i
        return None, None

    x_name, i1 = find_x(lambda t: "项目名称" in t)
    x_spec, i2 = find_x(lambda t: "规格" in t, i1 + 1)
    x_unit, i3 = find_x(lambda t: t in ("单", "位") or "单位" in t, i2 + 1)
    x_qty, i4 = find_x(lambda t: t in ("数", "量") or "数量" in t, i3 + 1)
    x_price, i5 = find_x(lambda t: t in ("单", "价") or "单价" in t, i4 + 1)
    x_amount, i6 = find_x(lambda t: t in ("金", "额") or "金额" in t, i5 + 1)
    x_rate, i7 = find_x(lambda t: "税率" in t or "征收率" in t, i6 + 1)
    x_tax, i8 = find_x(lambda t: t in ("税", "额") or "税额" in t, i7 + 1)
    if None in (x_name, x_spec, x_unit, x_qty, x_price, x_amount, x_rate, x_tax):
        return []

    col_xs = [x_name, x_spec, x_unit, x_qty, x_price, x_amount, x_rate, x_tax]
    # 列边界取相邻表头首字 x0 的中点（数据 token 按中心点归属列，
    # 名称列允许向左越出表头范围）
    bounds = [-float("inf")]
    bounds.extend((col_xs[i] + col_xs[i + 1]) / 2 for i in range(len(col_xs) - 1))
    bounds.append(float("inf"))
    header_bottom = max(w[3] for w in header_words)

    # 合计行位置：表头下方第一个"合"字 token
    total_candidates = [w for w in words if w[1] > header_bottom + 6 and "合" in w[4]]
    total_y = min(w[1] for w in total_candidates) - 2 if total_candidates else None
    if total_y is None:
        return []

    rows: List[dict] = []
    current: Optional[dict] = None

    item_words = [
        w for w in words if header_bottom < w[1] and w[3] < total_y and w[4].strip()
    ]
    item_words.sort(key=lambda w: (w[1], w[0]))

    def column_of(word: Tuple[float, float, float, float, str]) -> int:
        center = (word[0] + word[2]) / 2
        for index in range(len(bounds) - 1):
            if bounds[index] <= center < bounds[index + 1]:
                return index
        return len(bounds) - 2

    def band_has_data(band: List[Tuple[float, float, float, float, str]]) -> bool:
        return any(column_of(w) > 0 for w in band)

    band: List[Tuple[float, float, float, float, str]] = []
    band_y: Optional[float] = None
    for word in item_words:
        if band_y is None or abs(word[1] - band_y) <= 3:
            band.append(word)
            band_y = word[1] if band_y is None else band_y
            continue
        # 新的一带：处理旧 band
        if band_has_data(band):
            row = _build_item_row(band, bounds)
            current = row
            rows.append(row)
        elif current is not None:
            # 只有名称列的续行 → 并入上一行的名称
            name_extra = "".join(w[4].strip() for w in band if column_of(w) == 0)
            current["name"] = (current["name"] or "") + name_extra
        band = [word]
        band_y = word[1]
    if band:
        if band_has_data(band):
            rows.append(_build_item_row(band, bounds))
        elif current is not None:
            name_extra = "".join(w[4].strip() for w in band if column_of(w) == 0)
            current["name"] = (current["name"] or "") + name_extra

    return rows


def _build_item_row(
    band: List[Tuple[float, float, float, float, str]], bounds: Sequence[float]
) -> dict:
    def column_of(word: Tuple[float, float, float, float, str]) -> int:
        center = (word[0] + word[2]) / 2
        for index in range(len(bounds) - 1):
            if bounds[index] <= center < bounds[index + 1]:
                return index
        return len(bounds) - 2

    cells: List[List[str]] = [[] for _ in range(8)]
    for word in sorted(band, key=lambda w: (w[0], -w[2])):
        cells[column_of(word)].append(word[4].strip())

    def cell_text(index: int) -> str:
        return "".join(cells[index]).strip()

    raw_name = cell_text(0)
    classification, name = _split_classification(raw_name)

    def as_number(index: int) -> Optional[float]:
        text = cell_text(index).replace("¥", "")
        return _number_token(text)

    return {
        "classification": classification,
        "name": name,
        "spec": cell_text(1),
        "unit": cell_text(2),
        "quantity": as_number(3),
        "unit_price": as_number(4),
        "amount": as_number(5),
        "tax_rate": _parse_tax_rate(cell_text(6)),
        "tax": as_number(7),
    }


def _split_classification(raw_name: str) -> Tuple[str, str]:
    match = re.match(r"^\*([^*]*)\*(.*)$", raw_name)
    if match:
        classification, name = match.group(1).strip(), match.group(2).strip()
        return classification, name or classification
    return "", raw_name


def _parse_tax_rate(text: str) -> Optional[float]:
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*%", text)
    if not match:
        return None
    try:
        return float(match.group(1)) / 100.0
    except ValueError:
        return None


def extract_invoice_fields(pdf_path: Path) -> InvoiceFields:
    with fitz.open(pdf_path) as document:
        if not document:
            raise ValueError(f"PDF has no pages: {pdf_path}")
        page = document[0]
        words = _word_rows(page)
        buyer = _name_from_label(words, "buyer", page.rect.width)
        seller = _name_from_label(words, "seller", page.rect.width)
        if not buyer or not seller:
            raise ValueError(
                f"Could not identify both buyer and seller names in {pdf_path.name}; "
                "use a text-based electronic invoice or add the names manually."
            )
        invoice_number = _find_invoice_number(page, words)
        total, subtotal = _find_amounts(page, words)
        buyer_tax_id = _tax_id_from_label(words, "buyer", page.rect.width)
        seller_tax_id = _tax_id_from_label(words, "seller", page.rect.width)
        invoice_date = _find_invoice_date(page)
        items = _find_items(page, words)
        return InvoiceFields(
            buyer,
            seller,
            invoice_number,
            total,
            subtotal,
            buyer_tax_id,
            seller_tax_id,
            invoice_date,
            items,
        )


def _line_from_item(item: object) -> Optional[Tuple[float, float, float, float]]:
    if not isinstance(item, (tuple, list)) or len(item) < 3:
        return None
    if item[0] != "l":
        return None
    p0, p1 = item[1], item[2]
    try:
        return float(p0.x), float(p0.y), float(p1.x), float(p1.y)
    except (AttributeError, TypeError, ValueError):
        return None


def find_remark_box(page: fitz.Page) -> RemarkBox:
    """Detect the large remarks rectangle from vector lines when available."""

    width, height = page.rect.width, page.rect.height
    horizontal: List[Tuple[float, float, float]] = []
    vertical: List[Tuple[float, float, float]] = []
    for drawing in page.get_drawings():
        for item in drawing.get("items", []):
            line = _line_from_item(item)
            if not line:
                continue
            x0, y0, x1, y1 = line
            if abs(y1 - y0) <= 1.2 and abs(x1 - x0) >= width * 0.75:
                horizontal.append((min(y0, y1), min(x0, x1), max(x0, x1)))
            elif abs(x1 - x0) <= 1.2 and abs(y1 - y0) >= height * 0.08:
                vertical.append((min(x0, x1), min(y0, y1), max(y0, y1)))

    # Ignore the totals row above the remarks section.
    # The totals row often has a full-width border immediately above the
    # remarks box.  Start below roughly 72% of the page height so that border
    # is not mistaken for the remarks top edge.
    lower_lines = sorted({round(y, 2) for y, _, _ in horizontal if y > height * 0.72})
    if len(lower_lines) >= 2:
        top, bottom = lower_lines[0], lower_lines[-1]
        if bottom - top >= height * 0.05:
            left_candidates = [
                x
                for x, y0, y1 in vertical
                if y0 <= top + 2 and y1 >= bottom - 2 and x < width * 0.25
            ]
            right_candidates = [
                x
                for x, y0, y1 in vertical
                if y0 <= top + 2 and y1 >= bottom - 2 and x > width * 0.75
            ]
            x0 = min(left_candidates) if left_candidates else width * 0.02
            x1 = max(right_candidates) if right_candidates else width * 0.98
            label_dividers = [
                x
                for x, y0, y1 in vertical
                if y0 <= top + 2 and y1 >= bottom - 2 and width * 0.03 < x < width * 0.20
            ]
            content_x0 = (max(label_dividers) if label_dividers else x0) + 2.5
            return RemarkBox(x0, top, x1, bottom, content_x0)

    # Fallback for flattened/scanned PDFs that do not retain vector borders.
    return RemarkBox(
        width * 0.02,
        height * 0.745,
        width * 0.98,
        height * 0.887,
        width * 0.05,
    )


def _occupied_bottom(page: fitz.Page, box: RemarkBox) -> float:
    words = _word_rows(page)
    bottoms = [
        y1
        for x0, y0, x1, y1, _ in words
        if x1 > box.content_x0 and x0 < box.x1 and y0 >= box.y0 and y0 < box.y0 + 22
    ]
    return max(bottoms) if bottoms else box.y0 + 3


def _rect_intersects_words(
    rect: fitz.Rect,
    words: Sequence[Tuple[float, float, float, float, str]],
    padding: float = 0.5,
) -> Optional[str]:
    expanded = fitz.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding, rect.y1 + padding)
    for x0, y0, x1, y1, text in words:
        if fitz.Rect(x0, y0, x1, y1).intersects(expanded):
            return text
    return None


def _signature_png(signature_path: Path, max_width_px: int = 600) -> bytes:
    with Image.open(signature_path) as source:
        rgb = source.convert("RGB")
    gray = ImageOps.grayscale(rgb)

    # Convert the white paper background to transparency while retaining
    # anti-aliased dark strokes.
    alpha = gray.point(lambda value: 0 if value >= 245 else max(0, min(255, int((245 - value) * 255 / 225))))
    bbox = alpha.getbbox()
    if not bbox:
        raise ValueError(f"Signature image has no dark pixels: {signature_path}")
    pad = max(8, int(min(rgb.size) * 0.01))
    bbox = (
        max(0, bbox[0] - pad),
        max(0, bbox[1] - pad),
        min(rgb.width, bbox[2] + pad),
        min(rgb.height, bbox[3] + pad),
    )
    alpha = alpha.crop(bbox)
    rgba = Image.new("RGBA", alpha.size, (0, 0, 0, 0))
    rgba.putalpha(alpha)
    if rgba.width > max_width_px:
        new_height = max(1, round(rgba.height * max_width_px / rgba.width))
        resample = getattr(Image, "Resampling", Image).LANCZOS
        rgba = rgba.resize((max_width_px, new_height), resample)
    buffer = io.BytesIO()
    rgba.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def _find_label_font(explicit: Optional[Path]) -> Optional[Path]:
    candidates: List[Path] = []
    if explicit:
        candidates.append(explicit.expanduser())
    env_path = os.environ.get("INVOICE_LABEL_FONT")
    if env_path:
        candidates.append(Path(env_path).expanduser())
    candidates.extend(
        [
            Path.home() / "Library/Fonts/SourceHanSansCN-Regular#1.otf",
            Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
            Path("/System/Library/Fonts/STHeiti Medium.ttc"),
            Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
            Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        ]
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _label_png(text: str, font_path: Path, max_width_pt: float, max_height_pt: float) -> Tuple[bytes, float, float]:
    scale = 4
    chosen_font: Optional[ImageFont.FreeTypeFont] = None
    chosen_bbox: Optional[Tuple[int, int, int, int]] = None
    for point_size in [7.5, 7.0, 6.5, 6.0, 5.5, 5.0, 4.5]:
        font = ImageFont.truetype(str(font_path), max(1, round(point_size * scale)))
        bbox = font.getbbox(text)
        width_pt = (bbox[2] - bbox[0]) / scale
        height_pt = (bbox[3] - bbox[1]) / scale
        if width_pt <= max_width_pt and height_pt <= max_height_pt:
            chosen_font, chosen_bbox = font, bbox
            break
    if chosen_font is None or chosen_bbox is None:
        point_size = 4.0
        chosen_font = ImageFont.truetype(str(font_path), round(point_size * scale))
        chosen_bbox = chosen_font.getbbox(text)

    pad = 2 * scale
    width = chosen_bbox[2] - chosen_bbox[0] + 2 * pad
    height = chosen_bbox[3] - chosen_bbox[1] + 2 * pad
    image = Image.new("RGBA", (max(1, width), max(1, height)), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    draw.text(
        (pad - chosen_bbox[0], pad - chosen_bbox[1]),
        text,
        font=chosen_font,
        fill=(25, 25, 25, 235),
    )
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue(), image.width / scale, image.height / scale


def _safe_filename_component(value: str) -> str:
    value = re.sub(r"\s+", "", value.strip())
    value = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", value)
    return value or "未命名公司"


def _render_preview(pdf_path: Path, preview_dir: Path, target_width: int = 1600) -> Path:
    preview_dir.mkdir(parents=True, exist_ok=True)
    stat = pdf_path.stat()
    cache_key = hashlib.sha256(
        f"{pdf_path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}".encode("utf-8")
    ).hexdigest()[:20]
    output_path = preview_dir / f"{cache_key}.png"
    if output_path.exists():
        return output_path

    with fitz.open(pdf_path) as document:
        if not document:
            raise ValueError(f"PDF has no pages: {pdf_path}")
        page = document[0]
        scale = max(1.0, min(3.0, target_width / page.rect.width))
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False, annots=True)
        pixmap.save(str(output_path))
    return output_path


def _parse_rect(values: Optional[Sequence[str]]) -> Optional[fitz.Rect]:
    if not values:
        return None
    if len(values) != 4:
        raise ValueError("--signature-rect requires four numbers: x0 y0 x1 y1")
    try:
        x0, y0, x1, y1 = (float(v) for v in values)
    except ValueError as exc:
        raise ValueError("--signature-rect values must be numbers") from exc
    if x1 <= x0 or y1 <= y0:
        raise ValueError("--signature-rect must have x1>x0 and y1>y0")
    return fitz.Rect(x0, y0, x1, y1)


def add_signature(
    input_pdf: Path,
    signature_path: Path,
    output_pdf: Path,
    company_side: str = "seller",
    amount_field: str = "total",
    add_label: bool = True,
    label_font: Optional[Path] = None,
    signature_rect_override: Optional[fitz.Rect] = None,
    overwrite: bool = False,
) -> Tuple[InvoiceFields, fitz.Rect, Optional[fitz.Rect]]:
    if output_pdf.resolve() == input_pdf.resolve():
        raise ValueError("Output PDF must differ from the source PDF")
    if output_pdf.exists() and not overwrite:
        raise FileExistsError(f"Output already exists (use --overwrite): {output_pdf}")
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    fields = extract_invoice_fields(input_pdf)
    company = fields.seller if company_side == "seller" else fields.buyer
    amount = fields.total if amount_field == "total" else fields.subtotal
    if amount is None:
        raise ValueError("The requested amount field is not present in this invoice")
    archive_label = f"{company}-{_money_text(amount)}-dzfp{fields.invoice_number}"

    signature_stream = _signature_png(signature_path)
    label_stream: Optional[bytes] = None
    label_display_size: Optional[Tuple[float, float]] = None
    if add_label:
        font = _find_label_font(label_font)
        if font is None:
            raise FileNotFoundError(
                "No Chinese-capable label font found. Pass --label-font PATH or use --no-label."
            )
        # The label is rasterized so the PDF remains readable on systems where
        # a CJK font is not installed in the PDF viewer.
        label_stream, label_w, label_h = _label_png(archive_label, font, 400, 11)
        label_display_size = (label_w, label_h)

    temporary_path: Optional[Path] = None
    document = fitz.open(input_pdf)
    try:
        page = document[0]
        box = find_remark_box(page)
        occupied_bottom = _occupied_bottom(page, box)
        safe_top = max(box.y0 + 3, occupied_bottom + 2)
        safe_bottom = box.y1 - 2
        if safe_bottom - safe_top < 20:
            raise ValueError(
                f"The remarks box has less than 20 pt of free vertical space: {input_pdf.name}"
            )

        if signature_rect_override is None:
            # Keep the bottom/right edges inside the remarks border and size the
            # mark from the actual available height, preserving its aspect.
            with Image.open(io.BytesIO(signature_stream)) as signature_image:
                ratio = signature_image.width / signature_image.height
            height = min(44.0, safe_bottom - safe_top)
            width = min(125.0, height * ratio)
            signature_rect = fitz.Rect(
                box.x1 - 6 - width,
                safe_bottom - height,
                box.x1 - 6,
                safe_bottom,
            )
        else:
            signature_rect = signature_rect_override

        occupied_word = _rect_intersects_words(signature_rect, _word_rows(page))
        if occupied_word:
            raise ValueError(
                f"The requested signature area overlaps existing invoice text ({occupied_word!r}); "
                "use --signature-rect to choose another blank area."
            )

        page.insert_image(signature_rect, stream=signature_stream, keep_proportion=True, overlay=True)

        label_rect: Optional[fitz.Rect] = None
        if label_stream is not None and label_display_size is not None:
            label_w, label_h = label_display_size
            label_right = signature_rect.x0 - 8
            label_left = box.content_x0 + 2
            available_width = max(1.0, label_right - label_left)
            label_w = min(label_w, available_width)
            label_h = min(label_h, 10.0)
            # Keep a small bottom gap so the annotation never touches the red
            # remarks border.
            label_rect = fitz.Rect(
                label_left,
                safe_bottom - 4 - label_h,
                label_left + label_w,
                safe_bottom - 4,
            )
            occupied_word = _rect_intersects_words(label_rect, _word_rows(page))
            if occupied_word:
                raise ValueError(
                    f"The archive label area overlaps existing invoice text ({occupied_word!r}); "
                    "use --no-label or a custom --signature-rect."
                )
            page.insert_image(label_rect, stream=label_stream, keep_proportion=True, overlay=True)

        metadata = document.metadata or {}
        metadata["keywords"] = archive_label
        metadata["subject"] = "Reimbursement invoice with electronic signature"
        document.set_metadata(metadata)

        with tempfile.NamedTemporaryFile(
            prefix=output_pdf.stem + ".", suffix=".tmp.pdf", dir=output_pdf.parent, delete=False
        ) as handle:
            temporary_path = Path(handle.name)
        document.save(
            str(temporary_path),
            garbage=4,
            clean=1,
            deflate=True,
            deflate_images=True,
            preserve_metadata=True,
        )
        document.close()
        os.replace(temporary_path, output_pdf)
        temporary_path = None
        return fields, signature_rect, label_rect
    finally:
        if not document.is_closed:
            document.close()
        if temporary_path and temporary_path.exists():
            temporary_path.unlink()


def _collect_inputs(values: Iterable[str]) -> List[Path]:
    paths: List[Path] = []
    for raw in values:
        path = Path(raw).expanduser()
        if path.is_dir():
            paths.extend(sorted(p for p in path.glob("*.pdf") if p.is_file()))
        elif path.is_file() and path.suffix.lower() == ".pdf":
            paths.append(path)
        else:
            raise FileNotFoundError(f"Input PDF not found: {path}")
    # Preserve order while removing duplicates.
    unique: List[Path] = []
    seen = set()
    for path in paths:
        key = path.resolve()
        if key not in seen:
            unique.append(path)
            seen.add(key)
    return unique


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Add a transparent reimbursement signature and archive label to electronic invoices."
    )
    parser.add_argument("inputs", nargs="+", help="PDF file(s) or directories containing PDFs")
    parser.add_argument(
        "--signature",
        type=Path,
        default=None,
        help="Signature JPG/PNG (defaults to the bundled signature in this directory)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for generated PDFs (default: output/pdf)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Exact output path; only valid with one input PDF",
    )
    parser.add_argument(
        "--company-side",
        "--company-field",
        choices=("seller", "buyer"),
        default="seller",
        help="Company name used in the archive label and filename (default: seller)",
    )
    parser.add_argument(
        "--amount-field",
        choices=("total", "subtotal"),
        default="total",
        help="Amount used in the archive label and filename (default: total)",
    )
    parser.add_argument("--label-font", type=Path, default=None, help="Chinese-capable font for the visible label")
    parser.add_argument(
        "--no-label",
        action="store_true",
        help="Only add the signature; omit the visible archive label in the remarks box",
    )
    parser.add_argument(
        "--signature-rect",
        nargs=4,
        metavar=("X0", "Y0", "X1", "Y1"),
        help="Override signature rectangle in PDF points (top-left origin)",
    )
    parser.add_argument(
        "--inspect-json",
        action="store_true",
        help="Inspect invoice fields without modifying files and print one JSON array",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="Print machine-readable JSON lines while processing",
    )
    parser.add_argument(
        "--preview-dir",
        type=Path,
        default=None,
        help="Render first-page PNG previews into this cache directory",
    )
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing generated PDF")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        inputs = _collect_inputs(args.inputs)
        if args.output is not None and len(inputs) != 1:
            parser.error("--output can only be used with one input PDF")

        if args.inspect_json:
            results = []
            for input_pdf in inputs:
                try:
                    fields = extract_invoice_fields(input_pdf)
                    company = fields.seller if args.company_side == "seller" else fields.buyer
                    amount = fields.total if args.amount_field == "total" else fields.subtotal
                    if amount is None:
                        raise ValueError(f"No {args.amount_field} amount found")
                    preview_path = None
                    preview_error = None
                    if args.preview_dir is not None:
                        try:
                            preview_path = _render_preview(input_pdf, args.preview_dir)
                        except (FileNotFoundError, ValueError, RuntimeError) as exc:
                            preview_error = str(exc)
                    results.append(
                        {
                            "ok": True,
                            "input": str(input_pdf.resolve()),
                            **_fields_dict(fields),
                            "suggested_filename": (
                                f"{_safe_filename_component(company)}-{_money_text(amount)}-"
                                f"dzfp{fields.invoice_number}-已签字.pdf"
                            ),
                            "preview_path": str(preview_path.resolve()) if preview_path else None,
                            "preview_error": preview_error,
                        }
                    )
                except (FileNotFoundError, ValueError, RuntimeError) as exc:
                    results.append({"ok": False, "input": str(input_pdf.resolve()), "error": str(exc)})
            print(json.dumps(results, ensure_ascii=False))
            return 0 if all(result["ok"] for result in results) else 2

        signature = args.signature or (Path(__file__).resolve().parent / DEFAULT_SIGNATURE)
        if not signature.is_file():
            raise FileNotFoundError(f"未找到电子签图片：{signature}（请通过 --signature 指定）")
        rect_override = _parse_rect(args.signature_rect)

        for input_pdf in inputs:
            fields = extract_invoice_fields(input_pdf)
            company = fields.seller if args.company_side == "seller" else fields.buyer
            amount = fields.total if args.amount_field == "total" else fields.subtotal
            if amount is None:
                raise ValueError(f"No {args.amount_field} amount found in {input_pdf.name}")
            filename = (
                f"{_safe_filename_component(company)}-{_money_text(amount)}-"
                f"dzfp{fields.invoice_number}-已签字.pdf"
            )
            output = args.output if args.output is not None else args.output_dir / filename
            result_fields, sig_rect, label_rect = add_signature(
                input_pdf=input_pdf,
                signature_path=signature,
                output_pdf=output,
                company_side=args.company_side,
                amount_field=args.amount_field,
                add_label=not args.no_label,
                label_font=args.label_font,
                signature_rect_override=rect_override,
                overwrite=args.overwrite,
            )
            label_status = "added" if label_rect is not None else "omitted"
            preview_path = None
            if args.preview_dir is not None:
                try:
                    preview_path = _render_preview(output, args.preview_dir)
                except (FileNotFoundError, ValueError, RuntimeError):
                    preview_path = None
            if args.json_output:
                print(
                    json.dumps(
                        {
                            "ok": True,
                            "input": str(input_pdf.resolve()),
                            "output": str(output.resolve()),
                            **_fields_dict(result_fields),
                            "signature_rect": [round(v, 2) for v in sig_rect],
                            "label": label_status,
                            "preview_path": str(preview_path.resolve()) if preview_path else None,
                        },
                        ensure_ascii=False,
                    )
                )
            else:
                print(
                    f"created: {output}\n"
                    f"  buyer={result_fields.buyer}\n"
                    f"  seller={result_fields.seller}\n"
                    f"  total={_money_text(result_fields.total)}\n"
                    f"  invoice_number={result_fields.invoice_number}\n"
                    f"  signature_rect={tuple(round(v, 2) for v in sig_rect)}\n"
                    f"  label={label_status}"
                )
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        if getattr(args, "json_output", False):
            print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        else:
            print(f"error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
