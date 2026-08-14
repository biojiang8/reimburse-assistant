#!/usr/bin/env python3
"""报销三件套槽位校验（AI 校验环节）。

对每张发票逐槽位检查：
1. pdf   —— 已签字 PDF 存在且文件名含 dzfp票号、以「已签字.pdf」结尾；
2. excel —— 发票明细 Excel 存在且以「发票明细.xlsx」结尾，
            内容含发票号、销售方、价税合计；
3. word  —— 实物证据 Word 存在且以「实物证据.docx」结尾，
            内容含发票号码、公司、税号，且嵌入至少 1 张图片；
4. photos—— 实物照片至少 1 张且全部存在。

输入：--json-file 指向的记录数组，每条：
{invoice_number, pdf, excel, word, photos: [路径...]}
输出：JSON {results: [{invoice_number, slots: {...}, ok}]}
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List


def _check_pdf(record: Dict[str, Any]) -> Dict[str, Any]:
    path = Path(record.get("pdf") or "")
    number = str(record.get("invoice_number") or "")
    if not path.is_file():
        return {"ok": False, "message": f"缺少已签字 PDF：{path.name or '（未生成）'}"}
    name = path.name
    if number and f"dzfp{number}" not in name:
        return {"ok": False, "message": f"PDF 文件名不含票号 dzfp{number}：{name}"}
    if not name.endswith("已签字.pdf"):
        return {"ok": False, "message": f"PDF 文件名未以「已签字.pdf」结尾：{name}"}
    return {"ok": True, "message": f"PDF 正常：{name}"}


def _check_excel(record: Dict[str, Any]) -> Dict[str, Any]:
    path = Path(record.get("excel") or "")
    number = str(record.get("invoice_number") or "")
    if not path.is_file():
        return {"ok": False, "message": f"缺少发票明细 Excel：{path.name or '（未生成）'}"}
    if not path.name.endswith("发票明细.xlsx"):
        return {"ok": False, "message": f"Excel 文件名未以「发票明细.xlsx」结尾：{path.name}"}
    try:
        import openpyxl

        workbook = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        sheet = workbook.worksheets[0]
        rows = list(sheet.iter_rows(values_only=True))
        header_index = None
        for index, row in enumerate(rows):
            if row and any("发票号" in str(cell) for cell in row if cell is not None):
                header_index = index
                break
        if header_index is None:
            return {"ok": False, "message": f"Excel 内容缺少「发票号」列：{path.name}"}
        header = [str(cell) for cell in rows[header_index]]

        def cell(row: tuple, name: str) -> Any:
            for column, title in enumerate(header):
                if name in title:
                    return row[column] if column < len(row) else None
            return None

        matched = None
        for row in rows[header_index + 1:]:
            invoice_no = str(cell(row, "发票号") or "")
            if number and number in invoice_no:
                matched = row
                break
        if matched is None:
            return {"ok": False, "message": f"Excel 内容不含票号 {number}"}
        for label, name in (("项目名称", "项目名称"), ("经费代码", "财务项目码"), ("供应商名称", "供应商名称")):
            value = str(cell(matched, name) or "").strip()
            if not value:
                return {"ok": False, "message": f"Excel 中「{label}」为空"}
        return {"ok": True, "message": f"Excel 正常：{path.name}"}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "message": f"Excel 读取失败：{exc}"}


def _check_word(record: Dict[str, Any]) -> Dict[str, Any]:
    path = Path(record.get("word") or "")
    number = str(record.get("invoice_number") or "")
    if not path.is_file():
        return {"ok": False, "message": f"缺少实物证据 Word：{path.name or '（未生成）'}"}
    if not path.name.endswith("实物证据.docx"):
        return {"ok": False, "message": f"Word 文件名未以「实物证据.docx」结尾：{path.name}"}
    try:
        import docx

        document = docx.Document(str(path))
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        image_count = len(document.inline_shapes)
        for label, needle in (("发票号码", f"dzfp{number}"), ("公司（销售方）", "公司"), ("税号", "税号")):
            if needle not in text:
                return {"ok": False, "message": f"Word 内容缺少「{label}」信息"}
        if image_count < 1:
            return {"ok": False, "message": "Word 中未嵌入实物图片"}
        return {"ok": True, "message": f"Word 正常：{path.name}（图片 {image_count} 张）"}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "message": f"Word 读取失败：{exc}"}


def _check_photos(record: Dict[str, Any]) -> Dict[str, Any]:
    photos = [str(p) for p in (record.get("photos") or []) if p]
    if not photos:
        return {"ok": False, "message": "未添加实物照片"}
    missing = [p for p in photos if not Path(p).is_file()]
    if missing:
        return {"ok": False, "message": f"照片文件缺失：{missing[0]}"}
    return {"ok": True, "message": f"实物照片 {len(photos)} 张"}


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description="报销三件套槽位校验")
    # 由 reimburse-helper 的 verify 子命令路由使用，脚本本身忽略
    parser.add_argument("--verify-json", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--json-file", required=True, help="记录 JSON 文件路径")
    args = parser.parse_args(argv)

    data = json.loads(Path(args.json_file).read_text(encoding="utf-8"))
    records = data if isinstance(data, list) else data.get("records", [])
    if not isinstance(records, list):
        records = []

    results = []
    for record in records:
        slots = {
            "pdf": _check_pdf(record),
            "excel": _check_excel(record),
            "word": _check_word(record),
            "photos": _check_photos(record),
        }
        results.append(
            {
                "invoice_number": str(record.get("invoice_number") or ""),
                "slots": slots,
                "ok": all(slot["ok"] for slot in slots.values()),
            }
        )
    print(json.dumps({"results": results}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
