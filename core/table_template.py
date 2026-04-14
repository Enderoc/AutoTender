"""
table_template.py
Минимальная логика шаблонов таблиц
"""

import json
import os
import re
from datetime import datetime


class TableTemplate:

    @staticmethod
    def create_template(table, name=""):
        rows = len(table.rows)
        cols = max(len(r.cells) for r in table.rows)

        structure = []
        for r_idx, row in enumerate(table.rows):
            row_data = []
            for c_idx, cell in enumerate(row.cells):
                # Проверяем, есть ли фон у текста в ячейке
                has_background = False
                for p in cell.paragraphs:
                    for run in p.runs:
                        rPr = run._r.rPr
                        if rPr is not None:
                            shd = rPr.find('.//w:shd', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                            if shd is not None:
                                fill = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
                                if fill and fill.upper() != 'FFFFFF' and fill.upper() != 'AUTO':
                                    has_background = True
                                    break
                    if has_background:
                        break

                row_data.append({
                    "row": r_idx,
                    "col": c_idx,
                    "text": cell.text,
                    "has_background": has_background
                })
            structure.append(row_data)

        return {
            "name": name or f"template_{rows}x{cols}",
            "rows": rows,
            "cols": cols,
            "created": datetime.now().isoformat(),
            "structure": structure
        }

    @staticmethod
    def save_template(template, filepath):
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)

    @staticmethod
    def load_template(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            t = json.load(f)
            t["filepath"] = filepath
            return t

    @staticmethod
    def extract_placeholders(template):
        placeholders = set()
        for row in template["structure"]:
            for cell in row:
                matches = re.findall(r"\{(\w+)\}", cell["text"])
                placeholders.update(matches)
        return placeholders

    @staticmethod
    def fill_table(table, template, data_map, formatter):
        for row in template["structure"]:
            for cell_data in row:
                r = cell_data["row"]
                c = cell_data["col"]

                text = cell_data["text"]
                for key, value in data_map.items():
                    text = text.replace(f"{{{key}}}", str(value))

                # Если в ячейке был фон, используем специальный метод
                if cell_data.get("has_background", False):
                    formatter.replace_text_with_background(table.rows[r].cells[c], text)
                else:
                    formatter.replace_text(table.rows[r].cells[c], text)