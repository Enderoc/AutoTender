import zipfile
import os
import shutil
import re
import copy
from uuid import uuid4
from typing import List, Tuple, Optional
from lxml import etree

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
DEFAULT_FORM_REGEX = re.compile(r"Форма\s+\d+[а-яА-Яa-zA-Z]*", re.IGNORECASE)
END_FORM_REGEX = re.compile(r"конец\s+формы", re.IGNORECASE)


class DocxSplitter:
    def __init__(self):
        self.method_regex_map = {
            "Роснефть": re.compile(r"Форма\s+\d+[а-яА-Я]*", re.IGNORECASE)
        }

    # ---------------------------------------------------------
    # Очистка скрытых символов
    # ---------------------------------------------------------
    def clean_hidden_chars(self, text: str) -> str:
        if not text:
            return text
        return (
            text.replace('&nbsp;', ' ')
            .replace('\xa0', ' ')
            .replace('\u200b', '')
            .replace('\u200e', '')
            .replace('\u200f', '')
            .replace('\ufeff', '')
        )

    def clean_xml_content(self, tree: etree.ElementTree) -> etree.ElementTree:
        for node in tree.xpath('//w:t', namespaces=NAMESPACE):
            if node.text:
                cleaned = self.clean_hidden_chars(node.text)
                if cleaned != node.text:
                    node.text = cleaned
        return tree

    # ---------------------------------------------------------
    # Удаление shading
    # ---------------------------------------------------------
    def _remove_background_shading(self, tree: etree.ElementTree):
        for rPr in tree.xpath('//w:rPr', namespaces=NAMESPACE):
            shd = rPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                rPr.remove(shd)

        for pPr in tree.xpath('//w:pPr', namespaces=NAMESPACE):
            shd = pPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                pPr.remove(shd)

        for tcPr in tree.xpath('//w:tcPr', namespaces=NAMESPACE):
            shd = tcPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                tcPr.remove(shd)

    # ---------------------------------------------------------
    # Удаление highlight
    # ---------------------------------------------------------
    def _remove_highlight(self, tree: etree.ElementTree):
        for rPr in tree.xpath('//w:rPr', namespaces=NAMESPACE):
            highlight = rPr.find('w:highlight', namespaces=NAMESPACE)
            if highlight is not None:
                rPr.remove(highlight)

    # ---------------------------------------------------------
    # Удаление highlight/shading из styles.xml
    # ---------------------------------------------------------
    def _remove_highlight_from_styles(self, styles_xml_path: str):
        if not os.path.exists(styles_xml_path):
            return

        tree = etree.parse(styles_xml_path)

        for rPr in tree.xpath('//w:rPr', namespaces=NAMESPACE):
            highlight = rPr.find('w:highlight', namespaces=NAMESPACE)
            if highlight is not None:
                rPr.remove(highlight)

            shd = rPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                rPr.remove(shd)

        for pPr in tree.xpath('//w:pPr', namespaces=NAMESPACE):
            shd = pPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                pPr.remove(shd)

        for tcPr in tree.xpath('//w:tcPr', namespaces=NAMESPACE):
            shd = tcPr.find('w:shd', namespaces=NAMESPACE)
            if shd is not None:
                tcPr.remove(shd)

        tree.write(styles_xml_path, xml_declaration=True, encoding='UTF-8', standalone="yes")

    # ---------------------------------------------------------
    # НОВОЕ: извлечение токенов из run (включая табуляции)
    # ---------------------------------------------------------
    def _extract_run_tokens(self, run):
        tokens = []
        for node in run:
            tag = etree.QName(node).localname
            if tag == "t":
                tokens.append(("text", node.text or ""))
            elif tag == "tab":
                tokens.append(("tab", None))
            elif tag == "br":
                tokens.append(("br", None))
            elif tag == "cr":
                tokens.append(("cr", None))
        return tokens

    # ---------------------------------------------------------
    # Нормализация runs с сохранением табуляций
    # ---------------------------------------------------------
    def _normalize_runs(self, tree: etree.ElementTree):

        paragraphs = tree.xpath('//w:p', namespaces=NAMESPACE)

        for p in paragraphs:
            if p.xpath('.//w:hyperlink', namespaces=NAMESPACE):
                continue

            runs = p.xpath('./w:r', namespaces=NAMESPACE)
            if len(runs) <= 1:
                continue

            tokens = []
            for r in runs:
                tokens.extend(self._extract_run_tokens(r))

            for r in runs:
                p.remove(r)

            new_r = etree.SubElement(p, '{%s}r' % NAMESPACE['w'])

            for kind, value in tokens:
                if kind == "text":
                    t = etree.SubElement(new_r, '{%s}t' % NAMESPACE['w'])
                    t.text = value

                    # сохраняем пробелы
                    if value and (value.startswith(" ") or value.endswith(" ")):
                        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

                elif kind == "tab":
                    etree.SubElement(new_r, '{%s}tab' % NAMESPACE['w'])
                elif kind == "br":
                    etree.SubElement(new_r, '{%s}br' % NAMESPACE['w'])
                elif kind == "cr":
                    etree.SubElement(new_r, '{%s}cr' % NAMESPACE['w'])

        # --- таблицы ---
        cells = tree.xpath('//w:tc', namespaces=NAMESPACE)

        for cell in cells:
            paragraphs = cell.xpath('.//w:p', namespaces=NAMESPACE)

            for p in paragraphs:
                if p.xpath('.//w:hyperlink', namespaces=NAMESPACE):
                    continue

                runs = p.xpath('./w:r', namespaces=NAMESPACE)
                if len(runs) <= 1:
                    continue

                tokens = []
                for r in runs:
                    tokens.extend(self._extract_run_tokens(r))

                for r in runs:
                    p.remove(r)

                new_r = etree.SubElement(p, '{%s}r' % NAMESPACE['w'])

                for kind, value in tokens:
                    if kind == "text":
                        t = etree.SubElement(new_r, '{%s}t' % NAMESPACE['w'])
                        t.text = value

                        if value and (value.startswith(" ") or value.endswith(" ")):
                            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

                    elif kind == "tab":
                        etree.SubElement(new_r, '{%s}tab' % NAMESPACE['w'])
                    elif kind == "br":
                        etree.SubElement(new_r, '{%s}br' % NAMESPACE['w'])
                    elif kind == "cr":
                        etree.SubElement(new_r, '{%s}cr' % NAMESPACE['w'])

    # ---------------------------------------------------------
    # Гиперссылки
    # ---------------------------------------------------------
    def _fix_hyperlink_structure(self, tree: etree.ElementTree):
        hyperlinks = tree.xpath('//w:hyperlink', namespaces=NAMESPACE)

        for hyperlink in hyperlinks:
            runs = hyperlink.xpath('./w:r', namespaces=NAMESPACE)
            if not runs:
                new_r = etree.SubElement(hyperlink, '{%s}r' % NAMESPACE['w'])
                new_t = etree.SubElement(new_r, '{%s}t' % NAMESPACE['w'])
                new_t.text = ''

    # ---------------------------------------------------------
    # Удаление стилей run
    # ---------------------------------------------------------
    def _remove_run_styles(self, tree: etree.ElementTree):
        for rPr in tree.xpath('//w:rPr', namespaces=NAMESPACE):
            rStyle = rPr.find('w:rStyle', namespaces=NAMESPACE)
            if rStyle is not None:
                rPr.remove(rStyle)

    # ---------------------------------------------------------
    # Обрезка по "конец формы"
    # ---------------------------------------------------------
    def _trim_form_at_end_marker(self, form_blocks: list) -> list:
        trimmed_blocks = []

        for block in form_blocks:
            block_text = ''.join(block.xpath('.//w:t/text()', namespaces=NAMESPACE)).strip()

            if END_FORM_REGEX.search(block_text):
                block_copy = copy.deepcopy(block)

                text_elements = block_copy.xpath('.//w:t', namespaces=NAMESPACE)

                full_text = ''
                text_positions = []

                for elem in text_elements:
                    if elem.text:
                        start = len(full_text)
                        full_text += elem.text
                        end = len(full_text)
                        text_positions.append((elem, start, end, elem.text))

                match = END_FORM_REGEX.search(full_text)
                if match:
                    cut = match.start()

                    for elem, start, end, original in text_positions:
                        if end <= cut:
                            continue
                        elif start >= cut:
                            elem.text = ''
                        else:
                            keep = cut - start
                            elem.text = original[:keep] if keep > 0 else ''

                    trimmed_blocks.append(block_copy)
                    break
            else:
                trimmed_blocks.append(block)

        return trimmed_blocks

    # ---------------------------------------------------------
    # Сохранение формы
    # ---------------------------------------------------------
    def _save_form_docx(
            self,
            base_dir: str,
            output_dir: str,
            form_name: str,
            form_blocks: list,
            clean_hidden_chars: bool
    ) -> str:

        temp_form_dir = os.path.join(output_dir, f"tmp_{uuid4().hex}")
        shutil.copytree(base_dir, temp_form_dir)

        doc_xml = os.path.join(temp_form_dir, "word", "document.xml")
        tree = etree.parse(doc_xml)
        body = tree.find('.//w:body', namespaces=NAMESPACE)

        for el in list(body):
            body.remove(el)

        trimmed_blocks = self._trim_form_at_end_marker(form_blocks)

        for el in trimmed_blocks:
            body.append(copy.deepcopy(el))

        if clean_hidden_chars:
            tree = self.clean_xml_content(tree)

        self._normalize_runs(tree)
        self._fix_hyperlink_structure(tree)
        self._remove_background_shading(tree)
        self._remove_highlight(tree)
        self._remove_run_styles(tree)

        tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone="yes")

        output_path = os.path.join(output_dir, f"{form_name}.docx")
        self._zip_dir(temp_form_dir, output_path)
        shutil.rmtree(temp_form_dir)

        return output_path

    # ---------------------------------------------------------
    # Извлечение форм
    # ---------------------------------------------------------
    def _extract_forms(self, xml_path: str, form_regex: re.Pattern) -> List[Tuple[str, list]]:
        tree = etree.parse(xml_path)
        body = tree.find('.//w:body', namespaces=NAMESPACE)
        blocks = list(body)
        forms, current_blocks, current_name = [], [], None

        for el in blocks:
            text = ''.join(el.xpath('.//w:t/text()', namespaces=NAMESPACE)).strip()
            match = form_regex.search(text)

            if match:
                if current_blocks and current_name:
                    forms.append((current_name, current_blocks))
                    current_blocks = []
                current_name = self._safe_filename(match.group())

            if current_name:
                current_blocks.append(el)

        if current_blocks and current_name:
            forms.append((current_name, current_blocks))

        return forms

    # ---------------------------------------------------------
    # Основной метод
    # ---------------------------------------------------------
    def split_document(
            self,
            source_docx: str,
            output_dir: str,
            method: Optional[str] = None,
            clean_hidden_chars: bool = True,
            progress_callback=None
    ) -> List[str]:

        if not os.path.exists(source_docx):
            raise FileNotFoundError(source_docx)

        regex = self.method_regex_map.get(method, DEFAULT_FORM_REGEX)
        os.makedirs(output_dir, exist_ok=True)
        temp_dir = self._create_temp_dir(output_dir)

        try:
            if progress_callback:
                progress_callback(10, 100, "Распаковка DOCX...")

            self._unzip_docx(source_docx, temp_dir)

            styles_xml = os.path.join(temp_dir, "word", "styles.xml")
            self._remove_highlight_from_styles(styles_xml)

            if progress_callback:
                progress_callback(30, 100, "Очистка XML...")

            xml_path = os.path.join(temp_dir, "word", "document.xml")
            if clean_hidden_chars:
                self._clean_document_xml(xml_path)

            if progress_callback:
                progress_callback(50, 100, "Извлечение форм...")

            forms = self._extract_forms(xml_path, regex)

            if progress_callback:
                progress_callback(60, 100, f"Найдено {len(forms)} форм")

            result_files = []
            total = max(len(forms), 1)

            for i, (name, blocks) in enumerate(forms):
                if progress_callback:
                    progress_callback(
                        60 + int(30 * i / total),
                        100,
                        f"Сохранение формы {i + 1}/{total}: {name}"
                    )

                result_file = self._save_form_docx(
                    base_dir=temp_dir,
                    output_dir=output_dir,
                    form_name=name,
                    form_blocks=blocks,
                    clean_hidden_chars=clean_hidden_chars
                )
                result_files.append(result_file)

            if progress_callback:
                progress_callback(100, 100, "Разделение завершено")

            return result_files

        finally:
            self._cleanup_temp_dir(temp_dir)

    # ---------------------------------------------------------
    # Вспомогательные методы
    # ---------------------------------------------------------
    def _clean_document_xml(self, xml_path: str):
        tree = etree.parse(xml_path)
        tree = self.clean_xml_content(tree)
        tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone="yes")

    def _safe_filename(self, name: str) -> str:
        return re.sub(r'[\\/*?:"<>|]', '_', name).strip().replace(' ', '_')[:120]

    def _create_temp_dir(self, base_dir: str) -> str:
        path = os.path.join(base_dir, f"tmp_{uuid4().hex}")
        os.makedirs(path, exist_ok=True)
        return path

    def _cleanup_temp_dir(self, path: str):
        if os.path.exists(path):
            shutil.rmtree(path)

    def _unzip_docx(self, docx_path: str, extract_dir: str):
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(extract_dir)

    def _zip_dir(self, folder: str, out_path: str):
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(folder):
                for f in files:
                    full = os.path.join(root, f)
                    z.write(full, os.path.relpath(full, folder))
