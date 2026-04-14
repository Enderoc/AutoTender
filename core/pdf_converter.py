import os
import shutil
import subprocess
import logging


class PdfConverter:
    def __init__(self):
        self.backend = self._detect_backend()

    def _detect_backend(self):
        try:
            import docx2pdf  # noqa
            logging.info("Используется backend: docx2pdf")
            return "docx2pdf"
        except ImportError:
            logging.info("docx2pdf не найден, пробуем LibreOffice...")

        if shutil.which("soffice"):
            logging.info("Используется backend: LibreOffice (soffice)")
            return "libreoffice"

        logging.warning("Не найден ни docx2pdf, ни LibreOffice. Используется заглушка.")
        return "stub"

    def convert(self, docx_path, pdf_path):
        if self.backend == "docx2pdf":
            self._convert_docx2pdf(docx_path, pdf_path)
        elif self.backend == "libreoffice":
            self._convert_libreoffice(docx_path, pdf_path)
        else:
            shutil.copy2(docx_path, pdf_path)
            logging.warning("PDF создан как копия DOCX (заглушка, не настоящий PDF).")

    def _convert_docx2pdf(self, docx_path, pdf_path):
        from docx2pdf import convert
        out_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
        os.makedirs(out_dir, exist_ok=True)
        convert(docx_path, pdf_path)

    def _convert_libreoffice(self, docx_path, pdf_path):
        out_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
        os.makedirs(out_dir, exist_ok=True)

        cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            docx_path,
        ]
        subprocess.run(cmd, check=True)

        base = os.path.splitext(os.path.basename(docx_path))[0]
        generated = os.path.join(out_dir, base + ".pdf")
        if os.path.abspath(generated) != os.path.abspath(pdf_path):
            shutil.move(generated, pdf_path)
