import tempfile
import os
import pythoncom
import win32com.client as win32

def convert_doc_to_pdf(docx_bytes: bytes, output_base_name: str) -> bytes | None:
    """
    Recebe um arquivo DOCX (em bytes) e converte para PDF usando
    Microsoft Word via COM. Retorna os bytes do PDF gerado,
    ou None em caso de erro.

    - docx_bytes: conteúdo binário de um .docx
    - output_base_name: nome-base do arquivo temporário (sem extensão)
    """
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # --- 1) Criar arquivo .docx temporário ---
            caminho_docx = os.path.join(tmpdir, f"{output_base_name}.docx")
            with open(caminho_docx, "wb") as f_doc:
                f_doc.write(docx_bytes)

            # --- 2) Definir caminho de saída .pdf ---
            caminho_pdf = os.path.join(tmpdir, f"{output_base_name}.pdf")

            # --- 3) Inicializar COM e converter via Word ---
            pythoncom.CoInitialize()
            word = win32.DispatchEx("Word.Application")
            word.Visible = False

            doc = word.Documents.Open(caminho_docx)
            # FileFormat=17 é wdFormatPDF
            doc.SaveAs(caminho_pdf, FileFormat=17)
            doc.Close()
            word.Quit()

            # --- 4) Ler os bytes do PDF gerado ---
            with open(caminho_pdf, "rb") as f_pdf:
                return f_pdf.read()

    except Exception:
        return None
