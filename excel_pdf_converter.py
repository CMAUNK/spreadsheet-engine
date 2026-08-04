"""Exporta uma cópia XLSX para PDF usando o Microsoft Excel do computador."""
from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory


class PdfConversionError(RuntimeError):
    """The local computer could not use Excel to create the PDF."""


def convert_workbook_to_pdf(workbook_bytes: bytes, sheet_name: str) -> bytes:
    """Create a faithful PDF of one sheet, retaining Excel's own layout/styles."""
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PdfConversionError("O componente de exportação do Microsoft Excel não está instalado.") from exc

    pythoncom.CoInitialize()
    with TemporaryDirectory(prefix="smartsheet_") as temp_dir:
        source_path = Path(temp_dir) / "planilha_preenchida.xlsx"
        pdf_path = Path(temp_dir) / "planilha_preenchida.pdf"
        source_path.write_bytes(workbook_bytes)
        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(source_path.resolve()), ReadOnly=True)
            try:
                worksheet = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise PdfConversionError(f"A aba '{sheet_name}' não foi encontrada para exportar o PDF.") from exc
            # Alguns modelos possuem uma área de impressão antiga começando
            # na linha dos cabeçalhos. Para a cópia gerada, usamos a área
            # efetivamente ocupada, incluindo título e data na primeira linha.
            worksheet.PageSetup.PrintArea = worksheet.UsedRange.Address
            # 0 = xlTypePDF. A exportação é feita pelo próprio Excel, portanto
            # respeita fontes, cores, bordas e configurações de impressão.
            worksheet.ExportAsFixedFormat(0, str(pdf_path.resolve()))
            if not pdf_path.exists():
                raise PdfConversionError("O Excel não gerou o arquivo PDF.")
            return pdf_path.read_bytes()
        except PdfConversionError:
            raise
        except Exception as exc:
            raise PdfConversionError(f"Não foi possível exportar o PDF pelo Excel: {exc}") from exc
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
