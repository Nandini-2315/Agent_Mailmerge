import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Optional

from .logger import logger


class PDFConversionError(Exception):
    """Raised when DOCX to PDF conversion fails."""
    pass


def _convert_with_word_vbs(docx_path: Path, pdf_path: Path, timeout: int = 60) -> bool:
    """
    Convert DOCX to PDF on Windows using Microsoft Word via late-binding COM script.
    Zero external pip dependency; uses Windows built-in cscript.exe host.
    """
    abs_docx = str(docx_path.resolve())
    abs_pdf = str(pdf_path.resolve())

    # VBScript late binding script (avoids early-bound TypeLib registration bugs)
    vbs_content = f"""
On Error Resume Next
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Unable to initialize Word.Application: " & Err.Description
    WScript.Quit 1
End If

wordApp.Visible = False
wordApp.DisplayAlerts = 0

Set doc = wordApp.Documents.Open("{abs_docx}", False, True)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Unable to open document: " & Err.Description
    wordApp.Quit
    WScript.Quit 2
End If

' 17 corresponds to wdFormatPDF
doc.SaveAs2 "{abs_pdf}", 17
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Unable to save as PDF: " & Err.Description
    doc.Close False
    wordApp.Quit
    WScript.Quit 3
End If

doc.Close False
wordApp.Quit
WScript.Echo "CONVERSION_SUCCESS"
"""
    # Write temporary VBS script
    temp_dir = Path(tempfile.gettempdir())
    vbs_file = temp_dir / f"convert_{uuid.uuid4().hex}.vbs"
    vbs_file.write_text(vbs_content, encoding="utf-8")

    try:
        cmd = ["cscript", "//nologo", str(vbs_file)]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout
        )
        stdout = result.stdout.strip()
        stderr = result.stderr.strip()

        if result.returncode == 0 and "CONVERSION_SUCCESS" in stdout and pdf_path.exists():
            return True

        err_msg = f"Word VBS script returned code {result.returncode}. Output: {stdout} {stderr}"
        logger.error(err_msg)
        raise PDFConversionError(err_msg)
    except subprocess.TimeoutExpired:
        raise PDFConversionError(f"Word PDF conversion timed out after {timeout} seconds.")
    finally:
        if vbs_file.exists():
            try:
                vbs_file.unlink()
            except OSError:
                pass


def _convert_with_libreoffice(docx_path: Path, pdf_path: Path, timeout: int = 60) -> bool:
    """
    Convert DOCX to PDF using headless LibreOffice (cross-platform).
    """
    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd:
        return False

    out_dir = pdf_path.parent
    cmd = [soffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)]
    
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        expected_pdf = out_dir / f"{docx_path.stem}.pdf"
        if expected_pdf.exists():
            if expected_pdf != pdf_path:
                shutil.move(str(expected_pdf), str(pdf_path))
            return True
        raise PDFConversionError(f"LibreOffice conversion failed: {res.stdout} {res.stderr}")
    except Exception as e:
        raise PDFConversionError(f"LibreOffice conversion exception: {e}")


def convert_docx_to_pdf(
    docx_path: Path,
    pdf_path: Path,
    backend: str = "auto",
    timeout: int = 60
) -> Path:
    """
    Convert a DOCX file to PDF using the configured backend.
    """
    if not docx_path.exists():
        raise FileNotFoundError(f"Source DOCX does not exist: {docx_path}")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    success = False
    last_error: Optional[Exception] = None

    if backend in ("auto", "word") and sys.platform == "win32":
        try:
            success = _convert_with_word_vbs(docx_path, pdf_path, timeout=timeout)
        except Exception as e:
            last_error = e
            logger.warning(f"Word conversion failed: {e}")

    if not success and backend in ("auto", "libreoffice"):
        try:
            success = _convert_with_libreoffice(docx_path, pdf_path, timeout=timeout)
        except Exception as e:
            last_error = e
            logger.warning(f"LibreOffice conversion failed: {e}")

    if not success:
        raise PDFConversionError(
            f"Failed to convert {docx_path.name} to PDF. "
            f"Ensure Microsoft Word or LibreOffice is installed. Error: {last_error}"
        )

    return pdf_path


def convert_docx_to_pdf_atomic(
    docx_path: Path,
    final_pdf_path: Path,
    old_pdf_path: Optional[Path] = None,
    backend: str = "auto",
    cleanup_docx: bool = True
) -> Path:
    """
    Safely and atomically generates a PDF certificate:
    1. Render to temporary PDF (e.g. temp_<uuid>_<filename>.pdf).
    2. Verify generation succeeded and file size > 0.
    3. If old_pdf_path exists and differs from final_pdf_path (e.g. student name changed), remove old PDF.
    4. Atomically replace final_pdf_path using os.replace.
    5. Clean up temporary DOCX.
    """
    final_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    temp_pdf = final_pdf_path.parent / f"temp_{uuid.uuid4().hex}_{final_pdf_path.name}"

    try:
        # Step 1: Generate temporary PDF
        convert_docx_to_pdf(docx_path, temp_pdf, backend=backend)

        # Step 2: Verify generation succeeded
        if not temp_pdf.exists() or temp_pdf.stat().st_size == 0:
            raise PDFConversionError(f"Generated PDF at {temp_pdf} is empty or missing.")

        # Step 3: Handle old PDF removal if filename changed (e.g., student name edited)
        if old_pdf_path and old_pdf_path.exists():
            try:
                resolved_old = old_pdf_path.resolve()
                resolved_new = final_pdf_path.resolve()
                if resolved_old != resolved_new:
                    old_pdf_path.unlink()
                    logger.info(f"Removed outdated certificate file: {old_pdf_path.name}")
            except Exception as e:
                logger.warning(f"Could not remove old certificate {old_pdf_path}: {e}")

        # Step 4: Atomic file replacement
        os.replace(str(temp_pdf), str(final_pdf_path))
        logger.info(f"Successfully generated certificate PDF: {final_pdf_path.name}")

    finally:
        # Clean up temp PDF if left over
        if temp_pdf.exists():
            try:
                temp_pdf.unlink()
            except OSError:
                pass

        # Step 5: Clean up temporary DOCX
        if cleanup_docx and docx_path.exists():
            try:
                docx_path.unlink()
            except OSError:
                pass

    return final_pdf_path
