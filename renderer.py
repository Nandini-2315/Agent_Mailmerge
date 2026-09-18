import os
import re
from pathlib import Path
from typing import Dict, Any, Optional
import docx

from .logger import logger


def sanitize_filename_component(text: str) -> str:
    """
    Sanitize a string component for safe use in Windows/Unix filenames.
    Preserves actual student data in memory while creating safe filenames.
    Characters like / \\ : * ? " < > | are stripped or converted.
    """
    if not text:
        return "Unknown"
    # Replace spaces with underscores
    cleaned = str(text).strip()
    # Replace invalid filename characters
    cleaned = re.sub(r'[\\/*?:"<>|]', "", cleaned)
    cleaned = re.sub(r"\s+", "_", cleaned)
    # Remove leading/trailing dots or underscores
    cleaned = cleaned.strip("._")
    return cleaned or "Student"


def get_certificate_filename(student_name: str, cert_id: str) -> str:
    """
    Construct safe output PDF filename: e.g. 'Ravi_Kumar_CERT-001.pdf'
    """
    safe_name = sanitize_filename_component(student_name)
    safe_id = sanitize_filename_component(cert_id)
    return f"{safe_name}_{safe_id}.pdf"


def replace_in_paragraph(paragraph: docx.text.paragraph.Paragraph, replacements: Dict[str, str]) -> None:
    """
    Robust placeholder replacement across runs in a python-docx paragraph.
    Handles placeholders split across multiple Word runs while preserving font styling.
    """
    if not paragraph.runs:
        return

    for placeholder, replacement in replacements.items():
        while True:
            full_text = "".join(r.text for r in paragraph.runs)
            idx = full_text.find(placeholder)
            if idx == -1:
                break

            end_idx = idx + len(placeholder)

            # Map full_text character positions to (run_index, char_index_within_run)
            char_map = []
            for r_idx, run in enumerate(paragraph.runs):
                for c_idx in range(len(run.text)):
                    char_map.append((r_idx, c_idx))

            if not char_map or end_idx > len(char_map):
                break

            start_r, start_c = char_map[idx]
            end_r, end_c = char_map[end_idx - 1]  # inclusive end index

            if start_r == end_r:
                # Placeholder is contained entirely within one run
                r = paragraph.runs[start_r]
                r.text = r.text[:start_c] + str(replacement) + r.text[end_c + 1:]
            else:
                # Placeholder spans multiple runs
                first_run = paragraph.runs[start_r]
                last_run = paragraph.runs[end_r]

                # First run gets the prefix and the entire replacement text
                first_run.text = first_run.text[:start_c] + str(replacement)

                # Intermediate runs are cleared
                for mid_r in range(start_r + 1, end_r):
                    paragraph.runs[mid_r].text = ""

                # Last run keeps only its suffix
                last_run.text = last_run.text[end_c + 1:]


def render_certificate_docx(
    template_path: Path,
    student_data: Dict[str, Any],
    output_docx_path: Path
) -> Path:
    """
    Loads template DOCX, performs deterministic placeholder substitution,
    and saves to output_docx_path.
    Preserves template layout, fonts, borders, background, and visual formatting.
    """
    if not template_path.exists():
        raise FileNotFoundError(f"Template DOCX not found at: {template_path}")

    doc = docx.Document(template_path)

    # 1. Prepare replacements dictionary
    replacements: Dict[str, str] = {}
    for k, v in student_data.items():
        if k.startswith("_"):
            continue
        replacements[f"{{{{{k}}}}}"] = str(v)

    # Add common alias mappings
    cert_id = student_data.get("Certificate_ID", "")
    roll_no = student_data.get("Roll_Number", "")
    if cert_id and "{{Roll_Number}}" not in replacements:
        replacements["{{Roll_Number}}"] = str(cert_id)
    if roll_no and "{{Certificate_ID}}" not in replacements:
        replacements["{{Certificate_ID}}"] = str(roll_no)

    # 2. Process all paragraphs in body (including text boxes, shapes, drawingML, tables)
    # Using xpath(".//w:p") captures paragraphs inside w:body, w:tbl, w:txbxContent, v:textbox, w:drawing
    all_p_elements = doc._element.xpath(".//w:p")
    for p_elem in all_p_elements:
        p = docx.text.paragraph.Paragraph(p_elem, doc)
        replace_in_paragraph(p, replacements)

    # 3. Process headers and footers across all sections
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header and not header.is_linked_to_previous:
                for p_elem in header._element.xpath(".//w:p"):
                    p = docx.text.paragraph.Paragraph(p_elem, doc)
                    replace_in_paragraph(p, replacements)

        for footer in (section.footer, section.first_page_footer, section.even_page_footer):
            if footer and not footer.is_linked_to_previous:
                for p_elem in footer._element.xpath(".//w:p"):
                    p = docx.text.paragraph.Paragraph(p_elem, doc)
                    replace_in_paragraph(p, replacements)

    # 4. Save to temporary or target DOCX
    output_docx_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx_path))
    logger.debug(f"Rendered temporary certificate DOCX: {output_docx_path}")
    return output_docx_path
