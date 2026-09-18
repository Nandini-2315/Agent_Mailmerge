import os
import tempfile
import uuid
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

from .logger import logger
from .excel_manager import load_config, read_excel_records, compute_row_hash
from .validator import (
    validate_environment,
    validate_dataset_structure,
    validate_students,
    extract_template_placeholders,
    ValidationError
)
from .manifest import (
    load_manifest,
    save_manifest,
    update_record,
    mark_deleted,
    remove_record,
    get_active_records
)
from .change_detector import detect_changes, ChangeReport
from .renderer import (
    get_certificate_filename,
    render_certificate_docx
)
from .pdf_converter import (
    convert_docx_to_pdf_atomic,
    PDFConversionError
)


def validate_inputs(
    excel_path: Optional[str] = None,
    template_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Validate Excel file and certificate template without generating certificates.
    Checks file presence, required columns, duplicate IDs, and template placeholders.
    """
    cfg = load_config()
    e_path = Path(excel_path) if excel_path else None
    t_path = Path(template_path) if template_path else None

    result: Dict[str, Any] = {
        "valid": True,
        "excel_path": "",
        "template_path": "",
        "records_count": 0,
        "placeholders": [],
        "errors": [],
        "warnings": []
    }

    try:
        res_e, res_t, _ = validate_environment(excel_path=e_path, template_path=t_path, config=cfg)
        result["excel_path"] = str(res_e.resolve())
        result["template_path"] = str(res_t.resolve())
    except Exception as e:
        result["valid"] = False
        result["errors"].append(str(e))
        return result

    # Validate template placeholders
    try:
        placeholders = extract_template_placeholders(res_t)
        result["placeholders"] = sorted(list(placeholders))
    except Exception as e:
        result["valid"] = False
        result["errors"].append(f"Template error: {e}")

    # Validate Excel data
    try:
        records = read_excel_records(res_e, config=cfg)
        result["records_count"] = len(records)
        validate_dataset_structure(records, config=cfg)
        valid_students, failed = validate_students(records, config=cfg)
        if failed:
            for ident, reason in failed.items():
                result["warnings"].append(f"{ident}: {reason}")
    except ValidationError as ve:
        result["valid"] = False
        result["errors"].append(str(ve))
    except Exception as e:
        result["valid"] = False
        result["errors"].append(str(e))

    return result


def read_student_data(excel_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Read and parse student data from Excel file.
    """
    cfg = load_config()
    e_path = Path(excel_path) if excel_path else Path(cfg["paths"]["excel_file"])
    if not e_path.exists():
        raise FileNotFoundError(f"Excel file not found at: {e_path}")

    records = read_excel_records(e_path, config=cfg)
    return records


def generate_certificate(
    student: Union[Dict[str, Any], str],
    excel_path: Optional[str] = None,
    template_path: Optional[str] = None,
    output_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
    Generate one certificate PDF for one student record or Certificate_ID.
    Renders DOCX, converts atomically to PDF, and updates the manifest.
    """
    cfg = load_config()
    e_path = Path(excel_path) if excel_path else None
    t_path = Path(template_path) if template_path else None
    o_dir = Path(output_dir) if output_dir else None

    res_e, res_t, res_o = validate_environment(e_path, t_path, o_dir, config=cfg)
    manifest_path = Path(cfg["paths"]["manifest_file"])
    manifest = load_manifest(manifest_path)

    # If student is passed as string ID, find in Excel
    if isinstance(student, str):
        cert_id_search = student.strip()
        all_students = read_excel_records(res_e, config=cfg)
        matched = [s for s in all_students if s.get("Certificate_ID") == cert_id_search]
        if not matched:
            raise ValueError(f"Student with Certificate_ID '{cert_id_search}' not found in Excel.")
        student_data = matched[0]
    elif isinstance(student, dict):
        student_data = student
    else:
        raise TypeError("Student must be a dictionary or a Certificate_ID string.")

    # Validate student fields
    valid_students, failed = validate_students([student_data], config=cfg)
    if failed:
        reason = list(failed.values())[0]
        raise ValueError(f"Invalid student record: {reason}")

    student_data = valid_students[0]
    cert_id = student_data["Certificate_ID"]
    student_name = student_data.get("Name", "Student")
    row_hash = student_data.get("_row_hash") or compute_row_hash(student_data)

    # Determine filenames
    output_pdf_filename = get_certificate_filename(student_name, cert_id)
    final_pdf_path = res_o / output_pdf_filename

    # Check previous manifest for old output filename
    old_pdf_path = None
    if cert_id in manifest:
        old_output = manifest[cert_id].get("output_file")
        if old_output:
            old_pdf_path = res_o / old_output

    # Render temporary DOCX
    temp_docx = Path(tempfile.gettempdir()) / f"cert_{uuid.uuid4().hex}.docx"
    try:
        render_certificate_docx(res_t, student_data, temp_docx)
        
        # Convert atomically to PDF
        backend = cfg["pdf_conversion"].get("backend", "auto")
        convert_docx_to_pdf_atomic(
            docx_path=temp_docx,
            final_pdf_path=final_pdf_path,
            old_pdf_path=old_pdf_path,
            backend=backend
        )

        # Update manifest
        update_record(
            manifest=manifest,
            cert_id=cert_id,
            student_name=student_name,
            row_hash=row_hash,
            output_file=output_pdf_filename,
            extra_data=student_data,
            status="active"
        )
        save_manifest(manifest, manifest_path)

        return {
            "status": "success",
            "certificate_id": cert_id,
            "student_name": student_name,
            "output_file": str(final_pdf_path.name),
            "output_path": str(final_pdf_path.resolve())
        }
    except Exception as e:
        logger.error(f"Failed to generate certificate for {cert_id} ({student_name}): {e}")
        raise


def generate_all_certificates(
    excel_path: Optional[str] = None,
    template_path: Optional[str] = None,
    output_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
    Generate certificates for every valid student in the Excel file.
    Does not halt entire generation if one individual student fails.
    Updates the manifest and returns a comprehensive report.
    """
    cfg = load_config()
    res_e, res_t, res_o = validate_environment(
        Path(excel_path) if excel_path else None,
        Path(template_path) if template_path else None,
        Path(output_dir) if output_dir else None,
        config=cfg
    )
    manifest_path = Path(cfg["paths"]["manifest_file"])
    manifest = load_manifest(manifest_path)

    # 1. Read records
    students = read_excel_records(res_e, config=cfg)
    logger.info(f"===== CERTIFICATE GENERATION =====\nExcel records: {len(students)}")

    # 2. Check dataset structure (duplicate IDs will raise ValidationError)
    validate_dataset_structure(students, config=cfg)

    # 3. Validate student rows
    valid_students, failed_validation = validate_students(students, config=cfg)

    generated: List[str] = []
    failed: Dict[str, str] = dict(failed_validation)

    backend = cfg["pdf_conversion"].get("backend", "auto")

    # 4. Generate certificates for valid students
    for student_data in valid_students:
        cert_id = student_data["Certificate_ID"]
        student_name = student_data.get("Name", "Student")
        row_hash = student_data.get("_row_hash") or compute_row_hash(student_data)

        output_filename = get_certificate_filename(student_name, cert_id)
        final_pdf_path = res_o / output_filename

        old_pdf_path = None
        if cert_id in manifest:
            old_output = manifest[cert_id].get("output_file")
            if old_output:
                old_pdf_path = res_o / old_output

        temp_docx = Path(tempfile.gettempdir()) / f"cert_{uuid.uuid4().hex}.docx"

        try:
            render_certificate_docx(res_t, student_data, temp_docx)
            convert_docx_to_pdf_atomic(
                docx_path=temp_docx,
                final_pdf_path=final_pdf_path,
                old_pdf_path=old_pdf_path,
                backend=backend
            )

            update_record(
                manifest=manifest,
                cert_id=cert_id,
                student_name=student_name,
                row_hash=row_hash,
                output_file=output_filename,
                extra_data=student_data,
                status="active"
            )
            generated.append(f"{student_name} ({cert_id})")
            logger.info(f"✓ Generated: {student_name} ({cert_id})")

        except Exception as e:
            err_msg = str(e)
            failed[cert_id] = err_msg
            logger.error(f"✗ Failed generating {cert_id} ({student_name}): {err_msg}")
        finally:
            if temp_docx.exists():
                try:
                    temp_docx.unlink()
                except OSError:
                    pass

    # 5. Save updated manifest
    save_manifest(manifest, manifest_path)

    # Format summary matching Section 22
    summary_lines = [
        "===== CERTIFICATE GENERATION =====",
        "",
        f"Excel records: {len(students)}",
        "",
        "Generated:"
    ]
    for gen in generated:
        summary_lines.append(f"✓ {gen}")

    if failed:
        summary_lines.append("\nFailed:")
        for cid, reason in failed.items():
            summary_lines.append(f"✗ {cid}: {reason}")

    summary_lines.append(f"\nTotal generated: {len(generated)}")
    summary_lines.append(f"Failed: {len(failed)}")

    summary_text = "\n".join(summary_lines)
    logger.debug(summary_text)

    return {
        "total_records": len(students),
        "generated_count": len(generated),
        "failed_count": len(failed),
        "generated": generated,
        "failed": failed,
        "summary": summary_text
    }


def detect_excel_changes(
    excel_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Compare current Excel records against the manifest state.
    Returns categorized NEW, MODIFIED, UNCHANGED, and REMOVED records.
    """
    cfg = load_config()
    res_e, _, res_o = validate_environment(
        Path(excel_path) if excel_path else None,
        config=cfg
    )
    manifest_path = Path(cfg["paths"]["manifest_file"])
    manifest = load_manifest(manifest_path)

    students = read_excel_records(res_e, config=cfg)
    validate_dataset_structure(students, config=cfg)
    valid_students, _ = validate_students(students, config=cfg)

    report = detect_changes(valid_students, manifest, output_dir=res_o)

    return {
        "new_count": len(report.new_records),
        "modified_count": len(report.modified_records),
        "unchanged_count": len(report.unchanged_records),
        "removed_count": len(report.removed_records),
        "has_changes": report.has_changes,
        "summary": report.summary(),
        "report": report
    }


def process_excel_changes(
    excel_path: Optional[str] = None,
    delete_removed: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Intelligently synchronize certificates based on Excel changes:
    - Generates only NEW student certificates.
    - Regenerates MODIFIED student certificates and replaces existing PDFs.
    - Leaves UNCHANGED certificates completely untouched.
    - Handles REMOVED students according to deletion policy.
    """
    cfg = load_config()
    res_e, res_t, res_o = validate_environment(
        Path(excel_path) if excel_path else None,
        config=cfg
    )
    manifest_path = Path(cfg["paths"]["manifest_file"])
    manifest = load_manifest(manifest_path)

    should_delete_removed = (
        delete_removed if delete_removed is not None
        else cfg["policy"].get("delete_removed_certificates", False)
    )

    students = read_excel_records(res_e, config=cfg)
    validate_dataset_structure(students, config=cfg)
    valid_students, failed_validation = validate_students(students, config=cfg)

    report = detect_changes(valid_students, manifest, output_dir=res_o)
    logger.info(report.summary())

    created: List[str] = []
    regenerated: List[str] = []
    failed: Dict[str, str] = dict(failed_validation)

    backend = cfg["pdf_conversion"].get("backend", "auto")

    # 1. Process NEW records
    for student_data in report.new_records:
        cert_id = student_data["Certificate_ID"]
        student_name = student_data.get("Name", "Student")
        row_hash = student_data.get("_row_hash") or compute_row_hash(student_data)

        output_filename = get_certificate_filename(student_name, cert_id)
        final_pdf_path = res_o / output_filename
        temp_docx = Path(tempfile.gettempdir()) / f"cert_new_{uuid.uuid4().hex}.docx"

        try:
            render_certificate_docx(res_t, student_data, temp_docx)
            convert_docx_to_pdf_atomic(
                docx_path=temp_docx,
                final_pdf_path=final_pdf_path,
                backend=backend
            )
            update_record(
                manifest=manifest,
                cert_id=cert_id,
                student_name=student_name,
                row_hash=row_hash,
                output_file=output_filename,
                extra_data=student_data,
                status="active"
            )
            created.append(f"{cert_id} - {student_name}")
            logger.info(f"✓ Created NEW certificate: {cert_id} ({student_name})")
        except Exception as e:
            failed[cert_id] = str(e)
            logger.error(f"✗ Failed creating {cert_id}: {e}")
        finally:
            if temp_docx.exists():
                try:
                    temp_docx.unlink()
                except OSError:
                    pass

    # 2. Process MODIFIED records (replace old PDF)
    for student_data, old_rec in report.modified_records:
        cert_id = student_data["Certificate_ID"]
        student_name = student_data.get("Name", "Student")
        row_hash = student_data.get("_row_hash") or compute_row_hash(student_data)

        output_filename = get_certificate_filename(student_name, cert_id)
        final_pdf_path = res_o / output_filename

        old_pdf_path = None
        old_output = old_rec.get("output_file")
        if old_output:
            old_pdf_path = res_o / old_output

        temp_docx = Path(tempfile.gettempdir()) / f"cert_mod_{uuid.uuid4().hex}.docx"

        try:
            render_certificate_docx(res_t, student_data, temp_docx)
            convert_docx_to_pdf_atomic(
                docx_path=temp_docx,
                final_pdf_path=final_pdf_path,
                old_pdf_path=old_pdf_path,
                backend=backend
            )
            update_record(
                manifest=manifest,
                cert_id=cert_id,
                student_name=student_name,
                row_hash=row_hash,
                output_file=output_filename,
                extra_data=student_data,
                status="active"
            )
            regenerated.append(f"{cert_id} - {student_name}")
            logger.info(f"↻ Regenerated MODIFIED certificate: {cert_id} ({student_name})")
        except Exception as e:
            failed[cert_id] = str(e)
            logger.error(f"✗ Failed regenerating {cert_id}: {e}")
        finally:
            if temp_docx.exists():
                try:
                    temp_docx.unlink()
                except OSError:
                    pass

    # 3. Process REMOVED records
    removed_reported: List[str] = []
    for cid, old_rec in report.removed_records:
        student_name = old_rec.get("student_name", "Unknown")
        removed_reported.append(f"{cid} - {student_name}")

        if should_delete_removed:
            # Delete physical PDF
            old_output = old_rec.get("output_file")
            if old_output:
                old_pdf = res_o / old_output
                if old_pdf.exists():
                    try:
                        old_pdf.unlink()
                        logger.info(f"Deleted PDF for removed student {cid}: {old_output}")
                    except Exception as e:
                        logger.warning(f"Could not delete PDF {old_pdf}: {e}")
            remove_record(manifest, cid)
        else:
            # Mark as deleted in manifest
            mark_deleted(manifest, cid)
            logger.info(f"Marked certificate {cid} as deleted_from_excel in manifest (PDF preserved).")

    # 4. Save updated manifest
    save_manifest(manifest, manifest_path)

    # Format output matching Section 22
    summary_lines = [
        report.summary(),
        "",
        "Certificates updated successfully.",
        "",
        f"Created: {len(created)}",
        f"Regenerated: {len(regenerated)}",
        f"Unchanged: {len(report.unchanged_records)}",
        f"Failed: {len(failed)}"
    ]

    summary_text = "\n".join(summary_lines)
    logger.debug(summary_text)

    return {
        "created_count": len(created),
        "regenerated_count": len(regenerated),
        "unchanged_count": len(report.unchanged_records),
        "removed_count": len(report.removed_records),
        "failed_count": len(failed),
        "created": created,
        "regenerated": regenerated,
        "removed": removed_reported,
        "failed": failed,
        "summary": summary_text
    }