from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
from .manifest import get_active_records
from .logger import logger


@dataclass
class ChangeReport:
    new_records: List[Dict[str, Any]] = field(default_factory=list)
    modified_records: List[Tuple[Dict[str, Any], Dict[str, Any]]] = field(default_factory=list)
    unchanged_records: List[Dict[str, Any]] = field(default_factory=list)
    removed_records: List[Tuple[str, Dict[str, Any]]] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return bool(self.new_records or self.modified_records or self.removed_records)

    def summary(self) -> str:
        lines = [
            "===== EXCEL CHANGE DETECTION =====",
            "",
            f"New records: {len(self.new_records)}",
            f"Modified records: {len(self.modified_records)}",
            f"Unchanged records: {len(self.unchanged_records)}",
            f"Removed records: {len(self.removed_records)}",
        ]

        if self.new_records:
            lines.append("\nNEW:")
            for s in self.new_records:
                cid = s.get("Certificate_ID", "Unknown")
                name = s.get("Name", "Unknown")
                lines.append(f"  ✓ {cid} - {name}")

        if self.modified_records:
            lines.append("\nMODIFIED:")
            for s, old_rec in self.modified_records:
                cid = s.get("Certificate_ID", "Unknown")
                name = s.get("Name", "Unknown")
                lines.append(f"  ↻ {cid} - {name}")

        if self.removed_records:
            lines.append("\nREMOVED:")
            for cid, old_rec in self.removed_records:
                name = old_rec.get("student_name", "Unknown")
                lines.append(f"  ✗ {cid} - {name}")

        return "\n".join(lines)


def detect_changes(
    current_students: List[Dict[str, Any]],
    manifest: Dict[str, Any],
    output_dir: Optional[Path] = None
) -> ChangeReport:
    """
    Compare current Excel student records against the manifest state.
    Classifies every student into NEW, MODIFIED, UNCHANGED, or REMOVED.
    """
    report = ChangeReport()
    current_ids = set()

    active_manifest = get_active_records(manifest)

    for student in current_students:
        cid = str(student.get("Certificate_ID", "")).strip()
        if not cid:
            continue

        current_ids.add(cid)
        current_hash = student.get("_row_hash")

        if cid not in active_manifest:
            # Check if it was previously marked as deleted
            report.new_records.append(student)
            logger.info(f"Change Detection: NEW record detected -> {cid} ({student.get('Name')})")
        else:
            old_record = active_manifest[cid]
            old_hash = old_record.get("row_hash")
            old_output = old_record.get("output_file")
            
            # Verify if the physical PDF actually exists
            pdf_exists = False
            if output_dir and old_output:
                pdf_path = output_dir / old_output
                pdf_exists = pdf_path.exists() and pdf_path.stat().st_size > 0
            else:
                pdf_exists = True

            if current_hash != old_hash:
                report.modified_records.append((student, old_record))
                logger.info(f"Change Detection: MODIFIED record detected -> {cid} ({student.get('Name')})")
            elif not pdf_exists:
                # Hash is the same, but PDF is missing from output directory -> Regenerate!
                logger.info(f"Change Detection: UNCHANGED hash but missing PDF for {cid}. Marking for regeneration.")
                report.modified_records.append((student, old_record))
            else:
                report.unchanged_records.append(student)
                logger.debug(f"Change Detection: UNCHANGED record -> {cid}")

    # Check for removed records: active in manifest but not present in current Excel
    for cid, old_rec in active_manifest.items():
        if cid not in current_ids:
            report.removed_records.append((cid, old_rec))
            logger.info(f"Change Detection: REMOVED record detected -> {cid} ({old_rec.get('student_name')})")

    return report
