import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Any, Optional
from .logger import logger

DEFAULT_MANIFEST_PATH = Path("state/certificate_manifest.json")


def load_manifest(manifest_path: Optional[os.PathLike] = None) -> Dict[str, Any]:
    """
    Load the certificate manifest from disk.
    Returns an empty dict if the file does not exist or is corrupted.
    """
    path = Path(manifest_path or DEFAULT_MANIFEST_PATH)
    if not path.exists():
        logger.debug(f"Manifest file not found at {path}. Returning empty manifest.")
        return {}
    
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
            logger.warning(f"Manifest at {path} is not a valid JSON object. Returning empty manifest.")
            return {}
    except Exception as e:
        logger.error(f"Error loading manifest from {path}: {e}")
        return {}


def save_manifest(manifest: Dict[str, Any], manifest_path: Optional[os.PathLike] = None) -> None:
    """
    Atomically save the certificate manifest to disk.
    Uses a temporary file before replacement to prevent corruption.
    """
    path = Path(manifest_path or DEFAULT_MANIFEST_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    
    temp_path = path.with_suffix(".tmp")
    try:
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2, ensure_ascii=False)
        
        # Atomic file replacement
        if temp_path.exists():
            os.replace(temp_path, path)
            logger.debug(f"Manifest successfully saved to {path} with {len(manifest)} records.")
    except Exception as e:
        logger.error(f"Failed to save manifest to {path}: {e}")
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
        raise


def update_record(
    manifest: Dict[str, Any],
    cert_id: str,
    student_name: str,
    row_hash: str,
    output_file: str,
    extra_data: Optional[Dict[str, Any]] = None,
    status: str = "active"
) -> Dict[str, Any]:
    """
    Update or create a student certificate entry in the manifest.
    """
    record = {
        "student_name": student_name,
        "row_hash": row_hash,
        "output_file": output_file,
        "status": status,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    if extra_data:
        record["data"] = {k: v for k, v in extra_data.items() if not k.startswith("_")}
    manifest[cert_id] = record
    return record


def mark_deleted(manifest: Dict[str, Any], cert_id: str) -> bool:
    """
    Mark a certificate as deleted_from_excel in the manifest.
    """
    if cert_id in manifest:
        manifest[cert_id]["status"] = "deleted_from_excel"
        manifest[cert_id]["deleted_at"] = datetime.now(timezone.utc).isoformat()
        return True
    return False


def remove_record(manifest: Dict[str, Any], cert_id: str) -> bool:
    """
    Completely remove a certificate record from the manifest.
    """
    if cert_id in manifest:
        del manifest[cert_id]
        return True
    return False


def get_active_records(manifest: Dict[str, Any]) -> Dict[str, Any]:
    """
    Return all records that are currently active (not marked as deleted_from_excel).
    """
    return {
        cert_id: record
        for cert_id, record in manifest.items()
        if isinstance(record, dict) and record.get("status") != "deleted_from_excel"
    }
