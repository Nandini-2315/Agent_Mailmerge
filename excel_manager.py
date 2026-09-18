import os
import hashlib
from datetime import datetime, date
from pathlib import Path
from typing import List, Dict, Any, Optional
import pandas as pd
import yaml

from .logger import logger

DEFAULT_CONFIG_PATH = Path("config.yaml")


def load_config(config_path: Optional[os.PathLike] = None) -> Dict[str, Any]:
    """
    Load configuration from YAML file or return defaults.
    """
    path = Path(config_path or DEFAULT_CONFIG_PATH)
    defaults = {
        "paths": {
            "excel_file": "input/studentsdata.xlsx",
            "template_file": "input/certificatetemp.docx",
            "fallback_template_file": "input/Student_Certificate.docx",
            "output_dir": "output",
            "manifest_file": "state/certificate_manifest.json",
            "log_file": "logs/certificate_agent.log",
        },
        "columns": {
            "id_column": "Certificate_ID",
            "fallback_id_columns": ["Certificate_ID", "Roll_Number", "Student_ID", "ID"],
            "required_columns": ["Certificate_ID", "Name"],
        },
        "policy": {
            "delete_removed_certificates": False,
        },
        "watcher": {
            "debounce_seconds": 2.0,
        },
        "pdf_conversion": {
            "backend": "auto",
            "timeout_seconds": 60,
        },
    }

    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                user_config = yaml.safe_load(f)
                if isinstance(user_config, dict):
                    # Deep merge
                    for section, values in user_config.items():
                        if isinstance(values, dict) and section in defaults:
                            defaults[section].update(values)
                        else:
                            defaults[section] = values
        except Exception as e:
            logger.warning(f"Error loading config from {path}: {e}. Using defaults.")

    # Allow environment variable override for deletion policy
    env_del = os.getenv("DELETE_REMOVED_CERTIFICATES")
    if env_del is not None:
        defaults["policy"]["delete_removed_certificates"] = env_del.strip().lower() in ("true", "1", "yes")

    return defaults


def compute_row_hash(student: Dict[str, Any]) -> str:
    """
    Calculate deterministic SHA-256 hash for a student record.
    Uses canonical sorted key-value pairs, ignoring internal keys starting with '_'.
    """
    normalized_items = []
    for k in sorted(student.keys()):
        if k.startswith("_"):
            continue
        v = student[k]
        val_str = str(v).strip() if v is not None else ""
        normalized_items.append(f"{k.strip()}:{val_str}")
    
    canonical_repr = "|".join(normalized_items)
    return hashlib.sha256(canonical_repr.encode("utf-8")).hexdigest()


def read_excel_records(
    excel_path: Optional[os.PathLike] = None,
    config: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """
    Read, normalize, and parse student data from Excel file.
    Returns a list of student dictionaries.
    """
    cfg = config or load_config()
    path = Path(excel_path or cfg["paths"]["excel_file"])

    if not path.exists():
        raise FileNotFoundError(f"Excel file not found at: {path}")

    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file at {path}: {e}")

    if df.empty:
        logger.warning(f"Excel file at {path} contains 0 rows.")
        return []

    # 1. Normalize column names (strip whitespace)
    df.columns = [str(c).strip() for c in df.columns]

    # 2. Normalize and resolve the ID column
    id_col = cfg["columns"].get("id_column", "Certificate_ID")
    fallback_ids = cfg["columns"].get("fallback_id_columns", ["Roll_Number", "Student_ID", "ID"])

    active_id_col = None
    if id_col in df.columns:
        active_id_col = id_col
    else:
        for fb in fallback_ids:
            if fb in df.columns:
                active_id_col = fb
                logger.info(f"Primary ID column '{id_col}' not found. Using fallback ID column '{fb}'.")
                break

    # 3. Clean and convert rows
    students = []
    for idx, row in df.iterrows():
        student_dict: Dict[str, Any] = {}
        for col in df.columns:
            val = row[col]
            if pd.isna(val) or val is None:
                student_dict[col] = ""
            elif isinstance(val, (datetime, date, pd.Timestamp)):
                student_dict[col] = val.strftime("%Y-%m-%d")
            else:
                student_dict[col] = str(val).strip()

        # Map to standard Certificate_ID if missing
        if id_col not in student_dict and active_id_col and active_id_col in student_dict:
            student_dict[id_col] = student_dict[active_id_col]
        elif active_id_col and active_id_col not in student_dict and id_col in student_dict:
            student_dict[active_id_col] = student_dict[id_col]

        # Attach original row number (1-indexed, +2 accounting for header row in Excel)
        student_dict["_row_number"] = idx + 2
        student_dict["_row_hash"] = compute_row_hash(student_dict)
        students.append(student_dict)

    logger.info(f"Loaded {len(students)} student records from {path}.")
    return students
