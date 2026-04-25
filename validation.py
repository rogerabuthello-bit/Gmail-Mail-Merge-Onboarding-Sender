import re
from email.utils import parseaddr

import pandas as pd


EMAIL_PATTERN = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.IGNORECASE)


def validate_required_columns(dataframe: pd.DataFrame) -> list[str]:
    errors: list[str] = []
    if "Email" not in dataframe.columns:
        errors.append("The uploaded file must include an `Email` column.")
    return errors


def normalize_email(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().casefold()


def is_valid_email(value: object) -> bool:
    normalized = normalize_email(value)
    _, parsed_email = parseaddr(normalized)
    return bool(parsed_email and EMAIL_PATTERN.match(parsed_email))


def annotate_recipients(dataframe: pd.DataFrame) -> pd.DataFrame:
    annotated = dataframe.copy()
    annotated["Email"] = annotated["Email"].astype(str).str.strip()
    annotated["_normalized_email"] = annotated["Email"].map(normalize_email)
    annotated["_is_valid_email"] = annotated["Email"].map(is_valid_email)
    duplicate_mask = annotated["_normalized_email"].duplicated(keep="first")
    annotated["_is_duplicate"] = duplicate_mask & annotated["_normalized_email"].ne("")
    return annotated


def get_duplicate_emails(dataframe: pd.DataFrame) -> list[str]:
    normalized = dataframe["Email"].map(normalize_email)
    duplicate_values = normalized[normalized.duplicated(keep=False) & normalized.ne("")]
    return sorted(duplicate_values.unique().tolist())
