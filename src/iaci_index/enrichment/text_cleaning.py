"""Basic text cleaning helpers used across enrichment steps."""
import re


def normalize_whitespace(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return re.sub(r"\s+", " ", text).strip()


__all__ = ["normalize_whitespace"]
