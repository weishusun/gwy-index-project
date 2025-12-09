"""Utilities for common data loading/saving paths."""
from pathlib import Path
from typing import Union

import pandas as pd

from .config import DATA_DIR

PathLike = Union[str, Path]


def ensure_dir(path: PathLike) -> Path:
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def read_excel(path: PathLike, **kwargs):
    return pd.read_excel(path, **kwargs)


def write_excel(df: pd.DataFrame, path: PathLike, **kwargs) -> None:
    path = Path(path)
    ensure_dir(path.parent)
    df.to_excel(path, **kwargs)


__all__ = ["ensure_dir", "read_excel", "write_excel", "DATA_DIR"]
