"""Redaction toolkit for .docx documents (standard library only)."""

import os
from typing import List, Tuple

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def load_list(filename: str) -> List[str]:
    """Read a one-entry-per-line data file, ignoring blanks and comments."""
    path = os.path.join(DATA_DIR, filename)
    entries = []
    with open(path, "r", encoding="utf-8") as handle:
        for line in handle:
            line = line.split("#", 1)[0].strip()
            if line:
                entries.append(line)
    return entries


def load_gazetteers() -> Tuple[List[str], List[str]]:
    return load_list("person_names.txt"), load_list("organisations.txt")
