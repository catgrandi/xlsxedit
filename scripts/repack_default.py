#!/usr/bin/env python3
"""Repack default-xlsx-template/ into templates/default.xlsx (maintainer tool)."""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = ROOT / "src" / "xlsxedit" / "templates" / "default-xlsx-template"
OUTPUT = ROOT / "src" / "xlsxedit" / "templates" / "default.xlsx"


def main() -> None:
    if not (TEMPLATE_DIR / "[Content_Types].xml").is_file():
        raise SystemExit(f"missing template dir: {TEMPLATE_DIR}")
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(TEMPLATE_DIR):
            for name in files:
                if name == ".DS_Store":
                    continue
                path = Path(root) / name
                arc = path.relative_to(TEMPLATE_DIR).as_posix()
                zf.write(path, arc)
    print(f"packed {TEMPLATE_DIR} → {OUTPUT}")


if __name__ == "__main__":
    main()
