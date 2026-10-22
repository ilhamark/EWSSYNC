#!/usr/bin/env python3
"""Static checks for VBA source files kept outside the Excel workbook.

This script deliberately does not attempt to execute VBA. Excel Desktop is
required for runtime tests; these checks catch repository mistakes before a
module or UserForm is imported into a workbook.
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE_EXTENSIONS = {".bas", ".frm", ".frx"}
TEXT_EXTENSIONS = {".bas", ".frm"}
VERSIONED_NAME = re.compile(r".+_v(?P<version>\d+\.\d+)\.(?:bas|frm|frx)$", re.IGNORECASE)
VBA_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', re.IGNORECASE | re.MULTILINE)
OPTION_EXPLICIT = re.compile(r"^Option\s+Explicit\s*$", re.IGNORECASE | re.MULTILINE)
OLE_OBJECT_BLOB = re.compile(r'OleObjectBlob\s*=\s*"([^"]+\.frx)"', re.IGNORECASE)


def fail(errors: list[str], message: str) -> None:
    errors.append(message)


def main() -> int:
    errors: list[str] = []
    modules: dict[str, list[Path]] = defaultdict(list)
    versions: set[str] = set()
    source_files = sorted(
        path for path in ROOT.iterdir() if path.is_file() and path.suffix.lower() in SOURCE_EXTENSIONS
    )

    if not source_files:
        fail(errors, "Tidak ada file source VBA (.bas, .frm, atau .frx).")

    for path in source_files:
        match = VERSIONED_NAME.fullmatch(path.name)
        if not match:
            fail(errors, f"Nama file harus menggunakan akhiran _vX.YY: {path.name}")
        else:
            versions.add(match.group("version"))

        if path.suffix.lower() not in TEXT_EXTENSIONS:
            continue

        text = path.read_text(encoding="utf-8-sig", errors="replace")
        name_match = VBA_NAME.search(text)
        if not name_match:
            fail(errors, f"Attribute VB_Name tidak ditemukan: {path.name}")
        else:
            modules[name_match.group(1)].append(path)

        if not OPTION_EXPLICIT.search(text):
            fail(errors, f"Option Explicit tidak ditemukan: {path.name}")

        if path.suffix.lower() == ".frm":
            blob_match = OLE_OBJECT_BLOB.search(text)
            if not blob_match:
                fail(errors, f"Referensi OleObjectBlob (.frx) tidak ditemukan: {path.name}")
            else:
                resource = path.parent / blob_match.group(1)
                if not resource.is_file():
                    fail(errors, f"Resource UserForm tidak ditemukan: {path.name} -> {resource.name}")

    if len(versions) != 1:
        fail(errors, "Versi file source harus seragam; ditemukan: " + ", ".join(sorted(versions)))

    for name, paths in sorted(modules.items()):
        if len(paths) > 1:
            fail(errors, "Attribute VB_Name ganda '" + name + "': " + ", ".join(path.name for path in paths))

    catalog = ROOT / "docs" / "MODULES.md"
    if not catalog.is_file():
        fail(errors, "Katalog modul tidak ditemukan: docs/MODULES.md")
    else:
        catalog_text = catalog.read_text(encoding="utf-8", errors="replace")
        for path in source_files:
            if path.name not in catalog_text:
                fail(errors, f"File source belum tercatat pada docs/MODULES.md: {path.name}")

    if errors:
        print("VALIDASI VBA GAGAL:")
        for error in errors:
            print(f"- {error}")
        return 1

    print("VALIDASI VBA BERHASIL")
    print(f"- {len(source_files)} file source diperiksa")
    print(f"- versi source: v{next(iter(versions))}")
    print(f"- {len(modules)} nama modul/UserForm unik")
    return 0


if __name__ == "__main__":
    sys.exit(main())
