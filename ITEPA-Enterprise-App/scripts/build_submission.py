"""Create a clean, integrity-checked ITEPA submission ZIP."""

from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
import zipfile


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "Joshua_Nehohwa_ITEPA3-33_Submission.zip"
PACKAGE_NAME = "Joshua_Nehohwa_ITEPA3-33_Submission"
EXCLUDED_PARTS = {".venv", ".pytest_cache", "__pycache__", "tmp", "submission", ".git"}
EXCLUDED_NAMES = {".coverage", ".DS_Store"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo", ".prof"}


def should_include(path: Path) -> bool:
    relative = path.relative_to(ROOT)
    if any(part in EXCLUDED_PARTS for part in relative.parts):
        return False
    if path.name in EXCLUDED_NAMES or path.suffix in EXCLUDED_SUFFIXES:
        return False
    if path == OUTPUT:
        return False
    return path.is_file()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> None:
    files = sorted(path for path in ROOT.rglob("*") if should_include(path))
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="itepa-submission-") as temp_directory:
        manifest = Path(temp_directory) / "MANIFEST_SHA256.txt"
        manifest.write_text(
            "\n".join(f"{sha256(path)}  {path.relative_to(ROOT)}" for path in files) + "\n",
            encoding="utf-8",
        )
        with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for path in files:
                archive.write(path, Path(PACKAGE_NAME) / path.relative_to(ROOT))
            archive.write(manifest, Path(PACKAGE_NAME) / manifest.name)

    with zipfile.ZipFile(OUTPUT) as archive:
        bad_file = archive.testzip()
        if bad_file is not None:
            raise RuntimeError(f"ZIP integrity failure: {bad_file}")
        names = archive.namelist()
        required = {
            f"{PACKAGE_NAME}/README.md",
            f"{PACKAGE_NAME}/AI_USE_LOG.md",
            f"{PACKAGE_NAME}/pyproject.toml",
            f"{PACKAGE_NAME}/output/pdf/Joshua_Nehohwa_ITEPA3-33_Practical_Report.pdf",
            f"{PACKAGE_NAME}/MANIFEST_SHA256.txt",
        }
        missing = required.difference(names)
        if missing:
            raise RuntimeError(f"Required submission files missing: {sorted(missing)}")

    print(f"Created {OUTPUT} with {len(files) + 1} files")
    print(f"SHA-256: {sha256(OUTPUT)}")


if __name__ == "__main__":
    main()

