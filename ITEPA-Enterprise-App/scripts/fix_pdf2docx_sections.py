"""Repair spurious blank pages created by pdf2docx footer-column sections."""

from __future__ import annotations

import argparse
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = {"w": WORD_NAMESPACE}
TYPE_ATTRIBUTE = f"{{{WORD_NAMESPACE}}}val"


def repair_sections(source: Path, destination: Path) -> int:
    """Change pdf2docx's body-level next-column breaks to continuous breaks."""
    with ZipFile(source, "r") as archive:
        document_xml = archive.read("word/document.xml")
        root = etree.fromstring(document_xml)

        repaired = 0
        for section_type in root.xpath("//w:sectPr/w:type", namespaces=NAMESPACES):
            if section_type.get(TYPE_ATTRIBUTE) == "nextColumn":
                section_type.set(TYPE_ATTRIBUTE, "continuous")
                repaired += 1

        # The sixth source-page boundary contains enough vertical spacing to
        # advance naturally. Its additional default next-page transition is
        # the single remaining blank page in LibreOffice/Word rendering.
        default_sections = root.xpath(
            "//w:pPr/w:sectPr[not(w:type)]", namespaces=NAMESPACES
        )
        for section in default_sections[5:6]:
            section_type = etree.Element(f"{{{WORD_NAMESPACE}}}type")
            section_type.set(TYPE_ATTRIBUTE, "continuous")
            section.insert(0, section_type)
            repaired += 1

        updated_xml = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        destination.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(
            prefix="itepa-docx-", suffix=".docx", dir=destination.parent, delete=False
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        try:
            with ZipFile(temporary_path, "w", compression=ZIP_DEFLATED) as output:
                for item in archive.infolist():
                    data = updated_xml if item.filename == "word/document.xml" else archive.read(item)
                    output.writestr(item, data)
            temporary_path.replace(destination)
        finally:
            if temporary_path.exists():
                temporary_path.unlink()

    return repaired


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    arguments = parser.parse_args()

    repaired = repair_sections(arguments.source, arguments.destination)
    print(f"Repaired {repaired} section breaks in {arguments.destination}")


if __name__ == "__main__":
    main()
