"""Extract bounded text from local OOXML attachment bytes on stdin."""

import io
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile


MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024
MAX_PART_BYTES = 8 * 1024 * 1024
MAX_TOTAL_PART_BYTES = 32 * 1024 * 1024
MAX_ZIP_ENTRIES = 4096
MAX_MARKUP_TOKENS = 400000


def tag_name(value):
    return value.rsplit("}", 1)[-1]


class OfficeArchive:
    def __init__(self, archive):
        self.archive = archive
        members = archive.infolist()
        self.entries = {item.filename: item for item in members}
        if len(members) > MAX_ZIP_ENTRIES or len(members) != len(self.entries):
            raise ValueError("Invalid OOXML archive members")
        self.remaining = MAX_TOTAL_PART_BYTES

    def xml(self, name, required=True):
        entry = self.entries.get(name)
        if entry is None:
            if required:
                raise ValueError("Required OOXML part is missing")
            return None
        budget = min(MAX_PART_BYTES, self.remaining)
        if entry.file_size > budget or entry.flag_bits & 1:
            raise ValueError("OOXML part is encrypted or exceeds the extraction limit")
        with self.archive.open(entry) as file:
            data = file.read(budget + 1)
        if len(data) != entry.file_size or len(data) > budget:
            raise ValueError("Invalid OOXML part size")
        self.remaining -= len(data)
        encoding = "utf-16" if data.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
        source = data.decode(encoding)
        lowered = source.lower()
        if "<!doctype" in lowered or "<!entity" in lowered or source.count("<") > MAX_MARKUP_TOKENS:
            raise ValueError("Unsupported OOXML markup")
        return ET.fromstring(source)


def descendants(root, name):
    return (node for node in root.iter() if tag_name(node.tag) == name)


def word_lines(archive):
    document = archive.xml("word/document.xml")
    if tag_name(document.tag) != "document":
        raise ValueError("Invalid Word document part")
    paragraphs = {"t": None, "tab": "\t", "br": "\n", "cr": "\n"}
    for paragraph in descendants(document, "p"):
        parts = [paragraphs[tag_name(node.tag)] if tag_name(node.tag) != "t" else node.text or ""
                 for node in paragraph.iter() if tag_name(node.tag) in paragraphs]
        yield "".join(parts)


def worksheet_targets(archive):
    relations = archive.xml("xl/_rels/workbook.xml.rels")
    result = {}
    for relation in relations:
        kind = relation.get("Type", "").rsplit("/", 1)[-1]
        if kind != "worksheet" or relation.get("TargetMode") == "External":
            continue
        identifier, target = relation.get("Id"), relation.get("Target")
        if not identifier or not target or identifier in result:
            raise ValueError("Ambiguous workbook relationship")
        if target.startswith("/"):
            name = posixpath.normpath(target.lstrip("/"))
        else:
            name = posixpath.normpath(posixpath.join("xl", target))
        if not name.startswith("xl/worksheets/") or not name.endswith(".xml") or "\\" in name:
            raise ValueError("Unsupported worksheet location")
        result[identifier] = name
    return result


def worksheet_value(cell, shared):
    kind = cell.get("t", "")
    if kind == "inlineStr":
        return "".join(node.text or "" for node in descendants(cell, "t"))
    value = next((child.text or "" for child in cell if tag_name(child.tag) == "v"), "")
    if kind == "s":
        if not value.isascii() or not value.isdecimal() or len(value) > 7 or int(value) >= len(shared):
            raise ValueError("Invalid shared string reference")
        return shared[int(value)]
    return value


def excel_lines(archive):
    workbook = archive.xml("xl/workbook.xml")
    if tag_name(workbook.tag) != "workbook":
        raise ValueError("Invalid workbook part")
    targets = worksheet_targets(archive)
    strings = archive.xml("xl/sharedStrings.xml", required=False)
    shared = ["".join(node.text or "" for node in descendants(entry, "t"))
              for entry in descendants(strings, "si")] if strings is not None else []
    for sheet in descendants(workbook, "sheet"):
        identifier = next((value for name, value in sheet.attrib.items() if tag_name(name) == "id"), None)
        if identifier not in targets:
            # Chart sheets and external workbook relationships are not cells.
            continue
        worksheet = archive.xml(targets[identifier])
        if tag_name(worksheet.tag) != "worksheet":
            raise ValueError("Invalid worksheet part")
        yield sheet.get("name") or "Worksheet"
        for row in descendants(worksheet, "row"):
            cells = [worksheet_value(cell, shared) for cell in row if tag_name(cell.tag) == "c"]
            yield "\t".join(cells)


def bounded_text(lines, maximum):
    chunks = []
    remaining = maximum
    for line in lines:
        line = ("" if not chunks else "\n") + line
        if len(line) > remaining:
            chunks.append(line[:remaining])
            return {"text": "".join(chunks), "truncated": True}
        chunks.append(line)
        remaining -= len(line)
    return {"text": "".join(chunks), "truncated": False}


def main(format_name, maximum):
    if format_name not in {"docx", "xlsx"} or not 1 <= maximum <= 100000:
        raise ValueError("Invalid text extraction options")
    data = sys.stdin.buffer.read(MAX_ATTACHMENT_BYTES + 1)
    if len(data) > MAX_ATTACHMENT_BYTES:
        raise ValueError("Attachment exceeds the extraction limit")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        office = OfficeArchive(archive)
        lines = word_lines(office) if format_name == "docx" else excel_lines(office)
        print(json.dumps(bounded_text(lines, maximum), ensure_ascii=True))


if __name__ == "__main__":
    try:
        main(sys.argv[1], int(sys.argv[2]))
    except (IndexError, OSError, RuntimeError, ValueError, zipfile.BadZipFile, ET.ParseError):
        sys.exit(1)
