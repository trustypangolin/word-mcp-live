import asyncio
import zipfile

import pytest
from docx import Document
from lxml import etree

from word_document_server.tools.footnote_tools import add_footnote, delete_footnote
from word_document_server.utils.file_utils import get_file_lock


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def run(coro):
    return asyncio.run(coro)


@pytest.fixture
def document(tmp_path):
    path = tmp_path / "source.docx"
    doc = Document()
    doc.add_paragraph("Beginning match ending")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Table paragraph"
    doc.add_paragraph("Last paragraph")
    doc.save(path)
    return str(path)


def xml_part(path, name):
    with zipfile.ZipFile(path) as archive:
        return etree.fromstring(archive.read(name))


def paragraphs(path):
    return xml_part(path, "word/document.xml").xpath("//w:p", namespaces=NS)


def footnotes(path):
    root = xml_part(path, "word/footnotes.xml")
    return {int(node.get(f"{{{W}}}id")): "".join(node.xpath(".//w:t/text()", namespaces=NS))
            for node in root.xpath("//w:footnote", namespaces=NS)}


def paragraph_runs(para):
    return ["ref" if run.xpath("./w:footnoteReference", namespaces=NS)
            else "".join(run.xpath(".//w:t/text()", namespaces=NS))
            for run in para.xpath("./w:r", namespaces=NS)]


@pytest.mark.parametrize("selector,position,index,expected", [
    ({"search_text": "match"}, "before", 0, ["ref", "Beginning match ending"]),
    ({"search_text": "match"}, "after", 0, ["Beginning match ending", "ref"]),
    ({"paragraph_index": 1}, "before", 1, ["ref", "Table paragraph"]),
    ({"paragraph_index": 2}, "after", 2, ["Last paragraph", "ref"]),
])
def test_add_at_paragraph_boundary(document, selector, position, index, expected):
    result = run(add_footnote(document, "actual note", position=position, **selector))
    assert result["success"], result
    assert paragraph_runs(paragraphs(document)[index]) == expected
    assert "actual note" in footnotes(document)[result["details"]["footnote_id"]]
    Document(document)  # The resulting package is readable.


def test_search_uses_first_matching_paragraph(document):
    doc = Document(document)
    doc.add_paragraph("match again")
    doc.save(document)
    result = run(add_footnote(document, "note", search_text="match"))
    assert result["success"]
    assert paragraph_runs(paragraphs(document)[0])[-1] == "ref"
    assert paragraph_runs(paragraphs(document)[-1]) == ["match again"]


@pytest.mark.parametrize("selector", ["id", "search"])
def test_delete_by_id_or_search(document, selector):
    added = run(add_footnote(document, "note", search_text="match"))
    note_id = added["details"]["footnote_id"]
    kwargs = {"footnote_id": note_id} if selector == "id" else {"search_text": "match"}
    result = run(delete_footnote(document, **kwargs))
    assert result["success"], result
    assert result["details"]["footnote_id"] == note_id
    assert not paragraphs(document)[0].xpath(".//w:footnoteReference", namespaces=NS)
    assert note_id not in footnotes(document)
    Document(document)


def test_delete_preserves_text_sharing_reference_run(document):
    added = run(add_footnote(document, "note", paragraph_index=0))
    with zipfile.ZipFile(document) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    root = etree.fromstring(parts["word/document.xml"])
    reference = root.xpath("//w:footnoteReference", namespaces=NS)[0]
    etree.SubElement(reference.getparent(), f"{{{W}}}t").text = "Keep me"
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
    with zipfile.ZipFile(document, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    result = run(delete_footnote(document, footnote_id=added["details"]["footnote_id"]))
    assert result["success"]
    assert "Keep me" in Document(document).paragraphs[0].text


def add_orphan(path, note_id=99):
    with zipfile.ZipFile(path) as archive:
        parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    root = etree.fromstring(parts["word/footnotes.xml"])
    orphan = etree.SubElement(root, f"{{{W}}}footnote", {f"{{{W}}}id": str(note_id)})
    etree.SubElement(orphan, f"{{{W}}}p")
    parts["word/footnotes.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


@pytest.mark.parametrize("clean", [True, False])
def test_clean_orphans_option(document, clean):
    added = run(add_footnote(document, "note", paragraph_index=0))
    add_orphan(document)
    result = run(delete_footnote(document, footnote_id=added["details"]["footnote_id"], clean_orphans=clean))
    assert result["success"], result
    assert (99 in footnotes(document)) is not clean
    assert ("99" in result["details"]["orphans_removed"]) is clean


@pytest.mark.parametrize("options", [
    {}, {"search_text": "match", "paragraph_index": 0},
    {"search_text": ""}, {"search_text": "", "paragraph_index": 0},
    {"paragraph_index": -1}, {"paragraph_index": 1.5},
    {"paragraph_index": True}, {"paragraph_index": "1"},
    {"paragraph_index": 0, "position": "middle"},
])
def test_invalid_add_options_do_not_modify_files(document, tmp_path, options):
    original = open(document, "rb").read()
    output = tmp_path / "output.docx"
    output.write_bytes(b"existing")
    result = run(add_footnote(document, "note", output_filename=str(output), **options))
    assert not result["success"]
    assert result["details"] is None
    assert open(document, "rb").read() == original
    assert output.read_bytes() == b"existing"


@pytest.mark.parametrize("options", [
    {}, {"footnote_id": 2, "search_text": "match"}, {"search_text": ""},
    {"footnote_id": 0}, {"footnote_id": -1}, {"footnote_id": True},
    {"footnote_id": "2"}, {"footnote_id": 2, "clean_orphans": "false"},
])
def test_invalid_delete_options_do_not_modify_files(document, tmp_path, options):
    original = open(document, "rb").read()
    output = tmp_path / "output.docx"
    output.write_bytes(b"existing")
    result = run(delete_footnote(document, output_filename=str(output), **options))
    assert not result["success"]
    assert result["details"] is None
    assert open(document, "rb").read() == original
    assert output.read_bytes() == b"existing"


def test_output_add_and_delete_keep_source_unchanged(document, tmp_path):
    original = open(document, "rb").read()
    added_path = str(tmp_path / "added.docx")
    deleted_path = str(tmp_path / "deleted.docx")
    added = run(add_footnote(document, "note", paragraph_index=0, output_filename=added_path))
    assert added["success"], added
    assert open(document, "rb").read() == original
    added_bytes = open(added_path, "rb").read()
    deleted = run(delete_footnote(added_path, footnote_id=added["details"]["footnote_id"],
                                 output_filename=deleted_path))
    assert deleted["success"], deleted
    assert open(added_path, "rb").read() == added_bytes
    assert added["details"]["footnote_id"] not in footnotes(deleted_path)


@pytest.mark.parametrize("operation", ["add", "delete"])
def test_output_is_preserved_when_target_missing(document, tmp_path, operation):
    output = tmp_path / "existing.docx"
    output.write_bytes(b"previous contents")
    if operation == "add":
        result = run(add_footnote(document, "note", search_text="missing", output_filename=str(output)))
    else:
        result = run(delete_footnote(document, footnote_id=999, output_filename=str(output)))
    assert not result["success"]
    assert output.read_bytes() == b"previous contents"


def test_missing_source_and_out_of_range(document, tmp_path):
    result = run(add_footnote(str(tmp_path / "missing"), "note", paragraph_index=0))
    assert not result["success"] and "not found" in result["message"].lower()
    result = run(add_footnote(document, "note", paragraph_index=50))
    assert not result["success"] and "out of range" in result["message"].lower()


def test_output_lock_is_observed(document, tmp_path):
    output = str(tmp_path / "output.docx")

    async def check():
        async with get_file_lock(output):
            task = asyncio.create_task(add_footnote(document, "note", paragraph_index=0, output_filename=output))
            await asyncio.sleep(0)
            assert not task.done()
        return await task

    assert run(check())["success"]