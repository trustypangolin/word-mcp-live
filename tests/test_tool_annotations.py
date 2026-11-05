"""Check annotations exposed to MCP clients, including mixed-purpose tools."""

import asyncio

import pytest
from docx import Document
from fastmcp import Client, FastMCP

from word_document_server import main
from word_document_server.tools import document_tools, format_tools, footnote_tools


@pytest.fixture(scope="module")
def annotations():
    # Register on a separate server so the tests do not mutate the global server.
    original = main.mcp
    server = FastMCP("Annotation test")
    main.mcp = server
    try:
        main.register_tools()
    finally:
        main.mcp = original

    async def list_annotations():
        async with Client(server) as client:
            tools = await client.list_tools()
        return {tool.name: tool.annotations.model_dump(exclude_none=True) for tool in tools}

    return asyncio.run(list_annotations())


@pytest.mark.parametrize("name", [
    "get_document_text", "word_live_list_open", "word_live_get_comments",
    "word_live_get_page_text", "list_tracked_changes",
])
def test_read_tools_are_read_only(annotations, name):
    assert annotations[name]["readOnlyHint"] is True


@pytest.mark.parametrize("name", [
    "add_paragraph", "word_live_insert_paragraphs", "word_live_format_text",
    "word_live_format_table", "word_live_add_comment", "word_live_reply_to_comment",
    "word_live_resolve_comment", "word_live_toggle_track_changes",
    "word_live_insert_image", "word_live_insert_cross_reference",
    "word_live_insert_equation", "word_live_set_page_layout",
    "word_live_set_paragraph_spacing", "word_live_add_watermark",
    "set_page_layout", "set_paragraph_spacing", "add_section_break",
    "add_table_of_contents",
])
def test_non_destructive_writers_have_explicit_hints(annotations, name):
    assert annotations[name]["readOnlyHint"] is False
    assert annotations[name]["destructiveHint"] is False


@pytest.mark.parametrize("name", [
    "create_document", "copy_document", "merge_documents", "convert_to_pdf",
    "delete_paragraph", "search_and_replace", "merge_table_cells",
    "format_table_cell_text", "add_footnote_after_text", "add_footnote_before_text",
    "add_footnote_enhanced", "word_live_insert_text", "word_live_add_table",
    "word_live_apply_list", "word_live_setup_heading_numbering",
    "word_live_modify_table", "word_live_replace_text", "word_live_delete_text",
    "word_live_add_header_footer", "word_live_add_page_numbers",
    "word_live_set_core_properties", "word_live_save", "word_live_undo",
    "word_screen_capture", "word_live_take_snapshot", "word_live_get_diff",
])
def test_destructive_or_mixed_tools_remain_strict(annotations, name):
    assert annotations[name].get("readOnlyHint", False) is False
    assert annotations[name]["destructiveHint"] is True


def test_idempotence_is_only_declared_for_fixed_setting(annotations):
    assert annotations["word_live_resolve_comment"]["idempotentHint"] is True
    assert annotations["word_live_toggle_track_changes"].get("idempotentHint", False) is False
    assert annotations["word_live_insert_paragraphs"].get("idempotentHint", False) is False


def test_create_document_can_overwrite_existing_file(tmp_path):
    path = tmp_path / "existing.docx"
    doc = Document()
    doc.add_paragraph("Existing content")
    doc.save(path)
    result = asyncio.run(document_tools.create_document(str(path)))
    assert "created successfully" in result
    assert "Existing content" not in [p.text for p in Document(path).paragraphs]


def test_cell_formatting_can_replace_text(tmp_path):
    path = tmp_path / "table.docx"
    doc = Document()
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "Original"
    doc.save(path)
    asyncio.run(format_tools.format_table_cell_text(str(path), 0, 0, 0, text_content="Replacement"))
    assert Document(path).tables[0].cell(0, 0).text == "Replacement"


def test_footnote_output_can_overwrite_another_file(tmp_path):
    source = tmp_path / "source.docx"
    destination = tmp_path / "destination.docx"
    doc = Document()
    doc.add_paragraph("Source text")
    doc.save(source)
    doc = Document()
    doc.add_paragraph("Destination text")
    doc.save(destination)
    asyncio.run(footnote_tools.add_footnote_after_text(
        str(source), "Source text", "A footnote", str(destination)
    ))
    assert Document(destination).paragraphs[0].text == "Source text"