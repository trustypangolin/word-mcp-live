"""Toolset discovery and payload budgets via the MCP protocol."""

import asyncio
import json
import re

import pytest
from fastmcp import Client, FastMCP

from word_document_server.main import register_tools
from word_document_server.toolsets import LEGACY_TOOLS, TOOLSETS, parse_legacy_tools, parse_toolsets, tool_groups


def tools_for(selection=None):
    server = FastMCP("Toolset test")
    register_tools(server, selection)

    async def list_tools():
        async with Client(server) as client:
            return await client.list_tools()

    return asyncio.run(list_tools())


@pytest.fixture(scope="module")
def all_tools():
    return tools_for()


def payload_size(tools):
    payload = [{"name": tool.name, "description": tool.description,
                "inputSchema": tool.inputSchema} for tool in tools]
    return len(json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode())


@pytest.mark.parametrize("value,expected", [
    (None, None), ("", None), (" all ", None),
    (" LIVE, comments,live ", frozenset({"live", "comments"})),
])
def test_parse_toolsets(value, expected):
    assert parse_toolsets(value) == expected


@pytest.mark.parametrize("value", ["live,", ",live", "live,,file", "all,live", "unknown"])
def test_invalid_toolsets_fail(value):
    with pytest.raises(ValueError, match="WORD_MCP_TOOLSETS"):
        parse_toolsets(value)


@pytest.mark.parametrize("group,budget", [
    ("live", 40000), ("file", 39500), ("comments", 4000), ("tables", 14000),
    ("footnotes", 7000), ("tracked_changes", 5000), ("protection", 2000),
    ("layout", 10000),
])
def test_toolsets_have_expected_membership_and_payload_budget(all_tools, group, budget):
    selected = tools_for(frozenset({group}))
    assert {tool.name for tool in selected} == {
        tool.name for tool in all_tools if group in tool_groups(tool.name)
    }
    assert selected
    assert payload_size(selected) <= budget


def test_groups_are_union_and_include_screen_capture(all_tools):
    selected = tools_for(frozenset({"live", "comments"}))
    names = {tool.name for tool in selected}
    assert names == {t.name for t in all_tools if tool_groups(t.name) & {"live", "comments"}}
    assert {"word_screen_capture", "word_live_get_comments", "get_all_comments"} <= names
    assert "add_paragraph" not in names
    assert "tables" not in tool_groups("add_table_of_contents")


def test_descriptions_are_compact_and_parameter_help_is_in_schema(all_tools):
    for tool in all_tools:
        assert not re.search(r"\b(?:Args|Arguments|Parameters|Returns):", tool.description or "")
    by_name = {tool.name: tool for tool in all_tools}
    assert "1-indexed" in by_name["word_live_format_text"].inputSchema["properties"]["start_paragraph"]["description"]
    assert payload_size(all_tools) < 78000


def test_excluded_tool_cannot_be_called():
    server = FastMCP("Restricted test")
    register_tools(server, frozenset({"live"}))

    async def call():
        async with Client(server) as client:
            with pytest.raises(Exception, match="(?i)(unknown|not found)"):
                await client.call_tool("add_paragraph", {"filename": "test.docx", "text": "test"})

    asyncio.run(call())


def test_registered_file_tool_executes_async_backend(tmp_path):
    server = FastMCP("Dispatch test")
    register_tools(server, frozenset({"file"}))
    path = tmp_path / "created.docx"

    async def call():
        async with Client(server) as client:
            result = await client.call_tool("create_document", {"filename": str(path)})
            assert "created successfully" in str(result)

    asyncio.run(call())
    assert path.is_file()


def test_selection_is_read_from_environment(monkeypatch):
    monkeypatch.setenv("WORD_MCP_TOOLSETS", "comments")
    assert {t.name for t in tools_for()} == {t.name for t in tools_for(frozenset({"comments"}))}
    monkeypatch.setenv("WORD_MCP_TOOLSETS", "invalid")
    with pytest.raises(ValueError, match="WORD_MCP_TOOLSETS"):
        tools_for()


def test_legacy_tools_are_opt_in_to_hide(monkeypatch):
    monkeypatch.delenv("WORD_MCP_TOOLSETS", raising=False)
    monkeypatch.delenv("WORD_MCP_LEGACY_TOOLS", raising=False)
    compatible = {t.name for t in tools_for()}
    assert LEGACY_TOOLS <= compatible
    monkeypatch.setenv("WORD_MCP_LEGACY_TOOLS", "false")
    compact = {t.name for t in tools_for()}
    assert compatible - compact == LEGACY_TOOLS
    assert {"add_footnote", "delete_footnote", "merge_table_cells"} <= compact
    with pytest.raises(ValueError, match="WORD_MCP_LEGACY_TOOLS"):
        parse_legacy_tools("disabled")


def test_canonical_footnotes_via_mcp(tmp_path):
    from docx import Document

    path = tmp_path / "test.docx"
    doc = Document()
    doc.add_paragraph("Example text")
    doc.save(path)
    server = FastMCP("Canonical footnotes")
    register_tools(server, frozenset({"footnotes"}), legacy_tools=False)

    async def call():
        async with Client(server) as client:
            tools = {tool.name: tool for tool in await client.list_tools()}
            assert tools["add_footnote"].annotations.destructiveHint is True
            assert tools["delete_footnote"].annotations.destructiveHint is True
            added = await client.call_tool("add_footnote", {
                "filename": str(path), "footnote_text": "Native note", "paragraph_index": 0,
            })
            result = added.data if hasattr(added, "data") else json.loads(added[0].text)
            assert result["success"], result
            deleted = await client.call_tool("delete_footnote", {
                "filename": str(path), "footnote_id": result["details"]["footnote_id"],
            })
            result = deleted.data if hasattr(deleted, "data") else json.loads(deleted[0].text)
            assert result["success"]

    asyncio.run(call())