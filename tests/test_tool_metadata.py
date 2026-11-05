"""Tests for opt-in, non-mutating tool metadata preparation."""

import asyncio
import inspect
from typing import Annotated, get_args, get_type_hints

import pytest
from fastmcp import Client, FastMCP
from pydantic import Field
from pydantic.fields import FieldInfo

from word_document_server.tool_metadata import prepare_tool


def _tool(function, **kwargs):
    server = FastMCP("Metadata test")
    prepared, options = prepare_tool(function, **kwargs)
    server.tool(**options)(prepared)

    async def get_tool():
        async with Client(server) as client:
            return (await client.list_tools())[0]

    return prepared, options, asyncio.run(get_tool())


def test_wrapper_docstring_is_compacted_and_parameters_are_exposed():
    def example(filename: str, count: int = 2, *, color: str = "red") -> str:
        """Process a document. Select a color. This long explanation is redundant.

        WARNING: This can overwrite the document. Requires Word running.

        Args:
            filename: Path to the file (must not be open in Word).
            count (int): Number of passes.
                Values above one repeat the operation.
            color: Color as a hex value or a named color.
            missing: Not in the wrapper signature.

        Returns:
            A status message.
        """
        return f"{filename}:{count}:{color}"

    original_hints = dict(example.__annotations__)
    prepared, options, tool = _tool(example, name="compact")
    assert prepared.__name__ == example.__name__
    assert prepared("x", color="blue") == example("x", color="blue")
    assert inspect.signature(prepared).parameters["count"].default == 2
    assert inspect.signature(prepared).parameters["color"].kind == inspect.Parameter.KEYWORD_ONLY
    assert example.__annotations__ == original_hints
    assert "Args:" not in tool.description
    assert "redundant" not in tool.description
    assert "overwrite" in tool.description and "Word running" in tool.description
    assert options["description"] == prepared.__doc__ == tool.description
    props = tool.inputSchema["properties"]
    assert props["filename"]["description"] == "Path to the file (must not be open in Word)."
    assert props["count"]["description"] == "Number of passes. Values above one repeat the operation."
    assert props["count"]["default"] == 2
    assert "missing" not in props
    assert tool.inputSchema["required"] == ["filename"]


def test_explicit_implementation_description_and_unmatched_parameters():
    def wrapper(filename: str, flag: bool = False):
        """The wrapper's docstring should not be chosen."""
        return filename, flag

    description = """[Windows only] Edit an open document. Use the active document if needed.
    Other context is not necessary.

    Args:
        filename: Document name or path (None = active document).
        flag: bool
        extra: Only in the implementation.
    """
    prepared, kwargs, tool = _tool(wrapper, description=description)
    assert "wrapper's" not in kwargs["description"]
    assert "[Windows only]" in tool.description
    assert "Other context" not in tool.description
    assert tool.inputSchema["properties"]["filename"]["description"].startswith("Document name")
    assert "description" not in tool.inputSchema["properties"]["flag"]
    assert prepared("test") == ("test", False)


def test_existing_annotated_and_field_metadata_are_preserved():
    marker = object()

    def example(
        existing: Annotated[int, Field(description="Original", ge=1)],
        metadata: Annotated[str, marker],
        default_field: int = Field(default=3, description="Existing default", ge=0),
    ):
        """Use existing metadata.

        Args:
            existing: Must not replace original.
            metadata: A useful new description.
            default_field: Must not replace the default field.
        """
        return existing, metadata, default_field

    prepared, _, tool = _tool(example)
    hints = get_type_hints(prepared, include_extras=True)
    assert get_args(hints["metadata"])[1] is marker
    assert any(isinstance(item, FieldInfo) for item in get_args(hints["metadata"])[1:])
    props = tool.inputSchema["properties"]
    assert props["existing"]["description"] == "Original"
    assert props["existing"]["minimum"] == 1
    assert props["metadata"]["description"] == "A useful new description."
    assert props["default_field"]["description"] == "Existing default"
    assert props["default_field"]["minimum"] == 0
    assert prepared(2, "a", 3) == example(2, "a", 3)


def test_inline_type_only_args_are_removed_without_invented_descriptions():
    def example(filename: str, position: str = "after"):
        """Insert a line. Specify by text. Args: filename (str), position (str, optional)."""
        return filename, position

    _, _, tool = _tool(example)
    assert "Args:" not in tool.description
    assert "description" not in tool.inputSchema["properties"]["filename"]
    assert "description" not in tool.inputSchema["properties"]["position"]


def test_late_notes_and_windows_boilerplate():
    def example(filename: str):
        """[Windows only] Edit a Word document that is open in Word. Make changes.

        Args:
            filename: Document filename.

        Notes:
            Must save explicitly. Unsaved changes may be lost.
        """
        return filename

    _, _, tool = _tool(example)
    assert "[Windows only]" in tool.description
    assert "an open Word document" in tool.description
    assert "Must save explicitly" in tool.description
    assert "Unsaved changes may be lost" in tool.description
    assert "filename:" not in tool.description


def test_real_implementation_docs_preserve_addressing_and_limits():
    from word_document_server.tools import live_tools

    def wrapper(filename: str = None, start: int = None, end: int = None,
                start_paragraph: int = None, end_paragraph: int = None):
        return filename, start, end, start_paragraph, end_paragraph

    _, _, tool = _tool(wrapper, description=live_tools.word_live_format_text.__doc__)
    assert "addressing modes" in tool.description
    assert "start/end" in tool.description
    assert "start_paragraph/end_paragraph" in tool.description
    assert "Args:" not in tool.description
    assert tool.inputSchema["properties"]["start_paragraph"]["description"].startswith("First paragraph")


def test_unannotated_arguments_are_left_unchanged_and_empty_doc_is_supported():
    def example(value, other: int = 1):
        return value + other

    prepared, kwargs = prepare_tool(example)
    assert kwargs["description"] == ""
    assert inspect.signature(prepared).parameters["value"].annotation is inspect.Parameter.empty
    assert prepared(3) == 4


def test_future_annotations_and_non_field_metadata_are_resolved():
    namespace = {}
    exec("from __future__ import annotations\n"
         "def example(value: int = 1) -> int:\n"
         "    '''Do work.\n\n    Args:\n        value: The number of passes.\n    '''\n"
         "    return value\n", namespace)
    prepared, _, tool = _tool(namespace["example"])
    assert get_args(get_type_hints(prepared, include_extras=True)["value"])[0] is int
    assert tool.inputSchema["properties"]["value"]["type"] == "integer"
    assert tool.inputSchema["properties"]["value"]["description"] == "The number of passes."


def test_sync_function_returning_awaitable_is_not_awaited_by_wrapper():
    async def result():
        return 42

    def example():
        return result()

    prepared, _ = prepare_tool(example)
    assert asyncio.run(prepared()) == 42


def test_unsupported_functions_and_unresolvable_annotations_fail_explicitly():
    async def asynchronous():
        pass

    def generator():
        yield 1

    def unresolved(value: "MissingType"):
        return value

    for function in (asynchronous, generator, unresolved):
        with pytest.raises(TypeError):
            prepare_tool(function)