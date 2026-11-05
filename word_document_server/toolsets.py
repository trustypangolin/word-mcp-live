"""Startup-time tool filtering, independent of FastMCP's version-specific APIs."""

from .tool_metadata import prepare_tool


TOOLSETS = frozenset({
    "live", "file", "comments", "tables", "footnotes", "tracked_changes",
    "protection", "layout",
})

# Legacy names retain their original contracts when compatibility mode is on.
LEGACY_TOOLS = frozenset({
    "add_footnote_to_document", "add_footnote_after_text", "add_footnote_before_text",
    "add_footnote_enhanced", "add_footnote_robust", "delete_footnote_from_document",
    "delete_footnote_robust", "merge_table_cells_horizontal", "merge_table_cells_vertical",
})


def parse_legacy_tools(value):
    """Whether compatibility tool names should be registered."""
    if value is None or value.strip().lower() in {"", "1", "true", "yes"}:
        return True
    if value.strip().lower() in {"0", "false", "no"}:
        return False
    raise ValueError("WORD_MCP_LEGACY_TOOLS must be true or false")


def parse_toolsets(value):
    """Parse a comma-separated union of toolsets; None means all tools."""
    if value is None or not value.strip():
        return None
    parts = [part.strip().lower() for part in value.split(",")]
    if any(not part for part in parts):
        raise ValueError("WORD_MCP_TOOLSETS contains an empty toolset name")
    if parts == ["all"]:
        return None
    unknown = set(parts) - TOOLSETS
    if unknown:
        raise ValueError(
            "Unknown WORD_MCP_TOOLSETS value(s): " + ", ".join(sorted(unknown))
            + ". Choose from all, " + ", ".join(sorted(TOOLSETS))
        )
    return frozenset(parts)


def tool_groups(name):
    """Classify tools into one mode and any applicable capability groups."""
    live = name.startswith("word_live_") or name == "word_screen_capture"
    groups = {"live" if live else "file"}
    if "comment" in name:
        groups.add("comments")
    if name != "add_table_of_contents" and any(
        word in name for word in ("table", "cell", "column", "alternating_rows")
    ):
        groups.add("tables")
    if any(word in name for word in ("footnote", "endnote")):
        groups.add("footnotes")
    if (name.startswith("track_") or "tracked_changes" in name
            or "revision" in name or "track_changes" in name):
        groups.add("tracked_changes")
    if name in {"protect_document", "unprotect_document", "add_restricted_editing",
                "add_digital_signature", "verify_document"}:
        groups.add("protection")
    if any(word in name for word in (
        "layout", "header_footer", "page_numbers", "section_break", "paragraph_spacing",
        "bookmark", "watermark",
    )):
        groups.add("layout")
    return frozenset(groups)


class ToolRegistrar:
    """Register selected tools with compact descriptions and parameter schemas."""

    def __init__(self, server, selected, legacy_tools=True):
        self.server = server
        self.selected = selected
        self.legacy_tools = legacy_tools

    def tool(self, **kwargs):
        def register(function):
            if ((self.legacy_tools or function.__name__ not in LEGACY_TOOLS)
                    and (self.selected is None or tool_groups(function.__name__) & self.selected)):
                prepared, options = prepare_tool(function, **kwargs)
                self.server.tool(**options)(prepared)
            return function
        return register