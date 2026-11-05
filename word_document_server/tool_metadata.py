"""Prepare compact FastMCP tool metadata without changing the callable.

Use ``prepared, kwargs = prepare_tool(function, description=..., annotations=...)``
then ``mcp.tool(**kwargs)(prepared)``. An explicit description takes precedence
over the function's docstring, just as it does in FastMCP.

This is deliberately a conservative Google-style docstring parser, not a general
documentation parser. Existing Pydantic fields are left alone. Unresolvable type
hints fail early rather than producing a subtly different tool schema.
"""

from __future__ import annotations

import functools
import inspect
import re
from typing import Annotated, Any, Callable, get_args, get_origin, get_type_hints

from pydantic import Field
from pydantic.fields import FieldInfo


_SECTION = re.compile(r"^\s*(?:Args|Arguments|Parameters|Returns|Yields|Raises|Examples?|Notes?|Warnings?|Cautions?):\s*$", re.IGNORECASE)
_ARG = re.compile(r"^(\s*)(\w+)(?:\s*\([^\n]*?\))?:\s*(.*)$")
_INLINE_ARG = re.compile(r"(?:^|,\s*)(\w+)\s*\(([^()]*)\)")
_IMPORTANT = re.compile(
    r"\b(?:warning|caution|important|note|must|requires?|only|cannot|can't|"
    r"does not|doesn't|not supported|unsupported|instead|do not|never|"
    r"overwrit\w*|destruct\w*|unsaved|running|addressing modes?)\b",
    re.IGNORECASE,
)
_GENERIC = re.compile(
    r"^(?:optional|required|str|string|int|integer|float|bool|boolean|list|dict|"
    r"path|file(?:name)?|text|value|default|none|true|false|of|to|a|the|"
    r"word|document|parameter|argument|and|or|[\s,.:/\-])+$",
    re.IGNORECASE,
)


def _parse_doc(doc: str) -> tuple[str, dict[str, str]]:
    """Separate prose from Google-style Args, including wrapped descriptions."""
    lines = inspect.cleandoc(doc).splitlines()
    prose: list[str] = []
    args: dict[str, str] = {}
    in_args = False
    current: str | None = None
    arg_indent: int | None = None
    skip_section = False

    for line in lines:
        # A few wrapper docstrings put a type-only Args list on the same line.
        if not in_args and re.search(r"\bArgs:\s*\S", line):
            before, inline = re.split(r"\bArgs:\s*", line, maxsplit=1)
            prose.append(before)
            for match in _INLINE_ARG.finditer(inline):
                args[match.group(1)] = match.group(2)
            continue
        if _SECTION.match(line):
            heading = line.strip().lower().rstrip(":")
            in_args = heading in {"args", "arguments", "parameters"}
            skip_section = heading in {"returns", "yields", "raises", "example", "examples"}
            if heading in {"note", "notes", "warning", "warnings", "caution", "cautions"}:
                prose.append(heading.capitalize() + ":")
            current = None
            arg_indent = None
            continue
        if in_args:
            if line.strip() and arg_indent is not None and len(line) - len(line.lstrip()) < arg_indent:
                in_args = False
                prose.append(line)
                continue
            match = _ARG.match(line)
            if match and (arg_indent is None or len(match.group(1)) == arg_indent):
                arg_indent = len(match.group(1))
                current = match.group(2)
                args[current] = match.group(3).strip()
            elif current and line.strip():
                args[current] += " " + line.strip()
        elif not skip_section:
            prose.append(line)

    return "\n".join(prose), args


def _description(prose: str) -> str:
    # Keep the first two sentences, then add operational caveats even when they
    # occur later. Newlines are significant for addressing-mode bullet lists.
    lines = [line.strip() for line in prose.splitlines() if line.strip()]
    text = " ".join(lines)
    text = re.sub(r"\ba Word document (?:that is )?open in Word\b", "an open Word document", text, flags=re.IGNORECASE)
    # Avoid splitting common abbreviations or decimal numbers.
    sentences = re.split(r"(?<!\be\.g)(?<!\bi\.e)(?<!\betc)(?<=[.!?])\s+(?=[A-Z\[])|(?<=\.)\s+(?=\[Windows)", text)
    chosen = [sentence.strip() for sentence in sentences[:2] if sentence.strip()]
    for sentence in sentences[2:]:
        if _IMPORTANT.search(sentence) and sentence.strip() not in chosen:
            chosen.append(sentence.strip())
    # A labelled addressing-mode block often has details in bullet points that
    # do not themselves contain a caveat keyword.
    for index, line in enumerate(lines):
        if re.search(r"addressing modes?", line, re.IGNORECASE):
            for bullet in lines[index + 1:]:
                if not bullet.startswith(("- ", "* ")):
                    break
                if not any(bullet in part for part in chosen):
                    chosen.append(bullet)
    return " ".join(chosen)


def _useful_arg(name: str, description: str) -> bool:
    text = description.strip().strip(".")
    if not text or _GENERIC.fullmatch(text):
        return False
    # Avoid adding descriptions which only repeat the parameter's name.
    words = re.sub(r"[^a-z0-9]", "", text.lower())
    return words != re.sub(r"[^a-z0-9]", "", name.lower())


def _has_field(annotation: Any, default: Any) -> bool:
    if isinstance(default, FieldInfo):
        return True
    return get_origin(annotation) is Annotated and any(
        isinstance(item, FieldInfo) for item in get_args(annotation)[1:]
    )


def prepare_tool(function: Callable[..., Any], **registration_kwargs: Any) -> tuple[Callable[..., Any], dict[str, Any]]:
    """Return a synchronous wrapper and registration kwargs with compact metadata.

    Parameter names, kinds, defaults, return values and underlying type hints
    are unchanged. Existing Field metadata is never replaced. Only descriptions
    for named, annotated parameters are transferred. Coroutine functions are
    unsupported; synchronous functions returning awaitables are passed through.
    """
    if inspect.iscoroutinefunction(function) or inspect.isgeneratorfunction(function):
        raise TypeError("prepare_tool requires a synchronous, non-generator function")

    signature = inspect.signature(function)
    try:
        hints = get_type_hints(function, include_extras=True)
    except (NameError, TypeError) as exc:
        raise TypeError("prepare_tool could not resolve the function's type hints") from exc

    doc = registration_kwargs.get("description")
    if doc is None:
        doc = function.__doc__ or ""
    if not isinstance(doc, str):
        raise TypeError("description must be a string or None")
    prose, arg_docs = _parse_doc(doc)
    description = _description(prose)

    parameters = []
    for parameter in signature.parameters.values():
        hint = hints.get(parameter.name, parameter.annotation)
        arg_doc = arg_docs.get(parameter.name, "")
        if (hint is not inspect.Parameter.empty and _useful_arg(parameter.name, arg_doc)
                and not _has_field(hint, parameter.default)):
            hint = Annotated[hint, Field(description=arg_doc)]
            hints[parameter.name] = hint
            parameter = parameter.replace(annotation=hint)
        parameters.append(parameter)

    @functools.wraps(function)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        return function(*args, **kwargs)

    wrapper.__doc__ = description
    wrapper.__annotations__ = hints
    wrapper.__signature__ = signature.replace(parameters=parameters)  # type: ignore[attr-defined]
    kwargs = dict(registration_kwargs)
    kwargs["description"] = description
    return wrapper, kwargs