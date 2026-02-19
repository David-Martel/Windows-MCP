"""Shared helper functions for MCP tool handlers."""

MAX_IMAGE_WIDTH, MAX_IMAGE_HEIGHT = 1920, 1080


def _coerce_bool(value: bool | str, default: bool = False) -> bool:
    """Convert a bool-or-string MCP parameter to a proper bool.

    MCP clients may send boolean parameters as strings ("true"/"false").
    This normalises both forms to a Python bool.
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() == "true"
    return default


def _validate_loc(loc: list, *, label: str = "loc") -> tuple[int, int]:
    """Validate and extract (x, y) integer coordinates from a list.

    MCP clients may send floats (from JSON) or strings; this coerces
    to int and validates the list length.
    """
    if not isinstance(loc, (list, tuple)) or len(loc) != 2:
        raise ValueError(f"{label} must be [x, y], got {loc!r}")
    try:
        return int(loc[0]), int(loc[1])
    except (TypeError, ValueError) as e:
        raise ValueError(f"{label} values must be numeric, got {loc!r}") from e
