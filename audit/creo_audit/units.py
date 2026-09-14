from __future__ import annotations

import re


_UNIT_ALIASES = {
    "mm": "millimeters",
    "millimeter": "millimeters",
    "millimeters": "millimeters",
    "cm": "centimeters",
    "centimeter": "centimeters",
    "centimeters": "centimeters",
    "m": "meters",
    "meter": "meters",
    "meters": "meters",
    "in": "inches",
    "inch": "inches",
    "inches": "inches",
    "ft": "feet",
    "foot": "feet",
    "feet": "feet",
}


def canonical_length_unit(value: str | None, fallback: str = "millimeters") -> str:
    if not value:
        return fallback
    key = re.sub(r"[^a-z]", "", value.lower())
    return _UNIT_ALIASES.get(key, value.lower())


def squared_unit(unit: str) -> str:
    return f"{unit}^2"


def cubed_unit(unit: str) -> str:
    return f"{unit}^3"

