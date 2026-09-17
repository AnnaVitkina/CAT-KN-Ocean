"""
City spelling rename + common-rating helpers.

1) Spelling rename (NOT common rating):
   Replace the written city with the full target string as-is.
   e.g. s-Hertogenbosch -> "s-Hertogenbosch, s-Hertogenbosch"
        Keelung (Chilung) -> "Keelung, Chilung"
        Saint-Étienne -> "Saint-Étienne; Saint-Etienne"

2) Common rating (separate):
   Primary (alias1, alias2), omitting aliases that already exist as their own lane.
"""

from __future__ import annotations

# Written form -> full replacement string (used verbatim in output).
CITY_RENAME_RAW: dict[str, str] = {
    "Nhava Sheva (JNPT)": "Nhava Sheva",
    "Keelung (Chilung)": "Keelung, Chilung",
    "Andrézieux-Boutheon": "Andrézieux-Bouthéon, Andrezieux-Boutheon, Andrzieux-Bouthon",
    "AndrÃ©zieux-Boutheon": "Andrézieux-Bouthéon, Andrezieux-Boutheon, Andrzieux-Bouthon",
    "Námestovo": "Námestovo, Namestovo, Nmestovo",
    "NÃ¡mestovo": "Námestovo, Namestovo, Nmestovo",
    "Tielt (West Flanders)": "Tielt",
    "Pointe des Galets (deprecated)": "Pointe des Galets",
    "Tamatave (Toamasina)": "Tamatave, Toamasina",
    "Waipahu (Oahu)": "Waipahu",
    "Pinghu (Zhejiang)": "Pinghu, Zhejiang",
    "s-Hertogenbosch": "s-Hertogenbosch, s-Hertogenbosch",
    "'s-Hertogenbosch": "s-Hertogenbosch, s-Hertogenbosch",
    "Hertogenbosch": "s-Hertogenbosch, s-Hertogenbosch",
    "Qingzhou (Shandong)": "Qingzhou, Shandong",
    "Machelen (Zulte)": "Machelen",
    "Chilgok-gun, 47": "Chilgok-gun",
    "Gimpo, 41": "Gimpo",
    "Goryeong-gun, 47": "Goryeong-gun",
    "Gyeongsan, 47": "Gyeongsan",
    "Jangheung-gun, 46": "Jangheung-gun",
    "Yeongcheon, 47": "Yeongcheon",
    "Goyang, 41": "Goyang",
    "Haman-gun, 48": "Haman-gun",
    "Saint-Étienne": "Saint-Étienne; Saint-Etienne",
    "Saint-Ã‰tienne": "Saint-Étienne; Saint-Etienne",
    "Varces-Allières-et-Risset": (
        "Varces-Allières-et-Risset; Varces-Allieres-et-Risset; Varces-Allires-et-Risset"
    ),
    "Varces-AlliÃ¨res-et-Risset": (
        "Varces-Allières-et-Risset; Varces-Allieres-et-Risset; Varces-Allires-et-Risset"
    ),
    "Newark (in US)": "New York",
    "Newark (in US) ": "New York",
}

# Primary city -> aliases that are common-rated to it (accepted as the same value).
COMMON_RATING: dict[str, list[str]] = {
    "London": ["London Gateway Port", "Southampton"],
    "Durban": ["Johannesburg"],
    "Tianjin": ["Tianjin Xingang", "Tianjinxingang"],
    "Tianjin Xingang": ["Tianjin", "Tianjinxingang"],
    "Jebel Ali": ["Dubai"],
    "New Delhi": ["Delhi"],
    "Kempton Park": ["Witfontein"],
}


def _first_spelling_token(text: str) -> str:
    """First name in a spelling list (comma or semicolon separated)."""
    text = (text or "").strip()
    if not text:
        return ""
    if ";" in text:
        return text.split(";", 1)[0].strip()
    if "," in text:
        return text.split(",", 1)[0].strip()
    return text


def _spelling_tokens(text: str) -> list[str]:
    text = (text or "").strip()
    if not text:
        return []
    sep = ";" if ";" in text else ","
    if sep not in text:
        return [text]
    return [p.strip() for p in text.split(sep) if p.strip()]


def _build_rename_index() -> dict[str, str]:
    """source.casefold -> full target spelling string."""
    by_source: dict[str, str] = {}
    for source, target in CITY_RENAME_RAW.items():
        full = (target or "").strip()
        if not full:
            continue
        by_source[source.strip().casefold()] = full
    return by_source


_RENAME_BY_SOURCE = _build_rename_index()

_COMMON_INDEX: dict[str, list[str]] = {
    k.casefold(): list(v) for k, v in COMMON_RATING.items()
}

# Full target strings (for identity checks after a prior rename pass)
_FULL_TARGETS = {v.casefold(): v for v in CITY_RENAME_RAW.values() if v.strip()}


def lookup_city_spelling(city: str) -> str | None:
    """Return full spelling replacement string, or None if no rename."""
    if city is None:
        return None
    text = str(city).strip()
    if not text:
        return None

    hit = _RENAME_BY_SOURCE.get(text.casefold())
    if hit is not None:
        return hit

    # Already a full multi-value spelling string from a prior pass
    if text.casefold() in _FULL_TARGETS:
        return _FULL_TARGETS[text.casefold()]

    # e.g. "Pinghu (Zhejiang), ZJ" / "Chilgok-gun, 47"
    if "," in text:
        left = text.split(",", 1)[0].strip()
        hit = _RENAME_BY_SOURCE.get(left.casefold())
        if hit is not None:
            return hit

    return None


def apply_city_rename(city: str) -> tuple[str, list[str]]:
    """
    Apply spelling rename. Returns (display_value, []).
    display_value is the full target string (commas/semicolons kept).
    Rename aliases are NOT returned — spelling is not common-rating.
    """
    if city is None:
        return "", []
    text = str(city).strip()
    if not text:
        return "", []
    spelling = lookup_city_spelling(text)
    if spelling is not None:
        return spelling, []
    return text, []


def city_match_key(city: str) -> str:
    """Normalized key: first spelling token after rename (for lane matching)."""
    display, _ = apply_city_rename(city)
    return _first_spelling_token(display).casefold()


def common_rating_aliases(primary: str) -> list[str]:
    key = _first_spelling_token(primary or "").casefold()
    return list(_COMMON_INDEX.get(key, []))


def format_city_with_aliases(
    city: str,
    *,
    side: str,
    from_key: str,
    from_service: str,
    to_key: str,
    to_service: str,
    lane_index: list[tuple[str, str, str, str]],
    rename_aliases: list[str] | None = None,
) -> str:
    """
    Spelling rename is already applied on `city` (full string, possibly with commas).
    Only common-rating aliases are added in parentheses:
      London (London Gateway Port, Southampton)
    Spelling variants stay as: s-Hertogenbosch, s-Hertogenbosch
    """
    del rename_aliases  # spelling is inline; not paren aliases
    display, _ = apply_city_rename(city)
    if not display:
        return ""

    primary_key = _first_spelling_token(display)
    aliases = common_rating_aliases(primary_key)

    seen = {t.casefold() for t in _spelling_tokens(display)}
    filtered: list[str] = []
    for alias in aliases:
        key = alias.casefold()
        if not alias or key in seen:
            continue
        seen.add(key)
        filtered.append(alias)

    if not filtered:
        return display

    kept: list[str] = []
    for alias in filtered:
        if _alias_has_own_lane(
            alias,
            side=side,
            from_key=from_key,
            from_service=from_service,
            to_key=to_key,
            to_service=to_service,
            lane_index=lane_index,
        ):
            continue
        kept.append(alias)

    if not kept:
        return display
    return f"{display} ({', '.join(kept)})"


def _alias_has_own_lane(
    alias: str,
    *,
    side: str,
    from_key: str,
    from_service: str,
    to_key: str,
    to_service: str,
    lane_index: list[tuple[str, str, str, str]],
) -> bool:
    alias_key = city_match_key(alias)
    fs = (from_service or "").strip().casefold()
    ts = (to_service or "").strip().casefold()
    for lf, lfs, lt, lts in lane_index:
        if lfs != fs or lts != ts:
            continue
        if side == "from":
            if lf == alias_key and lt == to_key:
                return True
        else:
            if lt == alias_key and lf == from_key:
                return True
    return False


def build_lane_city_index(
    rows: list[tuple[str, str, str, str]],
) -> list[tuple[str, str, str, str]]:
    """rows: (from_city, from_svc, to_city, to_svc) -> match keys."""
    index: list[tuple[str, str, str, str]] = []
    for from_city, from_svc, to_city, to_svc in rows:
        index.append(
            (
                city_match_key(from_city),
                (from_svc or "").strip().casefold(),
                city_match_key(to_city),
                (to_svc or "").strip().casefold(),
            )
        )
    return index
