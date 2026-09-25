from __future__ import annotations

import math
from pathlib import Path
import unicodedata

import pandas as pd


DEFAULT_COLUMN_ALIASES: dict[str, list[str]] = {
    "前台车型": ["车型名", "车姓名"],
    "最终尺码": ["对应尺码"],
}

DEFAULT_DERIVED_COLUMNS: dict[str, dict[str, object]] = {
    "主车型": {
        "join": ["品牌", "前台车型"],
        "sep": " ",
    },
}

DEFAULT_COLUMN_VALUES: dict[str, str] = {
    "分类": "",
}


def normalize_text(value: object) -> str:
    if value is None or value is pd.NA or value is pd.NaT:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""

    text = str(value)
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C")
    return text.strip()


def load_field_profile(path: Path | None) -> dict[str, object]:
    if path is None:
        return {}

    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError("使用 --field-profile 需要先安装 PyYAML：pip install -r requirements.txt") from exc

    if not path.exists():
        raise FileNotFoundError(f"Field profile not found: {path}")

    with path.open("r", encoding="utf-8-sig") as profile_file:
        profile = yaml.safe_load(profile_file) or {}
    if not isinstance(profile, dict):
        raise ValueError("Field profile must be a YAML mapping.")
    return profile


def _as_list(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [normalize_text(item) for item in value if normalize_text(item)]
    text = normalize_text(value)
    return [text] if text else []


def _profile_column_aliases(profile: dict[str, object]) -> dict[str, list[str]]:
    configured = profile.get("columns", profile.get("aliases", {}))
    if configured is None:
        configured = {}
    if not isinstance(configured, dict):
        raise ValueError("Field profile 'columns' must be a mapping.")

    aliases: dict[str, list[str]] = {
        normalize_text(canonical): _as_list(values)
        for canonical, values in DEFAULT_COLUMN_ALIASES.items()
    }
    for canonical_raw, values in configured.items():
        canonical = normalize_text(canonical_raw)
        if not canonical:
            continue
        aliases.setdefault(canonical, [])
        for alias in _as_list(values):
            if alias not in aliases[canonical]:
                aliases[canonical].append(alias)
    return aliases


def _profile_derived_columns(profile: dict[str, object]) -> dict[str, dict[str, object]]:
    configured = profile.get("derived", {})
    if configured is None:
        configured = {}
    if not isinstance(configured, dict):
        raise ValueError("Field profile 'derived' must be a mapping.")

    derived = {key: value.copy() for key, value in DEFAULT_DERIVED_COLUMNS.items()}
    for column, rule in configured.items():
        column_name = normalize_text(column)
        if not column_name:
            continue
        if not isinstance(rule, dict):
            raise ValueError(f"Field profile derived rule for {column_name} must be a mapping.")
        derived[column_name] = rule.copy()
    return derived


def _profile_default_values(profile: dict[str, object]) -> dict[str, str]:
    configured = profile.get("defaults", {})
    if configured is None:
        configured = {}
    if not isinstance(configured, dict):
        raise ValueError("Field profile 'defaults' must be a mapping.")

    defaults = DEFAULT_COLUMN_VALUES.copy()
    for column, value in configured.items():
        column_name = normalize_text(column)
        if column_name:
            defaults[column_name] = normalize_text(value)
    return defaults


def apply_field_profile(df: pd.DataFrame, profile: dict[str, object] | None = None) -> pd.DataFrame:
    work = df.copy()
    work.columns = [normalize_text(column) for column in work.columns]
    profile = profile or {}

    for canonical, aliases in _profile_column_aliases(profile).items():
        for alias in aliases:
            if alias not in work.columns:
                continue
            if canonical not in work.columns:
                work[canonical] = work[alias]
                break
            canonical_empty = work[canonical].map(normalize_text) == ""
            if canonical_empty.any():
                work.loc[canonical_empty, canonical] = work.loc[canonical_empty, alias]
                break

    for column, rule in _profile_derived_columns(profile).items():
        if column in work.columns:
            continue
        join_columns = _as_list(rule.get("join"))
        if not join_columns or not all(join_column in work.columns for join_column in join_columns):
            continue
        sep = normalize_text(rule.get("sep", " "))
        if work.empty:
            work[column] = work.apply(
                lambda row: sep.join(
                    part
                    for part in [normalize_text(row.get(join_column, "")) for join_column in join_columns]
                    if part
                ),
                axis=1,
            )
            continue
        parts = [work[join_column].map(normalize_text) for join_column in join_columns]
        work[column] = [sep.join(part for part in values if part) for values in zip(*parts)]

    for column, value in _profile_default_values(profile).items():
        if column not in work.columns:
            work[column] = value

    return work
