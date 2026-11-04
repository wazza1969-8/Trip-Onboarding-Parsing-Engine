"""
Destination field manifests
===========================

A *manifest* describes one destination for parsed data -- what fields it
has, in what order, under what labels, and which parser field feeds each
one. It is deliberately kept OUT of parsing_engine.py: the engine's job
is to get facts out of a source document, and a manifest's job is to say
where those facts go. Adding a second destination (a portal, a different
workbook) means adding a JSON file here, not touching the engine.

Manifests live in `manifests/*.json` next to this module and are loaded
by `target_id`. Two ship today:

  operational_information_workbook -- VERIFIED. Sections and order come
      from parsing_engine.SUBJECTS ("must match the Operational
      Information grid") and are confirmed against the workbook in this
      repository; field-level source labels come verbatim from
      parsing_engine.FIELD_WHITELIST.

  destination_portal -- PLACEHOLDER, deliberately empty. No portal
      specification has been obtained. `is_usable()` returns False for
      it, and build_entry_ready_records() refuses to run against it, so
      an unpopulated portal manifest fails closed rather than silently
      producing an empty or invented form layout.

Verification vocabulary
-----------------------
Every field carries `verification`:

  "verified"    -- the label, source mapping and any permitted values
                   are traceable to something authoritative (the
                   engine's own whitelist, or a real specification).
  "provisional" -- inferred, usually from observed sample data. Rendered
                   with a warning and excluded from automatic mapping
                   accuracy claims. `permitted_values_verified: false`
                   marks enum lists that were inferred rather than
                   specified.

Nothing in a manifest may be invented. A field whose destination label
or permitted values are unknown belongs in `present_in_source_but_not_
extracted` or in the portal manifest's `required_before_use` list -- not
in `sections` with a plausible-looking guess.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field as dc_field
from pathlib import Path
from typing import Any

MANIFEST_DIR = Path(__file__).resolve().parent / "manifests"

# Criticality buckets. Accuracy is reported separately per bucket --
# a 95% average is meaningless if the 5% that is wrong is an ICAO code
# or a customs declaration, so "critical" is never averaged in with the
# hotel-preference fields.
CRITICALITY_LEVELS = ("critical", "required_noncritical", "optional")

VERIFICATION_LEVELS = ("verified", "provisional")


class ManifestError(ValueError):
    """Raised when a manifest file is structurally invalid. Loading fails
    closed: a malformed manifest is never partially applied."""


@dataclass(frozen=True)
class FieldSpec:
    field_id: str
    target_label: str
    order: int
    section_id: str
    section_label: str
    source_section: str | None
    source_label: str | None
    data_type: str
    required: bool
    criticality: str
    normalizer: str
    validators: tuple[str, ...]
    availability: str
    verification: str
    permitted_values: tuple[str, ...] | None = None
    permitted_values_verified: bool = True
    permitted_values_source: str | None = None
    conditional_on: dict[str, Any] | None = None
    repeating: bool = False
    notes: str | None = None
    # Present only on manifests for an interactive target (a web form).
    # {"type": "text|select|checkbox|textarea", "locator": {...},
    #  "option_map": {normalized value -> exact option label}}
    # A field with no control block can be displayed and validated but
    # cannot be filled automatically -- autofill skips it rather than
    # inventing a way to reach it.
    control: dict[str, Any] | None = None

    @property
    def is_provisional(self) -> bool:
        return self.verification != "verified" or not self.permitted_values_verified

    @property
    def is_fillable(self) -> bool:
        return bool(self.control and self.control.get("locator"))


@dataclass(frozen=True)
class SectionSpec:
    section_id: str
    label: str
    order: int
    fields: tuple[FieldSpec, ...]
    static_only: bool = False
    field_order_is_fixed: bool = False


@dataclass(frozen=True)
class Manifest:
    manifest_version: str
    target_id: str
    target_name: str
    target_kind: str
    target_verified: bool
    target_verification_note: str
    source_form: str | None
    sections: tuple[SectionSpec, ...]
    not_covered_by_source_form: tuple[str, ...] = ()
    present_in_source_but_not_extracted: tuple[dict[str, str], ...] = ()
    required_before_use: tuple[str, ...] = ()
    # Structural snapshot of the target form taken when the manifest was
    # approved: [{"label", "type", "options": [...]}, ...]. Compared against
    # the live form before any field is touched -- see autofill.detect_drift().
    # Empty for non-interactive targets such as the workbook.
    form_fingerprint: tuple[dict[str, Any], ...] = ()
    form_url: str | None = None

    def is_usable(self) -> bool:
        """A manifest is usable for mapping only if it actually describes
        fields. An empty manifest (the portal placeholder) is not an
        empty form -- it is an unknown form, and callers must treat it
        as such rather than rendering nothing and implying completeness."""
        return bool(self.all_fields())

    def all_fields(self) -> list[FieldSpec]:
        return [f for s in self.sections for f in s.fields]

    def field(self, field_id: str) -> FieldSpec | None:
        return next((f for f in self.all_fields() if f.field_id == field_id), None)

    def provisional_fields(self) -> list[FieldSpec]:
        return [f for f in self.all_fields() if f.is_provisional]


def _require(cond: bool, message: str) -> None:
    if not cond:
        raise ManifestError(message)


def _parse_field(raw: dict, section_id: str, section_label: str) -> FieldSpec:
    for key in ("field_id", "target_label", "order", "data_type", "criticality",
                "normalizer", "availability", "verification"):
        _require(key in raw, f"field in section '{section_id}' is missing '{key}'")

    _require(
        raw["criticality"] in CRITICALITY_LEVELS,
        f"field '{raw['field_id']}' has criticality '{raw['criticality']}'; "
        f"expected one of {CRITICALITY_LEVELS}",
    )
    _require(
        raw["verification"] in VERIFICATION_LEVELS,
        f"field '{raw['field_id']}' has verification '{raw['verification']}'; "
        f"expected one of {VERIFICATION_LEVELS}",
    )

    source = raw.get("source") or {}
    permitted = raw.get("permitted_values")

    # An enum whose permitted values were inferred rather than specified
    # must say so. Silently treating an inferred list as authoritative is
    # exactly how a wrong dropdown value reaches a customer record.
    permitted_verified = raw.get("permitted_values_verified", True)
    if permitted is not None and not isinstance(permitted, list):
        raise ManifestError(f"field '{raw['field_id']}' has non-list permitted_values")

    return FieldSpec(
        field_id=raw["field_id"],
        target_label=raw["target_label"],
        order=int(raw["order"]),
        section_id=section_id,
        section_label=section_label,
        source_section=source.get("section"),
        source_label=source.get("label"),
        data_type=raw["data_type"],
        required=bool(raw.get("required", False)),
        criticality=raw["criticality"],
        normalizer=raw["normalizer"],
        validators=tuple(raw.get("validators", ())),
        availability=raw["availability"],
        verification=raw["verification"],
        permitted_values=tuple(permitted) if permitted else None,
        permitted_values_verified=bool(permitted_verified),
        permitted_values_source=raw.get("permitted_values_source"),
        conditional_on=raw.get("conditional_on"),
        repeating=bool(raw.get("repeating", False)),
        notes=raw.get("notes"),
        control=raw.get("control"),
    )


def _parse_manifest(raw: dict) -> Manifest:
    for key in ("manifest_version", "target_id", "target_name", "target_verified"):
        _require(key in raw, f"manifest is missing '{key}'")

    sections: list[SectionSpec] = []
    seen_field_ids: set[str] = set()
    for s in raw.get("sections", []):
        for key in ("section_id", "label", "order"):
            _require(key in s, f"section is missing '{key}'")
        fields = [_parse_field(f, s["section_id"], s["label"]) for f in s.get("fields", [])]
        for f in fields:
            _require(
                f.field_id not in seen_field_ids,
                f"duplicate field_id '{f.field_id}' -- field ids must be unique "
                f"across the whole manifest, since they key the mapping report",
            )
            seen_field_ids.add(f.field_id)
        # Sort by declared order so display order is a property of the
        # data, not of JSON key order.
        fields.sort(key=lambda f: f.order)
        sections.append(SectionSpec(
            section_id=s["section_id"],
            label=s["label"],
            order=int(s["order"]),
            fields=tuple(fields),
            static_only=bool(s.get("static_only", False)),
            field_order_is_fixed=bool(s.get("field_order_is_fixed", False)),
        ))
    sections.sort(key=lambda s: s.order)

    # Conditional dependencies must point at fields that exist, or the
    # condition can never be evaluated and would silently pass.
    for sec in sections:
        for f in sec.fields:
            if f.conditional_on:
                dep = f.conditional_on.get("field_id")
                _require(
                    dep in seen_field_ids,
                    f"field '{f.field_id}' is conditional on unknown field '{dep}'",
                )

    return Manifest(
        manifest_version=raw["manifest_version"],
        target_id=raw["target_id"],
        target_name=raw["target_name"],
        target_kind=raw.get("target_kind", "unknown"),
        target_verified=bool(raw["target_verified"]),
        target_verification_note=raw.get("target_verification_note", ""),
        source_form=raw.get("source_form"),
        sections=tuple(sections),
        not_covered_by_source_form=tuple(raw.get("not_covered_by_source_form", ())),
        present_in_source_but_not_extracted=tuple(raw.get("present_in_source_but_not_extracted", ())),
        required_before_use=tuple(raw.get("required_before_use", ())),
        form_fingerprint=tuple(raw.get("form_fingerprint", ())),
        form_url=raw.get("form_url"),
    )


_cache: dict[str, Manifest] = {}


def available_manifests(manifest_dir: Path | None = None) -> list[str]:
    """target_ids of every manifest file present, usable or not."""
    directory = manifest_dir or MANIFEST_DIR
    if not directory.is_dir():
        return []
    out = []
    for path in sorted(directory.glob("*.json")):
        try:
            out.append(json.loads(path.read_text())["target_id"])
        except (json.JSONDecodeError, KeyError):
            continue
    return out


def load_manifest(target_id: str, manifest_dir: Path | None = None) -> Manifest:
    """Load and validate one manifest by target_id. Raises ManifestError
    if the file is missing or structurally invalid -- never returns a
    partially-parsed manifest."""
    directory = manifest_dir or MANIFEST_DIR
    cache_key = f"{directory}::{target_id}"
    if cache_key in _cache:
        return _cache[cache_key]

    if not directory.is_dir():
        raise ManifestError(f"manifest directory not found: {directory}")

    for path in sorted(directory.glob("*.json")):
        try:
            raw = json.loads(path.read_text())
        except json.JSONDecodeError as e:
            raise ManifestError(f"{path.name} is not valid JSON: {e}") from e
        if raw.get("target_id") == target_id:
            manifest = _parse_manifest(raw)
            _cache[cache_key] = manifest
            return manifest

    raise ManifestError(
        f"no manifest with target_id '{target_id}' in {directory} "
        f"(found: {', '.join(available_manifests(directory)) or 'none'})"
    )


def clear_cache() -> None:
    """Test hook -- manifests are cached per (directory, target_id)."""
    _cache.clear()
