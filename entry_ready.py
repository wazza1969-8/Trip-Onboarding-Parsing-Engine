"""
Entry-ready records: normalization, validation and manifest mapping
===================================================================

Takes the parser's raw extracted `sections` (the same list
backend.run_parsing builds before it calls parsing_engine.map_to_subjects)
and a destination manifest, and produces one record per destination
field, in the destination's own order, normalized and validated, with
the original extracted value kept alongside.

Two rules govern everything here:

1. **Never invent a value.** If a value cannot be normalized
   unambiguously, the original is preserved verbatim and the record is
   flagged AMBIGUOUS for a human to resolve. A missing value is reported
   as missing. Nothing is defaulted, inferred, rounded into place, or
   "helpfully" corrected.

2. **Every value stays traceable.** Each record carries the source
   section heading, the source field label, the page it came from and
   the untouched original string, so a reviewer can always answer "where
   did this come from?" without re-reading the PDF.

This module performs no I/O, holds no state and calls no model. It is
pure and deterministic, which is what makes it testable offline and
reusable later by either an API adapter or a Playwright adapter without
change.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field as dc_field
from typing import Any, Callable

import parsing_engine as pe
from field_manifest import FieldSpec, Manifest

# --- record status vocabulary ------------------------------------------------
# Deliberately small. Anything a reviewer must look at is AMBIGUOUS or
# INVALID; anything they can trust is OK; anything absent is MISSING.
OK = "ok"
AMBIGUOUS = "ambiguous"
INVALID = "invalid"
MISSING = "missing"
NOT_APPLICABLE = "not_applicable"

# Statuses that block a field from being considered ready for entry.
BLOCKING = (AMBIGUOUS, INVALID)


@dataclass
class EntryField:
    """One destination field, resolved against one parsed document."""
    field_id: str
    section_id: str
    section_label: str
    target_label: str
    order: int
    status: str
    value: str                     # normalized, display/copy ready ("" if none)
    original_value: str | None     # exactly as extracted, never altered
    source_section: str | None
    source_label: str | None
    source_page: int | None
    data_type: str
    required: bool
    criticality: str
    verification: str
    is_provisional: bool
    reason: str | None = None      # why it is ambiguous/invalid/missing
    occurrences: list[str] = dc_field(default_factory=list)  # repeating fields

    @property
    def needs_review(self) -> bool:
        return self.status in BLOCKING or (self.required and self.status == MISSING)

    @property
    def is_ready(self) -> bool:
        return self.status in (OK, NOT_APPLICABLE)


@dataclass
class NormResult:
    status: str
    value: str
    reason: str | None = None


# --- normalizers -------------------------------------------------------------
# Each takes the raw extracted string and returns a NormResult. They never
# raise, never guess, and never return a value the input did not support.

_EMPTY = {"", "n/a", "na", "none", "none listed", "none held", "not listed", "-", "--"}


def _blank(raw: str) -> bool:
    return raw.strip().lower() in _EMPTY


def norm_text(raw: str) -> NormResult:
    """Whitespace tidy only. Collapses runs of spaces and trims; does not
    change casing, punctuation or content."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    return NormResult(OK, re.sub(r"\s+", " ", raw.strip()))


def norm_strip_yes_prefix(raw: str) -> NormResult:
    """Source form answers often read 'Yes — Iran, Russia, ...'. Keep the
    substance, drop the leading Yes/No marker. If stripping would leave
    nothing, the answer was a bare Yes with no detail -- that is ambiguous,
    not empty, because the form implies detail should follow."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    text = raw.strip()

    # A bare "Yes" with nothing after it is not an answer to a field that
    # asks *which* countries/permissions -- it is an incomplete answer, and
    # passing it through as the value would misrepresent the form.
    if re.fullmatch(r"yes\b\.?", text, flags=re.IGNORECASE):
        return NormResult(AMBIGUOUS, text, "answered Yes but no detail was given")

    stripped = re.sub(r"^(yes|no)\s*[—\-–:,]\s*", "", text, flags=re.IGNORECASE).strip()
    if not stripped:
        return NormResult(MISSING, "", "answered No")
    return NormResult(OK, re.sub(r"\s+", " ", stripped))


_YES = {"yes", "y", "true", "required", "yes always", "always"}
_NO = {"no", "n", "false", "not required", "never"}


def norm_yes_no(raw: str) -> NormResult:
    """Canonical Yes/No. A qualified answer ('Yes, International only')
    keeps its qualifier and is flagged ambiguous rather than being
    flattened to 'Yes' -- the qualifier usually carries the operational
    meaning and dropping it silently changes what was agreed."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    text = raw.strip()
    low = text.lower().rstrip(".")
    if low in _YES:
        return NormResult(OK, "Yes")
    if low in _NO:
        return NormResult(OK, "No")
    if low.startswith(("yes", "no")):
        return NormResult(AMBIGUOUS, text, "qualified yes/no answer -- qualifier must be carried across by hand")
    return NormResult(AMBIGUOUS, text, "not recognisable as Yes or No")


def norm_checkbox(raw: str) -> NormResult:
    """Checked/Unchecked. The engine asks the extractor to always report
    this state explicitly, so an unrecognised value means something went
    wrong upstream rather than an empty box."""
    if raw is None or raw.strip() == "":
        return NormResult(MISSING, "", "checkbox state was not reported")
    low = raw.strip().lower()
    if low in {"checked", "ticked", "yes", "true", "x"}:
        return NormResult(OK, "Checked")
    if low in {"unchecked", "not checked", "no", "false", "empty"}:
        return NormResult(OK, "Unchecked")
    return NormResult(AMBIGUOUS, raw.strip(), "checkbox state not recognisable as Checked/Unchecked")


_ICAO_RE = re.compile(r"^[A-Z]{4}$")
_ICAO_LOOSE_RE = re.compile(r"^[A-Z0-9]{4}$")


def norm_icao(raw: str) -> NormResult:
    """Format validation only -- this does not confirm the aerodrome
    exists, which would need a navdata lookup and is deliberately out of
    scope for an offline, deterministic layer. Multiple codes are flagged
    rather than silently reduced to the first, matching the engine's own
    behaviour of withholding multi-airport answers."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    text = raw.strip().upper()
    parts = [p.strip() for p in re.split(r"[,/;&]| and ", text) if p.strip()]
    if len(parts) > 1:
        return NormResult(AMBIGUOUS, ", ".join(parts),
                          f"{len(parts)} identifiers given; destination expects a single home base")
    code = parts[0]
    if _ICAO_RE.match(code):
        return NormResult(OK, code)
    if _ICAO_LOOSE_RE.match(code):
        return NormResult(AMBIGUOUS, code, "4 characters but not all letters -- may be an FAA/IATA code rather than ICAO")
    return NormResult(INVALID, code, "not a 4-character ICAO identifier")


_REG_RE = re.compile(r"^[A-Z0-9]{1,7}(-[A-Z0-9]{1,5})?$")

# Placeholder text people write into a field they cannot answer yet. These
# pass a pure format check -- "TBC" is three alphanumeric characters and
# looks exactly like a short registration -- so they are caught by name.
# An unresolved placeholder reaching a handling or customs record is worse
# than a blank, because a blank is obviously incomplete and "TBC" is not.
_PLACEHOLDER_TOKENS = {
    "TBC", "TBD", "TBA", "NA", "N/A", "NONE", "UNKNOWN", "PENDING",
    "ASK", "VARIES", "SAME", "ALL", "XXX", "???",
}


def norm_registration_list(raw: str) -> NormResult:
    """One or more aircraft registrations. Format check only; no registry
    lookup. Anything that does not look like a registration is flagged
    rather than dropped, because a dropped tail number is invisible."""
    if _blank(raw):
        return NormResult(MISSING, "", "no registration given")
    # Note: "/" is NOT a separator here. Registrations are comma- or
    # and-separated on this form, and splitting on "/" would turn the
    # placeholder "N/A" into the two plausible-looking tokens "N" and "A".
    parts = [p.strip().upper() for p in re.split(r"[,;&]| and ", raw.strip()) if p.strip()]
    joined = ", ".join(parts)

    placeholders = [p for p in parts if p.strip(".?") in _PLACEHOLDER_TOKENS]
    if placeholders:
        return NormResult(AMBIGUOUS, joined,
                          f"placeholder text where a registration is expected: "
                          f"{', '.join(placeholders)}")

    bad = [p for p in parts if not _REG_RE.match(p)]
    if bad:
        return NormResult(AMBIGUOUS, joined, f"not recognisable as registration(s): {', '.join(bad)}")
    return NormResult(OK, joined)


_EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[A-Za-z]{2,}$")


def norm_email(raw: str) -> NormResult:
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    text = raw.strip()
    parts = [p.strip() for p in re.split(r"[,;/ ]+", text) if p.strip()]
    emails = [p for p in parts if "@" in p]
    if not emails:
        return NormResult(INVALID, text, "no email address found")
    bad = [e for e in emails if not _EMAIL_RE.match(e)]
    if bad:
        return NormResult(INVALID, ", ".join(emails), f"malformed address: {', '.join(bad)}")
    if len(emails) > 1:
        return NormResult(AMBIGUOUS, ", ".join(emails), "more than one address given")
    return NormResult(OK, emails[0])


def norm_contact(raw: str) -> NormResult:
    """Contact rows arrive as a single composed string built by the
    extraction prompt ('<Name>, <Role>, <Email>[, Office: ...][, Mobile: ...]').
    It is passed through intact -- splitting it into components here would
    mean guessing which comma-separated fragment is a role, and a wrong
    guess is worse than a reviewer reading one line."""
    if _blank(raw):
        return NormResult(MISSING, "", "no contact given")
    text = re.sub(r"\s+", " ", raw.strip())
    if "@" not in text:
        return NormResult(AMBIGUOUS, text, "no email address in the contact row")
    return NormResult(OK, text)


def norm_star_rating(raw: str) -> NormResult:
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    m = re.search(r"(\d(?:\.\d)?)", raw)
    if not m:
        return NormResult(AMBIGUOUS, raw.strip(), "no numeric rating found")
    value = float(m.group(1))
    if not (1 <= value <= 5):
        return NormResult(INVALID, raw.strip(), f"star rating {value} outside 1-5")
    return NormResult(OK, f"{int(value)} star" if value.is_integer() else f"{value} star")


def norm_duration_hours(raw: str) -> NormResult:
    """'1 hour', '2.0 hours prior', '90 minutes' -> canonical hours.
    A bare number is ambiguous: the form does not state a unit, and
    assuming hours would be inventing one."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    text = raw.strip().lower()
    m = re.search(r"(\d+(?:\.\d+)?)\s*(hours?|hrs?|h|minutes?|mins?|m)\b", text)
    if not m:
        if re.fullmatch(r"\d+(?:\.\d+)?", text):
            return NormResult(AMBIGUOUS, raw.strip(), "no unit given -- cannot tell hours from minutes")
        return NormResult(AMBIGUOUS, raw.strip(), "no recognisable duration")
    amount, unit = float(m.group(1)), m.group(2)
    hours = amount / 60 if unit.startswith(("m", "min")) and unit != "h" else amount
    rendered = f"{hours:g} hour" + ("" if hours == 1 else "s")
    # Preserve a directional qualifier ('prior', 'after') if present -- it
    # changes the meaning and belongs with the value.
    qualifier = ""
    for word in ("prior", "before", "after", "post"):
        if word in text:
            qualifier = f" {word}"
            break
    return NormResult(OK, rendered + qualifier)


def norm_flight_types(raw: str) -> NormResult:
    """Validated against parsing_engine.FLIGHT_TYPE_OPTIONS -- the engine's
    own list, read at runtime so the two cannot drift apart."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    options = {o.lower(): o for o in pe.FLIGHT_TYPE_OPTIONS}
    parts = [p.strip() for p in re.split(r"[,;/]| and ", raw.strip()) if p.strip()]
    resolved, unknown = [], []
    for p in parts:
        match = options.get(p.lower())
        if match:
            resolved.append(match)
        else:
            unknown.append(p)
    if unknown:
        return NormResult(AMBIGUOUS, ", ".join(parts),
                          f"not in the engine's flight-type list: {', '.join(unknown)}")
    return NormResult(OK, ", ".join(resolved))


def norm_who_list(raw: str) -> NormResult:
    """'Crew, Passengers' style answers. Permitted values are checked by
    the enum validator against the manifest, which flags them as
    provisional because they were inferred from sample data."""
    if _blank(raw):
        return NormResult(MISSING, "", "no value on the source form")
    parts = [p.strip().title() for p in re.split(r"[,;/&]| and ", raw.strip()) if p.strip()]
    return NormResult(OK, ", ".join(parts))


NORMALIZERS: dict[str, Callable[[str], NormResult]] = {
    "text": norm_text,
    "strip_yes_prefix": norm_strip_yes_prefix,
    "yes_no": norm_yes_no,
    "checkbox": norm_checkbox,
    "icao": norm_icao,
    "registration_list": norm_registration_list,
    "email": norm_email,
    "contact": norm_contact,
    "star_rating": norm_star_rating,
    "duration_hours": norm_duration_hours,
    "flight_types": norm_flight_types,
    "who_list": norm_who_list,
}


# --- validators --------------------------------------------------------------
# Run after normalization. They can only downgrade a status, never upgrade
# it -- a validator cannot rescue a value the normalizer rejected.

def _validate(spec: FieldSpec, result: NormResult) -> NormResult:
    if result.status in (MISSING, INVALID):
        return result

    for name in spec.validators:
        if name == "non_empty" and not result.value.strip():
            return NormResult(MISSING, "", "required value is empty")

        if name == "enum_member" and spec.permitted_values:
            allowed = {v.lower() for v in spec.permitted_values}
            given = [p.strip() for p in result.value.split(",") if p.strip()]
            unknown = [g for g in given if g.lower() not in allowed]
            if unknown:
                note = (
                    " (note: this list was inferred from sample data, not a "
                    "specification -- confirm before relying on it)"
                    if not spec.permitted_values_verified else ""
                )
                return NormResult(
                    AMBIGUOUS, result.value,
                    f"value(s) not in the permitted list: {', '.join(unknown)}{note}",
                )

        if name == "icao_format" and result.status == OK and not _ICAO_RE.match(result.value):
            return NormResult(INVALID, result.value, "not a 4-letter ICAO identifier")

    return result


# --- mapping -----------------------------------------------------------------

def _index_sections(sections: list[dict]) -> dict[tuple[str, str], list[tuple[str, int | None]]]:
    """(section heading lower, label lower) -> [(value, page), ...].

    Keeps every occurrence: repeating rows such as 'Key contact' legitimately
    appear more than once, and collapsing them would lose contacts."""
    index: dict[tuple[str, str], list[tuple[str, int | None]]] = {}
    for section in sections or []:
        heading = (section.get("heading") or "").strip().lower()
        page = section.get("page")
        for item in section.get("items", []) or []:
            label = (item.get("label") or "").strip().lower()
            value = (item.get("value") or "").strip()
            index.setdefault((heading, label), []).append((value, page))
    return index


def _lookup(index, source_section: str | None, source_label: str | None):
    """Exact match on the manifest's declared source names first, then a
    substring match on the label within the same section heading -- the
    engine's own whitelist matching is substring-based and case
    insensitive, so this mirrors it rather than being stricter."""
    if not source_section or not source_label:
        return []
    sec, lab = source_section.strip().lower(), source_label.strip().lower()
    if (sec, lab) in index:
        return index[(sec, lab)]
    hits: list[tuple[str, int | None]] = []
    for (h, l), values in index.items():
        if sec in h and (lab in l or l in lab):
            hits.extend(values)
    return hits


def _condition_met(spec: FieldSpec, resolved: dict[str, EntryField]) -> bool | None:
    """None when the condition cannot be evaluated yet (dependency not
    resolved). Fails open to 'evaluate the field anyway' only when the
    dependency itself is missing -- a field is never silently suppressed
    because of an unresolved dependency."""
    if not spec.conditional_on:
        return True
    dep_id = spec.conditional_on.get("field_id")
    expected = spec.conditional_on.get("equals")
    dep = resolved.get(dep_id)
    if dep is None or dep.status in (MISSING, AMBIGUOUS, INVALID):
        return None
    return dep.value.strip().lower() == str(expected).strip().lower()


def build_entry_ready_records(sections: list[dict], manifest: Manifest) -> list[EntryField]:
    """Resolve every manifest field against one parsed document.

    Raises ValueError if the manifest has no fields -- an unpopulated
    manifest (the portal placeholder) must fail closed rather than
    render an empty form that looks complete."""
    if not manifest.is_usable():
        raise ValueError(
            f"manifest '{manifest.target_id}' describes no fields and cannot be used "
            f"for mapping. {manifest.target_verification_note}"
        )

    index = _index_sections(sections)
    resolved: dict[str, EntryField] = {}
    records: list[EntryField] = []

    for section in manifest.sections:
        for spec in section.fields:
            hits = _lookup(index, spec.source_section, spec.source_label)
            occurrences = [v for v, _ in hits if v]
            page = next((p for _, p in hits if p is not None), None)

            if not occurrences:
                result = NormResult(MISSING, "", "field not present in the parsed document")
                original = None
            elif spec.repeating:
                normalized, reasons = [], []
                for raw in occurrences:
                    r = _validate(spec, NORMALIZERS[spec.normalizer](raw))
                    normalized.append(r.value or raw)
                    if r.status in BLOCKING:
                        reasons.append(f"{raw!r}: {r.reason}")
                result = NormResult(
                    AMBIGUOUS if reasons else OK,
                    "\n".join(normalized),
                    "; ".join(reasons) or None,
                )
                original = " | ".join(occurrences)
            else:
                original = occurrences[0]
                if len(occurrences) > 1:
                    result = NormResult(
                        AMBIGUOUS, " | ".join(occurrences),
                        f"{len(occurrences)} values found for a single-value field",
                    )
                    original = " | ".join(occurrences)
                else:
                    result = _validate(spec, NORMALIZERS[spec.normalizer](original))

            condition = _condition_met(spec, resolved)
            if condition is False and result.status == MISSING:
                dep = spec.conditional_on.get("field_id")
                result = NormResult(NOT_APPLICABLE, "",
                                    f"not applicable -- {dep} is not "
                                    f"{spec.conditional_on.get('equals')}")

            record = EntryField(
                field_id=spec.field_id,
                section_id=section.section_id,
                section_label=section.label,
                target_label=spec.target_label,
                order=spec.order,
                status=result.status,
                value=result.value,
                original_value=original,
                source_section=spec.source_section,
                source_label=spec.source_label,
                source_page=page,
                data_type=spec.data_type,
                required=spec.required,
                criticality=spec.criticality,
                verification=spec.verification,
                is_provisional=spec.is_provisional,
                reason=result.reason,
                occurrences=occurrences if spec.repeating else [],
            )
            resolved[spec.field_id] = record
            records.append(record)

    return records


# --- summary -----------------------------------------------------------------

@dataclass
class CompletionSummary:
    total: int
    ready: int
    needs_review: int
    missing_required: int
    provisional: int
    by_criticality: dict[str, dict[str, int]]

    @property
    def percent_ready(self) -> float:
        return (self.ready / self.total * 100) if self.total else 0.0

    @property
    def blocked_on_critical(self) -> int:
        c = self.by_criticality.get("critical", {})
        return c.get("needs_review", 0) + c.get("missing_required", 0)


def summarize(records: list[EntryField]) -> CompletionSummary:
    by_crit: dict[str, dict[str, int]] = {}
    for level in ("critical", "required_noncritical", "optional"):
        subset = [r for r in records if r.criticality == level]
        by_crit[level] = {
            "total": len(subset),
            "ready": sum(1 for r in subset if r.is_ready),
            "needs_review": sum(1 for r in subset if r.status in BLOCKING),
            "missing_required": sum(1 for r in subset if r.required and r.status == MISSING),
        }
    return CompletionSummary(
        total=len(records),
        ready=sum(1 for r in records if r.is_ready),
        needs_review=sum(1 for r in records if r.status in BLOCKING),
        missing_required=sum(1 for r in records if r.required and r.status == MISSING),
        provisional=sum(1 for r in records if r.is_provisional),
        by_criticality=by_crit,
    )
