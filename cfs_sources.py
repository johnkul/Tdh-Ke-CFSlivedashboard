"""Read-only historical sources and conservative Kobo version reconciliation."""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import uuid
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

VERSION = "cfs-combined-v1"
PUBLIC_SOURCE_FIELDS = ("data_source", "source_membership", "submission_conflict", "identity_unresolved")
INTERNAL_FIELDS = set(PUBLIC_SOURCE_FIELDS) | {"source_file", "source_row_number", "submission_key", "version_key"}
OPERATIONAL = {"id", "uuid", "root_uuid", "index", "version", "notes", "tags", "status",
               "validation_status", "submitted_by", "username", "deviceid", "phonenumber",
               "information_statement", "end_note", "_attachments", "_geolocation", "_xform_id_string",
               "_bamboo_dataset_id", "_edited", "_last_edited", "start", "end", "formhub/uuid"}


def blank(value):
    if isinstance(value, (list, tuple, dict, set)):
        return False
    return value is None or bool(pd.isna(value)) or (isinstance(value, str) and value.strip().casefold() in
        {"", "nan", "na", "n/a", "missing / unspecified", "[missing]", "[not recorded]"})


def uuid_value(value):
    if blank(value):
        return None
    text = str(value).strip()
    if text.casefold().startswith("uuid:"):
        text = text[5:]
    try:
        return str(uuid.UUID(text))
    except (ValueError, AttributeError):
        return None


def integer_id(value):
    if blank(value):
        return None
    try:
        number = Decimal(str(value).strip())
        return str(int(number)) if number.is_finite() and number > 0 and number == number.to_integral_value() else None
    except (InvalidOperation, ValueError):
        return None


def comparable(value):
    if blank(value):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).isoformat()
    if isinstance(value, (list, tuple, set, dict)):
        return json.dumps(value, sort_keys=True, default=str)
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float)):
        return str(Decimal(str(value)).normalize())
    return " ".join(str(value).split())


def file_signature(path):
    path = Path(path)
    try:
        stat = path.stat()
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        return str(path.resolve()), stat.st_mtime_ns, stat.st_size, digest.hexdigest()
    except OSError:
        return str(path.absolute()), None, None, None


def archive_signatures(primary, secondary):
    """One read per file/content even when an override selects an old snapshot."""
    output = []
    seen_paths, seen_digests = set(), set()
    for path in (primary, *secondary):
        signature = file_signature(path)
        if signature[0] in seen_paths or (signature[3] and signature[3] in seen_digests):
            continue
        output.append(signature)
        seen_paths.add(signature[0])
        if signature[3]:
            seen_digests.add(signature[3])
    return tuple(output)


@st.cache_data(show_spinner=False, max_entries=8)
def load_archive(signature, version=VERSION):
    del version
    path = Path(signature[0])
    if signature[1] is None:
        raise FileNotFoundError(path.name)
    frame = pd.read_excel(path, dtype=object) if path.suffix.lower() in {".xlsx", ".xls"} else pd.read_csv(path, dtype=object)
    if tuple(signature) != file_signature(path):
        raise RuntimeError("Archive changed while reading; retry against a stable file.")
    return frame


def normalize_source(raw, source, filename, mapping, norm_text, contract, archive_timezone="UTC", yes_no_parser=None):
    """Named mapping with private preservation of conflicting input aliases."""
    ZoneInfo(archive_timezone)  # Reject a bad timezone instead of emptying dates.
    exact = {str(k).strip(): v for k, v in mapping.items()}
    normalized = {norm_text(k): v for k, v in mapping.items()}
    output, alias_issues = {}, []
    original_types = {}
    for position, column in enumerate(raw.columns):
        name = str(column).strip()
        target = exact.get(name) or normalized.get(norm_text(name)) or name
        if target in INTERNAL_FIELDS or target == "record_id":
            target = "__reserved__" + target
        values = raw.iloc[:, position].copy().reset_index(drop=True)
        original_types[name] = sorted({type(v).__name__ for v in values if not blank(v)})
        if target not in output:
            output[target] = values
        else:
            current = output[target]
            both = ~current.map(blank) & ~values.map(blank)
            conflicts = both & (current.map(lambda v: json.dumps(comparable(v))) != values.map(lambda v: json.dumps(comparable(v))))
            if conflicts.any():
                output[f"__alias_conflict__{target}_{position}"] = values.where(conflicts)
                alias_issues.append({"field": target, "rows": int(conflicts.sum())})
            output[target] = current.where(~current.map(blank), values)
    out = pd.DataFrame(output, index=range(len(raw)))
    original_columns = set(out.columns)
    conversion_issues = []
    if yes_no_parser is not None:
        for column in ("consent", "child_living_with_disability", "first_visit_tdh_cfs",
                       "referral_made", "take5_activities_integrated"):
            if column in out:
                out[column] = out[column].map(lambda v: yes_no_parser(v) if yes_no_parser(v) in {"Yes", "No"} else v)
    for column in ("id", "uuid", "root_uuid", "caregiver_parent_phone", "child_individual_number"):
        if column in out:
            out[column] = out[column].map(lambda v: pd.NA if blank(v) else
                (str(int(v)) if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() else str(v).strip()))
    for column, parser in (("id", integer_id), ("uuid", uuid_value), ("root_uuid", uuid_value)):
        if column in out:
            invalid = ~out[column].map(blank) & out[column].map(parser).isna()
            if invalid.any():
                out[f"__invalid__{column}"] = out[column].where(invalid)
                conversion_issues.append({"field": column, "rows": int(invalid.sum())})
    for column in ("today", "date", "submission_time"):
        if column not in out:
            continue
        def parse(value):
            if blank(value):
                return pd.NaT
            try:
                stamp = (pd.Timestamp("1899-12-30") + pd.to_timedelta(value, unit="D")
                         if isinstance(value, (int, float)) else pd.Timestamp(value))
                if column == "submission_time":
                    if stamp.tzinfo is None:
                        stamp = stamp.tz_localize(archive_timezone if source == "Historical archive" else "UTC")
                    return stamp.tz_convert("UTC")
                return stamp.tz_localize(None) if stamp.tzinfo else stamp
            except (ValueError, TypeError, OverflowError):
                return pd.NaT
        parsed = out[column].map(parse)
        invalid = ~out[column].map(blank) & parsed.isna()
        if invalid.any():
            out[f"__invalid__{column}"] = out[column].where(invalid)
            conversion_issues.append({"field": column, "rows": int(invalid.sum())})
        out[column] = pd.to_datetime(parsed, utc=column == "submission_time")
    for column in original_columns:
        if column.startswith(("issue_", "support_")) and column in contract and column not in {"issue_other_specify", "support_offered_text"}:
            parsed = pd.to_numeric(out[column], errors="coerce")
            invalid = ~out[column].map(blank) & parsed.isna()
            if invalid.any():
                out[f"__invalid__{column}"] = out[column].where(invalid)
                conversion_issues.append({"field": column, "rows": int(invalid.sum())})
            out[column] = parsed
    out["data_source"] = source
    out["source_membership"] = source
    out["source_file"] = filename
    out["source_row_number"] = range(2, len(out) + 2)
    out["submission_conflict"] = False
    out["identity_unresolved"] = False
    out.attrs["diagnostics"] = {
        "source": source, "file": filename, "rows": len(out), "columns": sorted(original_columns),
        "missing_fields": sorted(set(contract) - original_columns),
        "extra_fields": sorted(original_columns - set(contract)), "alias_conflicts": alias_issues,
        "conversion_issues": conversion_issues, "original_types": original_types,
        "timezone_note": f"Excel submission timestamps are interpreted as {archive_timezone}; confirm the export timezone."
                         if source == "Historical archive" and len(out) else None,
    }
    return out


def reconcile(frames):
    """Equivalent UUID copies collapse; changed/root-linked versions remain."""
    diagnostics = [frame.attrs.get("diagnostics", {}) for frame in frames]
    copies = []
    for frame in frames:
        copy = frame.copy(); copy.attrs = {}; copies.append(copy)
    if not copies:
        copies = [pd.DataFrame(columns=list(INTERNAL_FIELDS))]
    combined = pd.concat(copies, ignore_index=True, sort=False)
    if combined.empty:
        for c in INTERNAL_FIELDS:
            combined[c] = pd.Series(dtype=object)
    versions = [uuid_value(v) for v in combined.get("uuid", pd.Series(None, index=combined.index, dtype=object))]
    roots = [uuid_value(v) for v in combined.get("root_uuid", pd.Series(None, index=combined.index, dtype=object))]
    ids = [integer_id(v) for v in combined.get("id", pd.Series(None, index=combined.index, dtype=object))]
    fields = sorted(set(combined.columns) - INTERNAL_FIELDS - OPERATIONAL)
    # Fingerprint only business answers, retaining null/unknown fields. Operational
    # metadata is kept in the private audit for equivalent copies.
    payloads = [{c: comparable(v) for c, v in zip(fields, row)}
                for row in combined.reindex(columns=fields).itertuples(index=False, name=None)]
    fingerprints = [hashlib.sha256(json.dumps(p, sort_keys=True, default=str).encode()).hexdigest() for p in payloads]
    groups, root_groups, id_groups = defaultdict(list), defaultdict(list), defaultdict(list)
    for i in combined.index:
        groups[versions[i] or f"unresolved:{i}"].append(i)
        if roots[i]: root_groups[roots[i]].append(i)
        if ids[i]: id_groups[ids[i]].append(i)
        combined.at[i, "version_key"] = versions[i] or f"unresolved:{i}"
        combined.at[i, "submission_key"] = roots[i] or versions[i] or f"unresolved:{i}"
        combined.at[i, "identity_unresolved"] = versions[i] is None
    root_revision_groups = sum(len({versions[i] for i in members if versions[i]}) > 1 for members in root_groups.values())
    id_collision_groups = sum(len({versions[i] for i in members if versions[i]}) > 1 for members in id_groups.values())
    flagged = set()
    for members in (*root_groups.values(), *id_groups.values()):
        if len({versions[i] for i in members if versions[i]}) > 1:
            flagged.update(members)
    kept, removed, duplicate_groups, conflicting_uuid_groups = [], 0, 0, 0
    private_audit = []
    for version, members in groups.items():
        duplicate_groups += int(len(members) > 1)
        differing = len({fingerprints[i] for i in members}) > 1
        identity_contradiction = len({ids[i] for i in members if ids[i]}) > 1 or len({roots[i] for i in members if roots[i]}) > 1
        invalid = any(any(c.startswith(("__alias_conflict__", "__invalid__")) and not blank(combined.at[i, c]) for c in combined) for i in members)
        if differing or identity_contradiction or invalid:
            flagged.update(members)
            kept.extend(members)
            conflicting_uuid_groups += 1
            private_audit.append({"kind": "conflict", "locators": [{"source": combined.at[i,"data_source"], "file": combined.at[i,"source_file"], "row": int(combined.at[i,"source_row_number"])} for i in members],
                                  "fields": [c for c in fields if len({json.dumps(payloads[i][c], default=str) for i in members}) > 1]})
        else:
            # Frame order is primary archive, T2, T1, then live. Live wins;
            # otherwise primary/new archive wins among equivalent copies.
            chosen = next((i for i in members if combined.at[i,"data_source"] == "Live Kobo"), members[0])
            combined.at[chosen,"source_membership"] = "; ".join(sorted({combined.at[i,"data_source"] for i in members}))
            kept.append(chosen); removed += len(members)-1
            if len(members)>1:
                private_audit.append({"kind":"equivalent", "locators":[{"source":combined.at[i,"data_source"],"file":combined.at[i,"source_file"],"row":int(combined.at[i,"source_row_number"])} for i in members],
                    "operational_metadata":[{c:comparable(combined.at[i,c]) for c in OPERATIONAL if c in combined} for i in members]})
    for i in flagged:
        combined.at[i,"submission_conflict"] = True
    result = combined.loc[kept].reset_index(drop=True)
    historical = sum(len(f) for f in frames if f.attrs.get("diagnostics",{}).get("source") == "Historical archive")
    live = sum(len(f) for f in frames if f.attrs.get("diagnostics",{}).get("source") == "Live Kobo")
    assert historical+live-removed == len(result)
    stamps = pd.to_datetime(result.get("submission_time",pd.Series(dtype=object)),utc=True,errors="coerce")
    without_ids = [fingerprints[i] for i in combined.index if not versions[i]]
    metadata = {
        "historical_rows":historical,"live_rows":live,"duplicates_removed":removed,"combined_rows":len(result),
        "per_file_rows": {d.get("file", "unknown"):d.get("rows",0) for d in diagnostics if d.get("source")=="Historical archive"},
        "duplicate_uuid_groups":duplicate_groups,"conflicting_uuid_groups":conflicting_uuid_groups,
        "root_revision_groups":root_revision_groups,"id_collision_groups":id_collision_groups,
        "conflict_rows":int(result["submission_conflict"].sum()),"identity_unresolved_rows":int(result["identity_unresolved"].sum()),
        "distinct_uuids":len({versions[i] for i in kept if versions[i]}),"distinct_root_uuids":len({roots[i] for i in kept if roots[i]}),
        "possible_duplicates_without_uuid":sum(n-1 for n in Counter(without_ids).values()),
        "source_breakdown":result["data_source"].value_counts().to_dict(),
        "source_membership_breakdown":result["source_membership"].value_counts().to_dict(),
        "earliest_submission_utc":stamps.min().isoformat() if stamps.notna().any() else None,
        "latest_submission_utc":stamps.max().isoformat() if stamps.notna().any() else None,
        "missing_submission_dates":int(stamps.isna().sum()),"source_diagnostics":diagnostics,
        "fingerprint":hashlib.sha256(json.dumps([(fingerprints[i],versions[i],roots[i],combined.at[i,"source_membership"]) for i in kept],default=str).encode()).hexdigest(),
    }
    result.attrs["private_reconciliation_audit"] = private_audit
    return result,metadata
