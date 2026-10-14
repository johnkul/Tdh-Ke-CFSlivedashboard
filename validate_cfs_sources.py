"""Read-only aggregate validation; no child values or credentials are printed."""
import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def dashboard_definitions(path=None):
    path = Path(path or ROOT / "app.py")
    text = path.read_text(encoding="utf-8-sig").split("\nst.set_page_config(", 1)[0]
    namespace = {"__file__": str(ROOT / "app.py"), "__name__": "cfs_validation_definitions"}
    exec(compile(text, str(path), "exec"), namespace)
    return namespace


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive", help="Override the primary historical file.")
    parser.add_argument("--no-secondary", action="store_true", help="Validate only the supplied archive, without older T1/T2.")
    parser.add_argument("--with-live", action="store_true", help="Read live Kobo using locally configured credentials.")
    parser.add_argument("--output", help="Save aggregate JSON report.")
    args = parser.parse_args()
    app = dashboard_definitions()
    import os
    primary = Path(args.archive or os.environ.get("CFS_DATA_PATH") or app["HISTORICAL_PRIMARY_PATH"])
    secondary = () if args.no_secondary else app["HISTORICAL_SECONDARY_PATHS"]
    signatures = app["archive_signatures"](primary, secondary)
    if args.with_live and not app["kobo_configured"]():
        parser.error("--with-live requires locally configured KOBO_TOKEN and KOBO_ASSET_UID.")
    context = (str(app["setting"]("KOBO_BASE_URL", "https://kf.kobotoolbox.org")).rstrip("/"),
               str(app["setting"]("KOBO_ASSET_UID", "")), str(app["setting"]("KOBO_TOKEN", "")))
    result = app["load_combined_dashboard"].__wrapped__(*context, signatures,
              str(app["setting"]("ARCHIVE_SUBMISSION_TIMEZONE", "UTC")), 0, args.with_live)
    report = dict(result[1].attrs["source_metadata"])
    report["archive_checksums_unchanged"] = all(app["file_signature"](s[0]) == s for s in signatures)
    report["table_rows"] = {n: len(f) for n, f in zip(("records", "issues", "supports", "games"), result[1:5])}
    if not report["archive_checksums_unchanged"]:
        raise RuntimeError("Source archive changed during validation; repeat with stable files.")
    text = json.dumps(report, indent=2, default=str, ensure_ascii=False)
    print(text)
    if args.output:
        path = Path(args.output); path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
