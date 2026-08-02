# Tdh Kenya CFS Dashboard — Kobo live setup

The supplied `app.py` reads submissions from the KoboToolbox KPI API v2. It caches API responses for 60 seconds, follows all pagination links, and provides a **Fetch latest Kobo data** button. The existing Excel/CSV loader remains an automatic fallback only when Kobo credentials are not configured.

## 1. Place the files

Use this structure:

```text
your-app/
  app.py
  requirements.txt
  assets/
    styles.css
    tdh-logo.png
    developer-logo.png
  .streamlit/
    secrets.toml
```

## 2. Configure secrets

Copy `secrets.toml.example` to `.streamlit/secrets.toml` and replace the placeholders. Never commit the real token to Git.

The asset UID is the alphanumeric value in a Kobo project URL, for example the `AbCd123...` part of `.../#/forms/AbCd123.../data/table`.

For the global Kobo humanitarian server use `https://kobo.humanitarianresponse.info` as `KOBO_BASE_URL`. For the global non-humanitarian server use `https://kf.kobotoolbox.org`.

## 3. Map Kobo XML field names

The JSON API returns XLSForm XML names, not necessarily the human-readable Excel question labels. The app already accepts:

- transformed dashboard names such as `child_age`;
- those names inside Kobo group paths, such as `child_details/child_age`;
- system columns such as `_id` and `_submission_time`;
- your existing raw human-readable labels when present.

If an XML name differs, add it to `KOBO_COLUMN_MAP` in the secrets file. The left side is Kobo's exact JSON key/path and the right side is the transformed analysis name used in `app.py`.

Use **Data Quality → Analysis Schema Readiness** after the first connection to identify unmapped fields. Do not map two source fields to the same target unless one is known to be empty, because duplicate names can make analysis ambiguous.

## 4. Install and run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Live-dashboard behaviour

- A browser interaction reruns Streamlit, but Kobo is called at most once per 60-second cache window.
- The manual refresh button clears only Kobo data caches and fetches immediately.
- The dashboard displays the last successful fetch time in East Africa Time.
- API errors are shown without printing the API token.
- All API pages are followed, so projects with more than 1,000 submissions are supported.

This is near-live polling, not a permanent streaming socket. A 30–60 second refresh window is usually a better fit for Streamlit and Kobo than querying on every widget click. For unattended wall displays, add `streamlit-autorefresh` and refresh every 60 seconds; avoid much shorter intervals because every active browser session can generate server load.

## Recommended next improvements

1. Add a small “new submissions since last refresh” KPI based on `_submission_time`.
2. Add data-quality alerts for missing consent, age, location, and unmapped schema fields.
3. Add role-based access before exposing row-level or PII-bearing views; keep public downloads privacy-safe.
4. For large projects, move ingestion to a scheduled database job and let Streamlit query aggregated tables. Kobo polling is appropriate at modest scale, but a database layer improves concurrency, auditability, and history.
5. If edits to existing submissions must appear, retain API polling. Kobo REST Services can push newly created submissions, but do not by themselves cover later edits.

