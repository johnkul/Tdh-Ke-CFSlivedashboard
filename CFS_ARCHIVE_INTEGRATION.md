# CFS historical + live integration

This change combines read-only historical archives with current Kobo submissions, then calls the existing CFS calculations once. It never changes the original Excel files or uploads archived records to Kobo.

## Implementation steps and full files

1. **cfs_sources.py**: cached historical loading, named schema/type normalization, UUID/root validation, conservative duplicate/revision handling and count diagnostics.
2. **app.py**: combine sources before existing prepare_data, retain the four linked tables and numeric record IDs, provide explicit archive-only mode and stale successful snapshots, and show routine diagnostics in a collapsed sidebar panel. Source/conflict fields are added to the existing safe Records view; the Data Quality section adds a collapsed conflict review. Private alternate values and reconciliation metadata are excluded from public exports.
3. **validate_cfs_sources.py**, **tests/test_cfs_sources.py**, **tests/test_cfs_ui.py**: read-only aggregate validation, synthetic cases and headless section/fallback checks.
4. **requirements.txt** pins the tested Streamlit 1.63 minor release. Other dependencies and domain calculations remain as before. **README.md** explains the combined workflow; **.gitignore** protects the new archive copy, credentials and generated caches.

All changed files are provided in full, with a separate code-only ZIP and Git patch. The local implementation branch is `feature/cfs-historical-live-data`, based on `fe18e1ac09d3743650c4a0c9d767e82eef90d06f`. Nothing has been pushed remotely by this implementation.

## Historical files and location

The supplied, newer backup has the same filename as an older workbook in the repository. Both originals are preserved. Put the supplied backup at the NEW path:

```text
data/historical/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx
```

The complete local source ZIP includes an unchanged copy there, plus the existing T1 and T2 snapshots. The changed-files ZIP contains no Excel records. The app loads:

1. The new primary archive above, or the explicit `CFS_DATA_PATH` override.
2. Existing `data/CFS_QUESTIONNAIRE_-_T2.xlsx`.
3. Existing `data/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx`.

Path/content checks prevent reading the same file twice if the override selects a secondary snapshot. The transformed-columns workbook and Excel lock file are not additional archive sources. If all explicit archives are absent, the former local-file resolver remains available as a clearly identified legacy fallback.

If the primary is missing, older snapshots/live remain usable but the app warns that the newer historical coverage is incomplete. Set any cloud override to a valid server path, never a Windows path on your personal computer.

## Record and version rules

- A valid canonical `_uuid` identifies a submission version. Whitespace, case and an optional `uuid:` prefix are normalized for comparison.
- Equivalent same-UUID copies become one row. Live Kobo wins; otherwise the primary/new archive wins, followed by T2 and older T1. Every contributing source/file locator is preserved in the private audit.
- `meta/rootUuid` links possible earlier/later versions. Different UUIDs sharing a root are retained and flagged. They are not automatically erased even if one version is newer.
- `_id` corroborates identity. Reused IDs with differing UUIDs are flagged, not used alone to delete rows.
- Changed answers, conflicting IDs/roots, bad parsed metadata/dates and contradictory aliases remain reviewable. Alternate raw values are held privately.
- Missing UUIDs stay separate. Exact payload fingerprints flag possible duplicates only; they do not delete records without reliable identity.
- Repeat visits with different genuine submission identities remain legitimate. Names, ages and phones are never submission-deduplication keys.
- Export/operational metadata such as validation status is ignored for answer equality but retained privately in the audit of equivalent copies.

Counts are **retained versions**, not independently verified original submissions or unique children, while conflict groups exist. The app displays a prominent warning when unresolved versions may inflate counts.

## Existing calculations preserved

The app retains numeric record_id values assigned once by prepare_data after reconciliation. All record/issue/support/game joins still use that ID.

Consent No stays excluded from analytical records. The reporting date still prefers **Kobo today**, then submission date, then entered date. Staff/CFS/location harmonization, demographics, first visits, disability, referrals, issue/support/game parsing and mention denominators stay in the original functions.

Dates, identifiers and phones are normalized before comparison. Numeric identifiers/phones become text without a numeric `.0`; leading zeros already lost by Excel cannot be restored. Private input representations may therefore differ from the old dataframe even though analytical results match.

Excel submission timestamps have no timezone metadata. `ARCHIVE_SUBMISSION_TIMEZONE` defaults to `UTC` to preserve the existing interpretation. Confirm this against your export; if it used Nairobi local time, use `Africa/Nairobi` in local/cloud Secrets. This notice lives inside reconciliation, not above the title.

## Source status and caching

- Archive loading uses bounded st.cache_data, keyed by path, nanosecond mtime, size, SHA256 and normalization version. Same-size replacements invalidate it.
- Live fetching retains the 30-minute cache and explicit refresh button. Combined keys include the archive set, normalization version, timezone, live context and refresh window/nonce.
- A failed same-source refresh keeps the last complete snapshot and displays its timestamp with a stale warning. Another asset/token never receives the previous snapshot.
- If initial live loading fails, the app stops and offers **Use historical archive only** when an archive is readable. **Retry live Kobo** returns to combined mode.
- No configured Kobo: explicitly labeled historical/local mode. Successful empty live data is a valid zero count. No usable source: actionable stop.
- Combined loading bypasses local-only prepared pickle files, so they cannot replace the merged dataset. It does not regenerate those pickle files or modify archives.

## Local VS Code commands

Open your existing CFS project in VS Code. Copy the changed-files ZIP contents into the matching paths, keeping both test files under **tests/**. Keep your existing secrets/authentication settings.

Run these one at a time from the project root in PowerShell (Python 3.12 was used for testing):

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_cfs_*.py" -v
.\.venv\Scripts\python.exe validate_cfs_sources.py --output validation-report-history.json
.\.venv\Scripts\python.exe -m streamlit run app.py
```

If .venv already exists, skip the first command. If multiple Python versions are installed, use `py -3.12 -m venv .venv`.

To validate just the supplied primary archive:

```powershell
.\.venv\Scripts\python.exe validate_cfs_sources.py --no-secondary --output validation-report-primary.json
```

To validate actual live overlap with your locally configured Kobo credentials:

```powershell
.\.venv\Scripts\python.exe validate_cfs_sources.py --with-live --output validation-report-combined.json
```

Do not paste tokens/passwords into chat or commit .streamlit/secrets.toml. The validator prints aggregate counts/fields only and refuses partial live retrieval.

To override the primary archive for this terminal session:

```powershell
$env:CFS_DATA_PATH = "C:\path\to\CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx"
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Alternatively use `validate_cfs_sources.py --archive "C:\path\to\file.xlsx"` for a validation-only override.

## Validated historical results

| Measure | Rows |
|---|---:|
| Supplied primary archive | 23,830 |
| Existing T2 snapshot | 16,551 |
| Existing T1 snapshot | 14,187 |
| Historical input total | 54,568 |
| Equivalent copies removed | 30,726 |
| Retained rows / UUID versions | 23,842 |
| Existing consent rule: analytical rows | 23,779 |
| Consent No excluded | 63 |
| Conflicting/revision rows retained | 22 |

There are 23,831 distinct canonical root UUIDs and 11 revision/collision groups. All 12 older-only UUID versions are preserved. No live records were fetched for this validation; live count and overlap remain unverified.

Reporting dates span 9 May 2025–3 October 2026. Submission timestamps as interpreted under UTC span 22 January 2026 04:16:07–30 September 2026 07:54:06.

Supplied-archive regression: the original and updated pipelines both produce 23,767 consented records, 46,376 issue mentions, 29,606 support mentions and 23,554 game mentions. Established analytical date/staff/gender/age/first-visit/disability/referral/location/CFS/activity values match. With the extra older versions retained, totals are 46,408 issues, 29,624 supports and 23,566 games.

The original supplied Excel checksum is unchanged:
`20ce6b266af2fb83a01b2e204b908137c2163c88e996bb9c3fbe9bce90c38ce6`.

Synthetic tests cover identities/revisions, alias conflicts, metadata/dates, empty sources, source failures, cache replacement, four-frame consistency, privacy-safe exports and all nine dashboard sections. Browser-only charts, CSS and map tiles still need a manual visual check.

Validation completed with **28 passing automated tests**, Python 3.12.10, Streamlit 1.63.0, pandas 2.3.3, openpyxl 3.1.5 and Plotly 6.9.0. Syntax and Git whitespace checks passed. Existing deprecated widget/component notices are non-failing warnings; this change does not redesign those controls.

## Push and deployment sequence

The CFS repository was public at inspection. **Do not push the newer archive until it is private.** The existing repository already contains historical child-data spreadsheets.

1. Make the GitHub repository private.
2. In Streamlit workspace Settings → Linked accounts, grant Streamlit private-repository access before rebooting. Keep existing Kobo/authentication Secrets.
3. In VS Code, create/switch to the feature branch, stage explicit code files, commit and push:

```powershell
git switch -c feature/cfs-historical-live-data
git add .gitignore README.md CFS_ARCHIVE_INTEGRATION.md requirements.txt app.py cfs_sources.py validate_cfs_sources.py tests/test_cfs_sources.py tests/test_cfs_ui.py
git diff --cached --name-only
git commit -m "Combine CFS historical archives with live Kobo"
git push -u origin feature/cfs-historical-live-data
```

If the branch already exists, use `git switch feature/cfs-historical-live-data` instead. For only the supplied matching repository base, a separate source patch can be checked/applied with `git apply --check` then `git apply`; it contains no Excel file.

4. After confirming the repository is private, include the new archive explicitly:

```powershell
git add -f data/historical/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx
git diff --cached --name-only
git commit -m "Add unchanged CFS historical backup"
git push
```

5. Create a GitHub pull request to main, review the changed files and merge. Your existing app uses app.py; no new app is needed. Reboot only after private access is confirmed if the Excel is not picked up automatically.
6. Open the deployed **Historical + live reconciliation** panel. Confirm every archive count, actual live count, duplicate equation, retained conflicts, dates and schema diagnostics.
7. Compare all dashboard sections, date/location filters, first visits, referrals, maps and protected device reporting. Do not assume a displayed retained-row count is a verified unique-child count.

If the cloud cannot clone the private repository, fix Linked accounts authorization rather than editing app code or Kobo secrets. If primary archive is missing, verify the cloud branch contains data/historical/... and that CFS_DATA_PATH is absent or valid for the server.
