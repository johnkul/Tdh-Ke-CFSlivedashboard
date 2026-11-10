"""Restricted, paginated child register. Never used by public exports."""
import hmac
import time
from io import BytesIO

import pandas as pd
import streamlit as st

PREFIX = "cfs_child_register_"


def filter_register_visits(data, selection):
    """Use the dashboard's cleaned first-visit response, never infer missing visits."""
    if selection == "All visits":
        return data
    status = data.get("first_visit_clean", pd.Series(pd.NA, index=data.index, dtype="string"))
    status = status.astype("string").fillna("")
    if selection == "First visits":
        mask = status.eq("Yes")
    elif selection == "Repeat visits":
        mask = status.eq("No")
    else:
        mask = ~status.isin(["Yes", "No"])
    return data.loc[mask]


def reset_visit_export():
    st.session_state.pop(PREFIX + "export", None)
    st.session_state.pop(PREFIX + "page", None)


def child_register(data):
    """Project recorded values and apply the approved nationality/profile rule."""
    def field(*names):
        result = pd.Series("", index=data.index, dtype=object)
        for name in names:
            if name in data:
                values = data[name].astype("string").fillna("").str.strip()
                valid = ~values.str.lower().isin(["", "nan", "none", "missing", "missing / unspecified", "not recorded"])
                result = result.mask(result.eq("") & valid, values)
        return result.replace("", "Not recorded")

    dates = pd.to_datetime(data.get("date", pd.Series(index=data.index, dtype=object)), errors="coerce")
    table = pd.DataFrame(index=data.index)
    table["Date"] = dates.dt.strftime("%d %b %Y").fillna("Not recorded")
    table["Child's Name"] = field("child_name")
    table["Individual Number"] = field("child_individual_number")
    table["Location/Telephone Number"] = field("caregiver_parent_phone")
    table["Gender"] = field("gender_clean", "child_gender")
    table["Age"] = field("age_clean", "child_age")
    table["Age group"] = field("age_group")
    table["Nationality"] = field("nationality")
    table["Disability (Yes/No)"] = field("disability_status_clean", "child_living_with_disability")
    table["If Yes, which disability"] = field("disability_type_display", "disability_type")
    table.loc[~table["Disability (Yes/No)"].str.casefold().eq("yes"), "If Yes, which disability"] = "—"
    table["Location (specific)"] = field("location_clean", "specific_camp_location", "exact_registered_location")
<<<<<<< Updated upstream
=======
    table["Dadaab — Section and Block residence of the information seeker"] = field("section_block_residence", "Section and Block residence of the information seeker")
    table["Kalobeyei Settlement — Exact location according to proof of registration"] = field("exact_registered_location", "Exact location of the information seeker (according to proof of registration)")
>>>>>>> Stashed changes
    nationality = table["Nationality"].str.strip().str.casefold()
    profile = pd.Series("Refugee", index=data.index)
    profile.loc[nationality.isin(["kenyan", "kenyans", "kenya"])] = "Host Community"
    profile.loc[nationality.eq("not recorded")] = "Not recorded"
    table["Profile (Refugee/Host Community)"] = profile
    return table.reset_index(drop=True)


def lock_register():
    for key in ("until", "credential", "password", "error", "page", "export"):
        st.session_state.pop(PREFIX + key, None)


def register_excel(data):
    """Export only the requested register fields; identifiers remain Excel text."""
    from openpyxl import Workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Filtered child register"
    table = child_register(data)
    sheet.append(list(table.columns))
    for row_number, row in enumerate(table.itertuples(index=False, name=None), 2):
        # Preserve leading zeros and treat user-entered formula-looking values
        # as literal text, rather than spreadsheet formulas.
        for column_number, value in enumerate(row, 1):
            cell = sheet.cell(row_number, column_number, value)
            cell.data_type = "s"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    from openpyxl.utils import get_column_letter
    for position, column in enumerate(table.columns, 1):
        sheet.column_dimensions[get_column_letter(position)].width = min(42, max(18, len(column) + 2))
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def unlock_register(expected):
    supplied = str(st.session_state.get(PREFIX + "password", ""))
    st.session_state[PREFIX + "password"] = ""
    if expected and hmac.compare_digest(supplied.encode(), expected.encode()):
        st.session_state[PREFIX + "until"] = time.time() + 600
        # Bind access to the configured credential without retaining plaintext.
        import hashlib
        st.session_state[PREFIX + "credential"] = hashlib.sha256(expected.encode()).hexdigest()
        st.session_state[PREFIX + "error"] = False
    else:
        lock_register()
        st.session_state[PREFIX + "error"] = True


@st.fragment(run_every="30s")
def render_protected_records(data, setting):
    """Check authorization before constructing or transmitting identifying rows."""
    import hashlib
    password = str(setting("CFS_RECORDS_PASSWORD", "") or "").strip()
    if password.upper().startswith("REPLACE_WITH"):
        password = ""
    until = st.session_state.get(PREFIX + "until", 0)
    authorized = bool(password and until > time.time() and
                      st.session_state.get(PREFIX + "credential") == hashlib.sha256(password.encode()).hexdigest())
    if not authorized and until:
        lock_register()
    with st.expander("🔒 Protected child register", expanded=authorized):
        st.caption("Follows the current date, location, CFS and staff filters. Telephone is the caregiver/parent telephone. Excel download includes all filtered register rows. Access lasts 10 minutes.")
        if not password:
            st.info("Configure CFS_RECORDS_PASSWORD in Streamlit Secrets to enable this register.")
            return
        if not authorized:
            with st.form(PREFIX + "form"):
                st.text_input("Child register password", type="password", key=PREFIX + "password")
                st.form_submit_button("Unlock child register", on_click=unlock_register, args=(password,))
            if st.session_state.get(PREFIX + "error"):
                st.error("Incorrect password. The register remains locked.")
            return
        st.button("Lock child register", on_click=lock_register, key=PREFIX + "lock")
        source_data = data
        visit_filter = st.radio(
            "Visit type", ["All visits", "First visits", "Repeat visits", "Visit not recorded"],
            horizontal=True, key=PREFIX + "visit_filter", on_change=reset_visit_export,
        )
        data = filter_register_visits(source_data, visit_filter)
        st.caption(f"{visit_filter}: {len(data):,} of {len(source_data):,} dashboard-filtered submission records. This selection applies to both the table and Excel download.")
        st.caption("First visits use first_visit_clean = Yes; repeat visits use No. Missing or unrecognized responses stay in All visits or Visit not recorded. Counts are submissions, not unique children.")
<<<<<<< Updated upstream
=======
        st.caption("Age preserves the recorded exact age or source age bracket. Dadaab uses the recorded Section and Block residence; Kalobeyei Settlement uses the recorded exact location according to proof of registration. Both columns are separate from the broader specific location; missing values are not inferred.")
>>>>>>> Stashed changes
        st.caption("Profile follows the agreed reporting rule: Kenyan nationality = Host Community; other recorded nationalities = Refugee. Missing nationality = Not recorded.")
        prepared = st.session_state.get(PREFIX + "export")
        # A parent rerun/filter change supplies a new frame. Never reuse a file
        # prepared for a different frame; no PII is placed in a shared cache.
        if prepared is not None and (prepared["source"] is not source_data or prepared.get("visit_filter") != visit_filter):
            st.session_state.pop(PREFIX + "export", None)
            prepared = None
        if st.button("Prepare protected Excel download", disabled=data.empty, key=PREFIX + "prepare"):
            prepared = {"source": source_data, "visit_filter": visit_filter, "bytes": register_excel(data)}
            st.session_state[PREFIX + "export"] = prepared
        if prepared is not None:
            st.caption("Contains child-identifying information. The workbook itself is not password-encrypted.")
            st.download_button(
                "Download protected register (.xlsx)", data=prepared["bytes"],
                file_name="CFS_filtered_child_register.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=PREFIX + "download", on_click="ignore",
            )
        # Build only the visible page. st.table provides no CSV/download toolbar.
        pages = max(1, (len(data) + 49) // 50)
        if st.session_state.get(PREFIX + "page", 1) > pages:
            st.session_state[PREFIX + "page"] = 1
        page = st.number_input("Register page", min_value=1, max_value=pages, step=1, key=PREFIX + "page")
        st.caption(f"Page {page} of {pages} · up to 50 rows per page")
        if data.empty:
            st.info("No records match the current filters.")
        else:
            st.table(child_register(data.iloc[(page - 1) * 50:page * 50]))
