"""Headless UI checks with synthetic records; no real Kobo calls."""
from pathlib import Path
import sys
import unittest

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from streamlit.testing.v1 import AppTest
from test_cfs_sources import sample


def script(live=False, fail=False, missing=False):
    content=(ROOT/"app.py").read_text(encoding="utf-8")
    row=sample(child_gender="Female",staff_filling_form="Synthetic staff",
        camp_of_information_seeking="Kalobeyei Integrated Settlement",cfs_visited="Village 1",
        first_visit_tdh_cfs="Yes",child_living_with_disability="No",referral_made="No",
        games_played="Football",support_offered_text="Psychological First Aid")
    insertion=f'''
kobo_configured=lambda:{live!r}
_old_setting=setting
setting=lambda name,default=None: "false" if name=="AUTH_REQUIRED" else _old_setting(name,default)
archive_signatures=lambda *args: (("synthetic-primary.xlsx",1,1,"synthetic"),)
def load_archive(*args):
    if {missing!r}: raise FileNotFoundError("synthetic")
    return pd.DataFrame([{row!r}])
def fetch_kobo_submissions(*args):
    if {fail!r}: raise RuntimeError("synthetic offline")
    return pd.DataFrame([{row!r}]),"2026-09-03T10:00:00Z",1
'''
    content=content.replace('\nst.set_page_config(',insertion+'\nst.set_page_config(',1)
    # Keep directory-relative assets bound to the real project during AppTest.
    position=content.index('from __future__ import annotations')+len('from __future__ import annotations')
    content=content[:position]+f'\n__file__={str(ROOT/"app.py")!r}\n'+content[position:]
    return content


class UITests(unittest.TestCase):
    def test_all_sections_render_and_routine_notices_are_in_panel(self):
        at=AppTest.from_string(script(),default_timeout=30).run()
        self.assertFalse(at.exception)
        panel=next(e for e in at.expander if e.label=="Historical + live reconciliation")
        self.assertFalse(panel.proto.expanded)
        for section in ("Monthly Trends","CPVs KPIs","Demographics","Games & Activities",
                        "Protection & Support","Referrals","Data Quality","Raw Data"):
            with self.subTest(section=section):
                at.selectbox(key="dashboard_section").set_value(section).run()
                self.assertFalse(at.exception)

    def test_failed_initial_live_load_requires_explicit_archive_choice(self):
        at=AppTest.from_string(script(live=True,fail=True),default_timeout=30).run()
        self.assertFalse(at.exception)
        self.assertTrue(at.error)
        at.button(key="cfs_choose_archive_only").click().run()
        self.assertFalse(at.exception)
        self.assertTrue(at.session_state["cfs_archive_only_requested"])
        self.assertTrue(any("Historical archive only" in w.value for w in at.warning))

    def test_missing_archives_live_only_still_renders(self):
        at=AppTest.from_string(script(live=True,missing=True),default_timeout=30).run()
        self.assertFalse(at.exception)
        self.assertTrue(any("primary historical archive is missing" in w.value for w in at.warning))


if __name__=="__main__":unittest.main()
