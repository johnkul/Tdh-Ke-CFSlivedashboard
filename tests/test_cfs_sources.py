"""Synthetic identity/normalization and shared pipeline tests."""
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from cfs_sources import archive_signatures, file_signature, load_archive, normalize_source, reconcile
from validate_cfs_sources import dashboard_definitions

A = "00000000-0000-4000-8000-000000000001"
B = "00000000-0000-4000-8000-000000000002"
R = "00000000-0000-4000-8000-000000000010"
CONTRACT = {"id", "uuid", "root_uuid", "consent", "today", "date", "submission_time", "child_age", "issue_food"}


def sample(**changes):
    row = {"id": 1, "uuid": A, "root_uuid": "uuid:"+R, "consent": "Yes", "today": "2026-09-01",
           "date": "2026-08-20", "submission_time": "2026-09-03T10:00:00Z", "child_age": 10, "issue_food": 1}
    row.update(changes)
    return row


def source(rows, live=False, filename="primary.xlsx"):
    return normalize_source(pd.DataFrame(rows), "Live Kobo" if live else "Historical archive", filename,
                            {c:c for c in CONTRACT}, lambda s:str(s).strip().casefold(), CONTRACT)


class SourceTests(unittest.TestCase):
    def test_equivalent_copies_prefer_live(self):
        data, meta = reconcile([source([sample()]), source([sample(uuid="uuid:"+A.upper())], True)])
        self.assertEqual(len(data), 1)
        self.assertEqual(meta["duplicates_removed"], 1)
        self.assertEqual(data.iloc[0].data_source, "Live Kobo")
        self.assertEqual(data.iloc[0].source_membership, "Historical archive; Live Kobo")

    def test_same_uuid_changed_answers_keeps_both(self):
        _, meta = reconcile([source([sample()]), source([sample(child_age=11)],True)])
        self.assertEqual(meta["combined_rows"], 2)
        self.assertEqual(meta["conflict_rows"], 2)

    def test_same_root_different_versions_are_not_deleted(self):
        _, meta = reconcile([source([sample()]), source([sample(uuid=B,child_age=11)],True)])
        self.assertEqual(meta["root_revision_groups"], 1)
        self.assertEqual(meta["combined_rows"], 2)
        self.assertEqual(meta["conflict_rows"], 2)

    def test_same_id_different_uuid_and_root_flags_collision(self):
        _, meta = reconcile([source([sample()]), source([sample(uuid=B,root_uuid=B)],True)])
        self.assertEqual(meta["id_collision_groups"], 1)
        self.assertEqual(meta["conflict_rows"], 2)

    def test_repeated_child_visit_with_new_identity_is_legitimate(self):
        _, meta = reconcile([source([sample()]),source([sample(id=2,uuid=B,root_uuid=B)],True)])
        self.assertEqual(meta["combined_rows"], 2)
        self.assertEqual(meta["conflict_rows"], 0)

    def test_missing_uuid_never_automatically_removed(self):
        _, meta = reconcile([source([sample(uuid=None,id=None,root_uuid=None)]*2)])
        self.assertEqual(meta["combined_rows"], 2)
        self.assertEqual(meta["identity_unresolved_rows"], 2)
        self.assertEqual(meta["possible_duplicates_without_uuid"], 1)

    def test_same_uuid_conflicting_ids_are_retained(self):
        _, meta = reconcile([source([sample()]),source([sample(id=99)],True)])
        self.assertEqual(meta["conflict_rows"], 2)

    def test_operational_metadata_differences_are_audited(self):
        data, meta = reconcile([source([sample(status="submitted")]),source([sample(status="edited")],True)])
        self.assertEqual(meta["combined_rows"], 1)
        self.assertIn("operational_metadata",data.attrs["private_reconciliation_audit"][0])

    def test_boolean_indicators_compare_to_excel_integer_choices(self):
        _,meta=reconcile([source([sample(issue_food=1)]),source([sample(issue_food=True)],True)])
        self.assertEqual(meta["duplicates_removed"],1)

    def test_new_populated_field_is_not_silently_lost(self):
        _, meta = reconcile([source([sample()]),source([sample(new_answer="Recorded")],True)])
        self.assertEqual(meta["combined_rows"], 2)

    def test_contradictory_alias_values_are_retained(self):
        frame = pd.DataFrame([["Yes","No"]],columns=["consent","Do you consent?"])
        data = normalize_source(frame,"Historical archive","test.xlsx",{"Do you consent?":"consent"},str,CONTRACT)
        self.assertIn("No",data.iloc[0].tolist())
        self.assertEqual(data.attrs["diagnostics"]["alias_conflicts"][0]["rows"],1)

    def test_bad_date_is_reported_and_preserved(self):
        data=source([sample(submission_time="bad date")])
        self.assertTrue(pd.isna(data.iloc[0].submission_time))
        self.assertEqual(data.iloc[0]["__invalid__submission_time"],"bad date")

    def test_invalid_root_metadata_prevents_automatic_copy_deletion(self):
        _,meta=reconcile([source([sample(root_uuid="malformed")]),source([sample(root_uuid="different-invalid")],True)])
        self.assertEqual(meta["combined_rows"],2)
        self.assertEqual(meta["conflict_rows"],2)

    def test_reserved_input_source_is_not_trusted_or_dropped(self):
        data=source([sample(data_source="Original answer")])
        self.assertEqual(data.iloc[0].data_source,"Historical archive")
        self.assertEqual(data.iloc[0]["__reserved__data_source"],"Original answer")

    def test_empty_sources(self):
        data,meta=reconcile([source([]),source([],True)])
        self.assertTrue(data.empty)
        self.assertEqual(meta["combined_rows"],0)

    def test_same_size_replacement_and_file_override(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/"archive.csv";path.write_text("answer\nfirst\n")
            before=file_signature(path)
            self.assertEqual(load_archive(before).iloc[0,0],"first")
            path.write_text("answer\nother\n")
            after=file_signature(path)
            self.assertNotEqual(before[3],after[3])
            self.assertEqual(len(archive_signatures(path,[path])),1)
            self.assertEqual(load_archive(after).iloc[0,0],"other")


class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app=dashboard_definitions()

    def fixture(self):
        return pd.DataFrame([sample(child_gender="Female",staff_filling_form="Synthetic staff",
            camp_of_information_seeking="Kalobeyei Integrated Settlement",cfs_visited="Village 1",
            first_visit_tdh_cfs="Yes",child_living_with_disability="No",referral_made="No",
            games_played="Football",support_offered_text="Psychological First Aid")])

    def test_all_four_frames_share_one_combined_snapshot_and_date_priority(self):
        raw=self.fixture()
        with patch.dict(self.app,{"load_archive":lambda *a:raw,
              "fetch_kobo_submissions":lambda *a:(raw,"2026-09-03T10:00:00Z",1)}):
            frames=self.app["load_combined_dashboard"].__wrapped__("server","asset","token",(("primary",1,1,"digest"),))
        self.assertEqual(frames[0],1)
        self.assertEqual(frames[1].iloc[0].date,pd.Timestamp("2026-09-01"))
        self.assertEqual(frames[1].iloc[0].record_id,1)
        self.assertEqual(frames[1].attrs["source_metadata"]["duplicates_removed"],1)
        for frame in frames[1:5]:
            self.assertEqual(frame.attrs["source_metadata"]["combined_rows"],1)

    def test_missing_archive_uses_live_and_reports_missing_primary(self):
        def missing(*a):raise FileNotFoundError("primary")
        with patch.dict(self.app,{"load_archive":missing,
            "fetch_kobo_submissions":lambda *a:(self.fixture(),"2026-09-03T10:00:00Z",1)}):
            frames=self.app["load_combined_dashboard"].__wrapped__("server","asset","token",(("primary",None,None,None),))
        self.assertEqual(frames[1].attrs["source_metadata"]["live_rows"],1)
        self.assertFalse(frames[1].attrs["source_metadata"]["primary_archive_available"])

    def test_successful_empty_live_and_empty_archive(self):
        with patch.dict(self.app,{"load_archive":lambda *a:pd.DataFrame(),
            "fetch_kobo_submissions":lambda *a:(pd.DataFrame(),"2026-09-03T10:00:00Z",1)}):
            frames=self.app["load_combined_dashboard"].__wrapped__("server","asset","token",(("primary",1,1,"digest"),))
        self.assertEqual(frames[0],0)
        self.assertTrue(all(f.empty for f in frames[1:5]))

    def test_consent_no_remains_excluded(self):
        raw=self.fixture(); raw.loc[0,"consent"]="No"
        with patch.dict(self.app,{"load_archive":lambda *a:raw}):
            frames=self.app["load_combined_dashboard"].__wrapped__("","","",(("primary",1,1,"digest"),),use_live=False)
        self.assertEqual(frames[0],1)
        self.assertEqual(len(frames[1]),0)

    def test_live_yes_no_codes_match_archive_labels(self):
        archived=self.fixture();live=archived.copy()
        live["consent"]="yes";live["referral_made"]="no"
        with patch.dict(self.app,{"load_archive":lambda *a:archived,
            "fetch_kobo_submissions":lambda *a:(live,"2026-09-03T10:00:00Z",1)}):
            frames=self.app["load_combined_dashboard"].__wrapped__("server","asset","token",(("primary",1,1,"digest"),))
        self.assertEqual(frames[1].attrs["source_metadata"]["duplicates_removed"],1)

    def test_kobo_failure_does_not_silently_fallback(self):
        def fail(*a):raise RuntimeError("offline")
        with patch.dict(self.app,{"load_archive":lambda *a:self.fixture(),"fetch_kobo_submissions":fail}):
            with self.assertRaisesRegex(RuntimeError,"offline"):
                self.app["load_combined_dashboard"].__wrapped__("server","asset","token",(("primary",1,1,"digest"),))

    def test_public_export_strips_private_reconciliation_data(self):
        raw=self.fixture();raw["__invalid__child_name"]="Synthetic child"
        raw["submission_key"]=R;raw.attrs["private_reconciliation_audit"]=[{"deviceid":"private"}]
        public=self.app["privacy_safe_export"](raw)
        self.assertNotIn("__invalid__child_name",public)
        self.assertNotIn("submission_key",public)
        self.assertNotIn("private_reconciliation_audit",public.attrs)

    def test_failed_refresh_preserves_same_source_snapshot(self):
        state={}
        original=self.app["load_combined_dashboard"]
        with patch.dict(self.app,{"st":SimpleNamespace(session_state=state),"load_combined_dashboard":lambda *a:("complete",)}):
            result=self.app["combined_session_snapshot"](("server","asset","token"),(),"UTC",0,True)
        def fail(*a):raise RuntimeError("offline")
        with patch.dict(self.app,{"st":SimpleNamespace(session_state=state),"load_combined_dashboard":fail}):
            next_result=self.app["combined_session_snapshot"](("server","asset","token"),(),"UTC",1,True)
        self.assertEqual(result,next_result)
        self.assertTrue(state["cfs_source_refresh_failed"])

    def test_changed_asset_does_not_reuse_previous_snapshot(self):
        state={}
        with patch.dict(self.app,{"st":SimpleNamespace(session_state=state),"load_combined_dashboard":lambda *a:("complete",)}):
            self.app["combined_session_snapshot"](("server","asset","token"),(),"UTC",0,True)
        def fail(*a):raise RuntimeError("offline")
        with patch.dict(self.app,{"st":SimpleNamespace(session_state=state),"load_combined_dashboard":fail}):
            with self.assertRaisesRegex(RuntimeError,"offline"):
                self.app["combined_session_snapshot"](("server","other","token"),(),"UTC",1,True)


if __name__=="__main__":unittest.main()
