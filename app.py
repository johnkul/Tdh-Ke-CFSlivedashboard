"""
Tdh Kenya CFS Dashboard - Streamlit Application

A professional, multi-section dashboard for analysing Child Friendly Spaces
(CFS) data from TDH Kenya operations. It includes data loading, harmonisation,
interactive filtering, charts, styled tables, and CSV exports.

Usage:
    streamlit run app.py

Data source:
    Configure KOBO_TOKEN and KOBO_ASSET_UID in Streamlit secrets. The app reads
    live submissions from Kobo KPI API v2. A local Excel/CSV remains an optional
    fallback when Kobo credentials are not configured.
"""

from __future__ import annotations

import base64
import hashlib
import html as html_lib
import os
import re
import time
import json
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_FILE_CANDIDATES = [
    "CFS_QUESTIONNAIRE_-_T2.xlsx",
    "CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx",
    "CFS_QUESTIONNAIRE_-_Tdh_Kenya_T1.xlsx",
]
DATA_FILE_NAME = DATA_FILE_CANDIDATES[0]
DEFAULT_DATA_PATH = DATA_DIR / DATA_FILE_NAME
LOGO_PATH = BASE_DIR / "assets" / "tdh-logo.png"
DEVELOPER_LOGO_PATH = BASE_DIR / "assets" / "developer-logo.png"
CSS_PATH = BASE_DIR / "assets" / "styles.css"
APP_VERSION = "Version 1.1 · Kobo live data · August 2026"
SCHEMA_CONTRACT_VERSION = "cfs-schema-2026.08"
# Keep prepared data hot for normal navigation. Administrators can bypass this
# window at any time with “Fetch latest Kobo data”.
KOBO_CACHE_TTL_SECONDS = 1800
KOBO_REQUEST_TIMEOUT_SECONDS = 45
PREPARED_DATA_PATH = DATA_DIR / "cfs_dashboard_prepared.pkl"
PREPARED_CACHE_PATH = BASE_DIR / ".cfs_dashboard_prepared_cache.pkl"
PREPARED_CACHE_VERSION = "cfs-dashboard-prepared-v16"

RAW_TO_TRANSFORMED_COLUMNS = {
    # System / metadata columns
    "today": "today",
    "username": "username",
    "deviceid": "deviceid",
    "phonenumber": "phonenumber",
    "Thank you for visiting the Tdh Child-Friendly Space. We keep a record of our conversations so that we can provide you with the best possible service. We also use this information to improve our services and to report to our donors. Your information will be kept confidential and will not be shared with anyone outside of Tdh without your permission.": "information_statement",
    "THANK YOU FOR PARTICIPATING.": "end_note",
    "Location of the information seeker.": "gps_location",
    "_id": "_id",
    "_uuid": "_uuid",
    "_submission_time": "submission_time",
    "_validation_status": "validation_status",
    "_notes": "notes",
    "_status": "status",
    "_submitted_by": "submitted_by",
    "__version__": "version",
    "_tags": "tags",
    "meta/rootUuid": "root_uuid",
    "_index": "index",

    # Core survey fields
    "Do you consent to participate?": "consent",
    "Staff filling the form": "staff_filling_form",
    "Enter a date": "date",
    "Name of the Child (two names only)": "child_name",
    "Age of the child.": "child_age",
    "How does the child seeking information identify.": "child_gender",
    "How does the child seeking information identify. ": "child_gender",
    "Is the child living with disability?": "child_living_with_disability",
    "What type of disability?": "disability_type",
    "If Other, please specify the type of disability.": "disability_type_other",
    "Individual Number of the child": "child_individual_number",
    "Child's Caregiver/Parent Names (two names)": "caregiver_parent_names",
    "Camp of information seeking.": "camp_of_information_seeking",
    "Specific camp location of the information seeker.": "specific_camp_location",
    "Section and Block residence of the information seeker": "section_block_residence",
    "Camp camp location of the information seeker.": "camp_location_alt",
    "Exact location of the information seeker (according to proof of registration)": "exact_registered_location",
    "Child's caregiver or parent telephone number (if available, this is optional)": "caregiver_parent_phone",
    "Name of the child friendly space the information seeker visited.": "child_friendly_space_visited",
    "Name of the CFS the information seeker visited.": "cfs_visited",
    "Which games did the child play/was engaged with?": "games_played",
    "If other, specify the type of game involved.": "game_other_specify",
    "Was Take 5 activities integrated into play sessions?": "take5_activities_integrated",
    " Was Take 5 activities integrated into play sessions?": "take5_activities_integrated",
    "Is this your first visit to any of Tdh`s CFS": "first_visit_tdh_cfs",
    "Is this your first visit to any of Tdh's CFS": "first_visit_tdh_cfs",
    "Is this your first visit to any of Tdh’s CFS": "first_visit_tdh_cfs",
    "Is this your first visit to any of Tdh CFS": "first_visit_tdh_cfs",
    "Is this your first visit to any of Tdh CFS?": "first_visit_tdh_cfs",

    # Issues
    "Nature of Issues Reported": "nature_issues_reported_text",
    "Nature of Issues Reported/New arrival/Lack of card": "issue_new_arrival_lack_of_card",
    "Nature of Issues Reported/Disability": "issue_disability",
    "Nature of Issues Reported/Basic needs": "issue_basic_needs",
    "Nature of Issues Reported/Deceased parent": "issue_deceased_parent",
    "Nature of Issues Reported/Education": "issue_education",
    "Nature of Issues Reported/Psychosocial support": "issue_psychosocial_support",
    "Nature of Issues Reported/Neglected": "issue_neglected",
    "Nature of Issues Reported/Parents seperated": "issue_parents_separated",
    "Nature of Issues Reported/Parents separated": "issue_parents_separated",
    "Nature of Issues Reported/Child out of wedlock": "issue_child_out_of_wedlock",
    "Nature of Issues Reported/Food": "issue_food",
    "Nature of Issues Reported/Clothing": "issue_clothing",
    "Nature of Issues Reported/Shelter": "issue_shelter",
    "Nature of Issues Reported/Reporting a protection concern": "issue_reporting_protection_concern",
    "Nature of Issues Reported/persons in need of profiling/registration by UNHCR": "issue_need_profiling_registration_unhcr",
    "Nature of Issues Reported/None": "issue_none",
    "Nature of Issues Reported/Other": "issue_other",
    "If Other, specify": "issue_other_specify",

    # Support / referral
    "Support offered": "support_offered_text",
    "Support offered/Psychological First Aid": "support_psychological_first_aid",
    "Support offered/Play and art therapy": "support_play_art_therapy",
    "Support offered/Psychoeducation": "support_psychoeducation",
    "Support offered/None": "support_none",
    "Was a referral on the issue reported made?": "referral_made",
    "If yes, where was it referred to?": "referral_destination",
    "If External, please specify the agency referred to:": "external_referral_agency",
    "THANK YOU FOR PARTICIPATING. ": "end_note",
}

# Additional schema aliases from the current online survey export. These keep
# the app stable when Kobo/ODK exports either raw question labels or transformed
# analysis-column names.
RAW_TO_TRANSFORMED_COLUMNS.update({
    "Thank you for visiting the Tdh Child-Friendly Space. We keep a record of our conversations by completing a form. This helps us follow up and provide the right support when needed. The information you share will remain confidential and will only be used for reporting or referral purposes. Do you give your consent for us to record this information?": "information_statement",
    "_id": "id",
    "_uuid": "uuid",
    "_submission_time": "submission_time",
    "_validation_status": "validation_status",
    "_submitted_by": "submitted_by",
    "__version__": "version",
    "*id*": "id",
    "*uuid*": "uuid",
    "*submission*time": "submission_time",
    "*validation*status": "validation_status",
    "*submitted*by": "submitted_by",
    "**version**": "version",
})

# Ensure transformed analysis-column names map to themselves. This makes the
# standardisation layer explicit and prevents future schema changes from being
# silently ignored when the dataset already arrives transformed.
ANALYSIS_COLUMN_NAMES = [
    "today", "username", "deviceid", "phonenumber", "information_statement", "consent",
    "staff_filling_form", "date", "child_name", "child_age", "child_gender",
    "child_living_with_disability", "disability_type", "disability_type_other",
    "child_individual_number", "caregiver_parent_names", "camp_of_information_seeking",
    "specific_camp_location", "section_block_residence", "camp_location_alt",
    "exact_registered_location", "caregiver_parent_phone", "child_friendly_space_visited",
    "cfs_visited", "games_played", "game_other_specify", "take5_activities_integrated",
    "first_visit_tdh_cfs", "nature_issues_reported_text", "issue_new_arrival_lack_of_card",
    "issue_disability", "issue_basic_needs", "issue_deceased_parent", "issue_education",
    "issue_psychosocial_support", "issue_neglected", "issue_parents_separated",
    "issue_child_out_of_wedlock", "issue_food", "issue_clothing", "issue_shelter",
    "issue_reporting_protection_concern", "issue_need_profiling_registration_unhcr",
    "issue_none", "issue_other", "issue_other_specify", "support_offered_text",
    "support_psychological_first_aid", "support_play_art_therapy", "support_psychoeducation",
    "support_none", "referral_made", "referral_destination", "external_referral_agency",
    "end_note", "gps_location", "id", "uuid", "submission_time", "validation_status",
    "notes", "status", "submitted_by", "version", "tags", "root_uuid", "index",
]
RAW_TO_TRANSFORMED_COLUMNS.update({col: col for col in ANALYSIS_COLUMN_NAMES})

MISSING = "Missing / unspecified"
REVIEW = "Needs review"
AGE_GROUP_ORDER = [
    "0-4 years",
    "5-11 years",
    "12-17 years",
    "18-24 years",
    "25-49 years",
    "50 years and above",
    MISSING,
]
YES_NO_ORDER = ["Yes", "No", MISSING]
GENDER_ORDER = ["Girls", "Boys", "Intersex", "Transgender", "Intersex / Transgender", "Non-binary", "Other / self-described", MISSING]
TOP_N_OPTIONS = [5, 10, 15, 25]
CHART_COLORS = ["#1d4ed8", "#16a34a", "#f97316", "#dc2626", "#7c3aed", "#0891b2", "#be123c", "#4338ca"]

# Direct and indirect identifiers must never be included in ordinary dashboard
# downloads. Keep them in the prepared frame only for controlled internal data
# quality checks.
PII_COLUMNS = {
    "child_name", "child_individual_number", "caregiver_parent_names",
    "caregiver_parent_phone", "phonenumber", "deviceid", "username",
    "gps_location", "exact_registered_location", "section_block_residence",
    "uuid", "root_uuid", "id", "notes", "tags", "submitted_by",
}

ISSUE_COLUMNS = {
    "issue_new_arrival_lack_of_card": "New arrival / Lack of card",
    "issue_disability": "Disability",
    "issue_basic_needs": "Basic needs",
    "issue_deceased_parent": "Deceased parent",
    "issue_education": "Education",
    "issue_psychosocial_support": "Psychosocial support",
    "issue_neglected": "Neglected",
    "issue_parents_separated": "Parents separated",
    "issue_child_out_of_wedlock": "Child out of wedlock",
    "issue_food": "Food",
    "issue_clothing": "Clothing",
    "issue_shelter": "Shelter",
    "issue_reporting_protection_concern": "Reporting a protection concern",
    "issue_need_profiling_registration_unhcr": "Needs profiling / registration by UNHCR",
    "issue_none": "None",
}

SUPPORT_COLUMNS = {
    "support_psychological_first_aid": "Psychological First Aid",
    "support_play_art_therapy": "Play and art therapy",
    "support_psychoeducation": "Psychoeducation",
    "support_none": "None",
}

CORE_ANALYSIS_COLUMNS = [
    "today", "username", "deviceid", "phonenumber", "information_statement",
    "consent", "staff_filling_form", "date", "child_name", "child_age", "child_gender",
    "child_living_with_disability", "disability_type", "disability_type_other",
    "child_individual_number", "caregiver_parent_names", "caregiver_parent_phone",
    "camp_of_information_seeking", "specific_camp_location", "section_block_residence",
    "camp_location_alt", "exact_registered_location", "child_friendly_space_visited", "cfs_visited",
    "games_played", "game_other_specify", "take5_activities_integrated", "first_visit_tdh_cfs",
    "nature_issues_reported_text", "issue_other", "issue_other_specify", "support_offered_text",
    "referral_made", "referral_destination", "external_referral_agency",
    "end_note", "gps_location", "id", "uuid", "submission_time", "validation_status",
    "notes", "status", "submitted_by", "version", "tags", "root_uuid", "index",
]

CORE_ANALYSIS_COLUMNS += list(ISSUE_COLUMNS.keys()) + list(SUPPORT_COLUMNS.keys())

STAFF_MAP = {
    "mohamed sidi": "Mohamed Sidi", "mohamed": "Mohamed Sidi",
    "halima amin": "Halima Amin",
    "moge garad": "Moge Garad", "moge": "Moge Garad", "more garad": "Moge Garad",
    "daud hussein": "Daud Hussein",
    "maslah mohamed hassan": "Maslah Mohamed Hassan", "maslah mohamed hasssan": "Maslah Mohamed Hassan",
    "maslah mohamed": "Maslah Mohamed Hassan", "maslsh mohamed hassan": "Maslah Mohamed Hassan",
    "maslish mohamed": "Maslah Mohamed Hassan", "maslah kohamed hassan": "Maslah Mohamed Hassan",
    "maslah kohamed": "Maslah Mohamed Hassan", "maslsh mohamed hasssan": "Maslah Mohamed Hassan",
    "ndayikeje ferdinand": "Ndayikeje Ferdinand", "ferdinand ndayikeje": "Ndayikeje Ferdinand",
    "ndayikeje": "Ndayikeje Ferdinand", "ferdinand": "Ndayikeje Ferdinand",
    "teresia natire thomas": "Teresia Natire Thomas", "teresia natire": "Teresia Natire Thomas",
    "haret derow ibrahim": "Haret Derow Ibrahim", "haret derow": "Haret Derow Ibrahim", "hared derow": "Haret Derow Ibrahim",
    "beatrice akwero": "Beatrice Akwero",
    "david otifo": "David Otifo", "dave otifo": "David Otifo",
    "farah mohamed hussein": "Farah Mohamed Hussein",
    "musdaf mohamed hassan": "Musdaf Mohamed Hassan", "musdaf mohamed": "Musdaf Mohamed Hassan",
    "hirwa gentille": "Hirwa Gentille", "hirwa": "Hirwa Gentille",
    "oliek omot": "Oliek Omot", "nelson amanya": "Nelson Amanya",
    "dimo justin": "Dimo Justin", "dimo": "Dimo Justin", "louis kyanza": "Louis Kyanza",
    "dominic nangiro lomil": "Dominic Nangiro Lomil", "dominic nangiro": "Dominic Nangiro Lomil",
    "leer biel leer": "Leer Biel Leer", "spora niyikiza": "Spora Niyikiza", "niyikiza spora": "Spora Niyikiza",
    "jean claude": "Jean Claude", "fowzia omar": "Fowzia Omar",
    "yaak akech": "Yaak Akech", "aketch yaak": "Yaak Akech", "akech yaak": "Yaak Akech",
    "peter kingombe": "Peter Kingombe",
    "dual ador arok": "Dual Ador Arok", "dual ador": "Dual Ador Arok", "ador arok dual": "Dual Ador Arok",
    "gatwech bayak": "Gatwech Bayak", "gatwech": "Gatwech Bayak",
    "armele ngakani": "Armele Ngakani", "armele": "Armele Ngakani", "john wani": "John Wani",
    "safari david": "Safari David", "safali": "Safari David", "epusie brenda": "Epusie Brenda",
    "zahara issack": "Zahara Issack", "zahara": "Zahara Issack",
    "nyakhor buob": "Nyakhor Buob Tang", "nyqkhor buob tang": "Nyakhor Buob Tang", "nyakhor": "Nyakhor Buob Tang",
    "halimo ahmed": "Halimo Ahmed", "halimo": "Halimo Ahmed",
    "agnes ingiara": "Agnes Ingiara", "agnes ingiara oreste": "Agnes Ingiara",
    "oweteshe mirindi": "Oweteshe Mirindi", "rahmo abdi": "Rahmo Abdi",
    "ongoro john": "Ongoro John", "ongoro john tadeo": "Ongoro John", "ongoro john tadeow": "Ongoro John",
    "abdikadir osman": "Abdikadir Osman", "salma said": "Salma Said", "abdiwakil ali": "Abdiwakil Ali",
    "manow muse": "Manow Muse", "jama mohamed": "Jama Mohamed",
    "lobono peter": "Lobono Peter", "peter lobono": "Lobono Peter", "lino lotino": "Lino Lotino",
    "magnifique ndayisenga": "Magnifique Ndayisenga", "ndayisenga magnifique": "Magnifique Ndayisenga", "magnifique": "Magnifique Ndayisenga",
    "rose akii": "Rose Akii", "adams odwa peter": "Adams Odwa Peter",
    "lokiro mazkil napao": "Lokiro Mazkil Napao", "lokiro mazkil": "Lokiro Mazkil Napao",
    "kwarto oliha": "Kwarto Oliha", "gabriella amani": "Gabriella Amani", "nyok jennifer": "Nyok Jennifer",
    "ohide akech viola": "Ohide Akech Viola",
    "21 04 2026": MISSING, "kalobeyei reception centre": MISSING, "kalobeyei reception center": MISSING,
}

ISSUE_OTHER_MAP = {
    # Documentation / registration / card issues
    "lack of card": "New arrival / Lack of card",
    "no card": "New arrival / Lack of card",
    "new arrival": "New arrival / Lack of card",
    "new arrivals": "New arrival / Lack of card",
    "card separation": "New arrival / Lack of card",
    "separation card": "New arrival / Lack of card",
    "separated from card": "New arrival / Lack of card",
    "registration": "Needs profiling / registration by UNHCR",
    "profiling": "Needs profiling / registration by UNHCR",
    "unhcr registration": "Needs profiling / registration by UNHCR",
    "needs registration": "Needs profiling / registration by UNHCR",
    "mandate": "Needs profiling / registration by UNHCR",
    "documentation": "Needs profiling / registration by UNHCR",

    # Disability / medical / basic-needs related
    "disability": "Disability",
    "disabled": "Disability",
    "child with disability": "Disability",
    "basic needs": "Basic needs",
    "nfi": "Basic needs",
    "nfis": "Basic needs",
    "material support": "Basic needs",
    "materials support": "Basic needs",
    "medical": "Basic needs",
    "medical concern": "Basic needs",
    "health": "Basic needs",
    "health issue": "Basic needs",
    "sick": "Basic needs",

    # Education
    "education": "Education",
    "school": "Education",
    "school fees": "Education",
    "school fee": "Education",
    "scholastic materials": "Education",
    "scholastic material": "Education",
    "school materials": "Education",
    "learning materials": "Education",
    "uniform": "Education",
    "books": "Education",

    # Food, clothing, shelter
    "food": "Food",
    "food assistance": "Food",
    "food issue": "Food",
    "bamba chakula": "Food",
    "bamba chakula issues": "Food",
    "alternative food collector": "Food",
    "ration": "Food",
    "clothing": "Clothing",
    "clothes": "Clothing",
    "shelter": "Shelter",
    "house": "Shelter",
    "housing": "Shelter",
    "tent": "Shelter",

    # Family / care arrangement
    "deceased parent": "Deceased parent",
    "orphan": "Deceased parent",
    "dead parent": "Deceased parent",
    "parents separated": "Parents separated",
    "parent separated": "Parents separated",
    "separated parents": "Parents separated",
    "family separation": "Parents separated",
    "unaccompanied": "Parents separated",
    "separated child": "Parents separated",
    "neglected": "Neglected",
    "neglect": "Neglected",
    "abandoned": "Neglected",
    "child out of wedlock": "Child out of wedlock",
    "out of wedlock": "Child out of wedlock",

    # Psychosocial / child engagement activities recorded as issue-other
    "pss": "Psychosocial support",
    "psychosocial": "Psychosocial support",
    "psychosocial support": "Psychosocial support",
    "counselling": "Psychosocial support",
    "counseling": "Psychosocial support",
    "stress": "Psychosocial support",
    "trauma": "Psychosocial support",
    "playing": "Psychosocial support",
    "playing with the kids": "Psychosocial support",
    "playing with other children": "Psychosocial support",
    "playing colouring book": "Psychosocial support",
    "playing and reading story book": "Psychosocial support",
    "take5": "Psychosocial support",
    "take 5": "Psychosocial support",
    "take5 training": "Psychosocial support",
    "take 5 training": "Psychosocial support",
    "take5 routine": "Psychosocial support",
    "take 5 routine": "Psychosocial support",
    "psychoeducation": "Psychosocial support",
    "psycho education": "Psychosocial support",

    # Protection concerns
    "protection": "Reporting a protection concern",
    "protection concern": "Reporting a protection concern",
    "reporting protection concern": "Reporting a protection concern",
    "abuse": "Reporting a protection concern",
    "child abuse": "Reporting a protection concern",
    "physical assault": "Reporting a protection concern",
    "assault": "Reporting a protection concern",
    "violence": "Reporting a protection concern",
    "gbv": "Reporting a protection concern",
    "gbv survivor": "Reporting a protection concern",
    "defilement": "Reporting a protection concern",
    "rape": "Reporting a protection concern",
    "abduction": "Reporting a protection concern",
    "threat of abduction": "Reporting a protection concern",
    "threat": "Reporting a protection concern",
    "child labour": "Reporting a protection concern",
    "child labor": "Reporting a protection concern",
    "early marriage": "Reporting a protection concern",
    "forced marriage": "Reporting a protection concern",

    # None / review noise
    "none": "None",
    "no issue": "None",
    "no issues": "None",
    "nil": "None",
    "n a": "None",
    "na": "None",
    "pl": REVIEW,
    "o00": REVIEW,
    "000": REVIEW,
}

CORE_ISSUE_VALUES = sorted(set(ISSUE_COLUMNS.values()))

ISSUE_KEYWORD_RULES = [
    (r"\b(new arrival|lack.*card|no card|card separation|separation card)\b", "New arrival / Lack of card"),
    (r"\b(profile|profiling|registration|register|unhcr|documentation|mandate)\b", "Needs profiling / registration by UNHCR"),
    (r"\b(disab|impair|special need)\b", "Disability"),
    (r"\b(basic need|nfi|material support|medical|health|sick|medicine|clinic|hospital)\b", "Basic needs"),
    (r"\b(school|education|scholastic|learning|uniform|book|fees?)\b", "Education"),
    (r"\b(food|bamba|chakula|ration|alternative food collector)\b", "Food"),
    (r"\b(cloth|clothes|clothing|shoe)\b", "Clothing"),
    (r"\b(shelter|house|housing|tent|accommodation)\b", "Shelter"),
    (r"\b(deceased|dead parent|orphan)\b", "Deceased parent"),
    (r"\b(parent.*separat|separat.*parent|family separation|unaccompanied|separated child)\b", "Parents separated"),
    (r"\b(neglect|abandon)\b", "Neglected"),
    (r"\b(out of wedlock|wedlock)\b", "Child out of wedlock"),
    (r"\b(pss|psychosocial|counsel|trauma|stress|play|playing|take\s*5|take5|psychoeducation)\b", "Psychosocial support"),
    (r"\b(protection|abuse|assault|violence|gbv|defile|rape|abduction|threat|labou?r|marriage|exploitation)\b", "Reporting a protection concern"),
]

LOCATION_MAP = {
    "hagadera camp": "Hagadera",
    "hagadera": "Hagadera",
    "ifo 1": "Ifo 1",
    "ifo one": "Ifo 1",
    "ifo 2": "Ifo 2",
    "ifo two": "Ifo 2",
    "dagahaley": "Dagahaley",
    "kalobeyei reception center": "Kalobeyei Reception Centre",
    "kalobeyei reception centre": "Kalobeyei Reception Centre",
    "kalobeyei village 1": "Kalobeyei Village 1",
    "kalobeyei village 2": "Kalobeyei Village 2",
    "kalobeyei village 3": "Kalobeyei Village 3",
    "kalobeyei host": "Kalobeyei Host Community",
}

CFS_MAP = {
    "kalobeyei reception center cfs": "Kalobeyei Reception Centre CFS",
    "kalobeyei reception centre cfs": "Kalobeyei Reception Centre CFS",
    "host community": "Host Community CFS",
    "host community cfs": "Host Community CFS",
    "ifo 2 mobile cfs": "Ifo 2 Mobile CFS",
    "ifo mobile cfs 1": "Ifo Mobile CFS 1",
}

GAME_MAP = {
    "take5": "Take 5", "take 5": "Take 5", "take five": "Take 5",
    "take5 training": "Take 5", "take 5 training": "Take 5", "take5 routine": "Take 5", "take 5 routine": "Take 5",
    "ismf": "I Support My Friend", "imsf": "I Support My Friend", "i support my friend": "I Support My Friend", "support my friend": "I Support My Friend",
    "psycho education": "Psychoeducation", "psychoeducation": "Psychoeducation", "psycho-education": "Psychoeducation", "pss": "Psychoeducation",
    "merry go round": "Merry-go-round", "merry-go-round": "Merry-go-round", "round about": "Merry-go-round",
    "see saw": "See-saw", "seesaw": "See-saw", "see-saw": "See-saw",
    "slider": "Slider", "slide": "Slider", "sliding": "Slider",
    "see saw and slider": "See-saw / Slider", "see saw and sliders": "See-saw / Slider", "slide and see saw": "See-saw / Slider",
    "modelling": "Modelling clay", "modeling": "Modelling clay", "modelling clay": "Modelling clay", "modeling clay": "Modelling clay", "meddling clay": "Modelling clay",
    "drawing": "Drawing / colouring", "colouring": "Drawing / colouring", "coloring": "Drawing / colouring", "painting": "Drawing / colouring", "crayons": "Drawing / colouring",
    "book reading": "Story / book reading", "reading story books": "Story / book reading", "story telling": "Story / book reading", "storytelling": "Story / book reading", "reading": "Story / book reading", "story book": "Story / book reading",
    "football": "Football", "foot ball": "Football", "soccer": "Football",
    "skipping": "Skipping rope", "skipping rope": "Skipping rope", "rope skipping": "Skipping rope",
    "singing": "Singing / music", "songs": "Singing / music", "music": "Singing / music",
    "dancing": "Dancing", "dance": "Dancing",
    "running": "Running games", "running game": "Running games", "racing": "Running games",
    "hide and seek": "Hide and seek", "hide seek": "Hide and seek",
    "playing": "Play / unspecified activity", "play": "Play / unspecified activity", "free play": "Play / unspecified activity", "indoor games": "Play / unspecified activity", "outdoor games": "Play / unspecified activity", "toys": "Play / unspecified activity",
    "none": "None", "no game": "None", "no games": "None", "nil": "None", "na": "None",
}

GAME_KEYWORD_RULES = [
    (r"\btake\s*5\b|\btake5\b|\btake five\b", "Take 5"),
    (r"\b(i support my friend|support my friend|ismf|imsf)\b", "I Support My Friend"),
    (r"\b(psycho\s*education|psychoeducation|pss)\b", "Psychoeducation"),
    (r"\bmerry\s*-?\s*go\s*-?\s*round\b|\bround about\b", "Merry-go-round"),
    (r"\bsee\s*-?\s*saw\b|\bseesaw\b", "See-saw"),
    (r"\bslider?\b|\bsliding\b", "Slider"),
    (r"\bfootball\b|\bfoot ball\b|\bsoccer\b", "Football"),
    (r"\bskipping\b|\brope skipping\b", "Skipping rope"),
    (r"\bmodelling\b|\bmodeling\b|\bclay\b", "Modelling clay"),
    (r"\bdrawing\b|\bcolou?r\w*\b|\bpainting\b|\bcrayon", "Drawing / colouring"),
    (r"\bread\w*\b|\bstory\b|\bbook", "Story / book reading"),
    (r"\bsing\w*\b|\bsongs?\b|\bmusic\b", "Singing / music"),
    (r"\bdanc\w*\b", "Dancing"),
    (r"\brunn?ing\b|\brace\b|\bracing\b", "Running games"),
    (r"\bhide\s*(and)?\s*seek\b", "Hide and seek"),
    (r"\bplay\w*\b|\btoys?\b|\bgames?\b", "Play / unspecified activity"),
]

GAME_PLACEHOLDERS = {
    "other", "others", "other option", "other specify", "specify other", "if other specify",
    "other please specify", "please specify", "specify", "other games", "other game",
    "game other specify", "games other specify", "game_other_specify", "games_played other",
}

AGENCY_MAP = {
    "dras": "DRS",
    "drs": "DRS",
    "dra": "DRS",
    "unhcr": "UNHCR",
    "uhcr": "UNHCR",
    "unchr": "UNHCR",
    "un": "UNHCR",
    "unhcr and dras": "UNHCR / DRS",
    "dras and unhcr": "UNHCR / DRS",
    "lwf": "LWF",
    "lwfl": "LWF",
    "lwf education": "LWF",
    "hi": "HI",
    "wfp": "WFP",
    "pwj": "PWJ",
    "peace wind japan": "PWJ",
    "nrc": "NRC",
    "drc": "DRC",
    "rck": "RCK",
    "krcs": "KRCS",
    "save the children": "Save the Children",
    "save children": "Save the Children",
    "000": "Unknown",
    "n": "Unknown",
}

UPPER_TOKENS = frozenset({"cfs", "tdh", "unhcr", "drs", "dras", "lwf", "hi", "wfp", "nrc", "drc", "rck", "krcs", "pss", "gbv"})
LOWER_TOKENS = frozenset({"and", "or", "of", "the", "for", "to", "in", "by", "with", "at"})

# -----------------------------------------------------------------------------
# Core helpers
# -----------------------------------------------------------------------------
def resolve_data_path() -> Path:
    if os.environ.get("CFS_DATA_PATH"):
        return Path(os.environ["CFS_DATA_PATH"])
    for file_name in DATA_FILE_CANDIDATES:
        candidate = DATA_DIR / file_name
        if candidate.exists():
            return candidate
    data_files = sorted(
        list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.xls")) + list(DATA_DIR.glob("*.csv")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if data_files:
        return data_files[0]
    return DEFAULT_DATA_PATH



def setting(name: str, default=None):
    """Read configuration from Streamlit secrets first, then environment."""
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    return os.environ.get(name, value)


def kobo_configured() -> bool:
    return bool(setting("KOBO_TOKEN") and setting("KOBO_ASSET_UID"))


def flatten_kobo_record(record: dict, prefix: str = "") -> dict:
    """Flatten nested Kobo groups while retaining their full XML paths."""
    flat = {}
    for key, value in record.items():
        full_key = f"{prefix}/{key}" if prefix else str(key)
        if isinstance(value, dict):
            flat.update(flatten_kobo_record(value, full_key))
        else:
            flat[full_key] = value
    return flat


def configured_kobo_column_map() -> dict:
    """Optional XML-name/path to dashboard-name mapping from secrets."""
    raw = setting("KOBO_COLUMN_MAP", {})
    if isinstance(raw, dict):
        return dict(raw)
    if isinstance(raw, str) and raw.strip():
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError("KOBO_COLUMN_MAP must be a JSON object") from exc
    return {}


@st.cache_data(show_spinner=False, ttl=3600, max_entries=4)
def fetch_kobo_schema_map(base_url: str, asset_uid: str, token: str) -> dict:
    """Build XML-name/path -> analysis-name mappings from Kobo form metadata.

    Kobo submission JSON uses XLSForm XML names while ordinary exports often
    show question labels. The dashboard mapping is label-aware, so use the
    asset's survey schema to connect those two representations automatically.
    """
    url = f"{base_url.rstrip('/')}/api/v2/assets/{asset_uid}/"
    headers = {"Authorization": f"Token {token}", "Accept": "application/json"}
    response = requests.get(url, headers=headers, timeout=KOBO_REQUEST_TIMEOUT_SECONDS)
    if response.status_code in {401, 403}:
        raise RuntimeError("Kobo rejected the credentials or this account cannot view the form schema.")
    response.raise_for_status()
    survey = response.json().get("content", {}).get("survey", []) or []

    stripped_map = {str(key).strip(): value for key, value in RAW_TO_TRANSFORMED_COLUMNS.items()}
    normalized_map = {norm_text(key): value for key, value in RAW_TO_TRANSFORMED_COLUMNS.items()}
    schema_map = {}

    for question in survey:
        if not isinstance(question, dict):
            continue
        name = str(question.get("name", "")).strip()
        xpath = str(question.get("$xpath", "")).strip().lstrip("/")
        label = question.get("label", "")
        if isinstance(label, list):
            label_candidates = [str(item) for item in label]
        elif isinstance(label, dict):
            label_candidates = [str(item) for item in label.values()]
        else:
            label_candidates = [str(label)]

        target = None
        for candidate in label_candidates:
            target = stripped_map.get(candidate.strip()) or normalized_map.get(norm_text(candidate))
            if target:
                break
        if not target:
            target = stripped_map.get(name) or normalized_map.get(norm_text(name))
        if not target and name in ANALYSIS_COLUMN_NAMES:
            target = name
        if not target:
            continue

        if name:
            schema_map[name] = target
        if xpath:
            schema_map[xpath] = target
            # Some Kobo deployments include a root node in $xpath while the
            # submissions endpoint omits it. Also retain the path without it.
            if "/" in xpath:
                schema_map[xpath.split("/", 1)[1]] = target

    return schema_map


@st.cache_data(show_spinner=False, ttl=3600, max_entries=4)
def fetch_kobo_support_choice_map(base_url: str, asset_uid: str, token: str) -> dict:
    """Return authoritative XML choice code -> displayed support label mappings."""
    url = f"{base_url.rstrip('/')}/api/v2/assets/{asset_uid}/"
    headers = {"Authorization": f"Token {token}", "Accept": "application/json"}
    response = requests.get(url, headers=headers, timeout=KOBO_REQUEST_TIMEOUT_SECONDS)
    response.raise_for_status()
    content = response.json().get("content", {}) or {}
    survey = content.get("survey", []) or []
    choices = content.get("choices", []) or []

    stripped_map = {str(key).strip(): value for key, value in RAW_TO_TRANSFORMED_COLUMNS.items()}
    normalized_map = {norm_text(key): value for key, value in RAW_TO_TRANSFORMED_COLUMNS.items()}
    support_lists = set()

    def label_values(label) -> List[str]:
        if isinstance(label, list):
            return [str(item) for item in label]
        if isinstance(label, dict):
            return [str(item) for item in label.values()]
        return [str(label)]

    for question in survey:
        if not isinstance(question, dict):
            continue
        target = None
        for label in label_values(question.get("label", "")):
            target = stripped_map.get(label.strip()) or normalized_map.get(norm_text(label))
            if target:
                break
        if target != "support_offered_text":
            continue
        list_name = str(question.get("select_from_list_name", "")).strip()
        question_type = str(question.get("type", "")).strip()
        if not list_name:
            match = re.search(r"select_(?:multiple|one)\s+(.+)$", question_type)
            if match:
                list_name = match.group(1).strip()
        if list_name:
            support_lists.add(list_name)

    choice_map = {}
    for choice in choices:
        if not isinstance(choice, dict):
            continue
        list_name = str(choice.get("list_name", "")).strip()
        if support_lists and list_name not in support_lists:
            continue
        code = str(choice.get("name", "")).strip()
        if not code:
            continue
        resolved = set()
        for label in label_values(choice.get("label", "")):
            resolved.update(extract_support_labels_from_text(label))
        # The displayed label is authoritative. Use the XML code only as a
        # fallback for older schemas whose choice label metadata is incomplete.
        if not resolved:
            resolved.update(extract_support_labels_from_text(code))
        if resolved:
            choice_map[code] = sorted(resolved)
            choice_map[norm_text(code)] = sorted(resolved)
    return choice_map


def translate_kobo_support_choices(value, choice_map: dict):
    """Translate a Kobo select_multiple code string into displayed labels."""
    if pd.isna(value) or not str(value).strip() or not choice_map:
        return value
    raw = str(value).strip()
    codes = [item for item in re.split(r"[\s;,|]+", raw) if item]
    labels = []
    for code in codes:
        matches = choice_map.get(code) or choice_map.get(norm_text(code))
        if matches:
            labels.extend(matches)
    return "; ".join(dict.fromkeys(labels)) if labels else value


@st.cache_data(show_spinner=False, ttl=KOBO_CACHE_TTL_SECONDS, max_entries=4)
def fetch_kobo_submissions(base_url: str, asset_uid: str, token: str, refresh_nonce: int = 0):
    """Fetch every page from Kobo KPI v2 and return raw records plus metadata."""
    del refresh_nonce  # Included only to provide an explicit cache-busting key.
    base_url = base_url.rstrip("/")
    url = f"{base_url}/api/v2/assets/{asset_uid}/data/"
    headers = {"Authorization": f"Token {token}", "Accept": "application/json"}
    records = []
    page_count = 0
    with requests.Session() as session:
        while url:
            response = session.get(url, headers=headers, timeout=KOBO_REQUEST_TIMEOUT_SECONDS)
            if response.status_code in {401, 403}:
                raise RuntimeError("Kobo rejected the credentials or this account cannot view submissions.")
            response.raise_for_status()
            payload = response.json()
            if isinstance(payload, list):  # Defensive compatibility with older responses.
                records.extend(payload)
                url = None
            else:
                records.extend(payload.get("results", []))
                url = payload.get("next")
            page_count += 1
            if page_count > 10000:
                raise RuntimeError("Kobo pagination exceeded the safety limit.")

    flattened = [flatten_kobo_record(item) for item in records]
    frame = pd.json_normalize(flattened, sep="/") if flattened else pd.DataFrame()
    # Automatically map Kobo XML names through the human-readable question
    # labels in the asset schema. Explicit configuration takes precedence.
    schema_map = fetch_kobo_schema_map(base_url, asset_uid, token)
    effective_map = {**schema_map, **configured_kobo_column_map()}
    applicable_map = {source: target for source, target in effective_map.items() if source in frame.columns}
    if applicable_map:
        frame = frame.rename(columns=applicable_map)

    # Resolve support select choices from the deployed Kobo choice list. This
    # is intentionally performed after field-name mapping so the source remains
    # a select_multiple field—not a guessed free-text classification.
    if "support_offered_text" in frame.columns:
        support_choice_map = fetch_kobo_support_choice_map(base_url, asset_uid, token)
        frame["support_offered_text"] = frame["support_offered_text"].map(
            lambda value: translate_kobo_support_choices(value, support_choice_map)
        )

    # Kobo JSON uses XLSForm XML names, often nested below group paths. If the
    # final path component is already one of our analysis names, use it safely.
    existing = set(frame.columns)
    leaf_renames = {}
    for column in frame.columns:
        leaf = str(column).split("/")[-1]
        if leaf in ANALYSIS_COLUMN_NAMES and leaf not in existing:
            leaf_renames[column] = leaf
            existing.add(leaf)
    if leaf_renames:
        frame = frame.rename(columns=leaf_renames)

    fetched_at = datetime.now(timezone.utc).isoformat()
    return frame, fetched_at, page_count


@st.cache_resource(show_spinner=False, ttl=KOBO_CACHE_TTL_SECONDS, max_entries=4)
def load_kobo_dashboard_data(base_url: str, asset_uid: str, token: str, refresh_nonce: int = 0):
    """Return one shared in-memory prepared dataset for all dashboard sessions.

    cache_resource avoids serializing and copying several DataFrames on every
    section/filter rerun, which is substantially faster for 17k+ submissions.
    The dashboard treats the returned frames as read-only.
    """
    raw_df, fetched_at, page_count = fetch_kobo_submissions(base_url, asset_uid, token, refresh_nonce)
    raw_count = int(len(raw_df))
    df, issue_long, support_long, game_long = prepare_data(raw_df)
    return raw_count, df, issue_long, support_long, game_long, fetched_at, page_count


def load_external_css(path: Path) -> None:
    if not path.exists():
        return
    try:
        css = path.read_text(encoding="utf-8")
        if css.strip():
            st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    except Exception as exc:  # pragma: no cover - visual fallback only
        st.caption(f"Custom CSS could not be loaded: {exc}")


def image_to_data_uri(path: Path) -> str:
    """Return a base64 data URI for a local image, or an empty string if absent."""
    if not path.exists():
        return ""
    suffix = path.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in {"jpg", "jpeg"} else suffix
    try:
        encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
        return f"data:image/{mime};base64,{encoded}"
    except Exception:
        return ""


def render_app_footer() -> None:
    logo_uri = image_to_data_uri(DEVELOPER_LOGO_PATH)
    logo_html = f'<img src="{logo_uri}" alt="ImpactLens Africa logo" />' if logo_uri else '<div class="developer-logo-fallback">ILA</div>'
    st.markdown(
        f"""
        <div class="app-footer">
            <div class="footer-brand">
                <div class="footer-logo">{logo_html}</div>
                <div>
                    <div class="footer-name">ImpactLens Africa</div>
                    <div class="footer-tagline">Turning Data Into Human Impact</div>
                </div>
            </div>
            <div class="footer-meta">
                <div>Developed by <strong>John Kul</strong>, MEAL Officer-Tdh</div>
                <div>{APP_VERSION}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def norm_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def smart_title(value) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return MISSING
    text = re.sub(r"\s+", " ", str(value).strip())
    parts = []
    for token in text.split():
        key = token.lower()
        if key in UPPER_TOKENS:
            parts.append("DRS" if key == "dras" else token.upper())
        elif key in LOWER_TOKENS:
            parts.append(key)
        else:
            parts.append(token.capitalize())
    return " ".join(parts)


def build_lookup(mapping: dict) -> dict:
    return {norm_text(k): v for k, v in mapping.items() if norm_text(k)}


STAFF_LOOKUP = build_lookup(STAFF_MAP)
LOCATION_LOOKUP = build_lookup(LOCATION_MAP)
CFS_LOOKUP = build_lookup(CFS_MAP)
GAME_LOOKUP = build_lookup(GAME_MAP)
ISSUE_OTHER_LOOKUP = build_lookup(ISSUE_OTHER_MAP)
AGENCY_LOOKUP = build_lookup(AGENCY_MAP)


def yes_no(value) -> str:
    if pd.isna(value):
        return MISSING
    num = pd.to_numeric(value, errors="coerce")
    if pd.notna(num):
        return "Yes" if float(num) == 1 else "No" if float(num) == 0 else MISSING
    key = norm_text(value)
    if key in {"yes", "y", "1", "true", "consented", "consent", "i consent", "i agree", "agree", "agreed", "accepted", "accept", "given"}:
        return "Yes"
    if key in {"no", "n", "0", "false", "not consented", "do not consent", "dont consent", "declined", "decline", "refused", "refuse"}:
        return "No"
    return MISSING


def is_yes(value) -> bool:
    return yes_no(value) == "Yes"


def clean_first_visit(value) -> str:
    """Clean first-visit responses that may be Yes/No or text labels."""
    yn = yes_no(value)
    if yn in {"Yes", "No"}:
        return yn
    key = norm_text(value)
    if not key:
        return MISSING

    no_patterns = [
        r"^no\b", r"\brepeat\b", r"\breturn", r"\breturning\b", r"\bvisited before\b",
        r"\bprevious", r"\balready", r"\bnot first\b", r"\bsecond\b", r"\bthird\b",
        r"\bfollow up\b", r"\bfollowup\b", r"\bsubsequent\b", r"\bagain\b",
        r"\bmore than once\b", r"\bvisited cfs before\b",
    ]
    if any(re.search(pattern, key) for pattern in no_patterns):
        return "No"

    yes_patterns = [
        r"^yes\b", r"\bfirst\b", r"\b1st\b", r"\bone time\b", r"\bnew visit\b",
        r"\bnew visitor\b", r"\bfirst time\b", r"\bfirst timer\b", r"\bnever visited\b",
    ]
    if any(re.search(pattern, key) for pattern in yes_patterns):
        return "Yes"

    return MISSING


def nonmissing_value(value) -> bool:
    if pd.isna(value):
        return False
    text = str(value).strip()
    if not text:
        return False
    return norm_text(text) not in {"", "missing unspecified", "missing", "nan", "none", "na", "n a"}


def first_visit_candidate_columns(df: pd.DataFrame) -> List[str]:
    """Find possible first-visit columns even if raw headers vary slightly."""
    candidates: List[str] = []
    for col in df.columns:
        key = norm_text(col)
        if col == "first_visit_tdh_cfs":
            candidates.insert(0, col)
            continue
        # Avoid unrelated columns but catch raw survey variants.
        if "first" in key and "visit" in key:
            candidates.append(col)
        elif "tdh" in key and "cfs" in key and "visit" in key:
            candidates.append(col)
    # preserve order and remove duplicates
    seen = set()
    out = []
    for col in candidates:
        if col not in seen:
            seen.add(col)
            out.append(col)
    return out


def combine_first_visit_source(row: pd.Series, candidate_cols: List[str]):
    for col in candidate_cols:
        if col in row.index and nonmissing_value(row.get(col)):
            return row.get(col)
    return None


def repair_first_visit_columns(data: pd.DataFrame) -> pd.DataFrame:
    """Ensure first_visit_clean is populated from the direct Yes/No field.

    This is intentionally conservative: it uses first_visit_tdh_cfs as the source
    of truth where available. If a changed XLSForm export left that field empty
    but kept a raw first-visit column, it falls back to that raw column.
    """
    if data.empty:
        return data
    out = data.copy()

    candidate_cols = []
    for preferred in ["first_visit_tdh_cfs", "first_visit_source_raw"]:
        if preferred in out.columns:
            candidate_cols.append(preferred)
    for col in out.columns:
        key = norm_text(col)
        if col not in candidate_cols and "first" in key and "visit" in key:
            candidate_cols.append(col)

    if not candidate_cols:
        return out

    def row_source(row):
        return combine_first_visit_source(row, candidate_cols)

    source = out.apply(row_source, axis=1)
    cleaned = source.map(yes_no)
    unresolved = cleaned.eq(MISSING) & source.map(nonmissing_value)
    if unresolved.any():
        cleaned.loc[unresolved] = source.loc[unresolved].map(clean_first_visit)

    # Only replace if the repair found usable Yes/No values.
    if cleaned.astype(str).isin(["Yes", "No"]).any():
        out["first_visit_source_raw"] = source
        out["first_visit_clean"] = cleaned
    return out


def clean_gender(value) -> str:
    key = norm_text(value)
    if not key:
        return MISSING
    # Kobo JSON returns XLSForm choice *names*, not always their displayed
    # labels. Substring checks therefore cover codes such as intersex_child,
    # child_transgender and gender_non_binary after norm_text normalisation.
    contains_intersex = "intersex" in key or "inter sex" in key or "intersexual" in key
    contains_transgender = "transgender" in key or "trans gender" in key
    if contains_intersex and contains_transgender:
        # Preserve a combined form choice; do not arbitrarily split one response
        # into two identities or force it into only one category.
        return "Intersex / Transgender"
    if contains_intersex:
        return "Intersex"
    if contains_transgender or key in {"trans", "tg"}:
        return "Transgender"
    if "non binary" in key or "nonbinary" in key or key in {"nb", "genderqueer"}:
        return "Non-binary"
    if key in {"girl", "girls", "female", "f", "girl child", "female child"} or key.startswith("girl "):
        return "Girls"
    if key in {"boy", "boys", "male", "m", "boy child", "male child"} or key.startswith("boy "):
        return "Boys"
    if key in {"other", "others", "other specify", "prefer to self describe", "self described"}:
        return "Other / self-described"
    return MISSING


def fuzzy_harmonize(value, lookup: dict, cutoff: float = 0.86, fallback_title: bool = True) -> str:
    key = norm_text(value)
    if not key:
        return MISSING
    if key in lookup:
        return lookup[key]
    candidates = sorted(set(lookup.values()))
    best_score, best_match = 0.0, None
    for candidate in candidates:
        score = SequenceMatcher(None, key, norm_text(candidate)).ratio()
        if score > best_score:
            best_score, best_match = score, candidate
    if best_match and best_score >= cutoff:
        return best_match
    return smart_title(value) if fallback_title else str(value)


def clean_staff(value) -> str:
    return fuzzy_harmonize(value, STAFF_LOOKUP, cutoff=0.88)


def clean_issue_other(value) -> str:
    """Map free-text issue-other responses back to the main core issue list.

    This prevents many small spelling/phrase variants from appearing as separate
    issues in the Nature of Issues Reported table.
    """
    key = norm_text(value)
    if not key:
        return MISSING

    if key in ISSUE_OTHER_LOOKUP:
        return ISSUE_OTHER_LOOKUP[key]

    for pattern, label in ISSUE_KEYWORD_RULES:
        if re.search(pattern, key):
            return label

    # Fuzzy match against both manually mapped variants and the official core issue labels.
    candidates = sorted(set(ISSUE_OTHER_LOOKUP.values()) | set(CORE_ISSUE_VALUES))
    best_score, best_match = 0.0, None
    for candidate in candidates:
        if candidate == REVIEW:
            continue
        score = SequenceMatcher(None, key, norm_text(candidate)).ratio()
        if score > best_score:
            best_score, best_match = score, candidate

    if best_match and best_score >= 0.72:
        return best_match

    # Last-resort: keep unknown protection-related free text under the core protection bucket
    # instead of creating one-off hanging issue categories.
    return "Reporting a protection concern"


def first_available(row: pd.Series, columns: Iterable[str]):
    for col in columns:
        if col in row.index:
            val = row.get(col)
            if pd.notna(val) and str(val).strip():
                return val
    return None


def extract_numeric_age(value):
    if pd.isna(value):
        return None
    num = pd.to_numeric(value, errors="coerce")
    if pd.notna(num):
        return float(num)

    # The source column contains both individual ages (for example, "7 yrs")
    # and pre-grouped responses (for example, "5-11 yrs"). Use the midpoint
    # for a stated range so it is assigned to the matching canonical band.
    text = str(value).strip().lower()
    range_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:-|–|—|to)\s*(\d+(?:\.\d+)?)", text)
    if range_match:
        lower = float(range_match.group(1))
        upper = float(range_match.group(2))
        return (lower + upper) / 2

    # Handle textual single-age values and open-ended bands such as
    # "50 yrs and above", "50+", or "aged 18 years".
    single_age_match = re.search(r"\d+(?:\.\d+)?", text)
    if single_age_match:
        return float(single_age_match.group(0))
    return None


def age_group(value) -> str:
    age = extract_numeric_age(value)
    if age is None:
        return MISSING
    if 0 <= age < 5:
        return "0-4 years"
    if 5 <= age < 12:
        return "5-11 years"
    if 12 <= age < 18:
        return "12-17 years"
    if 18 <= age < 25:
        return "18-24 years"
    if 25 <= age < 50:
        return "25-49 years"
    if age >= 50:
        return "50 years and above"
    return MISSING


def clean_referral_destination(value) -> str:
    key = norm_text(value)
    if not key:
        return MISSING
    if key == "pss" or "psychosocial" in key:
        return "PSS"
    if "case" in key:
        return "Case Management"
    if "empower" in key:
        return "Empowerment"
    if "external" in key:
        return "External referrals"
    return smart_title(value)


def clean_disability_type(value) -> str:
    key = norm_text(value)
    if not key:
        return MISSING
    if "down" in key and "syndrome" in key:
        return "Down syndrome"
    if any(word in key for word in ["neurological", "adhd", "autism", "dyslexia"]):
        return "Neurological impairments"
    if any(word in key for word in ["chronic", "diabetes", "blood pressure"]):
        return "Chronic illnesses"
    return smart_title(value)


def is_game_placeholder(value) -> bool:
    return norm_text(value) in GAME_PLACEHOLDERS


def clean_game_label(value) -> str:
    key = norm_text(value)
    if not key:
        return MISSING
    # "Other Option" is not an activity. It is only a trigger telling us to use
    # game_other_specify, so it must never appear as its own row in the table.
    if key in GAME_PLACEHOLDERS:
        return MISSING
    if key in GAME_LOOKUP:
        return GAME_LOOKUP[key]
    for pattern, label in GAME_KEYWORD_RULES:
        if re.search(pattern, key):
            return label
    harmonized = fuzzy_harmonize(value, GAME_LOOKUP, cutoff=0.78)
    return MISSING if norm_text(harmonized) in GAME_PLACEHOLDERS else harmonized


def extract_game_labels(*values) -> Set[str]:
    labels: Set[str] = set()
    real_values = [v for v in values if pd.notna(v) and not is_game_placeholder(v)]
    combined = " ".join(str(v) for v in real_values)
    key_full = norm_text(combined)
    if not key_full:
        return labels
    if re.search(r"\bsee\s*-?\s*saw\b|\bseesaw\b", key_full) and re.search(r"\bslider?\b|\bsliding\b", key_full):
        labels.add("See-saw / Slider")
    for pattern, label in GAME_KEYWORD_RULES:
        if re.search(pattern, key_full):
            labels.add(label)
    for value in real_values:
        for part in split_multi_response_text(value):
            label = clean_game_label(part)
            if label not in {MISSING, REVIEW, "None"}:
                labels.add(label)
    if "See-saw" in labels and "Slider" in labels:
        labels.discard("See-saw"); labels.discard("Slider"); labels.add("See-saw / Slider")
    labels.discard("None")
    labels.discard("Other Option")
    labels.discard("Other")
    return labels


def clean_game(row: pd.Series) -> str:
    labels = extract_game_labels(row.get("games_played", None), row.get("game_other_specify", None))
    return "; ".join(sorted(labels)) if labels else MISSING


def file_slug(value: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", str(value)).strip("_").lower()
    return slug[:90] or "dashboard_export"


def df_signature(df: pd.DataFrame) -> str:
    if df.empty:
        return "empty"
    safe = df.astype(str).fillna("")
    meta = f"{safe.shape}|{'|'.join(safe.columns)}".encode("utf-8", errors="ignore")
    hashes = pd.util.hash_pandas_object(safe, index=True).values.tobytes()
    return hashlib.md5(meta + hashes).hexdigest()[:12]


def privacy_safe_export(data: pd.DataFrame) -> pd.DataFrame:
    """Return an export frame with direct identifiers removed."""
    sensitive = [column for column in data.columns if column in PII_COLUMNS]
    return data.drop(columns=sensitive, errors="ignore").copy()


@st.cache_data(show_spinner=False, max_entries=16)
def source_content_fingerprint(path: str, file_size: int, modified_time_ns: int) -> str:
    """Return a deployment-stable identity for the source workbook.

    The size and timestamp are cache arguments so ordinary Streamlit reruns do
    not reread the file.  The returned identity is based on file content, so a
    prepared cache remains valid when a deployment changes the checkout path or
    timestamp without changing the workbook itself.
    """
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

# -----------------------------------------------------------------------------
# Data loading and preparation
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=300)
def read_file_cached(path: str, modified_time: float) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(p)
    return pd.read_csv(p)



def harmonize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise the dataset to the analysis-column schema.

    The online form can export either raw question labels, transformed analysis
    column names, or both. If both exist, the analysis column is kept and missing
    values are filled from the raw column. This prevents a blank transformed
    column from overriding a populated raw survey column.
    """
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]

    stripped_map = {str(k).strip(): v for k, v in RAW_TO_TRANSFORMED_COLUMNS.items()}
    normalized_map = {norm_text(k): v for k, v in RAW_TO_TRANSFORMED_COLUMNS.items()}

    def missing_mask(series: pd.Series) -> pd.Series:
        return (
            series.isna()
            | series.astype(str).str.strip().eq("")
            | series.astype(str).str.strip().str.lower().isin(["nan", "none", "na", "n/a", "missing / unspecified"])
        )

    # First pass: rename raw columns to analysis names when the analysis column
    # is not already present.
    rename_map = {}
    existing = set(out.columns)
    for col in list(out.columns):
        col_str = str(col).strip()
        target = stripped_map.get(col_str) or normalized_map.get(norm_text(col_str))
        if not target or col == target:
            continue
        if target not in existing:
            rename_map[col] = target
            existing.add(target)

    if rename_map:
        out = out.rename(columns=rename_map)

    # Second pass: if both raw and analysis columns exist, fill missing analysis
    # values from the raw column and then remove the duplicate raw column.
    for raw_col, target in stripped_map.items():
        raw_col = str(raw_col).strip()
        if raw_col in out.columns and target in out.columns and raw_col != target:
            mask = missing_mask(out[target])
            if mask.any():
                out.loc[mask, target] = out.loc[mask, raw_col]
            out = out.drop(columns=[raw_col])

    # Third pass: handle punctuation variants through normalized matching.
    for col in list(out.columns):
        target = normalized_map.get(norm_text(col))
        if target and target in out.columns and col != target:
            mask = missing_mask(out[target])
            if mask.any():
                out.loc[mask, target] = out.loc[mask, col]
            out = out.drop(columns=[col])
        elif target and target not in out.columns and col != target:
            out = out.rename(columns={col: target})

    # Enforce one deterministic column per analysis name. A newly introduced
    # Kobo attribute can therefore never shift fields by position or create
    # ambiguous duplicate targets. Earlier populated sources take precedence;
    # later duplicates only fill blanks.
    if out.columns.duplicated().any():
        coalesced = {}
        for position, column in enumerate(out.columns):
            source = out.iloc[:, position]
            if column not in coalesced:
                coalesced[column] = source.copy()
            else:
                mask = missing_mask(coalesced[column])
                coalesced[column] = coalesced[column].where(~mask, source)
        out = pd.DataFrame(coalesced, index=out.index)

    out.attrs["schema_contract_version"] = SCHEMA_CONTRACT_VERSION
    out.attrs["unmapped_source_columns"] = sorted(
        str(column) for column in out.columns if str(column) not in ANALYSIS_COLUMN_NAMES
    )

    return out


def ensure_column(df: pd.DataFrame, column: str, default=None) -> None:
    """Create a missing expected column so preparation does not fail."""
    if column not in df.columns:
        df[column] = default


def split_multi_response_text(value) -> List[str]:
    if pd.isna(value) or str(value).strip() == "":
        return []
    text = str(value)
    # Keep slash inside known labels by first normalizing common XLSForm separators.
    parts = re.split(r"[;|\n\r]+|,(?=\s*[A-Za-z])", text)
    return [p.strip() for p in parts if p and p.strip()]


ISSUE_TEXT_ALIASES = {
    "new arrival lack of card": "New arrival / Lack of card",
    "new arrival": "New arrival / Lack of card",
    "lack of card": "New arrival / Lack of card",
    "persons in need of profiling registration by unhcr": "Needs profiling / registration by UNHCR",
    "person in need of profiling registration by unhcr": "Needs profiling / registration by UNHCR",
    "need profiling registration unhcr": "Needs profiling / registration by UNHCR",
    "parents seperated": "Parents separated",
    "parents separated": "Parents separated",
    "psychosocial support": "Psychosocial support",
    "reporting a protection concern": "Reporting a protection concern",
}


def extract_issue_labels_from_text(value) -> Set[str]:
    labels: Set[str] = set()
    key_full = norm_text(value)
    if not key_full:
        return labels

    # Whole-cell alias and keyword checks first, because some exports keep multi-selects in one cell.
    for alias, label in ISSUE_TEXT_ALIASES.items():
        if alias in key_full:
            labels.add(label)
    for label in CORE_ISSUE_VALUES:
        if norm_text(label) in key_full:
            labels.add(label)
    for pattern, label in ISSUE_KEYWORD_RULES:
        if re.search(pattern, key_full):
            labels.add(label)

    # Then process individual separated tokens.
    for part in split_multi_response_text(value):
        cleaned = clean_issue_other(part)
        if cleaned not in {MISSING, REVIEW}:
            labels.add(cleaned)
    return labels


SUPPORT_TEXT_ALIASES = {
    "psychological first aid": "Psychological First Aid",
    "psychological firstaid": "Psychological First Aid",
    "psychological aid": "Psychological First Aid",
    "psychosocial first aid": "Psychological First Aid",
    "psycological first aid": "Psychological First Aid",
    "psychlogical first aid": "Psychological First Aid",
    "phychological first aid": "Psychological First Aid",
    "first aid": "Psychological First Aid",
    "pfa": "Psychological First Aid",
    "play and art therapy": "Play and art therapy",
    "play art therapy": "Play and art therapy",
    "art therapy": "Play and art therapy",
    "psychoeducation": "Psychoeducation",
    "psycho education": "Psychoeducation",
    "none": "None",
}


def extract_support_labels_from_text(value) -> Set[str]:
    labels: Set[str] = set()
    key_full = norm_text(value)
    if not key_full:
        return labels
    # Kobo select_multiple responses use XML choice codes and may combine them
    # in one space-separated string. Recognise both human labels and resilient
    # code/abbreviation patterns before processing individual display values.
    if re.search(r"\bpfa\b", key_full) or (
        "first aid" in key_full
        and any(token in key_full for token in ["psych", "psyc", "phych", "psychosocial"])
    ):
        labels.add("Psychological First Aid")
    for alias, label in SUPPORT_TEXT_ALIASES.items():
        if alias in key_full:
            labels.add(label)
    for part in split_multi_response_text(value):
        key = norm_text(part)
        if key in SUPPORT_TEXT_ALIASES:
            labels.add(SUPPORT_TEXT_ALIASES[key])
    return labels


def disability_type_source(row: pd.Series):
    dtype = row.get("disability_type", None)
    other = row.get("disability_type_other", None)
    if norm_text(dtype) in {"other", "others", "specify other"} and norm_text(other):
        return other
    return dtype if norm_text(dtype) else other


def prepare_data(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = harmonize_input_columns(raw_df)
    schema_extras = list(df.attrs.get("unmapped_source_columns", []))
    df.insert(0, "record_id", range(1, len(df) + 1))

    expected_defaults = {
        "today": None,
        "username": None,
        "deviceid": None,
        "phonenumber": None,
        "information_statement": None,
        "date": pd.NaT,
        "consent": "Yes",
        "staff_filling_form": MISSING,
        "child_gender": MISSING,
        "child_age": None,
        "child_living_with_disability": MISSING,
        "disability_type": None,
        "disability_type_other": None,
        "first_visit_tdh_cfs": MISSING,
        "referral_made": MISSING,
        "referral_destination": None,
        "external_referral_agency": None,
        "camp_of_information_seeking": None,
        "specific_camp_location": None,
        "camp_location_alt": None,
        "exact_registered_location": None,
        "child_friendly_space_visited": None,
        "cfs_visited": None,
        "games_played": None,
        "game_other_specify": None,
        "take5_activities_integrated": MISSING,
        "issue_other": None,
        "issue_other_specify": None,
        "nature_issues_reported_text": None,
        "support_offered_text": None,
        "child_name": None,
        "child_individual_number": None,
        "caregiver_parent_names": None,
        "caregiver_parent_phone": None,
        "end_note": None,
        "gps_location": None,
        "id": None,
        "uuid": None,
        "submission_time": None,
        "validation_status": None,
        "notes": None,
        "status": None,
        "submitted_by": None,
        "version": None,
        "tags": None,
        "root_uuid": None,
        "index": None,
    }
    for col, default in expected_defaults.items():
        ensure_column(df, col, default)

    first_visit_candidates = first_visit_candidate_columns(df)
    df["first_visit_source_raw"] = df.apply(lambda row: combine_first_visit_source(row, first_visit_candidates), axis=1)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["month"] = df["date"].dt.to_period("M").astype("string")
    df["consent_clean"] = df["consent"].map(yes_no)
    # Analyse only records where consent was given. Consent = No rows are skipped in the form.
    df = df[df["consent_clean"] == "Yes"].copy()

    df["staff_clean"] = df["staff_filling_form"].map(clean_staff)
    df["gender_clean"] = df["child_gender"].map(clean_gender)
    df["disability_status_clean"] = df["child_living_with_disability"].map(yes_no)
    # Use strict Yes/No semantics for this field: Yes = First visit, No = Repeat visit.
    df["first_visit_clean"] = df["first_visit_source_raw"].map(yes_no)
    unresolved_first_visit = df["first_visit_clean"].eq(MISSING) & df["first_visit_source_raw"].map(nonmissing_value)
    if unresolved_first_visit.any():
        df.loc[unresolved_first_visit, "first_visit_clean"] = df.loc[unresolved_first_visit, "first_visit_source_raw"].map(clean_first_visit)
    df["referral_made_clean"] = df["referral_made"].map(yes_no)
    df["referral_destination_grouped"] = df["referral_destination"].map(clean_referral_destination)
    # Fuzzy matching is relatively expensive. Most datasets repeat the same
    # locations/agencies many times, so harmonise each distinct value once.
    agency_values = df["external_referral_agency"].drop_duplicates()
    agency_map = {value: fuzzy_harmonize(value, AGENCY_LOOKUP, cutoff=0.86) for value in agency_values}
    df["external_referral_agency_clean"] = df["external_referral_agency"].map(agency_map)
    df["age_clean"] = df["child_age"].map(extract_numeric_age)
    df["age_group"] = df["child_age"].map(age_group)
    df["disability_type_source"] = df.apply(disability_type_source, axis=1)
    df["disability_type_display"] = df["disability_type_source"].map(clean_disability_type)
    df.loc[df["disability_status_clean"].astype(str) != "Yes", "disability_type_display"] = MISSING

    df["settlement_clean"] = df["camp_of_information_seeking"].map(smart_title)
    location_cols = ["camp_location_alt", "specific_camp_location", "exact_registered_location"]
    location_candidates = df[location_cols].replace(r"^\s*$", pd.NA, regex=True)
    df["location_raw"] = location_candidates.bfill(axis=1).iloc[:, 0]
    location_values = df["location_raw"].drop_duplicates()
    location_map = {value: fuzzy_harmonize(value, LOCATION_LOOKUP, cutoff=0.86) for value in location_values}
    df["location_clean"] = df["location_raw"].map(location_map)

    cfs_cols = ["child_friendly_space_visited", "cfs_visited"]
    cfs_candidates = df[cfs_cols].replace(r"^\s*$", pd.NA, regex=True)
    df["cfs_raw"] = cfs_candidates.bfill(axis=1).iloc[:, 0]
    cfs_values = df["cfs_raw"].drop_duplicates()
    cfs_map = {value: fuzzy_harmonize(value, CFS_LOOKUP, cutoff=0.86) for value in cfs_values}
    df["cfs_clean"] = df["cfs_raw"].map(cfs_map)
    df["games_played_clean"] = df.apply(clean_game, axis=1)
    df["take5_integrated_clean"] = df["take5_activities_integrated"].map(yes_no)

    for col, order in [
        ("age_group", AGE_GROUP_ORDER),
        ("gender_clean", GENDER_ORDER),
        ("first_visit_clean", YES_NO_ORDER),
        ("disability_status_clean", YES_NO_ORDER),
        ("referral_made_clean", YES_NO_ORDER),
        ("take5_integrated_clean", YES_NO_ORDER),
        ("consent_clean", YES_NO_ORDER),
    ]:
        df[col] = pd.Categorical(df[col].astype(str), categories=order, ordered=True)

    issue_frames = []
    for col, label in ISSUE_COLUMNS.items():
        if col in df.columns:
            tmp = df.loc[df[col].map(is_yes), ["record_id"]].copy()
            tmp["issue_clean"] = label
            issue_frames.append(tmp)
    if "issue_other" in df.columns and "issue_other_specify" in df.columns:
        mask = df["issue_other"].map(is_yes) & df["issue_other_specify"].notna()
        tmp = df.loc[mask, ["record_id", "issue_other_specify"]].copy()
        tmp["issue_clean"] = tmp["issue_other_specify"].map(clean_issue_other)
        tmp = tmp[~tmp["issue_clean"].isin([MISSING, REVIEW])]
        issue_frames.append(tmp[["record_id", "issue_clean"]])
    if "nature_issues_reported_text" in df.columns:
        text_rows = []
        for record_id, value in df[["record_id", "nature_issues_reported_text"]].itertuples(index=False):
            for label in extract_issue_labels_from_text(value):
                if label not in {MISSING, REVIEW}:
                    text_rows.append({"record_id": record_id, "issue_clean": label})
        if text_rows:
            issue_frames.append(pd.DataFrame(text_rows))
    issue_long = pd.concat(issue_frames, ignore_index=True).drop_duplicates() if issue_frames else pd.DataFrame(columns=["record_id", "issue_clean"])

    support_frames = []
    for col, label in SUPPORT_COLUMNS.items():
        if col in df.columns:
            tmp = df.loc[df[col].map(is_yes), ["record_id"]].copy()
            tmp["support_clean"] = label
            support_frames.append(tmp)
    if "support_offered_text" in df.columns:
        text_rows = []
        for record_id, value in df[["record_id", "support_offered_text"]].itertuples(index=False):
            for label in extract_support_labels_from_text(value):
                if label not in {MISSING, REVIEW}:
                    text_rows.append({"record_id": record_id, "support_clean": label})
        if text_rows:
            support_frames.append(pd.DataFrame(text_rows))
    support_long = pd.concat(support_frames, ignore_index=True).drop_duplicates() if support_frames else pd.DataFrame(columns=["record_id", "support_clean"])

    game_rows = []
    for record_id, raw_game, other_game in df[["record_id", "games_played", "game_other_specify"]].itertuples(index=False):
        for label in extract_game_labels(raw_game, other_game):
            if label not in {MISSING, REVIEW, "None"}:
                game_rows.append({"record_id": record_id, "game_clean": label})
    game_long = pd.DataFrame(game_rows).drop_duplicates().reset_index(drop=True) if game_rows else pd.DataFrame(columns=["record_id", "game_clean"])

    issues_combined = issue_long.groupby("record_id")["issue_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("issues_combined")
    support_combined = support_long.groupby("record_id")["support_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("support_combined")
    games_combined = game_long.groupby("record_id")["game_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("games_combined")
    df = df.merge(issues_combined, on="record_id", how="left").merge(support_combined, on="record_id", how="left").merge(games_combined, on="record_id", how="left")
    df["issues_combined"] = df["issues_combined"].fillna(MISSING)
    df["support_combined"] = df["support_combined"].fillna(MISSING)
    df["games_combined"] = df["games_combined"].fillna(MISSING)

    df.attrs["schema_contract_version"] = SCHEMA_CONTRACT_VERSION
    df.attrs["unmapped_source_columns"] = schema_extras
    return df, issue_long, support_long, game_long


@st.cache_data(show_spinner=False, persist="disk", max_entries=8)
def load_dashboard_data_cached(path: str, source_fingerprint: str) -> Tuple[int, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read and prepare the dataset in one cached step keyed by file content.

    Performance strategy:
    1. Use a persisted prepared-data pickle where possible. This avoids reading
       and re-harmonising the Excel file after app/server restarts.
    2. Use Streamlit cache for fast in-session reruns when users change filters
       or dashboard sections.
    3. Invalidate automatically when the source file content changes.
    """
    p = Path(path)
    source_path = str(p.resolve())
    source_file_name = p.name
    source_size = p.stat().st_size if p.exists() else None

    try:
        if PREPARED_DATA_PATH.exists():
            payload = pd.read_pickle(PREPARED_DATA_PATH)
            if (
                payload.get("cache_version") == PREPARED_CACHE_VERSION
                and payload.get("source_file_name") == source_file_name
                and payload.get("source_size") == source_size
                and payload.get("source_fingerprint") == source_fingerprint
            ):
                return (
                    int(payload.get("raw_count", len(payload["df"]))),
                    payload["df"],
                    payload["issue_long"],
                    payload["support_long"],
                    payload["game_long"],
                )
    except Exception:
        pass

    try:
        if PREPARED_CACHE_PATH.exists():
            payload = pd.read_pickle(PREPARED_CACHE_PATH)
            if (
                payload.get("cache_version") == PREPARED_CACHE_VERSION
                and payload.get("source_file_name") == source_file_name
                and payload.get("source_size") == source_size
                and payload.get("source_fingerprint") == source_fingerprint
            ):
                return (
                    int(payload.get("raw_count", len(payload["df"]))),
                    payload["df"],
                    payload["issue_long"],
                    payload["support_long"],
                    payload["game_long"],
                )
    except Exception:
        try:
            PREPARED_CACHE_PATH.unlink(missing_ok=True)
        except Exception:
            pass

    if p.suffix.lower() in {".xlsx", ".xls"}:
        raw_df = pd.read_excel(p)
    else:
        raw_df = pd.read_csv(p)

    raw_count = int(len(raw_df))
    df, issue_long, support_long, game_long = prepare_data(raw_df)
    payload = {
        "cache_version": PREPARED_CACHE_VERSION,
        "source_file_name": source_file_name,
        "source_path": source_path,
        "source_size": source_size,
        "source_fingerprint": source_fingerprint,
        "raw_count": raw_count,
        "df": df,
        "issue_long": issue_long,
        "support_long": support_long,
        "game_long": game_long,
    }

    try:
        pd.to_pickle(payload, PREPARED_CACHE_PATH)
        if Path(path).resolve().parent == DATA_DIR.resolve():
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            pd.to_pickle(payload, PREPARED_DATA_PATH)
    except Exception:
        pass

    return raw_count, df, issue_long, support_long, game_long

# -----------------------------------------------------------------------------
# Table and chart utilities
# -----------------------------------------------------------------------------
def ordered_unique(series: pd.Series, order: Optional[List] = None) -> list:
    values = [v for v in series.dropna().astype(str).unique().tolist() if str(v).strip()]
    # Missing/unspecified is intentionally excluded from user-facing dashboard filters.
    # Consent-refused forms are removed, and blank skipped fields should not drive analysis views.
    values = [v for v in values if v != MISSING]
    if order:
        order_map = {str(v): i for i, v in enumerate(order)}
        return sorted(values, key=lambda v: (order_map.get(str(v), 999), str(v)))
    return sorted(values, key=lambda v: str(v))


def remove_missing_dimension_values(data: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """Remove Missing/unspecified values from analytical dimensions.

    Missing values are still counted in the Data Quality section, but they are
    not shown as categories in dashboard analysis tables/charts.
    """
    if data.empty:
        return data
    out = data.copy()
    for col in columns:
        if col in out.columns:
            out = out[out[col].notna()]
            out = out[out[col].astype(str).str.strip().ne("")]
            out = out[out[col].astype(str).ne(MISSING)]
            out = out[out[col].astype(str).ne(REVIEW)]
    return out


def drop_empty_missing_axes(table: pd.DataFrame) -> pd.DataFrame:
    """Safety net to remove zero-total Missing/unspecified rows/columns."""
    if table.empty:
        return table
    out = table.copy()

    def label_has_missing(label) -> bool:
        if isinstance(label, tuple):
            return any(str(x) == MISSING for x in label)
        return str(label) == MISSING

    keep_rows = []
    for idx, row in out.iterrows():
        numeric_sum = pd.to_numeric(row, errors="coerce").fillna(0).sum()
        keep_rows.append(not label_has_missing(idx) and numeric_sum > 0)
    out = out.loc[keep_rows]

    keep_cols = []
    for col in out.columns:
        numeric_sum = pd.to_numeric(out[col], errors="coerce").fillna(0).sum()
        keep_cols.append(not label_has_missing(col) and numeric_sum > 0)
    out = out.loc[:, keep_cols]
    return out


def count_table(data: pd.DataFrame, columns: List[str], order_col: Optional[str] = None) -> pd.DataFrame:
    if data.empty or not all(c in data.columns for c in columns):
        return pd.DataFrame(columns=columns + ["Count"])
    clean = remove_missing_dimension_values(data, columns)
    if clean.empty:
        return pd.DataFrame(columns=columns + ["Count"])
    out = clean.groupby(columns, observed=True).size().reset_index(name="Count")
    out = out[out["Count"] > 0]
    if order_col and order_col in out.columns:
        if order_col == "age_group":
            out[order_col] = pd.Categorical(out[order_col].astype(str), categories=[v for v in AGE_GROUP_ORDER if v != MISSING], ordered=True)
        out = out.sort_values([order_col, "Count"], ascending=[True, False])
    else:
        out = out.sort_values("Count", ascending=False)
    return out.reset_index(drop=True)


def multi_response_tables(long_data: pd.DataFrame, value_col: str, label: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Show both unique records and total mentions for multi-response fields.

    The main issue/activity tables count mentions. These supporting tables now
    show the same language by including a Total mentions column, while also
    showing how many unique records/children produced those mentions.
    """
    dist_cols = ["Selection count", "Records", "Total mentions"]
    combo_cols = [f"{label} combination", "Records", "Total mentions"]
    if long_data.empty or value_col not in long_data.columns or "record_id" not in long_data.columns:
        return pd.DataFrame(columns=dist_cols), pd.DataFrame(columns=combo_cols)

    # Use the same analytical exclusions as the main by-gender tables so totals reconcile.
    required_cols = [value_col] + (["gender_clean"] if "gender_clean" in long_data.columns else [])
    analytical = remove_missing_dimension_values(long_data, required_cols)
    cleaned = analytical[["record_id", value_col]].dropna().drop_duplicates().copy()
    cleaned = cleaned[~cleaned[value_col].astype(str).isin([MISSING, REVIEW, "None"])]
    if cleaned.empty:
        return pd.DataFrame(columns=dist_cols), pd.DataFrame(columns=combo_cols)

    per_record = cleaned.groupby("record_id", observed=True)[value_col].agg(lambda vals: sorted(set(str(v) for v in vals if str(v).strip())))
    selection_counts = per_record.map(len)
    total_records = int(selection_counts.shape[0])
    total_mentions = int(selection_counts.sum())

    dist = selection_counts.value_counts().sort_index().rename_axis("selection_number").reset_index(name="Records")
    dist["Total mentions"] = dist["selection_number"].astype(int) * dist["Records"].astype(int)
    dist["Selection count"] = dist["selection_number"].map(lambda c: "1 selection" if int(c) == 1 else f"{int(c)} selections")
    dist = dist[["Selection count", "Records", "Total mentions"]]
    total_row = pd.DataFrame([{"Selection count": "Grand Total", "Records": total_records, "Total mentions": total_mentions}])
    dist = pd.concat([dist, total_row], ignore_index=True)

    combo_rows = []
    for values in per_record:
        combo = " + ".join(values)
        combo_rows.append({f"{label} combination": combo, "Selection number": len(values)})
    combo_src = pd.DataFrame(combo_rows)
    combos = combo_src.groupby([f"{label} combination", "Selection number"], observed=True).size().reset_index(name="Records")
    combos["Total mentions"] = combos["Selection number"].astype(int) * combos["Records"].astype(int)
    combos = combos.drop(columns=["Selection number"]).sort_values(["Records", "Total mentions"], ascending=False).head(30).reset_index(drop=True)
    combo_total = pd.DataFrame([{f"{label} combination": "Grand Total", "Records": total_records, "Total mentions": total_mentions}])
    combos = pd.concat([combos, combo_total], ignore_index=True)
    return dist, combos


def table_with_total(data: pd.DataFrame, index: List[str], columns: Optional[List[str]] = None, value_col: str = "record_id") -> pd.DataFrame:
    if data.empty or not all(c in data.columns for c in index):
        return pd.DataFrame()
    group_cols = index + (columns or [])
    if not all(c in data.columns for c in group_cols):
        return pd.DataFrame()

    clean = remove_missing_dimension_values(data, group_cols)
    if clean.empty:
        return pd.DataFrame()

    grouped = clean.groupby(group_cols, observed=True)[value_col].count()
    table = grouped.unstack(columns, fill_value=0) if columns else grouped.to_frame("Count")
    if len(index) == 1:
        table.index.name = index[0]
    table = drop_empty_missing_axes(table)
    if table.empty:
        return table

    if isinstance(table.columns, pd.MultiIndex):
        table[("Total",) + ("",) * (table.columns.nlevels - 1)] = table.sum(axis=1)
        total_col = ("Total",) + ("",) * (table.columns.nlevels - 1)
    elif "Count" in table.columns and len(table.columns) == 1:
        total_col = "Count"
    else:
        table["Total"] = table.sum(axis=1)
        total_col = "Total"
    table = table.sort_values(total_col, ascending=False)
    table.loc["Grand Total"] = table.sum(numeric_only=True)
    return table


FRIENDLY_COLUMN_NAMES = {
    "settlement_clean": "Camp Name",
    "location_clean": "Specific Camp Location",
    "cfs_clean": "CFS / Site",
    "staff_clean": "Staff / CPV",
    "gender_clean": "Gender",
    "age_group": "Age Group",
    "issue_clean": "Issue",
    "support_clean": "Support Type",
    "game_clean": "Game / Activity",
    "take5_integrated_clean": "Take 5 Integrated",
    "visit_type": "Visit Type",
    "first_visit_clean": "First Visit",
    "disability_status_clean": "Disability Status",
    "disability_type_display": "Type of Disability",
    "referral_destination_grouped": "Referral Destination",
    "external_referral_agency_clean": "External Referral Agency",
    "month_label": "Month",
}


def friendly_column_label(name: Optional[str], title: str = "") -> str:
    """Return a user-facing label; never return the generic word Category."""
    raw = "" if name is None else str(name).strip()
    if raw in FRIENDLY_COLUMN_NAMES:
        return FRIENDLY_COLUMN_NAMES[raw]
    if raw and raw.lower() not in {"category", "index", "none"}:
        # If the label is already presentation-friendly, keep its capitalization
        # e.g. "Staff / CPV" should not become "Staff / Cpv".
        if any(ch.isupper() for ch in raw) or "/" in raw:
            return raw.replace("_", " ")
        return raw.replace("_", " ").title()

    t = title.lower()
    if "camp" in t:
        return "Camp / Settlement"
    if "cfs" in t or "site" in t:
        return "CFS / Site"
    if "staff" in t or "cpv" in t:
        return "Staff / CPV"
    if "age" in t:
        return "Age Group"
    if "gender" in t:
        return "Gender"
    if "issue" in t:
        return "Issue"
    if "support" in t:
        return "Support Type"
    if "game" in t or "activit" in t:
        return "Game / Activity"
    if "disability" in t:
        return "Type of Disability"
    if "referral" in t and "agency" in t:
        return "External Referral Agency"
    if "referral" in t:
        return "Referral Destination"
    return "Item"


def add_display_numbering(display: pd.DataFrame, label_col: Optional[str] = None) -> pd.DataFrame:
    """Insert a separate 1-based No. column.

    Row labels remain in their own descriptive column; numbering is no longer
    joined to the row value.
    """
    out = display.copy().reset_index(drop=True)
    serials = []
    n = 1
    for _, row in out.iterrows():
        is_total = any(str(v) == "Grand Total" for v in row.values)
        if is_total:
            serials.append("")
        else:
            serials.append(n)
            n += 1
    out.insert(0, "No.", serials)
    return out


def flatten_table(table: pd.DataFrame, title: str = "") -> pd.DataFrame:
    """Flatten tables for CSV export using clean 1-based row numbering."""
    flat = table.copy()
    row_label_col = None

    if isinstance(flat.columns, pd.MultiIndex):
        flat.columns = [" | ".join(str(p) for p in col if str(p) != "") for col in flat.columns]
    else:
        flat.columns = [friendly_column_label(str(c), title) for c in flat.columns]

    # Do not render an unnamed pandas integer index as a table column. After sorting,
    # pandas often keeps original row numbers, but they are not analytical data.
    default_range = (
        isinstance(flat.index, pd.RangeIndex)
        or (flat.index.name is None and not isinstance(flat.index, pd.MultiIndex))
    )
    if not default_range:
        if isinstance(flat.index, pd.MultiIndex):
            index_values = [" | ".join(str(p) for p in idx if str(p) != "") for idx in flat.index]
            index_names = [friendly_column_label(str(n), title) for n in flat.index.names if n]
            index_name = " | ".join(index_names) if index_names else friendly_column_label(None, title)
        else:
            index_values = [str(v) for v in flat.index]
            index_name = friendly_column_label(str(flat.index.name) if flat.index.name else None, title)
        flat.insert(0, index_name, index_values)
        row_label_col = index_name

    return add_display_numbering(flat, label_col=row_label_col)


def table_display_frame(table: pd.DataFrame, title: str = "") -> pd.DataFrame:
    """Convert any normal/MultiIndex dataframe into a clean display dataframe.

    There is no separate generic Category column. For indexed summary tables,
    the 1-based row number is merged into the descriptive row-label column.
    """
    display = table.copy()
    row_label_col = None

    if isinstance(display.columns, pd.MultiIndex):
        display.columns = [" | ".join(str(p) for p in col if str(p) != "") for col in display.columns]
    else:
        display.columns = [friendly_column_label(str(c), title) for c in display.columns]

    # Do not render an unnamed pandas integer index as a table column. After sorting,
    # pandas often keeps original row numbers, but they are not analytical data.
    default_range = (
        isinstance(display.index, pd.RangeIndex)
        or (display.index.name is None and not isinstance(display.index, pd.MultiIndex))
    )
    if not default_range:
        if isinstance(display.index, pd.MultiIndex):
            index_values = [" | ".join(str(p) for p in idx if str(p) != "") for idx in display.index]
            index_names = [friendly_column_label(str(n), title) for n in display.index.names if n]
            index_name = " | ".join(index_names) if index_names else friendly_column_label(None, title)
        else:
            index_values = [str(v) for v in display.index]
            index_name = friendly_column_label(str(display.index.name) if display.index.name else None, title)
        display.insert(0, index_name, index_values)
        row_label_col = index_name

    return add_display_numbering(display, label_col=row_label_col)


def is_percentage_column(column_name: str) -> bool:
    name = str(column_name).lower()
    return "%" in name or "percent" in name or "percentage" in name or "rate" in name


def format_table_value(value, column_name: str = "", precision: int = 0) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            number = float(value)
            if is_percentage_column(column_name):
                return f"{number:,.1f}"
            return f"{number:,.0f}"
        except Exception:
            return str(value)
    return str(value)


def build_professional_table_html(table: pd.DataFrame, title: str, precision: int = 0) -> Tuple[str, int]:
    """Build a self-contained, iframe-safe professional HTML table.

    This avoids Streamlit's dataframe styling limitations and makes the table
    look the same for all users in the deployed dashboard.
    """
    display = table_display_frame(table, title=title)
    row_count, col_count = display.shape

    numeric_cols = []
    max_values = {}
    for col in display.columns:
        numeric = pd.to_numeric(display[col], errors="coerce")
        if numeric.notna().sum() and not str(col).lower() in {"record_id", "id", "no.", "no", "#"}:
            numeric_cols.append(col)
            max_values[col] = max(float(numeric.max()), 0.0)

    thead = "".join(f"<th>{html_lib.escape(str(col))}</th>" for col in display.columns)

    body_rows = []
    for i, row in display.iterrows():
        row_values = [str(v) for v in row.values]
        is_total = any(v == "Grand Total" for v in row_values)
        row_class = "total-row" if is_total else ("even-row" if i % 2 else "odd-row")
        cells = []
        for j, col in enumerate(display.columns):
            value = row[col]
            value_text = html_lib.escape(format_table_value(value, column_name=str(col), precision=precision))
            classes = []
            style_parts = []

            if j == 0:
                classes.append("first-col")

            if col in numeric_cols:
                classes.append("num-cell")
                numeric_value = pd.to_numeric(value, errors="coerce")

            col_name = str(col).lower()
            if "total" in col_name or "rate" in col_name or "%" in col_name:
                classes.append("emphasis-col")

            class_attr = f" class='{' '.join(classes)}'" if classes else ""
            style_attr = f" style='{''.join(style_parts)}'" if style_parts else ""
            cells.append(f"<td{class_attr}{style_attr}>{value_text}</td>")
        body_rows.append(f"<tr class='{row_class}'>" + "".join(cells) + "</tr>")

    tbody = "".join(body_rows)
    height = min(820, max(360, 164 + min(row_count, 12) * 42))
    scroll_height = max(220, height - 135)

    html_doc = f"""
    <!doctype html>
    <html>
    <head>
      <meta charset="utf-8" />
      <style>
        :root {{
          --primary:#1d4ed8; --primary-dark:#172554; --border:#d9e2ef;
          --soft:#f8fafc; --text:#0f172a; --muted:#64748b;
        }}
        * {{ box-sizing:border-box; }}
        body {{ margin:0; background:transparent; font-family:Inter, Segoe UI, Roboto, Arial, sans-serif; color:var(--text); }}
        .table-card {{
          width:100%; background:#fff; border:1px solid var(--border); border-radius:16px;
          box-shadow:0 12px 30px rgba(15,23,42,.08); overflow:hidden;
        }}
        .table-card-header {{
          display:flex; justify-content:space-between; align-items:center; gap:1rem;
          padding:1rem 1.15rem; background:linear-gradient(135deg,#fff 0%,#eff6ff 100%);
          border-bottom:1px solid var(--border);
        }}
        .table-title {{ color:var(--primary-dark); font-size:1.05rem; font-weight:950; letter-spacing:-.02em; }}
        .table-subtitle {{ margin-top:.2rem; color:var(--muted); font-size:.82rem; font-weight:650; }}
        .table-scroll {{ max-height:{scroll_height}px; overflow:auto; background:#fff; }}
        table {{ width:max-content; min-width:100%; table-layout:auto; border-collapse:separate; border-spacing:0; font-size:.86rem; line-height:1.35; }}
        thead th {{
          position:sticky; top:0; z-index:5; padding:11px 13px; text-align:left; white-space:nowrap;
          color:#fff; font-weight:950; background:linear-gradient(180deg,#1e40af,#172554);
          border-right:1px solid #1e3a8a; border-bottom:1px solid #1e3a8a;
        }}
        thead th:first-child {{ left:0; z-index:7; }}
        tbody td {{
          padding:10px 13px; border-right:1px solid #e5edf5; border-bottom:1px solid #e5edf5;
          color:#1f2937; vertical-align:middle; background:#fff; white-space:nowrap;
        }}
        tbody tr.even-row td {{ background-color:#fbfdff; }}
        tbody tr:hover td {{ background-color:#eff6ff !important; }}
        tbody td.first-col {{
          position:sticky; left:0; z-index:2; min-width:max-content;
          color:#172554; font-weight:850; background:#f8fafc; box-shadow:1px 0 0 #e5edf5;
        }}
        tbody tr:hover td.first-col {{ background:#eaf2ff !important; }}
        td.num-cell {{ text-align:right; font-variant-numeric:tabular-nums; font-weight:750; }}
        td.emphasis-col {{ color:#172554; font-weight:900; }}
        tr.total-row td {{
          background:#eef2ff !important; color:#0f172a; font-weight:950;
          border-top:2px solid #1d4ed8; border-bottom:2px solid #1d4ed8;
        }}
        tr.total-row td.first-col {{ background:#dbeafe !important; color:#172554; }}
        @media (max-width:900px) {{
          .table-card-header {{ flex-direction:column; align-items:flex-start; }}
          table {{ min-width:100%; font-size:.82rem; }}
          tbody td, thead th {{ padding:9px 11px; }}
        }}
      </style>
    </head>
    <body>
      <div class="table-card">
        <div class="table-card-header">
          <div>
            <div class="table-title">{html_lib.escape(title)}</div>
            <div class="table-subtitle">{row_count:,} rows · {col_count:,} columns · current filters applied</div>
          </div>
        </div>
        <div class="table-scroll">
          <table>
            <thead><tr>{thead}</tr></thead>
            <tbody>{tbody}</tbody>
          </table>
        </div>
      </div>
    </body>
    </html>
    """
    return html_doc, height


def render_table(table: pd.DataFrame, title: str, key: str, precision: int = 0) -> None:
    if table.empty:
        st.info("No records available for this table.")
        return

    flat, sig, html_doc, height, csv_data = cached_table_assets(table, title, precision)
    components.html(html_doc, height=height, scrolling=False)

    st.download_button(
        "⬇ Download table as CSV",
        data=csv_data,
        file_name=f"{file_slug(title)}.csv",
        mime="text/csv",
        key=f"download_{key}_{sig}",
        use_container_width=True,
    )


@st.cache_data(show_spinner=False, max_entries=128)
def cached_table_assets(table: pd.DataFrame, title: str, precision: int) -> Tuple[pd.DataFrame, str, str, int, bytes]:
    """Reuse table HTML, signature, and CSV bytes across unchanged reruns."""
    flat = flatten_table(table, title=title)
    sig = df_signature(flat)
    html_doc, height = build_professional_table_html(table, title=title, precision=precision)
    csv_data = flat.to_csv(index=False).encode("utf-8")
    return flat, sig, html_doc, height, csv_data


def add_top_n_control(data: pd.DataFrame, category_col: str, key: str, default: int = 10) -> pd.DataFrame:
    if data.empty or category_col not in data.columns or data[category_col].nunique(dropna=True) <= min(TOP_N_OPTIONS):
        return data
    c1, c2 = st.columns([1, 1])
    with c1:
        mode = st.radio("Rank", ["Highest", "Lowest"], horizontal=True, key=f"rank_{key}")
    with c2:
        n = st.selectbox("Number of categories", TOP_N_OPTIONS, index=TOP_N_OPTIONS.index(default) if default in TOP_N_OPTIONS else 1, key=f"topn_{key}")
    totals = data.groupby(category_col, observed=False)["Count"].sum().sort_values(ascending=(mode == "Lowest"))
    return data[data[category_col].isin(totals.head(n).index)].copy()


def horizontal_category_order(data: pd.DataFrame, category_col: str, value_col: str = "Count") -> List[str]:
    if data.empty or category_col not in data.columns or value_col not in data.columns:
        return []
    totals = (
        data.groupby(category_col, observed=False)[value_col]
        .sum()
        .sort_values(ascending=False)
    )
    # Plotly draws y-axis category arrays from bottom to top, so reverse the
    # descending list to keep the largest visible at the top of the chart.
    return [str(v) for v in reversed(totals.index.tolist())]


def bar_chart(data: pd.DataFrame, x: str, y: str, color: Optional[str] = None, title: str = "", horizontal: bool = False, height: int = 420, category_orders: Optional[Dict] = None) -> None:
    if data.empty:
        st.info("No records available for this chart.")
        return
    fig = px.bar(
        data,
        x=x,
        y=y,
        color=color,
        text="Count" if "Count" in data.columns else None,
        orientation="h" if horizontal else "v",
        title=title,
        color_discrete_sequence=CHART_COLORS,
        category_orders=category_orders or {},
    )
    fig.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
    fig.update_layout(
        height=height,
        margin=dict(l=12, r=30, t=58, b=24),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend_title_text="",
        title_font=dict(size=18, color="#172554"),
        font=dict(size=13, color="#111827"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#e9eef5", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    if horizontal:
        order = horizontal_category_order(data, y, "Count" if "Count" in data.columns else x)
        if order:
            fig.update_yaxes(categoryorder="array", categoryarray=order)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True, "responsive": True})


def pie_chart(data: pd.DataFrame, names: str, values: str, title: str) -> None:
    if data.empty:
        st.info("No records available for this chart.")
        return
    fig = px.pie(data, names=names, values=values, hole=0.42, title=title, color_discrete_sequence=CHART_COLORS)
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(height=430, margin=dict(l=12, r=24, t=58, b=18), paper_bgcolor="white", title_font=dict(size=18, color="#172554"), font=dict(size=13, color="#111827"), legend_title_text="")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True, "responsive": True})


def line_chart(data: pd.DataFrame, x: str, y: str, color: Optional[str] = None, title: str = "", height: int = 380) -> None:
    if data.empty:
        st.info("No records available for this chart.")
        return
    fig = px.line(data, x=x, y=y, color=color, markers=True, title=title, color_discrete_sequence=CHART_COLORS)
    fig.update_traces(line_shape="spline", line_smoothing=0.5)
    fig.update_layout(height=height, margin=dict(l=12, r=24, t=58, b=18), plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", title_font=dict(size=18, color="#172554"), font=dict(size=13, color="#111827"))
    fig.update_xaxes(showgrid=True, gridcolor="#e9eef5", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#e9eef5", zeroline=False)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True, "responsive": True})

# -----------------------------------------------------------------------------
# UI helpers
# -----------------------------------------------------------------------------
def page_header(title: str, subtitle: Optional[str] = None) -> None:
    """Render the title panel with the Tdh logo inside the panel."""
    logo_uri = image_to_data_uri(LOGO_PATH)
    logo_html = (
        f'<img src="{logo_uri}" alt="Tdh logo" '
        'style="width:clamp(76px,9vw,126px);height:auto;object-fit:contain;flex:0 0 auto;" />'
        if logo_uri
        else ""
    )
    subtitle_html = f"<p>{html_lib.escape(subtitle)}</p>" if subtitle else ""
    st.markdown(
        f"""
        <div class="dashboard-title" style="display:flex;align-items:center;justify-content:space-between;gap:1.5rem;">
            <div style="min-width:0;">
                <h1>{html_lib.escape(title)}</h1>
                {subtitle_html}
            </div>
            {logo_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: Optional[str] = None) -> None:
    st.markdown(f"<div class='section-heading'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='section-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def metric_card(label, value, helper=None, tone="primary") -> None:
    tone_map = {"primary": "#1d4ed8", "success": "#15803d", "warning": "#d97706", "danger": "#dc2626", "neutral": "#64748b"}
    border = tone_map.get(tone, tone_map["primary"])
    helper_html = f"<div class='metric-helper'>{helper}</div>" if helper else ""
    st.markdown(f"<div class='metric-card' style='border-top:4px solid {border};'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div>{helper_html}</div>", unsafe_allow_html=True)


def insight_card(label, value, helper=None) -> None:
    helper_html = f"<div class='insight-helper'>{helper}</div>" if helper else ""
    st.markdown(f"<div class='insight-card'><div class='insight-label'>{label}</div><div class='insight-value'>{value}</div>{helper_html}</div>", unsafe_allow_html=True)


SECTION_META = {
    "Overview": {"icon": "🏠", "label": "Overview", "desc": "Coverage, age profile, top issues and top activities for the active filters."},
    "Monthly Trends": {"icon": "📈", "label": "Trends Over Time", "desc": "Use this section to follow monthly visits, gender patterns, referrals and first visits."},
    "CPVs KPIs": {"icon": "👥", "label": "Staff / CPV Performance", "desc": "Use this section to review staff activity, submissions, referrals and operating coverage."},
    "Demographics": {"icon": "🧒", "label": "Demographics", "desc": "First vs repeat visits, gender, age group and site-level beneficiary profile."},
    "Games & Activities": {"icon": "🎲", "label": "Games & Activities", "desc": "Harmonised games, activities, multiple-activity combinations and engagement patterns."},
    "Protection & Support": {"icon": "🛡️", "label": "Protection Needs & Support", "desc": "Use this section to understand disability, reported issues and support provided."},
    "Referrals": {"icon": "🔁", "label": "Referrals", "desc": "Referral rates, destinations and external referral agency breakdowns."},
    "Data Quality": {"icon": "✅", "label": "Data Quality", "desc": "Completeness checks and harmonisation review for operational data quality."},
    "Raw Data": {"icon": "📄", "label": "Records & Export", "desc": "Use this section to inspect filtered records and prepare a privacy-safe export."},
}

SECTION_OPTIONS = ["Overview", "Monthly Trends", "CPVs KPIs", "Demographics", "Games & Activities", "Protection & Support", "Referrals", "Data Quality", "Raw Data"]
SECTION_CATEGORY = {
    "Overview": "Executive summary",
    "Monthly Trends": "Time and trends",
    "CPVs KPIs": "Workforce performance",
    "Demographics": "Child profile",
    "Games & Activities": "Participation and activities",
    "Protection & Support": "Protection needs and response",
    "Referrals": "Referral pathway",
    "Data Quality": "Data governance",
    "Raw Data": "Record review and export",
}


def section_label(option: str) -> str:
    meta = SECTION_META.get(option, {"icon": "📌", "label": option})
    return f"{meta['icon']} {meta['label']}"


def nav_menu(options: List[str], key: str = "dashboard_section") -> str:
    if key not in st.session_state or st.session_state[key] not in options:
        st.session_state[key] = options[0]

    return st.selectbox(
        "**Dashboard section**",
        options,
        index=options.index(st.session_state[key]),
        key=key,
        format_func=section_label,
        help="Choose the analysis you need. Active filters remain in place.",
    )


def go_to_overview() -> None:
    """Navigation callback executed before Streamlit recreates the widgets."""
    st.session_state["dashboard_section"] = "Overview"


def section_intro_card(section: str) -> None:
    meta = SECTION_META.get(section, {"icon": "📌", "label": section, "desc": "Explore the selected analytical view."})
    st.markdown(
        f"""
        <div class="section-intro-card">
            <div class="section-intro-icon">{meta['icon']}</div>
            <div>
                <div class="section-intro-title">{meta['label']}</div>
                <div class="section-intro-desc">{meta['desc']}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def selector_control(label: str, options: List[str], key: str, order: Optional[List] = None) -> List[str]:
    """Professional All/Custom selector used inside the sidebar filter panel."""
    mode_key = f"filter_mode_{key}"
    values_key = f"filter_values_{key}"

    if mode_key not in st.session_state:
        st.session_state[mode_key] = "All"
    if values_key not in st.session_state:
        st.session_state[values_key] = []

    # Remove stale selections when source data changes.
    st.session_state[values_key] = [v for v in st.session_state[values_key] if v in options]

    if not options:
        st.caption("No values available after previous filters.")
        return []

    if hasattr(st, "segmented_control"):
        try:
            mode = st.segmented_control("Scope", ["All", "Custom"], key=mode_key, label_visibility="collapsed")
        except Exception:
            mode = st.radio("Scope", ["All", "Custom"], horizontal=True, key=mode_key, label_visibility="collapsed")
    else:
        mode = st.radio("Scope", ["All", "Custom"], horizontal=True, key=mode_key, label_visibility="collapsed")

    if mode == "All":
        st.markdown(
            f"<div class='filter-all-note'>All <b>{len(options):,}</b> available values are included.</div>",
            unsafe_allow_html=True,
        )
        return []

    if hasattr(st, "pills") and len(options) <= 12:
        try:
            selected = st.pills(
                "Choose values",
                options,
                selection_mode="multi",
                key=values_key,
                label_visibility="collapsed",
            ) or []
        except Exception:
            selected = st.multiselect("Choose values", options, key=values_key, placeholder="Type to search", label_visibility="collapsed")
    else:
        selected = st.multiselect("Choose values", options, key=values_key, placeholder="Type to search", label_visibility="collapsed")

    st.caption(f"{len(selected):,} selected" if selected else "Custom mode: select one or more values.")
    return selected


def multiselect_filter(label: str, data: pd.DataFrame, column: str, order: Optional[List] = None) -> List[str]:
    options = ordered_unique(data[column], order=order) if column in data.columns else []
    key = re.sub(r"[^a-zA-Z0-9]+", "_", column).strip("_").lower()
    with st.sidebar.expander(label, expanded=False):
        available_records = len(data) if column in data.columns else 0
        st.markdown(
            f"<div class='filter-meta'><span>{len(options):,} values</span><span>{available_records:,} records in scope</span></div>",
            unsafe_allow_html=True,
        )
        return selector_control(label, options, key=key, order=order)


def filter_if_selected(data: pd.DataFrame, column: str, selected: List[str]) -> pd.DataFrame:
    if not selected or column not in data.columns:
        return data
    return data[data[column].astype(str).isin([str(v) for v in selected])].copy()


def top_category(data: pd.DataFrame, column: str, exclude: Optional[List] = None) -> Tuple[str, int]:
    if data.empty or column not in data.columns:
        return "N/A", 0
    s = data[column].astype(str)
    if exclude:
        s = s[~s.isin(set(exclude))]
    counts = s.value_counts(dropna=True)
    return (str(counts.index[0]), int(counts.iloc[0])) if not counts.empty else ("N/A", 0)


def yes_count(data: pd.DataFrame, column: str) -> int:
    return int(data[column].astype(str).eq("Yes").sum()) if column in data.columns else 0


def build_section_takeaways(
    section: str,
    records: pd.DataFrame,
    issues: pd.DataFrame,
    support: pd.DataFrame,
    games: pd.DataFrame,
) -> str:
    """Create one concise, auditable narrative for the selected section."""
    total = len(records)
    if total == 0:
        return "No records match the current filters, so no descriptive findings can be generated."

    findings: List[str] = []

    def leading(data: pd.DataFrame, column: str, excluded: Optional[Set[str]] = None) -> Tuple[str, int]:
        if data.empty or column not in data.columns:
            return "", 0
        values = data[column].dropna().astype(str)
        if excluded:
            values = values[~values.isin(excluded)]
        counts = values.value_counts()
        return (str(counts.index[0]), int(counts.iloc[0])) if not counts.empty else ("", 0)

    if section == "Overview":
        camp, camp_count = leading(records, "settlement_clean", {MISSING, REVIEW})
        if camp:
            findings.append(f"Of the {total:,} filtered records, {camp} contributes the largest share at {camp_count:,} ({camp_count / total:.1%}).")
        issue, issue_count = leading(issues, "issue_clean", {MISSING, REVIEW})
        if issue:
            findings.append(f"{issue} is the most frequently recorded issue with {issue_count:,} mentions; these are multi-response mentions rather than unique children.")

    elif section == "Monthly Trends":
        dated = records.dropna(subset=["date"]).copy()
        if not dated.empty:
            monthly = dated.assign(period=dated["date"].dt.to_period("M")).groupby("period", observed=False).size().sort_index()
            peak_period = monthly.idxmax()
            findings.append(f"Across {total:,} filtered records, the highest monthly volume is {int(monthly.max()):,} in {peak_period.strftime('%b %Y')}.")
            if len(monthly) >= 2 and int(monthly.iloc[-2]) > 0:
                change = (int(monthly.iloc[-1]) - int(monthly.iloc[-2])) / int(monthly.iloc[-2])
                direction = "higher" if change > 0 else "lower" if change < 0 else "unchanged"
                findings.append(f"The latest month is {abs(change):.1%} {direction} than the preceding month, indicating a descriptive change in recorded activity rather than a causal effect.")

    elif section == "CPVs KPIs":
        staff, staff_count = leading(records, "staff_clean", {MISSING, REVIEW})
        if staff:
            findings.append(f"Among {total:,} filtered records, {staff} has the largest submission count at {staff_count:,} ({staff_count / total:.1%}).")
        findings.append(f"Overall, {records['staff_clean'].astype(str).loc[lambda s: ~s.isin([MISSING, REVIEW])].nunique():,} staff / CPVs are represented.")

    elif section == "Demographics":
        age, age_count = leading(records, "age_group", {MISSING, REVIEW})
        if age:
            findings.append(f"The largest age group is {age}, comprising {age_count:,} records ({age_count / total:.1%}).")
        known_visit = records["first_visit_clean"].astype(str).isin(["Yes", "No"])
        if known_visit.sum() >= 5:
            rate = records.loc[known_visit, "first_visit_clean"].astype(str).eq("Yes").mean()
            findings.append(f"Among {int(known_visit.sum()):,} records with a valid Yes/No response, the first-visit rate is {rate:.1%}.")

    elif section == "Games & Activities":
        game, game_count = leading(games, "game_clean", {MISSING, REVIEW})
        if game:
            findings.append(f"{game} is the leading activity with {game_count:,} mentions among {len(games):,} activity mentions across {games['record_id'].nunique():,} records.")
            findings.append("Activity totals exceed record totals when children participated in more than one activity.")

    elif section == "Protection & Support":
        disability_rate = records["disability_status_clean"].astype(str).eq("Yes").mean()
        findings.append(f"Within the {total:,} filtered records, {disability_rate:.1%} are marked as living with a disability.")
        issue, issue_count = leading(issues, "issue_clean", {MISSING, REVIEW})
        support_label, support_count = leading(support, "support_clean", {MISSING, REVIEW})
        if issue:
            findings.append(f"{issue} is the leading reported issue with {issue_count:,} mentions.")
        if support_label:
            findings.append(f"{support_label} is the most recorded support response with {support_count:,} mentions.")

    elif section == "Referrals":
        referral_rate = records["referral_made_clean"].astype(str).eq("Yes").mean()
        referral_count = int(records["referral_made_clean"].astype(str).eq("Yes").sum())
        findings.append(f"Of {total:,} filtered records, {referral_count:,} are marked as referred, giving a descriptive referral rate of {referral_rate:.1%}.")
        referred = records[records["referral_made_clean"].astype(str).eq("Yes")]
        destination, destination_count = leading(referred, "referral_destination_grouped", {MISSING, REVIEW, "Unknown"})
        if destination:
            findings.append(f"{destination} is the most recorded destination, accounting for {destination_count:,} referred records.")

    elif section == "Data Quality":
        checks = {
            "date": records["date"].isna(),
            "CFS / site": records["cfs_clean"].astype(str).eq(MISSING),
            "location": records["location_clean"].astype(str).isin([MISSING, REVIEW]),
            "age": records["age_clean"].isna(),
        }
        worst_label, worst_count = max(((label, int(mask.sum())) for label, mask in checks.items()), key=lambda item: item[1])
        if worst_count:
            findings.append(f"Across {total:,} filtered records, the largest monitored completeness gap is {worst_label}, affecting {worst_count:,} records ({worst_count / total:.1%}).")
        else:
            findings.append(f"All {total:,} filtered records pass the four headline completeness checks for date, CFS/site, location and age.")

    elif section == "Raw Data":
        findings.append(f"This view contains {total:,} filtered records for validation and review rather than causal interpretation.")
        findings.append("The downloadable extract removes direct identifiers and sensitive location or metadata fields.")

    # Add one consistent gender-disaggregation sentence when usable values are
    # present. It reports the actual denominator and does not infer causation.
    gender_values = records["gender_clean"].dropna().astype(str)
    gender_values = gender_values[~gender_values.isin([MISSING, REVIEW])]
    gender_counts = gender_values.value_counts()
    gender_sentence = ""
    if not gender_counts.empty:
        gender_denominator = int(gender_counts.sum())
        gender_parts = [
            f"{label}: {int(count):,} ({int(count) / gender_denominator:.1%})"
            for label, count in gender_counts.items()
        ]
        gender_sentence = (
            f"By recorded gender, among {gender_denominator:,} records with a usable response, "
            + "; ".join(gender_parts)
            + "."
        )

    narrative_parts = findings[:2]
    if gender_sentence:
        narrative_parts.append(gender_sentence)
    return " ".join(narrative_parts) if narrative_parts else "The available filtered data do not support a sufficiently clear descriptive finding for this section."


def build_table_driven_narrative(
    section: str,
    records: pd.DataFrame,
    issues: pd.DataFrame,
    support: pd.DataFrame,
    games: pd.DataFrame,
) -> str:
    """Summarise the full family of analytical tables used by a section."""
    total = len(records)
    if total == 0:
        return "No records match the current filters, so the section tables do not support a finding."

    excluded = {MISSING, REVIEW, "Unknown", "None"}

    def counts(data: pd.DataFrame, column: str) -> pd.Series:
        if data.empty or column not in data.columns:
            return pd.Series(dtype="int64")
        values = data[column].dropna().astype(str)
        values = values[~values.isin(excluded)]
        return values.value_counts()

    def leader(data: pd.DataFrame, column: str) -> Tuple[str, int]:
        result = counts(data, column)
        return (str(result.index[0]), int(result.iloc[0])) if not result.empty else ("", 0)

    def gender_summary(data: pd.DataFrame = records) -> str:
        result = counts(data, "gender_clean")
        denominator = int(result.sum())
        if not denominator:
            return "Gender disaggregation is unavailable because no usable gender responses remain."
        parts = [f"{label}: {int(value):,} ({int(value) / denominator:.1%})" for label, value in result.items()]
        return f"Among {denominator:,} usable gender responses, " + "; ".join(parts) + "."

    def leaders_by_gender(data: pd.DataFrame, category: str, unit: str) -> str:
        if data.empty or category not in data.columns or "gender_clean" not in data.columns:
            return ""
        parts = []
        for gender in GENDER_ORDER:
            if gender == MISSING:
                continue
            scoped = data[data["gender_clean"].astype(str).eq(gender)]
            label, value = leader(scoped, category)
            if label:
                parts.append(f"for {gender}, {label} leads with {value:,} {unit}")
        return "; ".join(parts) + "." if parts else ""

    def positive_rate_by_gender(data: pd.DataFrame, column: str, positive_label: str) -> str:
        """Describe a meaningful positive outcome instead of saying Yes/No 'leads'."""
        if data.empty or column not in data.columns or "gender_clean" not in data.columns:
            return ""
        parts = []
        for gender in GENDER_ORDER:
            if gender == MISSING:
                continue
            scoped = data[
                data["gender_clean"].astype(str).eq(gender)
                & data[column].astype(str).isin(["Yes", "No"])
            ]
            denominator = len(scoped)
            if denominator < 5:
                continue
            positive = int(scoped[column].astype(str).eq("Yes").sum())
            parts.append(f"{gender}: {positive:,} of {denominator:,} ({positive / denominator:.1%})")
        if not parts:
            return "Gender-specific rates are not shown because each available group has fewer than five valid responses."
        return f"By gender, {positive_label} was recorded for " + "; ".join(parts) + "."

    def selection_pattern(data: pd.DataFrame, category: str, noun: str) -> str:
        if data.empty or "record_id" not in data.columns:
            return f"No usable {noun.lower()} selections are available."
        per_record = data.groupby("record_id").size()
        selected_records = int(per_record.size)
        multiple = int(per_record.gt(1).sum())
        multiple_rate = multiple / selected_records if selected_records else 0
        combos = (
            data.groupby("record_id")[category]
            .apply(lambda values: " + ".join(sorted(set(map(str, values)))))
            .value_counts()
        )
        combo_text = ""
        if not combos.empty:
            combo_text = f" The most common recorded combination is {combos.index[0]} ({int(combos.iloc[0]):,} records)."
        return (
            f"The {len(data):,} {noun.lower()} mentions come from {selected_records:,} records; "
            f"{multiple:,} ({multiple_rate:.1%}) contain more than one selection.{combo_text}"
        )

    blocks: List[Tuple[str, str]] = []

    if section == "Overview":
        camp, camp_n = leader(records, "settlement_clean")
        age, age_n = leader(records, "age_group")
        issue, issue_n = leader(issues, "issue_clean")
        overall = f"The section covers {total:,} filtered records"
        if camp:
            overall += f"; {camp} contributes the largest camp share at {camp_n:,} ({camp_n / total:.1%})"
        if age:
            overall += f", and {age} is the largest age group at {age_n:,} ({age_n / total:.1%})"
        overall += "."
        blocks.append(("Coverage and profile", overall))
        blocks.append(("Gender", gender_summary()))
        if issue:
            blocks.append(("Protection signal", f"{issue} leads the issue table with {issue_n:,} mentions. {leaders_by_gender(issues, 'issue_clean', 'mentions')}"))

    elif section == "Monthly Trends":
        dated = records.dropna(subset=["date"]).copy()
        if dated.empty:
            blocks.append(("Trend", "No valid dates remain for monthly analysis."))
        else:
            dated["period"] = dated["date"].dt.to_period("M")
            monthly = dated.groupby("period", observed=False).size().sort_index()
            peak = monthly.idxmax()
            trend = f"The peak month is {peak.strftime('%b %Y')} with {int(monthly.max()):,} visits."
            if len(monthly) >= 2 and int(monthly.iloc[-2]):
                delta = (int(monthly.iloc[-1]) - int(monthly.iloc[-2])) / int(monthly.iloc[-2])
                trend += f" The latest month is {abs(delta):.1%} {'higher' if delta > 0 else 'lower' if delta < 0 else 'unchanged'} than the previous month."
            blocks.append(("Visit trend", trend))
            latest = dated[dated["period"].eq(monthly.index[-1])]
            blocks.append(("Latest-month gender profile", gender_summary(latest)))
        known_first = records["first_visit_clean"].astype(str).isin(["Yes", "No"])
        known_referral = records["referral_made_clean"].astype(str).isin(["Yes", "No"])
        rate_text = []
        if known_first.any():
            rate_text.append(f"first visits are {records.loc[known_first, 'first_visit_clean'].astype(str).eq('Yes').mean():.1%} of {int(known_first.sum()):,} valid responses")
        if known_referral.any():
            rate_text.append(f"referrals are {records.loc[known_referral, 'referral_made_clean'].astype(str).eq('Yes').mean():.1%} of {int(known_referral.sum()):,} valid responses")
        if rate_text:
            blocks.append(("Service indicators", "; ".join(rate_text).capitalize() + "."))

    elif section == "CPVs KPIs":
        staff_counts = counts(records, "staff_clean")
        if not staff_counts.empty:
            blocks.append(("Workload", f"{len(staff_counts):,} staff / CPVs are represented; {staff_counts.index[0]} leads with {int(staff_counts.iloc[0]):,} submissions ({int(staff_counts.iloc[0]) / total:.1%})."))
        blocks.append(("Gender reach", gender_summary()))
        referral_by_staff = records.groupby("staff_clean", observed=False).agg(records=("record_id", "count"), referrals=("referral_made_clean", lambda s: s.astype(str).eq("Yes").sum()))
        referral_by_staff = referral_by_staff[referral_by_staff["records"] >= 5]
        if not referral_by_staff.empty:
            referral_by_staff["rate"] = referral_by_staff["referrals"] / referral_by_staff["records"]
            top = referral_by_staff["rate"].idxmax()
            blocks.append(("Referral pattern", f"Among staff with at least five records, {top} has the highest descriptive referral rate at {referral_by_staff.loc[top, 'rate']:.1%} ({int(referral_by_staff.loc[top, 'referrals']):,} of {int(referral_by_staff.loc[top, 'records']):,})."))

    elif section == "Demographics":
        age, age_n = leader(records, "age_group")
        blocks.append(("Population profile", f"{age} is the largest age group at {age_n:,} records ({age_n / total:.1%}). {gender_summary()}" if age else gender_summary()))
        known = records[records["first_visit_clean"].astype(str).isin(["Yes", "No"])]
        if not known.empty:
            first_n = int(known["first_visit_clean"].astype(str).eq("Yes").sum())
            blocks.append(("Visit history", f"First visits account for {first_n:,} of {len(known):,} valid responses ({first_n / len(known):.1%}). {positive_rate_by_gender(known, 'first_visit_clean', 'a first visit')}"))
        site, site_n = leader(records, "cfs_clean")
        if site:
            blocks.append(("Site reach", f"{site} records the largest beneficiary volume at {site_n:,} ({site_n / total:.1%})."))

    elif section == "Games & Activities":
        game, game_n = leader(games, "game_clean")
        if game:
            blocks.append(("Activity profile", f"{game} leads with {game_n:,} mentions. {leaders_by_gender(games, 'game_clean', 'mentions')}"))
        blocks.append(("Multiple activities", selection_pattern(games, "game_clean", "Activity")))
        take5 = records[records["take5_integrated_clean"].astype(str).isin(["Yes", "No"])]
        if not take5.empty:
            yes_n = int(take5["take5_integrated_clean"].astype(str).eq("Yes").sum())
            blocks.append(("Take 5", f"Take 5 integration is recorded as Yes in {yes_n:,} of {len(take5):,} valid responses ({yes_n / len(take5):.1%}). {positive_rate_by_gender(take5, 'take5_integrated_clean', 'Take 5 integration')}"))

    elif section == "Protection & Support":
        disability = records[records["disability_status_clean"].astype(str).eq("Yes")]
        disability_type, disability_type_n = leader(disability, "disability_type_display")
        text = f"Disability is recorded in {len(disability):,} of {total:,} records ({len(disability) / total:.1%})."
        if disability_type:
            text += f" {disability_type} is the leading recorded type ({disability_type_n:,})."
        blocks.append(("Disability", text + " " + positive_rate_by_gender(records, "disability_status_clean", "living with a disability")))
        issue, issue_n = leader(issues, "issue_clean")
        if issue:
            blocks.append(("Issues", f"{issue} leads with {issue_n:,} mentions. {leaders_by_gender(issues, 'issue_clean', 'mentions')} {selection_pattern(issues, 'issue_clean', 'Issue')}"))
        support_label, support_n = leader(support, "support_clean")
        if support_label:
            blocks.append(("Support", f"{support_label} is the leading response with {support_n:,} mentions. {leaders_by_gender(support, 'support_clean', 'mentions')}"))

    elif section == "Referrals":
        referred = records[records["referral_made_clean"].astype(str).eq("Yes")]
        blocks.append(("Referral rate", f"{len(referred):,} of {total:,} records are marked as referred ({len(referred) / total:.1%}). {positive_rate_by_gender(records, 'referral_made_clean', 'a referral')}"))
        site_rates = records.groupby("cfs_clean", observed=False).agg(total=("record_id", "count"), referred=("referral_made_clean", lambda s: s.astype(str).eq("Yes").sum()))
        site_rates = site_rates[site_rates["total"] >= 5]
        if not site_rates.empty:
            site_rates["rate"] = site_rates["referred"] / site_rates["total"]
            site = site_rates["rate"].idxmax()
            blocks.append(("Site comparison", f"Among sites with at least five records, {site} has the highest descriptive referral rate at {site_rates.loc[site, 'rate']:.1%}."))
        destination, destination_n = leader(referred, "referral_destination_grouped")
        agency, agency_n = leader(referred, "external_referral_agency_clean")
        detail = []
        if destination:
            detail.append(f"{destination} is the leading destination ({destination_n:,})")
        if agency:
            detail.append(f"{agency} is the leading named external agency ({agency_n:,})")
        if detail:
            blocks.append(("Referral destinations", "; ".join(detail) + ". " + leaders_by_gender(referred, "referral_destination_grouped", "records")))

    elif section == "Data Quality":
        fields = {
            "date": records["date"].notna(), "gender": ~records["gender_clean"].astype(str).isin([MISSING, REVIEW]),
            "age": records["age_clean"].notna(), "CFS / site": ~records["cfs_clean"].astype(str).isin([MISSING, REVIEW]),
            "location": ~records["location_clean"].astype(str).isin([MISSING, REVIEW]),
            "first visit": records["first_visit_clean"].astype(str).isin(["Yes", "No"]),
            "referral": records["referral_made_clean"].astype(str).isin(["Yes", "No"]),
        }
        completeness = sorted(((label, int(mask.sum()) / total) for label, mask in fields.items()), key=lambda item: item[1])
        blocks.append(("Completeness", f"The lowest headline completeness is {completeness[0][0]} at {completeness[0][1]:.1%}; the highest is {completeness[-1][0]} at {completeness[-1][1]:.1%}."))
        blocks.append(("Gender", gender_summary()))
        schema_missing = [column for column in CORE_ANALYSIS_COLUMNS if column not in records.columns]
        blocks.append(("Schema", "All required analysis columns are available." if not schema_missing else f"{len(schema_missing):,} required analysis columns are missing: {', '.join(schema_missing[:5])}."))

    else:  # Records & Export
        blocks.append(("Records", f"The table contains {total:,} rows after applying the current filters. {gender_summary()}"))
        blocks.append(("Export", "The optional CSV export excludes direct identifiers, sensitive locations, device/user identifiers, notes and UUID fields."))

    if not blocks:
        return "The available filtered tables do not support a sufficiently clear descriptive finding."
    return "\n\n".join(f"**{heading}.** {text.strip()}" for heading, text in blocks)


def schema_readiness_table(data: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in CORE_ANALYSIS_COLUMNS:
        present = col in data.columns
        filled = 0
        if present:
            filled = int((data[col].notna() & data[col].astype(str).str.strip().ne("") & data[col].astype(str).ne(MISSING)).sum())
        rows.append({
            "Analysis column": col,
            "Status": "Available" if present else "Missing",
            "Records filled": filled,
        })
    out = pd.DataFrame(rows)
    return out.sort_values(["Status", "Analysis column"], ascending=[True, True]).reset_index(drop=True)

# -----------------------------------------------------------------------------
# Streamlit app
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Tdh Kenya CFS Dashboard", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
load_external_css(CSS_PATH)

# Optional OIDC protection for public/multi-user hosting. Set AUTH_REQUIRED=true
# and configure Streamlit's [auth] secrets to activate it.
auth_required = str(setting("AUTH_REQUIRED", "false")).strip().lower() in {"1", "true", "yes", "on"}
if auth_required and not st.user.is_logged_in:
    st.title("Tdh Kenya CFS Dashboard")
    st.info("Sign in with your organisation account to access this dashboard.")
    st.button("Sign in", type="primary", on_click=st.login)
    st.stop()

page_header(
    "Tdh Kenya Child Friendly Spaces Dashboard",
    "Professional monitoring dashboard for CFS visits, demographics, protection issues, support, activities, referrals, and data quality.",
)

st.sidebar.markdown("<div class='sidebar-title'>Dashboard Controls</div>", unsafe_allow_html=True)
st.sidebar.caption(APP_VERSION)
if auth_required:
    st.sidebar.caption(f"Signed in: {getattr(st.user, 'email', getattr(st.user, 'name', 'Authenticated user'))}")
    st.sidebar.button("Sign out", on_click=st.logout, use_container_width=True)
if st.sidebar.button("↺ Reset view", use_container_width=True, help="Clear active filters and return to the default dashboard view."):
    for state_key in list(st.session_state.keys()):
        if state_key.startswith("filter_") or state_key in {"date_from_filter", "date_to_filter", "dashboard_section", "dashboard_section_category"}:
            st.session_state.pop(state_key, None)
    st.rerun()

data_path = resolve_data_path()
source_label = ""
source_modified = None
source_mode = "KoboToolbox API" if kobo_configured() else "Local file fallback"
source_fingerprint = "api-live"
source_stat = None
page_count = 0

if "kobo_refresh_nonce" not in st.session_state:
    st.session_state.kobo_refresh_nonce = 0

load_started_at = time.perf_counter()
startup_status = st.status("Loading prepared dashboard data…", expanded=False)
try:
    if kobo_configured():
        startup_status.update(label="Connecting to Kobo and loading cached submissions…", state="running")
        kobo_base_url = str(setting("KOBO_BASE_URL", "https://kf.kobotoolbox.org"))
        kobo_asset_uid = str(setting("KOBO_ASSET_UID"))
        kobo_token = str(setting("KOBO_TOKEN"))
        raw_count, df, issue_long, support_long, game_long, fetched_at, page_count = load_kobo_dashboard_data(
            kobo_base_url, kobo_asset_uid, kobo_token, st.session_state.kobo_refresh_nonce
        )
        fetched_dt = pd.to_datetime(fetched_at, utc=True).tz_convert("Africa/Nairobi")
        source_modified = fetched_dt.strftime("%d %b %Y %H:%M EAT")
        source_label = f"Kobo asset {kobo_asset_uid}"
    elif data_path.exists():
        startup_status.update(label="Loading cached local dashboard data…", state="running")
        source_stat = data_path.stat()
        modified_time = source_stat.st_mtime
        source_label = f"Source file: {data_path}"
        source_modified = time.strftime("%d %b %Y %H:%M", time.localtime(modified_time))
        source_fingerprint = source_content_fingerprint(str(data_path), source_stat.st_size, source_stat.st_mtime_ns)
        raw_count, df, issue_long, support_long, game_long = load_dashboard_data_cached(str(data_path), source_fingerprint)
    else:
        st.error("Kobo is not configured and no fallback data file was found.")
        st.info("Add KOBO_TOKEN and KOBO_ASSET_UID to .streamlit/secrets.toml. See README_KOBO.md.")
        st.stop()
except requests.RequestException as exc:
    startup_status.update(label="Kobo connection failed", state="error")
    st.error(f"Kobo could not be reached: {exc}")
    st.info("The app did not expose your token. Check connectivity, Kobo server URL, and project permissions.")
    st.stop()
except Exception as exc:
    startup_status.update(label="Dashboard preparation failed", state="error")
    st.error(f"The dashboard data could not be loaded/prepared: {exc}")
    st.stop()

load_elapsed_seconds = time.perf_counter() - load_started_at
startup_status.update(label=f"Dashboard data ready in {load_elapsed_seconds:.1f} seconds", state="complete")

# Enforce the approved age taxonomy after every load, including loads served
# from a prepared pickle or Streamlit's persisted cache. This prevents legacy
# age-band labels from resurfacing in any filter, table, chart, or narrative,
# even if an older transformed frame is encountered.
if not df.empty:
    if "age_clean" in df.columns:
        enforced_age_groups = df["age_clean"].map(age_group)
    elif "child_age" in df.columns:
        enforced_age_groups = df["child_age"].map(age_group)
    else:
        enforced_age_groups = pd.Series(MISSING, index=df.index)
    df["age_group"] = pd.Categorical(
        enforced_age_groups.astype(str),
        categories=AGE_GROUP_ORDER,
        ordered=True,
    )

if df.empty:
    st.error("The dataset was loaded, but no analysable records remained after preparation.")
    st.info("This usually means the consent field or transformed column mapping changed. Check Data Quality → Analysis Schema Readiness after verifying the source file.")

st.caption(f"Data source: {source_mode} · Last fetched/loaded: {source_modified}")

with st.sidebar.expander("ℹ Live data & performance", expanded=False):
    st.caption(f"Source: {source_label}")
    st.caption(f"Refresh window: {KOBO_CACHE_TTL_SECONDS} seconds" if kobo_configured() else "Refresh mode: file change")
    if kobo_configured():
        st.caption(f"API pages fetched: {page_count:,}")
    elif source_stat is not None:
        st.caption(f"Workbook size: {source_stat.st_size / (1024 * 1024):.2f} MB")
        st.caption(f"Source fingerprint: {source_fingerprint[:12]}")
    st.caption(f"Load/cache lookup: {load_elapsed_seconds:.3f} seconds")
    st.caption(f"Raw records: {raw_count:,}")
    st.caption(f"Analysable records: {len(df):,}")
    if st.button("Fetch latest Kobo data" if kobo_configured() else "Refresh data cache", key="admin_refresh_data_cache", use_container_width=True):
        if kobo_configured():
            st.session_state.kobo_refresh_nonce += 1
            fetch_kobo_submissions.clear()
            load_kobo_dashboard_data.clear()
            fetch_kobo_schema_map.clear()
            fetch_kobo_support_choice_map.clear()
        else:
            source_content_fingerprint.clear()
            load_dashboard_data_cached.clear()
            read_file_cached.clear()
            try:
                PREPARED_CACHE_PATH.unlink(missing_ok=True)
            except Exception:
                pass
        st.rerun()

# Reserve a persistent navigation position above the filters. The container is
# populated after filtering so it can also show the current record count.
section_nav_slot = st.sidebar.container()

# Sidebar filters
st.sidebar.markdown("---")
st.sidebar.markdown("<div class='sidebar-title'>Smart Filters</div>", unsafe_allow_html=True)
st.sidebar.caption("Filters cascade automatically: each selection updates the available options below it.")

valid_dates = df["date"].dropna()
filtered = df.copy()
if not valid_dates.empty:
    min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
    with st.sidebar.expander("📅 Date range", expanded=True):
        st.markdown(
            f"<div class='filter-meta'><span>Dataset range</span><span>{min_date:%d %b %Y} → {max_date:%d %b %Y}</span></div>",
            unsafe_allow_html=True,
        )
        c1, c2 = st.columns(2)
        with c1:
            from_date = st.date_input("From", value=min_date, min_value=min_date, max_value=max_date, key="date_from_filter")
        with c2:
            to_date = st.date_input("To", value=max_date, min_value=min_date, max_value=max_date, key="date_to_filter")
    if from_date > to_date:
        from_date, to_date = to_date, from_date
    filtered = filtered[filtered["date"].between(pd.to_datetime(from_date), pd.to_datetime(to_date), inclusive="both")].copy()
    date_label = f"{from_date:%d %b %Y} to {to_date:%d %b %Y}"
else:
    date_label = "All available dates"

with st.sidebar.expander("📍 Location filters", expanded=True):
    st.caption("Select from broad location to specific CFS; choices cascade automatically.")
    selected_camp = multiselect_filter("🏕 Camp", filtered, "settlement_clean")
    filtered = filter_if_selected(filtered, "settlement_clean", selected_camp)
    selected_location = multiselect_filter("📌 Specific location", filtered, "location_clean")
    filtered = filter_if_selected(filtered, "location_clean", selected_location)
    selected_cfs = multiselect_filter("🏛 CFS / site", filtered, "cfs_clean")
    filtered = filter_if_selected(filtered, "cfs_clean", selected_cfs)

with st.sidebar.expander("👥 Child and service profile", expanded=False):
    st.caption("Narrow the records by staff and child characteristics.")
    selected_staff = multiselect_filter("👤 Staff / CPV", filtered, "staff_clean")
    filtered = filter_if_selected(filtered, "staff_clean", selected_staff)
    selected_gender = multiselect_filter("⚧ Gender", filtered, "gender_clean", GENDER_ORDER)
    filtered = filter_if_selected(filtered, "gender_clean", selected_gender)
    selected_age = multiselect_filter("🎂 Age group", filtered.assign(age_group=filtered["age_group"].astype(str)), "age_group", AGE_GROUP_ORDER)
    filtered = filter_if_selected(filtered.assign(age_group=filtered["age_group"].astype(str)), "age_group", selected_age)
    selected_disability = multiselect_filter("♿ Living with disability", filtered.assign(disability_status_clean=filtered["disability_status_clean"].astype(str)), "disability_status_clean", YES_NO_ORDER)
    filtered = filter_if_selected(filtered.assign(disability_status_clean=filtered["disability_status_clean"].astype(str)), "disability_status_clean", selected_disability)

with st.sidebar.expander("🎯 Visit and service outcomes", expanded=False):
    st.caption("Focus the dashboard on visit status or referrals.")
    selected_first_visit = multiselect_filter("🔄 First visit", filtered.assign(first_visit_clean=filtered["first_visit_clean"].astype(str)), "first_visit_clean", YES_NO_ORDER)
    filtered = filter_if_selected(filtered.assign(first_visit_clean=filtered["first_visit_clean"].astype(str)), "first_visit_clean", selected_first_visit)
    selected_referral = multiselect_filter("🔁 Referral made", filtered.assign(referral_made_clean=filtered["referral_made_clean"].astype(str)), "referral_made_clean", YES_NO_ORDER)
    filtered = filter_if_selected(filtered.assign(referral_made_clean=filtered["referral_made_clean"].astype(str)), "referral_made_clean", selected_referral)
# First-visit fields are repaired once during cached data preparation; repeating
# the row-wise repair on every widget interaction needlessly slows navigation.

filter_summary_bits = [
    f"Date selected: {date_label}",
    f"Camp Name: {', '.join(selected_camp) if selected_camp else 'All camps'}",
    f"Specific camp location: {', '.join(selected_location) if selected_location else 'All specific camp locations'}",
    f"CFS / site: {', '.join(selected_cfs) if selected_cfs else 'All CFS / sites'}",
    f"Staff: {', '.join(selected_staff) if selected_staff else 'All staff'}",
    f"Gender: {', '.join(selected_gender) if selected_gender else 'All genders'}",
    f"Age group: {', '.join(selected_age) if selected_age else 'All age groups'}",
    f"Disability: {', '.join(selected_disability) if selected_disability else 'All disability statuses'}",
    f"First visit: {', '.join(selected_first_visit) if selected_first_visit else 'All visit statuses'}",
    f"Referral made: {', '.join(selected_referral) if selected_referral else 'All referral statuses'}",
]
with st.sidebar.expander("Current filter path", expanded=False):
    st.markdown(
        f"<div class='filter-summary'>{'<br>'.join(filter_summary_bits)}</div>",
        unsafe_allow_html=True,
    )
active_filter_count = sum(bool(values) for values in [selected_camp, selected_location, selected_cfs, selected_staff, selected_gender, selected_age, selected_disability, selected_first_visit, selected_referral])
st.sidebar.markdown(
    f"<div class='filter-result-card'><span>{active_filter_count} active filters · records</span><b>{len(filtered):,}</b></div>",
    unsafe_allow_html=True,
)

with section_nav_slot:
    st.markdown("<div class='sidebar-title'>Explore Dashboard</div>", unsafe_allow_html=True)
    st.caption("Choose a section. Your filters remain active as you move between views.")
    section = nav_menu(SECTION_OPTIONS)
    active_meta = SECTION_META[section]
    st.caption(f"{SECTION_CATEGORY[section]} · {len(filtered):,} filtered records")
    with st.expander("How to use this dashboard", expanded=False):
        st.markdown(
            "1. Choose a **dashboard section**.\n"
            "2. Apply filters below; selections carry across views.\n"
            "3. Read the section purpose shown above each analysis.\n\n"
            "**CPV:** Community-based protection volunteer  \n"
            "**First-visit rate:** Share recorded as visiting for the first time  \n"
            "**Mention count:** One record may contain several selected responses"
        )

ctx_cols = ["record_id", "settlement_clean", "location_clean", "cfs_clean", "staff_clean", "gender_clean", "age_group"]
# Build only the multi-response datasets required by the active section. These
# merges were previously performed three times on every navigation/filter rerun.
needs_issues = section in {"Overview", "Protection & Support"}
needs_support = section in {"Overview", "Protection & Support"}
needs_games = section in {"Overview", "Games & Activities"}
issue_context = issue_long.merge(filtered[ctx_cols], on="record_id", how="inner") if needs_issues and not filtered.empty else issue_long.iloc[0:0]
support_context = support_long.merge(filtered[ctx_cols], on="record_id", how="inner") if needs_support and not filtered.empty else support_long.iloc[0:0]
game_context = game_long.merge(filtered[ctx_cols], on="record_id", how="inner") if needs_games and not filtered.empty else game_long.iloc[0:0]

st.markdown(
    f"<div class='record-status'>Showing {len(filtered):,} of {len(df):,} records | Date selected: {date_label}</div>",
    unsafe_allow_html=True,
)

active_filter_labels = []
if date_label != "All available dates":
    active_filter_labels.append(f"Date: {date_label}")
for label, values in [
    ("Camp", selected_camp), ("Location", selected_location), ("CFS", selected_cfs),
    ("Staff", selected_staff), ("Gender", selected_gender), ("Age", selected_age),
    ("Disability", selected_disability), ("First visit", selected_first_visit),
    ("Referral", selected_referral),
]:
    if values:
        display_values = ", ".join(map(str, values[:2])) + (f" +{len(values) - 2}" if len(values) > 2 else "")
        active_filter_labels.append(f"{label}: {display_values}")
if active_filter_labels:
    filter_chips = "".join(
        f'<span style="display:inline-block;margin:0 .4rem .4rem 0;padding:.32rem .65rem;border-radius:999px;background:#eaf2ff;color:#172554;font-size:.82rem;font-weight:700;">{html_lib.escape(label)}</span>'
        for label in active_filter_labels
    )
    st.markdown(f'<div aria-label="Active filters">{filter_chips}</div>', unsafe_allow_html=True)
else:
    st.caption("No additional filters applied · showing the full analysable dataset")

# A concise, filter-aware warning surfaces material data-quality issues without
# changing any analytical section or excluding additional records.
dq_checks = {
    "missing dates": int(filtered["date"].isna().sum()),
    "missing CFS/site": int(filtered["cfs_clean"].astype(str).eq(MISSING).sum()),
    "unresolved locations": int(filtered["location_clean"].astype(str).isin([MISSING, REVIEW]).sum()),
    "invalid/missing ages": int(filtered["age_clean"].isna().sum()),
}
dq_alert_threshold = max(5, round(len(filtered) * 0.05)) if len(filtered) else 1
dq_issues = [f"{count:,} {label}" for label, count in dq_checks.items() if count >= dq_alert_threshold]
if dq_issues:
    st.warning("Data-quality notice for the current filters: " + "; ".join(dq_issues) + ". Review the Data Quality section for details.")

# The complete landing-page summary appears only on Overview. Other sections
# already retain the compact record/date status and active-filter chips above,
# allowing users to reach section-specific analysis without repeated scrolling.
if section == "Overview":
    referral_rate = filtered["referral_made_clean"].astype(str).eq("Yes").mean() if len(filtered) else 0
    first_visit_rate = filtered["first_visit_clean"].astype(str).eq("Yes").mean() if len(filtered) else 0
    disability_rate = filtered["disability_status_clean"].astype(str).eq("Yes").mean() if len(filtered) else 0
    avg_age = filtered["age_clean"].mean() if len(filtered) else None

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        metric_card("Total CFS visits", f"{len(filtered):,}", "Filtered records", "primary")
    with k2:
        metric_card("Referral rate", f"{referral_rate:.1%}", "Referral marked Yes", "success")
    with k3:
        metric_card("First visit rate", f"{first_visit_rate:.1%}", "First visit marked Yes", "primary")
    with k4:
        metric_card("Disability prevalence", f"{disability_rate:.1%}", "Children recorded as Yes", "warning")
    with k5:
        metric_card("Issue mentions", f"{len(issue_context):,}", "Multi-response count", "warning")
    with k6:
        metric_card("Avg child age", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A", "Years", "neutral")

    section_header("Quick Insights", "Fast outlook for the currently filtered dataset.")
    qi1, qi2, qi3, qi4, qi5 = st.columns(5)
    top_camp, top_camp_count = top_category(filtered, "settlement_clean", [MISSING])
    top_site, top_site_count = top_category(filtered, "cfs_clean", [MISSING])
    top_staff, top_staff_count = top_category(filtered, "staff_clean", [MISSING])
    top_issue, top_issue_count = top_category(issue_context, "issue_clean", [MISSING])
    top_game, top_game_count = top_category(game_context, "game_clean", [MISSING])
    staff_represented = filtered.loc[~filtered["staff_clean"].astype(str).isin([MISSING, REVIEW]), "staff_clean"].nunique() if "staff_clean" in filtered.columns else 0
    with qi1:
        insight_card("Camp with most records", top_camp, f"{top_camp_count:,} records")
    with qi2:
        insight_card("Leading CFS / site", top_site, f"{top_site_count:,} records")
    with qi3:
        insight_card("Staff represented", f"{staff_represented:,}", f"Top: {top_staff} ({top_staff_count:,})")
    with qi4:
        insight_card("Most reported issue", top_issue, f"{top_issue_count:,} mentions")
    with qi5:
        insight_card("Top game / activity", top_game, f"{top_game_count:,} records")

active_meta = SECTION_META[section]
breadcrumb_col, overview_col = st.columns([5, 1])
with breadcrumb_col:
    st.caption(f"Dashboard › {SECTION_CATEGORY[section]} › {active_meta['label']} · {len(filtered):,} records in view")
with overview_col:
    if section != "Overview":
        st.button("← Overview", key="back_to_overview", use_container_width=True, on_click=go_to_overview)
section_intro_card(section)

with st.expander("Findings from the current tables", expanded=(section == "Overview")):
    st.caption("Automatically generated, filter-aware descriptive findings. They summarize observed patterns and do not establish causes.")
    if section == "Overview":
        section_narrative = build_table_driven_narrative(section, filtered, issue_context, support_context, game_context)
        st.markdown(section_narrative)
    elif st.button("Generate findings", key=f"generate_findings_{section}", use_container_width=False):
        section_narrative = build_table_driven_narrative(section, filtered, issue_context, support_context, game_context)
        st.markdown(section_narrative)
    else:
        st.caption("Generate this narrative only when needed to keep section navigation fast.")

# -----------------------------------------------------------------------------
# Sections
# -----------------------------------------------------------------------------
if section == "Overview":
    st.caption("Records where consent was 'No' are excluded from all dashboard analytics.")

    with st.expander("Quick guide to dashboard views", expanded=False):
        st.markdown(
            "**Summary:** current results and change over time  \n"
            "**People & Services:** children reached, staff delivery and activities  \n"
            "**Protection Pathway:** reported needs, support and referrals  \n"
            "**Data Management:** quality review, records and privacy-safe exports"
        )

    st.divider()
    st.subheader("Coverage: Camp Records by Gender")
    camp_table = table_with_total(filtered, ["settlement_clean"], ["gender_clean"])
    render_table(camp_table, "Camp Records by Gender", "overview_camp")
    camp_chart = count_table(filtered, ["settlement_clean", "gender_clean"])
    camp_chart = add_top_n_control(camp_chart, "settlement_clean", "overview_camp", default=5)
    bar_chart(camp_chart, "Count", "settlement_clean", "gender_clean", "Overview: camp records by gender", horizontal=True, height=420)

    st.divider()
    st.subheader("Beneficiary Profile: Age Group by Gender")
    age_table = table_with_total(filtered, ["age_group"], ["gender_clean"])
    render_table(age_table, "Age Group by Gender", "overview_age")
    age_chart = count_table(filtered, ["age_group", "gender_clean"], order_col="age_group")
    if not age_chart.empty:
        # Plotly renders plain strings more reliably than pandas categorical
        # values after the expanded age-band ordering is applied.
        age_chart = age_chart.copy()
        age_chart["age_group"] = age_chart["age_group"].astype(str)
        age_chart["gender_clean"] = age_chart["gender_clean"].astype(str)
        age_chart = age_chart[
            age_chart["age_group"].isin(
                [value for value in AGE_GROUP_ORDER if value != MISSING]
            )
            & age_chart["gender_clean"].isin(
                [value for value in GENDER_ORDER if value != MISSING]
            )
            & age_chart["Count"].gt(0)
        ].copy()
    bar_chart(
        age_chart,
        "age_group",
        "Count",
        "gender_clean",
        "Overview: age group by gender",
        height=480,
        category_orders={
            "age_group": [value for value in AGE_GROUP_ORDER if value != MISSING],
            "gender_clean": [value for value in GENDER_ORDER if value != MISSING],
        },
    )

    st.divider()
    st.subheader("Protection: Top Reported Issues")
    issue_table = table_with_total(issue_context, ["issue_clean"], ["gender_clean"])
    render_table(issue_table, "Top Issues by Gender", "overview_issues")
    issue_chart = count_table(issue_context, ["issue_clean", "gender_clean"])
    issue_chart = add_top_n_control(issue_chart, "issue_clean", "overview_issues", default=10)
    bar_chart(issue_chart, "Count", "issue_clean", "gender_clean", "Overview: top issues by gender", horizontal=True, height=540)

elif section == "Monthly Trends":
    st.subheader("Monthly Visit Trends")
    monthly = filtered.dropna(subset=["date"]).copy()
    if monthly.empty:
        st.info("No dated records available for the current filters.")
    else:
        monthly["month_label"] = monthly["date"].dt.to_period("M").astype(str)
        monthly["Month"] = monthly["date"].dt.strftime("%b %Y")
        vpm = monthly.groupby(["month_label", "Month"], observed=False).size().rename("Visits").reset_index().sort_values("month_label")
        line_chart(vpm, "Month", "Visits", title="Total CFS visits per month")

        gm = monthly.groupby(["month_label", "Month", "gender_clean"], observed=False).size().rename("Visits").reset_index().sort_values("month_label")
        gm = gm[gm["Visits"] > 0]
        line_chart(gm, "Month", "Visits", color="gender_clean", title="CFS visits by gender per month")

        summary = monthly.groupby(["month_label", "Month"], observed=False).agg(
            Total_Visits=("record_id", "count"),
            Girls=("gender_clean", lambda s: s.astype(str).eq("Girls").sum()),
            Boys=("gender_clean", lambda s: s.astype(str).eq("Boys").sum()),
            First_Visit_Yes=("first_visit_clean", lambda s: s.astype(str).eq("Yes").sum()),
            Referrals_Yes=("referral_made_clean", lambda s: s.astype(str).eq("Yes").sum()),
            Avg_Age=("age_clean", "mean"),
        ).reset_index().sort_values("month_label")
        summary["First Visit Rate %"] = (summary["First_Visit_Yes"] / summary["Total_Visits"] * 100).round(1)
        summary["Referral Rate %"] = (summary["Referrals_Yes"] / summary["Total_Visits"] * 100).round(1)
        summary["Avg_Age"] = summary["Avg_Age"].round(1)
        summary = summary.drop(columns=["month_label"]).rename(columns={"Total_Visits": "Total Visits", "First_Visit_Yes": "First-Visit Children", "Referrals_Yes": "Referrals Made", "Avg_Age": "Avg Age (yrs)"}).set_index("Month")
        summary.loc["Grand Total", ["Total Visits", "Girls", "Boys", "First-Visit Children", "Referrals Made"]] = summary[["Total Visits", "Girls", "Boys", "First-Visit Children", "Referrals Made"]].sum()
        render_table(summary, "Monthly Trends Summary", "monthly_summary", precision=1)

elif section == "CPVs KPIs":
    st.subheader("CPV / Staff Performance Summary")
    cpv = filtered.groupby("staff_clean", observed=False).agg(
        Records=("record_id", "count"),
        First_Visit_Yes=("first_visit_clean", lambda s: s.astype(str).eq("Yes").sum()),
        Referrals_Yes=("referral_made_clean", lambda s: s.astype(str).eq("Yes").sum()),
        Disability_Yes=("disability_status_clean", lambda s: s.astype(str).eq("Yes").sum()),
        Unique_CFS=("cfs_clean", "nunique"),
        First_Date=("date", "min"),
        Last_Date=("date", "max"),
    ).reset_index().sort_values("Records", ascending=False)
    if not cpv.empty:
        cpv["First Visit Rate %"] = (cpv["First_Visit_Yes"] / cpv["Records"] * 100).round(1)
        cpv["Referral Rate %"] = (cpv["Referrals_Yes"] / cpv["Records"] * 100).round(1)
        cpv["First_Date"] = pd.to_datetime(cpv["First_Date"], errors="coerce").dt.strftime("%d %b %Y").fillna("")
        cpv["Last_Date"] = pd.to_datetime(cpv["Last_Date"], errors="coerce").dt.strftime("%d %b %Y").fillna("")
    render_table(cpv.rename(columns={"staff_clean": "Staff / CPV", "First_Visit_Yes": "First Visits", "Referrals_Yes": "Referrals", "Disability_Yes": "Disability Count", "Unique_CFS": "Unique CFS", "First_Date": "First Date", "Last_Date": "Last Date"}), "CPV Performance Summary", "cpv_summary", precision=1)
    staff_chart = count_table(filtered, ["staff_clean", "gender_clean"])
    staff_chart = add_top_n_control(staff_chart, "staff_clean", "staff_chart", default=15)
    bar_chart(staff_chart, "Count", "staff_clean", "gender_clean", "Staff submissions by gender", horizontal=True, height=700)

elif section == "Demographics":
    st.subheader("Demographics")
    demo_source = repair_first_visit_columns(filtered)
    fvr = demo_source[demo_source["first_visit_clean"].astype(str).isin(["Yes", "No"])].copy()
    if not fvr.empty:
        fvr["visit_type"] = fvr["first_visit_clean"].astype(str).map({"Yes": "First visit", "No": "Repeat visit"})
    else:
        candidate_cols = [c for c in demo_source.columns if "first" in norm_text(c) and "visit" in norm_text(c)]
        with st.expander("First-visit data diagnostic", expanded=True):
            st.warning("No Yes/No values were detected for first-visit analysis after applying the current filters.")
            st.caption("Candidate first-visit columns detected: " + (", ".join(map(str, candidate_cols)) if candidate_cols else "None"))
            if candidate_cols:
                diag_cols = ["record_id"] + candidate_cols[:5]
                st.dataframe(demo_source[diag_cols].head(20), use_container_width=True, hide_index=True)

    c1, c2 = st.columns([1, 1])
    with c1:
        visit_split = count_table(fvr, ["visit_type"]) if "visit_type" in fvr.columns else pd.DataFrame(columns=["visit_type", "Count"])
        pie_chart(visit_split, "visit_type", "Count", "First vs repeat visit split")
    with c2:
        pie_chart(count_table(filtered, ["gender_clean"]), "gender_clean", "Count", "Overall gender split")

    st.divider()
    st.subheader("First Visit and Repeat Visit Summary")
    st.caption("This table distinguishes children visiting the CFS for the first time from repeat visitors. Missing first-visit responses are excluded from this view.")
    render_table(table_with_total(fvr, ["visit_type"], ["gender_clean"]), "First Visit and Repeat Visit by Gender", "first_repeat_gender")

    st.divider()
    st.subheader("First Visit and Repeat Visit by CFS / Site")
    render_table(table_with_total(fvr, ["cfs_clean"], ["visit_type", "gender_clean"]), "First Visit and Repeat Visit by CFS Site", "first_repeat_site")
    fv_chart = count_table(fvr, ["cfs_clean", "visit_type"])
    fv_chart = add_top_n_control(fv_chart, "cfs_clean", "first_repeat_site_chart", default=15)
    bar_chart(fv_chart, "Count", "cfs_clean", "visit_type", "First and repeat visits by CFS / site", horizontal=True, height=620)

    st.divider()
    st.subheader("Age Group Breakdown by Gender")
    st.caption(
        "Uses the same valid First Visit and Repeat Visit records as the CFS / Site "
        "table above, with age group replacing CFS / site as the row dimension."
    )
    render_table(
        table_with_total(fvr, ["age_group"], ["visit_type", "gender_clean"]),
        "First Visit and Repeat Visit by Age Group and Gender",
        "age_gender",
    )
    achart = count_table(fvr, ["age_group", "visit_type"], order_col="age_group")
    bar_chart(
        achart,
        "age_group",
        "Count",
        "visit_type",
        "First and repeat visits by age group",
        category_orders={
            "age_group": AGE_GROUP_ORDER,
            "visit_type": ["First visit", "Repeat visit"],
        },
    )

    st.divider()
    st.subheader(
        "Disability Prevalence by Age Group for First Visit and Repeat Visit by CFS / Site"
    )
    st.caption(
        "Uses the same valid First Visit and Repeat Visit records as the CFS / Site "
        "analysis, then retains only children recorded as living with disability."
    )
    disability_fvr_age = fvr[
        fvr["disability_status_clean"].astype(str).eq("Yes")
    ].copy()
    if disability_fvr_age.empty:
        st.info(
            "No children with disability and a valid First/Repeat Visit response "
            "match the current filters."
        )
    else:
        render_table(
            table_with_total(
                disability_fvr_age,
                ["age_group"],
                ["visit_type", "gender_clean"],
            ),
            "Disability: First and Repeat Visits by Age Group and Gender",
            "disability_first_repeat_age_gender",
        )

elif section == "Games & Activities":
    st.subheader("Games / Activities Engagement")
    render_table(table_with_total(game_context, ["game_clean"], ["gender_clean"]), "Games Activities by Gender", "games")
    gchart = count_table(game_context, ["game_clean", "gender_clean"])
    gchart = add_top_n_control(gchart, "game_clean", "games_chart", default=15)
    bar_chart(gchart, "Count", "game_clean", "gender_clean", "Games / activities by gender", horizontal=True, height=680)

    st.divider()
    st.subheader("Take 5 Activities for Participants with Disabilities")
    st.caption(
        "Uses the same activity mentions as the Games / Activities Engagement "
        "table, retaining the harmonized Take 5 activity and only participants "
        "recorded as living with disability. The separate integration Yes/No "
        "field is not used for this table."
    )
    disability_participant_ids = set(
        filtered.loc[
            filtered["disability_status_clean"].astype(str).eq("Yes"),
            "record_id",
        ].astype(str)
    )
    take5_disability_games = game_context[
        game_context["record_id"].astype(str).isin(disability_participant_ids)
        & game_context["game_clean"].astype(str).eq("Take 5")
    ].copy()
    if take5_disability_games.empty:
        st.info(
            "No harmonized Take 5 activity mentions for participants with "
            "disabilities match the current filters."
        )
    else:
        render_table(
            table_with_total(
                take5_disability_games,
                ["game_clean"],
                ["gender_clean"],
            ),
            "Take 5 Activities for Participants with Disabilities by Gender",
            "take5_disability_games",
        )

    st.divider()
    st.subheader("Take 5 Integration into Play Sessions")
    st.caption("This uses the new survey field: 'Was Take 5 activities integrated into play sessions?'")
    take5_scope = filtered[filtered["take5_integrated_clean"].astype(str).isin(["Yes", "No"])].copy()
    render_table(
        table_with_total(take5_scope, ["take5_integrated_clean"], ["gender_clean"]),
        "Take 5 Integration by Gender",
        "take5_integration_gender",
    )
    take5_chart = count_table(take5_scope, ["take5_integrated_clean", "gender_clean"])
    bar_chart(
        take5_chart,
        "Count",
        "take5_integrated_clean",
        "gender_clean",
        "Take 5 integration into play sessions by gender",
        horizontal=True,
        height=360,
        category_orders={"take5_integrated_clean": YES_NO_ORDER, "gender_clean": GENDER_ORDER},
    )

    st.divider()
    st.subheader("Multiple Games / Activities per Child")
    st.caption("Activity totals are multi-response mention counts. The tables below show both Records and Total mentions so they reconcile with the Games / Activities totals.")
    game_dist, game_combos = multi_response_tables(game_context, "game_clean", "Game / activity")
    gdist_col, gcombo_col = st.columns([1, 1.8])
    with gdist_col:
        render_table(game_dist, "Game Activity Selection Count", "game_selection_count")
    with gcombo_col:
        render_table(game_combos, "Most Common Game Activity Combinations", "game_combinations")

elif section == "Protection & Support":
    st.subheader("Disability Prevalence")
    render_table(table_with_total(filtered, ["disability_status_clean"], ["gender_clean"]), "Disability Prevalence", "disability")
    dchart = count_table(filtered, ["disability_status_clean", "gender_clean"])
    bar_chart(dchart, "Count", "disability_status_clean", "gender_clean", "Disability prevalence by gender", horizontal=True, height=360)

    st.divider()
    st.subheader("Disability Prevalence by Age Group")
    st.caption(
        "Includes only children recorded as living with disability. Children "
        "recorded as No or with a missing disability response are excluded."
    )
    disability_age = filtered[
        filtered["disability_status_clean"].astype(str).eq("Yes")
    ].copy()
    if disability_age.empty:
        st.info("No children with disability match the current filters.")
    else:
        render_table(
            table_with_total(disability_age, ["age_group"], ["gender_clean"]),
            "Disability Prevalence by Age Group and Gender",
            "disability_age_gender",
        )

    st.divider()
    st.subheader("Type of Disability")
    st.caption("This view is limited to children recorded as living with a disability. Missing disability-type entries are excluded from the analytical table and chart, but remain visible through Data Quality completeness checks.")
    disability_yes = filtered[filtered["disability_status_clean"].astype(str).eq("Yes")].copy()
    disability_type_data = disability_yes[
        disability_yes["disability_type_display"].notna()
        & disability_yes["disability_type_display"].astype(str).str.strip().ne("")
        & ~disability_yes["disability_type_display"].astype(str).isin([MISSING, REVIEW])
    ].copy()
    render_table(
        table_with_total(disability_type_data, ["disability_type_display"], ["gender_clean"]),
        "Type of Disability by Gender",
        "disability_type_gender",
    )
    disability_type_chart = count_table(disability_type_data, ["disability_type_display", "gender_clean"])
    disability_type_chart = add_top_n_control(disability_type_chart, "disability_type_display", "disability_type_chart", default=10)
    bar_chart(
        disability_type_chart,
        "Count",
        "disability_type_display",
        "gender_clean",
        "Type of disability by gender",
        horizontal=True,
        height=500,
    )

    st.divider()
    st.subheader("Nature of Issues Reported")
    render_table(table_with_total(issue_context, ["issue_clean"], ["gender_clean"]), "Issues Reported by Gender", "issues")
    ichart = count_table(issue_context, ["issue_clean", "gender_clean"])
    ichart = add_top_n_control(ichart, "issue_clean", "issues_chart", default=15)
    bar_chart(ichart, "Count", "issue_clean", "gender_clean", "Issues reported by gender", horizontal=True, height=720)

    st.divider()
    st.subheader("Multiple Issues per Child / Record")
    st.caption("Issue totals are multi-response mention counts. The tables below show both Records and Total mentions so they reconcile with the Nature of Issues Reported totals.")
    issue_dist, issue_combos = multi_response_tables(issue_context, "issue_clean", "Issue")
    idist_col, icombo_col = st.columns([1, 1.8])
    with idist_col:
        render_table(issue_dist, "Issue Selection Count", "issue_selection_count")
    with icombo_col:
        render_table(issue_combos, "Most Common Issue Combinations", "issue_combinations")

    st.divider()
    st.subheader("Support Offered")
    render_table(table_with_total(support_context, ["support_clean"], ["gender_clean"]), "Support Offered by Gender", "support")
    schart = count_table(support_context, ["support_clean", "gender_clean"])
    bar_chart(schart, "Count", "support_clean", "gender_clean", "Support offered by gender", horizontal=True, height=420)

elif section == "Referrals":
    st.subheader("Referral Rate by CFS / Site")
    rrs = filtered.groupby("cfs_clean", observed=False).agg(Total_Records=("record_id", "count"), Referrals_Yes=("referral_made_clean", lambda s: s.astype(str).eq("Yes").sum())).reset_index()
    if not rrs.empty:
        rrs["Referrals_No"] = rrs["Total_Records"] - rrs["Referrals_Yes"]
        rrs["Referral Rate %"] = (rrs["Referrals_Yes"] / rrs["Total_Records"] * 100).round(1)
        rrs = rrs.sort_values("Referral Rate %", ascending=True)
    bar_chart(rrs.rename(columns={"Referral Rate %": "Count"}), "Count", "cfs_clean", title="Referral Rate (%) by CFS / Site", horizontal=True, height=max(360, len(rrs) * 34 + 100))
    render_table(rrs.rename(columns={"cfs_clean": "CFS / Site", "Total_Records": "Total Records", "Referrals_Yes": "Referrals Yes", "Referrals_No": "Referrals No"}), "Referral Rate by CFS Site", "referral_rate", precision=1)

    st.divider()
    st.subheader("Referral Destination")
    referrals = filtered[filtered["referral_made_clean"].astype(str).eq("Yes")].copy()
    render_table(table_with_total(referrals, ["referral_destination_grouped"], ["gender_clean"]), "Referral Destination by Gender", "ref_dest")
    rdchart = count_table(referrals, ["referral_destination_grouped", "gender_clean"])
    bar_chart(rdchart, "Count", "referral_destination_grouped", "gender_clean", "Referral destination by gender", horizontal=True, height=420)

    st.divider()
    st.subheader("External Referral Agencies")
    ext = referrals[~referrals["external_referral_agency_clean"].astype(str).isin([MISSING, "Unknown", REVIEW])].copy()
    render_table(table_with_total(ext, ["external_referral_agency_clean"], ["gender_clean"]), "External Referral Agencies by Gender", "ext_agencies")

elif section == "Data Quality":
    st.subheader("Data Quality & Harmonisation Review")
    st.info("Records where consent was 'No' are excluded because they typically contain skipped form fields.")

    st.markdown("#### Analysis Schema Readiness")
    st.caption("This checks whether the transformed analysis-column structure required by the dashboard is present after raw-to-analysis column harmonisation.")
    render_table(schema_readiness_table(df), "Analysis Schema Readiness", "schema_readiness")
    schema_extras = list(df.attrs.get("unmapped_source_columns", []))
    st.caption(f"Schema contract: {SCHEMA_CONTRACT_VERSION}")
    if schema_extras:
        st.info(
            f"Kobo contains {len(schema_extras):,} additional/unmapped field(s). They are isolated from existing analytics, "
            "so they cannot shift contracted dashboard fields. Review and map them only when they should become analytical fields."
        )
        render_table(
            pd.DataFrame({"Additional / unmapped Kobo field": schema_extras}).set_index("Additional / unmapped Kobo field"),
            "Kobo Schema Additions",
            "kobo_schema_additions",
        )
    else:
        st.success("No unmapped Kobo fields were detected against the current schema contract.")

    st.markdown("#### Gender Classification Audit")
    st.caption("This reconciles the exact Kobo response/code with the category used in every dashboard table and chart.")
    if "child_gender" in filtered.columns:
        gender_audit_source = filtered[["child_gender", "gender_clean"]].copy()
        gender_audit_source["Kobo raw response"] = gender_audit_source["child_gender"].fillna("<blank>").astype(str)
        gender_audit_source["Dashboard category"] = gender_audit_source["gender_clean"].astype(str)
        gender_audit = (
            gender_audit_source.groupby(["Kobo raw response", "Dashboard category"], observed=True)
            .size()
            .reset_index(name="Records")
            .sort_values("Records", ascending=False)
        )
        render_table(
            gender_audit.set_index(["Kobo raw response", "Dashboard category"]),
            "Raw Kobo Gender Responses → Dashboard Categories",
            "gender_classification_audit",
        )
        unresolved_gender = gender_audit[gender_audit["Dashboard category"].eq(MISSING)]
        if not unresolved_gender.empty:
            st.warning(
                f"{int(unresolved_gender['Records'].sum()):,} gender responses remain unmapped. "
                "Use the raw values in this table to extend clean_gender() without guessing."
            )

    st.markdown("#### Support Classification Audit")
    st.caption("This reconciles Kobo's raw Support offered response/code with the support categories used in tables and charts.")
    if "support_offered_text" in filtered.columns:
        support_audit_source = filtered[["support_offered_text", "support_combined"]].copy()
        support_audit_source["Kobo raw support response"] = support_audit_source["support_offered_text"].fillna("<blank>").astype(str)
        support_audit_source["Dashboard support categories"] = support_audit_source["support_combined"].fillna(MISSING).astype(str)
        support_audit = (
            support_audit_source.groupby(["Kobo raw support response", "Dashboard support categories"], observed=True)
            .size()
            .reset_index(name="Records")
            .sort_values("Records", ascending=False)
        )
        render_table(
            support_audit.set_index(["Kobo raw support response", "Dashboard support categories"]),
            "Raw Kobo Support Responses → Dashboard Categories",
            "support_classification_audit",
        )
        suspected_pfa = support_audit[
            support_audit["Kobo raw support response"].map(norm_text).str.contains(r"\bpfa\b|first aid", regex=True)
        ]
        unresolved_pfa = suspected_pfa[
            ~suspected_pfa["Dashboard support categories"].str.contains("Psychological First Aid", case=False, na=False)
        ]
        if not unresolved_pfa.empty:
            st.error(
                f"{int(unresolved_pfa['Records'].sum()):,} response(s) appear to contain PFA/First Aid but are not classified as Psychological First Aid."
            )

    q1, q2, q3, q4 = st.columns(4)
    with q1:
        metric_card("Missing location", f"{filtered['location_clean'].astype(str).eq(MISSING).sum():,}", tone="danger")
    with q2:
        metric_card("Missing CFS / site", f"{filtered['cfs_clean'].astype(str).eq(MISSING).sum():,}", tone="danger")
    with q3:
        metric_card("Missing age", f"{filtered['age_clean'].isna().sum():,}", tone="warning")
    with q4:
        metric_card("Raw records loaded", f"{raw_count:,}", tone="primary")

    fields = [
        ("Child name", "child_name"),
        ("Gender", "gender_clean"),
        ("Age", "age_clean"),
        ("Disability status", "disability_status_clean"),
        ("First visit", "first_visit_clean"),
        ("Referral made", "referral_made_clean"),
        ("Settlement / camp", "settlement_clean"),
        ("Location", "location_clean"),
        ("CFS / site", "cfs_clean"),
        ("Games played", "games_played_clean"),
        ("Take 5 integrated", "take5_integrated_clean"),
        ("Staff", "staff_clean"),
    ]
    rows = []
    total = len(filtered)
    for label, col in fields:
        filled = int((filtered[col].notna() & filtered[col].astype(str).ne("") & filtered[col].astype(str).ne(MISSING)).sum()) if col in filtered.columns else 0
        rows.append({"Field": label, "Records Filled": filled, "Total Records": total, "Completeness %": round(filled / total * 100, 1) if total else 0})
    completeness = pd.DataFrame(rows).sort_values("Completeness %")
    fig = px.bar(completeness, x="Completeness %", y="Field", orientation="h", title="Field Completeness (%)", text="Completeness %", color="Completeness %", color_continuous_scale=[(0, "#dc2626"), (0.6, "#f59e0b"), (0.8, "#16a34a"), (1, "#15803d")], range_color=[0, 100])
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_layout(height=max(360, len(completeness) * 38 + 90), margin=dict(l=12, r=24, t=58, b=18), plot_bgcolor="white", paper_bgcolor="white", coloraxis_showscale=False, title_font=dict(size=18, color="#172554"), font=dict(size=13, color="#111827"))
    fig.update_xaxes(range=[0, 105], showgrid=True, gridcolor="#e9eef5", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True, "responsive": True})
    render_table(completeness.set_index("Field"), "Field Completeness Summary", "completeness", precision=1)

elif section == "Raw Data":
    st.subheader("Filtered Raw Data")
    display_cols = [
        "record_id", "date", "settlement_clean", "location_clean", "cfs_clean", "staff_clean", "gender_clean", "age_clean", "age_group", "disability_status_clean", "first_visit_source_raw", "first_visit_clean", "referral_made_clean", "issues_combined", "support_combined", "games_played_clean", "take5_integrated_clean",
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)
    st.caption("Privacy protection: names, telephone numbers, individual identifiers, device/user identifiers, GPS data, notes, and UUID fields are excluded from this download.")
    prepare_export = st.checkbox("Prepare privacy-safe CSV download", value=False, help="CSV bytes are created only when requested, which keeps ordinary page opening faster.")
    if prepare_export:
        safe_filtered_export = privacy_safe_export(filtered)
        st.download_button("⬇ Download privacy-safe filtered data CSV", safe_filtered_export.to_csv(index=False).encode("utf-8"), file_name="filtered_cfs_data_privacy_safe.csv", mime="text/csv", use_container_width=True)


# Footer
render_app_footer()
