from datetime import date

from app.rulebook.primitives import (
    parse_date, format_date_slash, format_date_compact, add_one_year,
    to_number, find_col, find_header_key, tid_value, is_yes, is_no,
    default_test_id,
)


def test_default_test_id_is_ts_zero_padded():
    assert default_test_id(1) == "TS-001"
    assert default_test_id(12) == "TS-012"
    assert default_test_id(1, width=2) == "TS-01"


def test_parse_date_accepts_slash_compact_and_iso():
    assert parse_date("06/22/2026") == date(2026, 6, 22)
    assert parse_date("06222026") == date(2026, 6, 22)
    assert parse_date("2026-06-22") == date(2026, 6, 22)
    assert parse_date("not a date") is None
    assert parse_date(None) is None


def test_format_date_variants_are_distinct():
    d = date(2026, 6, 22)
    assert format_date_slash(d) == "06/22/2026"
    assert format_date_compact(d) == "06222026"


def test_add_one_year_handles_leap_day():
    assert add_one_year(date(2024, 2, 29)) == date(2025, 2, 28)
    assert add_one_year(date(2026, 6, 22)) == date(2027, 6, 22)


def test_to_number_strips_currency_and_keeps_int_vs_float():
    assert to_number("$25,000") == 25000
    assert to_number("1234.50") == 1234.5
    assert to_number("") is None
    assert to_number("abc") is None


def test_find_col_matches_all_keywords_lowercased():
    row = {"Effective Date": "x", "Total Value ($)": "y"}
    assert find_col(row, "effective") == "Effective Date"
    assert find_col(row, "total", "value") == "Total Value ($)"
    assert find_col(row, "missing") is None


def test_find_header_key_matches_any_candidate():
    row = {"Expiry Date": "x"}
    assert find_header_key(row, ["expiration", "expiry"]) == "Expiry Date"
    assert find_header_key(row, ["nope"]) is None


def test_tid_value_reads_test_id_column():
    assert tid_value({"Test ID": " TS-01 "}) == "TS-01"
    assert tid_value({"Other": "x"}) == ""


def test_yes_no_predicates():
    assert is_yes("Yes") and is_yes("y") and is_yes("1")
    assert is_no("No") and is_no("n") and is_no("0")
    assert not is_yes("No") and not is_no("Yes")


def test_date_paths_unified_to_slash():
    # S2 decision (QA sign-off): both the IMS handler path and the generic
    # llm_service path emit MM/DD/YYYY. Previously the generic path emitted
    # compact "06222026", which shipped Effective/Expiration as "07272026"
    # (DF-IM-001 sibling). Guard against silent regression to compact.
    from app.policies import ims as ims_mod
    from app import llm_service as llm_mod
    d = date(2026, 6, 22)
    assert ims_mod._format_mmddyyyy(d) == "06/22/2026"
    assert llm_mod._format_mmddyyyy(d) == "06/22/2026"
