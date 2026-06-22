from app.policies.spg_pl import _normalize_common


def test_normalize_common_snaps_dates_and_zip():
    rows = [{"Effective Date": "06222026", "Mailing Zip": "22030-1234", "Name": "Acme"}]
    out = _normalize_common(rows)
    assert out[0]["Effective Date"] == "06/22/2026"
    assert out[0]["Mailing Zip"] == "22030"
    assert out[0]["Name"] == "Acme"  # untouched open field


def test_normalize_common_is_idempotent():
    rows = [{"DOB": "01/02/1990", "Zip": "22030"}]
    once = _normalize_common([dict(r) for r in rows])
    twice = _normalize_common([dict(r) for r in once])
    assert once == twice
