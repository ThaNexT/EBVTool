import pytest
from evaluator import find_best_class

def test_arsen_feststoff_bm0():
    klasse, gw, fn, order = find_best_class("Arsen", "mg/kg", 5.0, "", 0.1, "BM_0_Sand")
    assert klasse == "BM-0"
    assert gw == 10

def test_arsen_feststoff_bmf0():
    klasse, gw, fn, order = find_best_class("Arsen", "mg/kg", 35.0, "", 0.1, "BM_0_Sand")
    assert klasse == "BM-F0*"
    assert gw == 40

def test_arsen_feststoff_deponie():
    klasse, gw, fn, order = find_best_class("Arsen", "mg/kg", 200.0, "", 0.1, "BM_0_Sand")
    assert "> BM-F3" in klasse

def test_arsen_kleiner_operator():
    klasse, gw, fn, order = find_best_class("Arsen", "mg/kg", 10.0, "<", 0.1, "BM_0_Sand")
    assert klasse == "BM-0"

def test_kein_messwert():
    klasse, gw, fn, order = find_best_class("Arsen", "mg/kg", None, "", 0.1, "BM_0_Sand")
    assert klasse == "Kein Messwert"

def test_ph_im_bereich():
    klasse, gw, fn, order = find_best_class("pH-Wert", "-", 7.5, "", 0.1, "BM_0_Sand")
    assert "BM" in klasse

def test_toc_klammerwert():
    klasse_niedrig, gw_niedrig, _, _ = find_best_class("Arsen", "µg/l", 10.0, "", 0.1, "BM_0_Sand")
    klasse_hoch, gw_hoch, _, _ = find_best_class("Arsen", "µg/l", 10.0, "", 0.6, "BM_0_Sand")
    assert gw_hoch >= gw_niedrig