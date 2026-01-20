import sys
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))

from formula_engine import FormulaEngine, FormulaError


def test_textjoin_ignores_empty_values():
    result = FormulaEngine.evaluate('=TEXTJOIN("-"; TRUE; "A"; ""; "B"; None; "C")')
    assert result == "A-B-C"


def test_textjoin_keeps_empty_values_when_requested():
    result = FormulaEngine.evaluate('=TEXTJOIN(", "; FALSE; "A"; ""; "B")')
    assert result == "A, , B"


def test_textjoin_flattens_nested_sequences():
    result = FormulaEngine.evaluate('=TEXTJOIN(""; TRUE; SPLIT("AA BB"; " "); "CC")')
    assert result == "AABBCC"


def test_arrayformula_passes_through_sequences():
    result = FormulaEngine.evaluate('=ARRAYFORMULA(SPLIT("AA BB"; " "))')
    assert result == ["AA", "BB"]


def test_arrayformula_flattens_multiple_arguments():
    result = FormulaEngine.evaluate('=ARRAYFORMULA("A"; SPLIT("B C"; " "))')
    assert result == ["A", "B", "C"]


def test_default_function_is_restored_if_missing():
    removed = FormulaEngine.FUNCTIONS.pop("ARRAYFORMULA", None)
    try:
        result = FormulaEngine.evaluate("=ARRAYFORMULA(\"value\")")
        assert result == "value"
    finally:
        if removed is not None:
            FormulaEngine.FUNCTIONS["ARRAYFORMULA"] = removed


def test_regexreplace_performs_substitution():
    result = FormulaEngine.evaluate(r'=REGEXREPLACE("AA-11-BB-22"; "\\d+"; "#")')
    assert result == "AA-#-BB-#"


def test_regexreplace_invalid_pattern_raises_error():
    with pytest.raises(FormulaError):
        FormulaEngine.evaluate('=REGEXREPLACE("text"; "("; "x")')


def test_textjoin_with_arrayformula_and_left_vectorizes_inputs():
    context = {
        "brand": "Fujifilm",
        "model": "X T3",
        "film_type": "Camera",
    }
    formula = (
        '=TEXTJOIN(""; TRUE; ARRAYFORMULA(LEFT(SPLIT(TRIM({{ brand }}&" "&{{ model }}&" "&{{ film_type }});" ");1)))'
    )
    result = FormulaEngine.evaluate(formula, context)
    assert result == "FXTC"
