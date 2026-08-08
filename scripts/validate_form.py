"""
Validates that FEDI_Baseline_v1.xlsx converts cleanly to XForm XML.
Mirrors the eHA assessment's part2_q3_odk/scripts/validate_form.py.

Run: python validate_form.py
Writes: ../form/FEDI_Baseline_v1.xml
        ../form/validation_result.txt
Exits non-zero on failure.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
FORM = HERE.parent / "form" / "FEDI_Baseline_v1.xlsx"
XML_OUT = HERE.parent / "form" / "FEDI_Baseline_v1.xml"
RESULT_OUT = HERE.parent / "form" / "validation_result.txt"


def main():
    try:
        import pyxform
        version = getattr(pyxform, "__version__", "unknown")
    except ImportError:
        msg = ("pyxform is not installed in this environment. Install with "
               "`pip install pyxform` and re-run before deployment.")
        print(msg, file=sys.stderr)
        RESULT_OUT.write_text("NOT RUN: " + msg + "\n")
        sys.exit(1)

    from pyxform.xls2xform import xls2xform_convert

    warnings = []
    try:
        xls2xform_convert(str(FORM), str(XML_OUT))
    except Exception as e:
        result = f"FAIL: pyxform {version}\n{e}\n"
        RESULT_OUT.write_text(result)
        print(result, file=sys.stderr)
        sys.exit(1)

    result = f"PASS: pyxform {version}\n"
    RESULT_OUT.write_text(result)
    print(result)


if __name__ == "__main__":
    main()
