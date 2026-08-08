# FEDI Baseline/Needs Assessment — XLSForm

Built from `Baseline_Needs_Assessment_Questionnaire_for_Kobo_Collect.docx`
(76 questions), applying the same validation discipline used for the eHA
technical assessment's Section 3 questionnaire: bilingual (English/Hausa),
explicit skip logic, values derived rather than re-asked wherever that
removes a double-entry disagreement risk, and every deviation from the
source document's literal field list stated plainly rather than silent.

**This is a separate, unrelated project from the eHA technical assessment
repo** - kept in its own repo deliberately so the two don't get tangled
together in anyone's review.

## Build

    python scripts/build_xlsform.py    # writes form/FEDI_Baseline_v1.xlsx
    python scripts/validate_form.py    # converts to XML, fails loudly if invalid

Both wired into `.github/workflows/ci.yml`.

## Status

Built and structurally self-consistent (no dangling field references, all
groups balanced, every select type has a matching choice list - checked
programmatically). **Not yet run through a real ODK Validate/pyxform** -
no network access in the authoring sandbox. See `DESIGN_NOTES.md` for
every deviation from the source document and the reasoning behind it, and
please run `scripts/validate_form.py` (or ODK Validate directly) before
deployment.

## Files

- `form/FEDI_Baseline_v1.xlsx` — the XLSForm (built by `build_xlsform.py`,
  not hand-edited)
- `DESIGN_NOTES.md` — every deviation from the source document's literal
  field list, every inferred skip-logic decision, and every judgement call,
  with reasoning
