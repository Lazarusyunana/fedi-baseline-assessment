# Design notes

This form follows the same discipline used for the eHA technical
assessment's Section 3 questionnaire: bilingual throughout, explicit skip
logic rather than left implicit, and values derived instead of re-asked
wherever re-asking would only create a chance for two answers to disagree.
That discipline means this form is **not** a literal field-for-field
transcription of `Baseline_Needs_Assessment_Questionnaire_for_Kobo_Collect.docx`
- every place it deviates is listed below, with the reasoning, so nothing
changed silently.

## Deviations from the source document's literal field list

| Source field(s) | Source type | What was built instead | Why |
|---|---|---|---|
| `children_under_two` | integer (asked) | `calculate` = `children_under_six_months + children_6_23_months` | The two components are already asked immediately above it and exactly determine it - asking a third time only invites disagreement to clean up later. |
| `girls_out_of_school` | integer (asked) | `calculate` = `school_age_girls_total - girls_attending_school` | Same reasoning - it is arithmetically fixed by two fields already captured. |
| `girls_school_attendance` | select_one (asked) | `calculate`, derived from `school_age_girls_total` vs `girls_attending_school` | A categorical judgement ("some/all/none attending") sitting next to the exact counts it describes is exactly the class of disagreement the eHA build's constraint register exists to prevent (see that repo's CR-111 and CR-120 for the same reasoning applied to different fields). |
| `vulnerability_summary` | select_multiple (manually completed by enumerator) | `calculate`, computed from ten already-answered fields, shown back to the enumerator via a read-only note for sanity-checking | Asking an enumerator to manually re-derive, from memory, a summary of facts already on the form is the same clerical-recount risk fixed repeatedly in the eHA build. The computed value is still visible (not hidden), so a field-level sanity check remains possible. |
| `respondent_age_group` | select_one (asked, kept as asked) | Asked as specified, but with a soft (non-blocking) note if it disagrees with the age entered at `respondent_age` | NOT auto-derived, unlike the four rows above - see "A gap in the source document" below for why. |

## A gap in the source document, left as a gap rather than silently fixed

`household_head_type` explicitly includes "Child-headed" as an option,
implying a respondent could be under 18. But `respondent_age_group`'s
stated response options are `18-24; 25-34; 35-49; 50 and above` - no
bucket below 18 exists. I did not invent a new bucket to paper over this,
since the options given were explicit and specific; I also did not
auto-derive the field the way the other four rows above were, because
doing so would have required silently deciding what to do with an
under-18 age with no correct bucket to put it in. Instead: `respondent_age`
allows down to 10, `respondent_age_group` is asked as its own question
exactly as specified, and a note flags disagreement only when
`respondent_age >= 18` (below that, the mismatch is the known, named gap,
not a data error to flag). Whoever owns this instrument should decide
explicitly whether to add an under-18 bucket or restrict the survey to
adult respondents - that is a program design decision, not something to
default silently either way.

## Skip logic inferred, not stated verbatim

The source document names six groups explicitly ("pregnant women,
lactating women, caregivers of children under two years, households with
school-age children, households affected by windstorm, or households with
garden space") but several dependent-question relationships within those
groups are implied rather than spelled out. Implemented, with reasoning:

- `pregnancy_status` / `lactating_status`: gated on `respondent_sex='female'`.
- `participation_support_needed`: gated on
  `respondent_participation_difficulty != 'no_difficulty'` - asking what
  support would help only makes sense once some difficulty is reported.
- `child_meals_yesterday` / `child_diet_diversity`: the question text
  itself names "aged 6-23 months," not the whole under-2 group, so both
  are gated on `children_under_two>0 AND youngest_child_age_months>=6`,
  not on `children_under_two>0` alone.
- `tom_brown_prepare`: gated on `tom_brown_awareness='yes'` - the question
  asks whether they can prepare something using local ingredients, which
  presupposes having heard of it.
- Home garden block (`veg_selection_note`, `preferred_vegetables`,
  `veg_choice_reason`): gated on `garden_space` being `'yes'` OR
  `'shared_space'`. This is a judgement call - the source document names
  "households with garden space" as the gating group in general terms but
  does not specify which of the five `garden_space` response codes count.
  `'temporary_space'` and `'not_sure'` were excluded as too uncertain a
  basis for committing seed/tool support to a specific household; this is
  stated as judgement, not drawn from anything in the source document.

## Cross-question consistency (flag, not block)

Two soft, non-blocking notes, matching the eHA build's reasoning for why
plausible-but-imperfect household counts should be surfaced rather than
gated: `female_household_members + male_household_members` vs
`household_size`, and `school_age_girls_total + school_age_boys_total` vs
`children_school_age_total`. A hard equality constraint on either would
pressure the enumerator into silently changing whichever number is more
convenient rather than genuinely resolving the discrepancy.

## Hard constraints (block)

- `preferred_vegetables`: `count-selected(.) >= 2 and count-selected(.) <= 3`
  - the explicit instruction in the source document's intro paragraph.
- `girls_attending_school <= school_age_girls_total`,
  `boys_attending_school <= school_age_boys_total` - a logical impossibility
  (cannot have more children attending than exist), not a judgement call.
- Integer range constraints (household size, meal counts, age fields, etc.)
  are candidate judgement anchored to plausible household/demographic
  ranges, not drawn from an external standard - stated as such, consistent
  with how range constraints without a cited source were flagged throughout
  the eHA build.

## Not implemented

- `respondent_name`: kept as a required text field per the source document,
  but the hint recommends initials only, consistent with the data-
  minimisation reasoning in the eHA build's `data_protection.md`. Not
  enforced as a constraint, since a legitimate project confidentiality
  protocol may already govern this outside the form itself.
- `gps_location`: deliberately left NOT required, directly following the
  source document's own hint ("Collect only if safe and appropriate") -
  making it required would contradict that hint outright.

## What was not attempted here

Unlike the eHA Section 3 build, this form has no repeating structure (no
household roster, no per-child module) - every question is a single
household-level or respondent-level answer, so there was no need for
`indexed-repeat()`, ODK Entities, or any of the more involved mechanisms
used there. The validation discipline (bilingual, explicit skip logic,
derive-don't-re-ask, flag-vs-block reasoning) carried over; the mechanisms
did not need to, and adding repeat/Entities machinery where the source
document does not call for it would have been unjustified complexity.

## Still unverified

Same caveat as the eHA build: no `pyxform`/network access in the sandbox
this was authored in. `scripts/validate_form.py` (mirrors the eHA build's
script of the same name) is provided but has not been run. Please run it
- or ODK Validate directly - before deploying, and report back anything it
flags the same way you did for the eHA form's Entities issue.
