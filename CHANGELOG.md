# Changelog

## 2026-09-22 (2)

- Replaced the old `scripts/`/`templates/`/`config.py` layout with two
  top-level project folders, `Top 10 automation/` and
  `Annexure K automation/`, each self-contained with its own generator
  script, template(s), and gitignored `db_secret.py`.
- Added the Annexure K generator (`generate_annexure_k.py`): fills the DBE's
  Annexure K (achievement-level distribution) template for both the "8-12"
  and "R-7" groupings straight from the database, Composite and Exam-mark
  blocks included. See `Annexure K automation/README.md` for the full rule
  set and known edge cases.
- The web interface (`Top 10 automation/webapp/`) now drives both
  generators from one page, with a TOP 10 / ANNEXURE K toggle at the top.
- `Launch Top 10 Club.vbs` now stops any earlier copy of the server before
  starting a new one, so relaunching after an update always runs the
  current code instead of silently talking to a stale background process.

## 2026-09-22

- Added a local web interface for running Top 10 generation — pick the database
  file and output folder from a browser page and click Generate, instead of
  editing the script and running it from the terminal.
- Fixed Grade 12 being skipped from the per-grade Top 10 report — it was
  missing from the grade range the generator looped over.
