# Changelog

## 2026-10-12

- Added `MCP server/`: an MCP server so Claude Code can query the SASAMS
  database directly. Ready-made tools cover subject results, term
  comparisons, at-risk learners, attendance, single-learner marks and the
  existing Top 10 / Annexure K generators; `run_select` handles anything
  else. Read-only and SELECT-only; learners are identified by internal ID,
  personal columns are blocked, and named lists for teachers are written to
  a local Excel file instead of being returned to Claude.
- Added a fake SASAMS database and an end-to-end test
  (`MCP server/tests/test_server.py`) that checks no names leak and unsafe
  queries are refused.

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
