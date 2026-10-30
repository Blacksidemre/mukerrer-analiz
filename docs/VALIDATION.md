# Validation record — 2026-10-31

## Executed in the build environment

- Both generated XLSX archives opened and passed ZIP integrity and XML parsing.
- Six expected sheets, filters, conditional row colors, and source-row metadata verified structurally.
- 6,300 formulas per dynamic workbook tokenized successfully. This checks syntax structure only; it is not Excel calculation.
- 370 source records retained; cached control values compared to all source columns without differences.
- Independently grouped supplied synthetic records using the four initial criteria: **180 unique, 90 original, 100 surplus duplicates, 90 repeated groups**. Reconciliation: **370 = 180 + 90 + 100**.
- Formula caches carry these expected results. The cache was authored during workbook generation and is **not evidence of Excel executing formulas**.
- Four-sheet static output opened; no cell formulas remain. It is an analysis snapshot, not dynamic.
- No external-link parts or VBA binaries in delivered XLSX files.
- See `artifact-validation.json` for hashes and actual formula counts. The uploaded source and its analysis are not committed to GitHub.

## Not executed / not established

- Excel LTSC 2024 formula recalculation and native sorting behavior.
- VBA compilation and the 17 assertions in `TestleriCalistir`.
- Import/criteria change/filter/sort/side-by-side/export end-to-end tests in desktop Excel.
- 500 / 5,000 / 20,000 / 50,000-row VBA benchmarks; no speed or responsiveness figures are claimed.
- The PowerShell-based GitHub Actions job has not been observed running at publication time. It checks workbook packaging only.
- No actual Excel screenshot, signed macro binary, or compiled XLSM has been produced.

## Required next validation

On Windows with Excel LTSC 2024, follow `README.tr.md`, import both modules, compile, run the tests and benchmarks, and complete `tests/ACCEPTANCE.md`. Do not describe this review release as production-tested until those checks pass. Packaging scripts used to author/inspect XLSX are not required to run the delivered solution; application logic consists only of Excel formulas and VBA.
