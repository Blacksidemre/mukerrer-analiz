# Duplicate Audit Studio · Excel Edition

**Deterministic, multi-column duplicate auditing inside Microsoft Excel LTSC 2024.**

[![Workbook integrity](https://github.com/Blacksidemre/m-kerrer-analiz/actions/workflows/integrity.yml/badge.svg)](https://github.com/Blacksidemre/m-kerrer-analiz/actions/workflows/integrity.yml)

[Turkish guide](README.tr.md) · [Download workbook](deliverables/Duplicate_Audit_Studio.xlsx) · [Validation report](docs/VALIDATION.md)

Designed for analysts who need to identify duplicate records without deleting data. Select columns with 1/0 switches, inspect full-row status colors, trace every group to its first source row, and export an auditable snapshot. There is no web application, backend, database, AI API, or JavaScript dependency.

> **Release status:** formula workbook and VBA source delivered. Desktop Excel is unavailable in the build environment. VBA execution, Excel formula recalculation, end-to-end behavior, and 50,000-row timings require validation in Excel LTSC 2024. This is a review release, not a certified production release. The XLSX contains no embedded VBA; import the two modules to create your XLSM.

## Choose a workflow

| Workflow | Delivered | Recalculation | Intended scale |
|---|---|---|---|
| Formula sample | `deliverables/Duplicate_Audit_Studio.xlsx` | Criteria changes recalculate formulas in Excel | Fixed 370-row, 12-column demonstration |
| VBA analysis | `vba/DuplicateAudit.bas` | Run `AnalizEt` after changing data/criteria | Up to 50,000 rows / 100 columns; unbenchmarked |
| VBA tests | `vba/AuditTests.bas` | Run explicitly in Excel | Assertions and 4 benchmark sizes |
| Export | `ExcelAktar` creates XLSX | Static values, not a dynamic analysis engine | All analyzed rows |

Formula mode uses quadratic exact comparisons. Do not extend it to 50,000 rows. VBA uses dictionaries with expected linear processing in rows × selected columns; this is an algorithmic property, not a measured performance claim.

## Quick start

1. Open the XLSX in Excel LTSC 2024. Set calculation to **Automatic**. Use **Ctrl+Alt+F9** to recalculate fully.
2. Open `02_Kriterler`. Set `B7:B18` to 1 or 0. Initial selection: file code, debtor code, product, principal.
3. Inspect `03_Kontrol`. Use header filters for status or group. Sort the **whole table**, never one column alone.
4. Check `04_Ozet`: 370 = 180 unique + 90 originals + 100 surplus duplicates. There are 90 duplicate groups.

For arbitrary file imports, exports, options and large datasets, install VBA:

1. Save the workbook as **Excel Macro-Enabled Workbook (.xlsm)** in a local folder.
2. Press **Alt+F11 → File → Import File** and import both `.bas` files from `vba/`.
3. Choose **Debug → Compile VBAProject**. Return to Excel, **Alt+F8 → Kurulum**.
4. Enable only reviewed macros according to your organization's policies. No access to the VBA project object model is required by the application.
5. Run **TestleriCalistir**, then **PerformansOlc**. Results appear in `05_Testler`.

The setup macro creates six functional buttons. Other macros are available through Alt+F8. If the file is renamed after setup, rerun `Kurulum` to bind buttons to the new filename.

## Import → select → analyze → inspect → export

- **DosyaYukle**: select an XLSX and a sheet number. First row is the header. The previous dataset is replaced only after a confirmation; old results are cleared. No records within the new dataset are removed, including interior blank rows. Duplicate/empty headers get unique names.
- **CSV**: use Excel's **Data → From Text/CSV**. Choose UTF-8 or Windows-1254 and comma/semicolon/tab explicitly. Set code columns to **Text**, load into a sheet, save as XLSX, then import. The macro deliberately does not guess CSV encodings or convert code columns automatically.
- **02_Kriterler**: every imported column appears with a 1/0 flag. `F7` trims edge ASCII spaces and `F8` enables Turkish case folding for VBA only. Defaults are exact.
- **AnalizEt**: analyzes all records using active criteria. No criteria is an error. Results include status, group, group size, original reference and source row. Source row numbering starts at 2, after the header.
- **GrubuIncele**: filters by exact group ID. Empty input clears the filter. Header dropdowns also support status filtering and text search.
- **YanYanaKarsilastir**: puts up to 20 members of a selected group side by side in `06_Karsilastirma`. The complete group remains available in the results table.
- **ExcelAktar**: exports the four analysis sheets as a values-only snapshot. Confirm that you have rerun analysis after edits. Output retains filters and conditional colors. It does not recalculate when exported criteria change.

## Matching contract

`TEKİL` = group of one. `ORİJİNAL` = first source record in a repeated group. `MÜKERRER` = every later record. The source sheet order, not display order, determines the original.

Keys are typed, length-prefixed tuples. `AB + C` differs from `A + BC`. Strings and numbers differ; text `001` remains different from text `1`. Empty strings and blank cells match. Wildcards are literal. Defaults preserve case and whitespace. No fuzzy matching, rounding of amounts, or Unicode canonical normalization is performed.

Excel dates are serial numbers; date formatting does not create a separate value type. Numeric `1` displayed with `0000` is still numeric `1`. Store identifiers as text before import. Excel's 15-significant-digit numeric precision cannot recover digits already lost in a source workbook. Formula cells are imported as their current values; calculate and save source workbooks before import. Source formulas, comments, charts, merged layouts and links are not archived as executable objects.

VBA uses binary dictionary comparison. Optional Turkish folding explicitly maps `I → ı` and `İ → i`. The formula demonstration always uses case-sensitive `EXACT`, independent of VBA options. Group IDs use the first source row (`GRP-000002`), unique within one analysis, not globally across files.

Sorting cannot change stored VBA original references. Formula mode uses immutable source-row cells and indexed source reads. Use Excel's whole-table sort and retain all metadata columns. Do not sort/edit the raw sheet after analysis without rerunning it.

## Workbook layout

| Sheet | Role |
|---|---|
| `00_Baslangic` | Navigation and macro buttons after installation |
| `01_Ham_Veri` | Source values, headers and display formats |
| `02_Kriterler` | Per-column flags and VBA options |
| `03_Kontrol` | Every source record plus audit metadata |
| `04_Ozet` | Counts, duplicate ratio and reconciliation |
| `05_Testler` | Actual desktop test status and benchmark results |
| `06_Karsilastirma` | Created on demand for comparison |

The workbook uses a restrained green palette, frozen headers, autofilters and full-row green/red conditional formatting. This is a desktop Excel UI, not a responsive website. Screenshots are intentionally not supplied: no desktop Excel render was captured in the build environment.

## Validation

See [the report](docs/VALIDATION.md) and [desktop acceptance checklist](tests/ACCEPTANCE.md). CI verifies ZIP/XML structure, expected worksheets, formulas, absence of external links and macro binaries in the XLSX, and module files. **CI does not run Excel, compile VBA or establish behavioral correctness.**

Run `TestleriCalistir` for matching assertions. `PerformansOlc` measures the array engine at 500, 5,000, 20,000 and 50,000 rows on your computer and writes seconds to the workbook. It excludes import, formatting, full Excel interaction, and export; measure those separately with the acceptance checklist. No measured latency or UI responsiveness guarantee is claimed.

## Data handling

All application processing happens inside Excel. The VBA has no network calls, telemetry or external AI dependencies. Imported sources are opened read-only, with macro execution disabled and link updates disabled. Existing source files are never saved by the importer. Treat untrusted workbooks according to organizational policy; this is not an Excel security sandbox. Formula-looking strings are written as literals.

The repository sample is independently generated synthetic data, not the user's uploaded workbook. No real user data is committed. Use local storage if you do not want Excel/OneDrive autosave to synchronize files. Local processing does not automatically establish regulatory compliance.

## Limitations and next steps

- Macro installation is manual because no Excel/VBA compiler is available here; a ready-to-run XLSM is **not** supplied.
- Formula sample has fixed bounds. Arbitrary file sizes require VBA. VBA results need explicit refresh after any edits.
- Excel/VBA operations are synchronous. 50,000-row execution and cancellation behavior have not been validated; 64-bit Excel is recommended.
- Source imports preserve values and cell formats, not entire source workbook semantics. Trailing formatted-only rows are excluded; interior blank rows are retained.
- Formula keys are constrained by Excel's 32,767-character cell limit. VBA dictionary keys are not stored in worksheet cells.
- CSV import uses Excel's built-in import UI. No drag/drop upload, web themes or live web demo apply to this Excel-only edition.
- Planned after validation: signed VBA release, automated invalidation events, richer import preview and a tested packaged XLSM.

## Repository

```text
vba/                    VBA application and self-tests
 deliverables/          Formula workbook (synthetic sample)
 docs/                  Validation record
 tests/                 Excel acceptance checklist
 .github/workflows/     Artifact integrity CI
```

## License

No open-source license is granted yet. Copyright remains with the repository owner. A license can be selected before public distribution.
