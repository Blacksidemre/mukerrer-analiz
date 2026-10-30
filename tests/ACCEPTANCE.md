# Excel LTSC 2024 desktop acceptance (NOT EXECUTED)

Record Excel version/build, 32/64-bit, OS, CPU, RAM and test date. Use a disposable copy.

- [ ] Full recalc of XLSX: no #NAME?, #VALUE!, #REF!, #DIV/0! errors; totals 370/180/90/100/90.
- [ ] Switch criterion flags; verify groups against a hand-checked sample; select zero criteria: explicit invalid state.
- [ ] Whole-table ascending/descending sort: source values, group and original references remain consistent.
- [ ] Filter each status; group filter; compare members side by side.
- [ ] Import both modules; Compile VBAProject; Kurulum twice creates exactly six buttons.
- [ ] Run TestleriCalistir; record assertion count and failure details, if any.
- [ ] Import an XLSX with 1 column/1 data row and another with 3 sheets; choose non-first sheet.
- [ ] Import empty, duplicate and collision-prone headers; confirm unique display names and all columns retained.
- [ ] Import a replacement file: old results and summary are invalidated.
- [ ] Reject zero criteria; single/multiple criteria; swapped columns with matching flags; no deleted rows.
- [ ] Turkish letters, padded codes, text vs number, serial dates, blanks, long keys and wildcard strings.
- [ ] Strings starting =, +, -, @ stay literal; source workbook is not modified.
- [ ] CSV UTF-8 BOM / Windows-1254; comma / semicolon / tab; quoted separators and multiline fields through Excel importer.
- [ ] Export snapshot; reopen XLSX; row count, all source values, date formats, filters, colors and source references agree.
- [ ] Set 1/0 criteria after export: exported results intentionally remain static.
- [ ] Run PerformansOlc at 500/5,000/20,000/50,000 rows; record engine seconds.
- [ ] Separately test complete import → analyze → filter → sort → export at all four sizes; record wall time, memory, UI response, reconciliation and output row count.

A passing artifact integrity workflow is not evidence that these tests passed.
