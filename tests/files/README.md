# Test fixtures

The formula tests compare what `pyxlsb` reads out of a `.xlsb` against the
formula text **Excel itself** writes for the same workbook as `.xlsx`. Keeping
a real Excel-authored pair means the expectations cannot silently drift into
agreeing with our own reading of the format.

| File | Provenance |
| --- | --- |
| `fixture.xlsb` / `fixture_excel.xlsx` | 107-formula corpus: operators, constants, every absolute/relative reference form, whole column and row references, set operators, fixed- and variable-arity functions, array constants, cross-sheet references, defined names, shared formulas, structured table references. |
| `fixture2.xlsb` / `fixture2_excel.xlsx` | Single- and multi-cell array (CSE) formulas, a sheet name containing an apostrophe, deeper nesting, and modern `_xlfn.` functions. |
| `fixture3.xlsb` / `fixture3_excel.xlsx` | Formulas carrying the whitespace their author typed, including spaces inside call parentheses, around prefix and postfix operators, and the space that *is* the intersection operator. |
| `simple.xlsb` | A minimal real-world workbook (`B2 =A2+A1`, `B3 =A3+B2`). |

## Regenerating

`generate_fixtures.py` writes the `.xlsx` sources (needs `openpyxl`). Excel
itself has to perform the binary conversion — no open-source writer produces
`.xlsb`:

```
python tests/generate_fixtures.py /tmp/fixtures
```

Then, for each source workbook, open it in Excel and save twice — once as
*Excel Binary Workbook (.xlsb)* and once as *Excel Workbook (.xlsx)* — so the
pair stays in step. On macOS this can be scripted:

```
osascript -e 'tell application "Microsoft Excel"
  open POSIX file "/path/fixture.xlsx"
  set wb to active workbook
  calculate
  save workbook as wb filename ((POSIX file "/path/fixture.xlsb") as text) ¬
    file format Excel binary file format with overwrite
end tell'
```

Excel is sandboxed, so it can only write where it has been granted access —
`~/Documents` works after approving the access prompt once.
