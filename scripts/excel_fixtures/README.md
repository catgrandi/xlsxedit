# Excel-authored fixtures

The scripts in this folder make Microsoft Excel write test workbooks, so every
part of them is Excel's own XML. They build:

- the fixtures in `fixtures/` that `build_fixtures.py` names: `TableRich.xlsx`,
  `TableWithContentBelow.xlsx`, `DataValidation.xlsx`,
  `ConditionalFormattingRich.xlsx`, `DrawingAnchors.xlsx`, `Formulas.xlsx` and
  `Preservation.xlsm`;
- the pairs in `fixtures/excel-reference/`: each folder holds `before.xlsx`,
  Excel's `after.xlsx` for one operation, and a README with the steps, what
  Excel did, and the Excel build. `tests/test_excel_reference.py` compares
  xlsxedit's result with them.

[`fixtures/MANIFEST.md`](../../fixtures/MANIFEST.md) lists what each workbook
contains. Each builder's docstring says what the workbook is for.

## Requirements

- Windows with desktop Excel recent enough for dynamic arrays, `XLOOKUP` and
  threaded comments, such as Microsoft 365.
- pywin32 in the development environment: `pip install pywin32`.

## Rebuild

From the repository root:

```bash
python scripts/excel_fixtures/build_fixtures.py
python scripts/excel_fixtures/build_reference.py
python scripts/describe_fixtures.py
python -m pytest -q
```

Pass names to rebuild only some, such as `build_fixtures.py TableRich` or
`build_reference.py table-append-row`. Then do the
[steps the scripts cannot do](#steps-the-scripts-cannot-do) for each workbook
you rebuilt.

Every save gives a workbook new revision ids and timestamps, so a rebuild
changes the files even when their content stays the same. The manifest only
changes when the content does.

## What the scripts do on your machine

- They start their own invisible Excel process and end it when they finish,
  so workbooks you have open in Excel are never touched. If COM hands them an
  Excel instance that already has a workbook open, they stop. They close only
  the workbooks they created or opened: if Windows opened one of yours in
  their instance, they show that instance, leave it running, and say so.
- Macros are disabled in that instance. The scripts change no Excel option and
  no Trust Center setting.
- Every workbook is saved with *Remove personal information from file
  properties on save* turned on (`filterPrivacy="1"` in `xl/workbook.xml`),
  which also keeps the folder it was saved in out of the file. Each saved
  file is then scanned for your Windows and Office user names, your git name
  and email, your home folder and the repository path, and the script stops
  if any of them appears.

`python scripts/excel_fixtures/build_fixtures.py --check` runs that scan over
every fixture without starting Excel. Run it after editing a fixture by hand.

## Last build

Every workbook here was last built with Excel 16.0.20430.20146 (Microsoft 365,
64-bit) on Windows 11 10.0.26200. The excel-reference READMEs record the build
that made each pair. `Preservation.xlsm` has no VBA project yet: see the first
step below.

## Steps the scripts cannot do

### A VBA module in `Preservation.xlsm`

Excel writes `xl/vbaProject.bin` only when the workbook's VBA project has
code. Adding code through Excel's object model needs *Trust access to the VBA
project object model*, which the scripts never turn on. Without it,
`build_fixtures.py` prints a notice and `Preservation.xlsm` has no
`vbaProject.bin`. Add the module in one of two ways.

By hand, after building:

1. Open `fixtures/Preservation.xlsm` in Excel and select **Enable Editing** if
   Excel asks.
2. Press Alt+F11 to open the Visual Basic editor.
3. Select **Insert > Module** and paste this code:

   ```vb
   Public Function Doubled(ByVal n As Double) As Double
       Doubled = n * 2
   End Function
   ```

4. Press F4 and set the module's **(Name)** to `FixtureMacros`.
5. Close the editor, press Ctrl+S, keep the `.xlsm` format, and close Excel.
6. Run `python scripts/excel_fixtures/build_fixtures.py --check`, then
   `python scripts/describe_fixtures.py` and the tests.

Or through the script:

1. In Excel, select **File > Options > Trust Center > Trust Center Settings >
   Macro Settings** and turn on **Trust access to the VBA project object
   model**.
2. Run `python scripts/excel_fixtures/build_fixtures.py Preservation`.
3. Turn the option off again, then run `python scripts/describe_fixtures.py`
   and the tests.

### Optional: check Tab against `ListRows.Add` in `table-append-row`

The pair was made with `ListRows.Add()`, which moved the note below the table
down a row although the row under the table was empty. Pressing Tab in the
table's last cell may not do that. To check, open
`fixtures/excel-reference/table-append-row/before.xlsx`, select E6, press Tab,
and see whether the note in A8 moves to A9. If it stays, Tab is a different
operation: save its result as a pair of its own and add it to `OPERATIONS` in
`tests/test_excel_reference.py`.

## The threaded comment has no author

The threaded comment in `Preservation.xlsm` has the all-zero person id and an
empty `xl/persons/person.xml`. That is what Excel writes for a workbook saved
with personal information removed. A named author would put that person's
display name and Office user id into the repository.
