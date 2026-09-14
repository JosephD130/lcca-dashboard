# Checking the workbook in real Excel

Everything in `verification/` outside this folder runs on Linux with LibreOffice and openpyxl.
That covers the arithmetic and the package structure, and it is what catches most things. Three
questions it cannot answer, because they only exist in desktop Excel on Windows:

- **Does Excel repair the file when it opens?** Only Excel decides that, and it writes a repair
  log when it does.
- **Do the macros work?** The Alternative Setup form, New Study and the KML export are VBA and
  ActiveX. LibreOffice will not run them and the Linux tooling cannot execute `vbaProject.bin`
  at all.
- **Does it look right?** Font metrics, auto row heights and chart rendering are Excel's own.
  Column-width arithmetic is a good proxy, not proof.

`Run-ExcelChecks.ps1` answers all three by driving Excel over COM.

## What you need

1. A Windows machine or VM with **desktop Excel** installed. Use the version TDOT standardises
   on. Excel for the web and the Microsoft Graph API will not do: they recalculate correctly but
   do not run VBA.
2. This folder, the workbook, and `GKT_LCCA_run.xlsx` (shipped here) on that machine.
3. Excel closed before you start. The script drives its own instance.

Two Trust Center settings, once per machine:

- **Macros.** File > Options > Trust Center > Trust Center Settings > Macro Settings >
  *Disable all macros with notification* is fine; the script enables macros for its own session.
  Easier still, add the folder under *Trusted Locations*.
- **The VBA project object model.** Same screen, tick *Trust access to the VBA project object
  model*. Without it the script still runs, but skips the checks that read the macro project and
  says so.

## Running it

```powershell
cd <the repo>\tdot-lcca
powershell -ExecutionPolicy Bypass -File .\verification\win\Run-ExcelChecks.ps1 -OutDir .\out
```

That is the automatic suite. No clicking, about a minute.

To also exercise the two macros that open a dialog:

```powershell
powershell -ExecutionPolicy Bypass -File .\verification\win\Run-ExcelChecks.ps1 -OutDir .\out -Interactive
```

It will stop twice and tell you what to click: once to save the New Study copy into a folder it
names, once to add an alternative through the Alternative Setup form. It checks the result of
each.

## What comes back

In `-OutDir`:

- `results.json` — every check, pass or fail, with detail. **This is the file to send back.**
- `TDOA_LCCA_Framework_*.xlsm` — the copy the run worked on. The script never opens or writes the
  file in the repo; it copies it here first and verifies the original's SHA256 is unchanged at the
  end. So nothing you run here can damage the workbook.
- `run.log` — the console transcript.
- `workbook.pdf` — every sheet as Excel draws it.
- any `error*.xml` repair log Excel wrote, copied out of the temp folder.

## What it checks

**Does Excel repair the file** — snapshots the places Excel writes repair logs, opens the
workbook, and fails if a new one appears. Also flags a workbook that comes up already dirty,
which is what a silent repair looks like.

**Macros** — that the Alternative Setup form and the New Study module are in the project, that
`NewStudy`, `ClearStudy`, `SetupSummaryWs` and `LoadAlternativeList` are defined, and that no
module or procedure name mentions the tool that wrote it.

**The pay-item pickers** — reads each ActiveX ComboBox at run time: column count, column widths,
dropdown width, the range it is filled from, and the first row of the list. This settles whether
the dropdown shows the item number and the description together without anyone squinting at it.

**Width** — measures each sheet's band with `Range.Width` in Excel's own points and fails
anything wider than 1,300 px, which is what fits a 1366-wide laptop.

**Calculation fidelity** — opens `GKT_LCCA_run.xlsx`, forces a full rebuild, and compares 57
values against `baseline.json`: the Summary results table, the six dashboard tiles, the RealCost
comparison block, and each alternative's initial construction total and net present worth. Those
numbers came from the LibreOffice run that `verification/examples/run_example.py` produced and
cross-checked in Python. If Excel and that baseline disagree anywhere, the two engines disagree
and I want to know.

**Rendering** — exports every sheet to one PDF.

With `-Interactive`, also: that New Study writes a cleared copy where you chose and leaves the
open study untouched, that the copy reopens clean with no alternatives and blank inputs, that
Alternative Setup builds a worksheet carrying its pickers, that the Summary picks it up, and that
saving afterwards still reopens clean.

## Regenerating the baseline

If the worked example changes, rebuild it on Linux and regenerate:

```bash
python3 verification/examples/run_example.py gkt <workbook.xlsm> /tmp/gkt 2071
python3 verification/win/make_baseline.py /tmp/gkt
cp /tmp/gkt/GKT_LCCA_run.xlsx verification/win/
```

## One thing to do by hand

The script cannot reliably trigger **Debug > Compile VBAProject**. Open the editor with Alt+F11
once after any macro change and run it; a project that will not compile fails only when someone
clicks a button.
