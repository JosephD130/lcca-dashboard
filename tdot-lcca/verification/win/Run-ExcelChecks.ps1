<#
.SYNOPSIS
  Drives real desktop Excel over COM to check the TDOT LCCA workbook.

.DESCRIPTION
  Everything the Linux tooling cannot reach: whether Excel repairs the file on open, whether its
  calculation engine agrees with the baseline, what the ActiveX pay-item pickers actually show,
  how wide each sheet really is in Excel's own metrics, and whether the macros are all present.
  Writes results.json, a log and a PDF of every sheet into -OutDir.

  The automatic suite needs no clicking. -Interactive adds the two steps that open a modal dialog
  (New Study, and Alternative Setup) and checks what they produced.

  The run works on a copy of the workbook and verifies the original is byte for byte unchanged
  when it finishes, so nothing here can damage the file.

.EXAMPLE
  Double-click Run.cmd

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Run-ExcelChecks.ps1 -Interactive
#>
[CmdletBinding()]
param(
  [string] $Workbook   = '',
  [string] $Baseline   = '',
  [string] $ExampleDir = '',
  [string] $OutDir     = '',
  [switch] $Interactive
)

$ErrorActionPreference = 'Stop'
$script:Results = @()
$script:Fails   = 0

function Resolve-Full([string]$p) {
  # Join-Path onto the working directory turns an already-rooted path into "C:\cwd\C:\..." and
  # GetFullPath then refuses it, so only join when the path is relative.
  if ([string]::IsNullOrWhiteSpace($p)) { return $p }
  if ([IO.Path]::IsPathRooted($p)) { return [IO.Path]::GetFullPath($p) }
  return [IO.Path]::GetFullPath((Join-Path (Get-Location).Path $p))
}

function Check([string]$label, $ok, $detail = '') {
  $ok = [bool]$ok
  $script:Results += [pscustomobject]@{ check = $label; ok = $ok; detail = "$detail" }
  if (-not $ok) { $script:Fails++ }
  $tag = if ($ok) { 'PASS  ' } else { 'FAIL  ' }
  if ($ok -or "$detail" -eq '') { Write-Host ($tag + $label) }
  else { Write-Host ($tag + $label + '   |   ' + $detail) }
}

function Section([string]$name) {
  Write-Host ''
  Write-Host ('=' * 78)
  Write-Host $name
  Write-Host ('=' * 78)
}

# One unexpected COM error costs that section only, not the rest of the run. The block runs in
# this function's scope, so anything a later section needs is assigned with $script:.
function Guard([string]$what, [scriptblock]$body) {
  try { . $body }
  catch { Check "$what ran without an unexpected error" $false $_.Exception.Message }
}

function Prop($obj, [string]$name, $fallback = $null) {
  try { return $obj.$name } catch { return $fallback }
}

function SetProp($obj, [string]$name, $value) {
  try { $obj.$name = $value; return $true } catch { return $false }
}

# The tables and their column names, read straight out of the package. A repair is Excel throwing
# something away, so comparing this between the file we handed Excel and the file Excel writes back
# catches a repair without depending on a log file turning up.
function Get-Tables([string]$xlsx) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
  $zip = [IO.Compression.ZipFile]::OpenRead($xlsx)
  try {
    $out = @{}
    foreach ($e in $zip.Entries) {
      if ($e.FullName -notmatch '^xl/tables/table\d+\.xml$') { continue }
      $sr = New-Object IO.StreamReader($e.Open())
      $x  = $sr.ReadToEnd(); $sr.Dispose()
      $nm = ([regex]::Match($x, '<table [^>]*?name="([^"]+)"')).Groups[1].Value
      $cols = @([regex]::Matches($x, '<tableColumn [^>]*?name="([^"]*)"') |
                ForEach-Object { $_.Groups[1].Value })
      if ($nm) { $out[$nm] = ($cols -join ' | ') }
    }
    return $out
  } finally { $zip.Dispose() }
}

# ---------------------------------------------------------------------------- preflight
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $OutDir)   { $OutDir   = Join-Path $here 'out' }
if (-not $Baseline) { $Baseline = Join-Path $here 'baseline.json' }

try {
  $OutDir   = Resolve-Full $OutDir
  $Baseline = Resolve-Full $Baseline
  if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
} catch {
  Write-Host ('FAIL  could not prepare the output folder: ' + $_.Exception.Message)
  Write-Host ('      it was: ' + $OutDir)
  exit 2
}

if (-not $Workbook) {
  foreach ($pattern in @((Join-Path $here 'TDOA_LCCA_Framework_*.xlsm'),
                         (Join-Path $here '..\..\TDOA_LCCA_Framework_*.xlsm'),
                         (Join-Path (Get-Location).Path 'TDOA_LCCA_Framework_*.xlsm'))) {
    $hit = @(Get-ChildItem $pattern -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    if ($hit.Count -gt 0) { $Workbook = $hit[0].FullName; break }
  }
}
if (-not $Workbook -or -not (Test-Path $Workbook)) {
  Write-Host 'FAIL  could not find TDOA_LCCA_Framework_*.xlsm beside this script.'
  Write-Host ('      looked in: ' + $here)
  Write-Host '      Put the workbook in this folder, or pass -Workbook <path>.'
  exit 2
}
$Workbook = Resolve-Full $Workbook

# Windows marks anything that came from the internet, which puts Excel into Protected View and
# stops the macros loading. Clear it for this folder before we start.
Get-ChildItem $here -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

$logPath = Join-Path $OutDir 'run.log'
try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
try { Start-Transcript -Path $logPath -Force | Out-Null } catch { Write-Host '(no transcript)' }
Write-Host ("PowerShell $($PSVersionTable.PSVersion) on $([Environment]::OSVersion.VersionString)")
Write-Host ("workbook: $Workbook")
Write-Host ("output:   $OutDir")

# Excel writes a repair log when it fixes a file on open. The location moves between versions, so
# snapshot every place it is known to land and diff after each Open.
$RepairDirs = @($env:TEMP, (Split-Path $Workbook), $OutDir, (Join-Path $env:USERPROFILE 'Documents')) |
              Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
function Repair-Snapshot {
  $h = @{}
  foreach ($d in $RepairDirs) {
    foreach ($f in (Get-ChildItem -Path $d -Filter 'error*.xml' -File -ErrorAction SilentlyContinue)) {
      $h[$f.FullName] = $f.LastWriteTimeUtc
    }
  }
  return $h
}
function Repair-New($before) {
  $after = Repair-Snapshot
  return @($after.Keys | Where-Object { -not $before.ContainsKey($_) -or $before[$_] -ne $after[$_] })
}

$excel = $null
$script:wb = $null
$script:openedBooks = @()
$shippedHash = (Get-FileHash $Workbook -Algorithm SHA256).Hash

try {
  # -------------------------------------------------------------------------- Excel
  Section 'EXCEL'
  try { $excel = New-Object -ComObject Excel.Application }
  catch {
    Check 'desktop Excel is installed and can be automated' $false $_.Exception.Message
    throw 'no Excel'
  }
  [void](SetProp $excel 'Visible' ([bool]$Interactive))
  # with no human watching, a modal repair dialog would hang the run; with -Interactive we want to
  # see it, so alerts are on and the script asks about it after the open
  [void](SetProp $excel 'DisplayAlerts' ([bool]$Interactive))
  [void](SetProp $excel 'ScreenUpdating' $false)
  [void](SetProp $excel 'EnableEvents' $true)
  [void](SetProp $excel 'AutomationSecurity' 1)   # msoAutomationSecurityLow: let the macros load
  # Application.Calculation does not exist until a workbook is open. Excel answers 0x800A03EC if
  # you set it before then, so hold a scratch workbook open across the switch.
  $scratch = $null
  try { $scratch = $excel.Workbooks.Add() } catch { }
  $manual = SetProp $excel 'Calculation' -4135     # xlCalculationManual
  Check "Excel $(Prop $excel 'Version') build $(Prop $excel 'Build') is driving this run" $true `
    ("calculation set to manual: $manual")

  # -------------------------------------------------------------------------- 1. repair on open
  Guard 'the open check' {
    Section 'DOES EXCEL REPAIR THE FILE'
    # a copy, so the interactive steps cannot touch the shipped workbook
    $working = Join-Path $OutDir ([IO.Path]::GetFileName($Workbook))
    Copy-Item $Workbook $working -Force
    $before = Repair-Snapshot
    $script:wb = $excel.Workbooks.Open($working, 0, $false)  # UpdateLinks = 0, ReadOnly = false
    $script:openedBooks += $script:wb
    if ($scratch) { try { $scratch.Close($false) } catch { } ; $scratch = $null }
    $logs = Repair-New $before
    Check 'Excel opened it without writing a repair log' ($logs.Count -eq 0) ($logs -join '; ')
    foreach ($l in $logs) {
      Copy-Item $l (Join-Path $OutDir (Split-Path $l -Leaf)) -Force -ErrorAction SilentlyContinue
      Write-Host ('      repair log says: ' + (Get-Content $l -Raw -ErrorAction SilentlyContinue))
    }
    Write-Host ("      (Saved=" + (Prop $script:wb 'Saved') +
                " after load; Workbook_Open hides two sheets, so False is expected here)")

    # A repair log is not guaranteed to appear with alerts suppressed, so also ask Excel to write
    # the workbook back out and compare the tables. Excel dropping or renaming a table column is
    # exactly the repair this workbook hit before.
    $resaved = Join-Path $OutDir 'excel-resaved.xlsm'
    if (Test-Path $resaved) { Remove-Item $resaved -Force -ErrorAction SilentlyContinue }
    $script:wb.SaveCopyAs($resaved)
    $wantTables = Get-Tables $Workbook
    $gotTables  = Get-Tables $resaved
    $diff = @()
    foreach ($k in $wantTables.Keys) {
      if (-not $gotTables.ContainsKey($k)) { $diff += "$k was dropped" }
      elseif ($gotTables[$k] -ne $wantTables[$k]) { $diff += "$k columns: '$($gotTables[$k])' vs '$($wantTables[$k])'" }
    }
    foreach ($k in $gotTables.Keys) { if (-not $wantTables.ContainsKey($k)) { $diff += "$k appeared" } }
    Check "all $($wantTables.Count) tables survive a round trip through Excel with their columns intact" `
      ($diff.Count -eq 0) (@($diff | Select-Object -First 4) -join '; ')
    if ($Interactive) {
      $said = Read-Host '  Did Excel show a repair or validation message when it opened? y / n'
      Check 'no repair or validation message appeared on screen' ("$said" -notmatch '^(y|Y)') "you answered '$said'"
    }
    $rehab = $null
    try { $rehab = $script:wb.Worksheets('TMP(HMARehab)') } catch { }
    Check 'the Workbook_Open macro ran and hid the two rehab templates' `
      (($null -ne $rehab) -and ((Prop $rehab 'Visible' 0) -eq 2)) `
      ("Visible=" + (Prop $rehab 'Visible' 'sheet not found') + ' (xlSheetVeryHidden is 2)')
  }
  if (-not $script:wb) { throw 'the workbook did not open' }

  # -------------------------------------------------------------------------- 2. macros
  Guard 'the macro check' {
    Section 'MACROS'
    # The definitive answer to whether Trust access to the VBA project object model is on. Excel
    # can hand back a VBProject object whose VBComponents collection is simply empty when it is
    # off, which looks like a workbook with no macros rather than a permissions problem.
    $ver = "$(Prop $excel 'Version')"
    $trust = $null
    foreach ($k in @("HKCU:\Software\Microsoft\Office\$ver\Excel\Security",
                     'HKCU:\Software\Microsoft\Office\16.0\Excel\Security')) {
      try { $trust = (Get-ItemProperty -Path $k -Name AccessVBOM -ErrorAction Stop).AccessVBOM; break } catch { }
    }
    Check 'Trust access to the VBA project object model is on' ($trust -eq 1) `
      ("AccessVBOM=$trust (Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings)")

    $proj = $null
    try { $proj = $script:wb.VBProject } catch { }
    if (-not $proj) {
      Check 'the macro project can be read' $false 'Workbook.VBProject is not available'
    } else {
      $count = 0
      try { $count = $proj.VBComponents.Count } catch { }
      Check 'the macro project lists its modules' ($count -gt 0) `
        ("VBComponents.Count=$count, Protection=" + (Prop $proj 'Protection' '?'))
      if ($count -gt 0) {
        # index rather than foreach: PowerShell does not always get an enumerator for this collection
        $comps = @(); $procs = @()
        for ($i = 1; $i -le $count; $i++) {
          $c = $null
          try { $c = $proj.VBComponents.Item($i) } catch { continue }
          $comps += "$($c.Name)"
          try {
            $cm = $c.CodeModule
            for ($j = 1; $j -le $cm.CountOfLines; $j++) {
              if ($cm.Lines($j, 1) -match '^\s*(?:Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+(\w+)') {
                $procs += $Matches[2]
              }
            }
          } catch { }
        }
        Check 'the Alternative Setup form and the New Study module are in the project' `
          (($comps -contains 'frmAlternativeSetup') -and ($comps -contains 'LCCA_NewStudy')) `
          ("$($comps.Count) modules: " + (($comps | Select-Object -First 12) -join ', '))
        foreach ($m in @('NewStudy', 'ClearStudy', 'SetupSummaryWs', 'LoadAlternativeList')) {
          Check "$m is defined" ($procs -contains $m) "$($procs.Count) procedures found"
        }
        Check 'no module or procedure mentions the tool that wrote it' `
          (@($comps + $procs | Where-Object { $_ -match '(?i)claude|anthropic' }).Count -eq 0) ''
      }
    }
  }

  # -------------------------------------------------------------------------- 3. the pickers
  Guard 'the pay-item picker check' {
    Section 'THE PAY-ITEM PICKERS (ActiveX, Windows only)'
    $checked = 0
    foreach ($ws in $script:wb.Worksheets) {
      if ($ws.Name -notlike 'TMP(*') { continue }
      foreach ($ole in $ws.OLEObjects()) {
        if ("$(Prop $ole 'progID')" -notlike '*ComboBox*') { continue }
        $o = Prop $ole 'Object'
        if (-not $o) { continue }
        if ($checked -eq 0) {
          Check 'the picker shows two columns, the item number and the description' `
            ((Prop $o 'ColumnCount' 0) -ge 2) `
            ("ColumnCount=" + (Prop $o 'ColumnCount') + " ColumnWidths='" + (Prop $o 'ColumnWidths') + "'")
          Check 'the dropdown is wide enough for both columns' `
            ((Prop $o 'ListWidth' 0) -gt 200) ("ListWidth=" + (Prop $o 'ListWidth') + " pt")
          Check 'it is filled from the Pay_Items list' `
            ("$(Prop $o 'ListFillRange')" -match 'Pay_Item') ("ListFillRange='" + (Prop $o 'ListFillRange') + "'")
          $n = Prop $o 'ListCount' 0
          Check 'the list has rows in it' ($n -gt 0) "ListCount=$n"
          if ($n -gt 0) {
            # the first entry is deliberately blank, the "Default" row of the picker, so look down
            $rows = @(); $both = $false
            for ($r = 0; $r -lt [Math]::Min(6, $n); $r++) {
              $cells = @()
              for ($k = 0; $k -lt [Math]::Min(2, (Prop $o 'ColumnCount' 1)); $k++) {
                try { $cells += "$($o.List($r, $k))" } catch { $cells += '' }
              }
              $rows += ($cells -join ' / ')
              if (@($cells | Where-Object { "$_" -ne '' }).Count -ge 2) { $both = $true }
            }
            Check 'the list reads back with the item number and the description together' `
              $both ((($rows | Select-Object -First 4) -join '   ') + "   (row 1 is the blank Default entry)")
          }
        }
        $checked++
      }
    }
    Check 'every alternative template carries its pickers' ($checked -ge 10) "$checked combo boxes found"
  }

  # -------------------------------------------------------------------------- 4. real widths
  Guard 'the width check' {
    Section "WIDTH IN EXCEL'S OWN METRICS (a 1366-wide laptop shows about 1300 px)"
    $bands = @(
      @('Summary',              'G1:R1'),
      @('Overview',             'B1:M1'),
      @('Instructions',         'B1:M1'),
      @('General Information',  'A1:J1'),
      @('Pay_Items',            'A1:J1'),
      @('Maintenance Policies', 'B1:F1'),
      @('Typical Values',       'A1:G1'),
      @('Method',               'A1:E1')
    )
    foreach ($b in $bands) {
      $ws = $null
      try { $ws = $script:wb.Worksheets($b[0]) } catch { Check "$($b[0]) is in the workbook" $false ''; continue }
      $pt = $ws.Range($b[1]).Width            # points
      $px = [Math]::Round($pt * 96.0 / 72.0)  # Excel lays out at 96 dpi
      Check "$($b[0]) fits a 1366-wide laptop" ($px -le 1300) `
        ("$px px, " + [Math]::Round($px / 96.0, 2) + " in, across $($b[1])")
    }
  }

  # -------------------------------------------------------------------------- 5. rendering
  Guard 'the PDF export' {
    Section 'RENDERING'
    $pdf = Join-Path $OutDir 'workbook.pdf'
    $script:wb.ExportAsFixedFormat(0, $pdf)          # xlTypePDF
    Check 'exported every sheet to PDF for review' (Test-Path $pdf) $pdf
  }

  # -------------------------------------------------------------------------- 6. the numbers
  Guard 'the calculation check' {
    Section 'CALCULATION FIDELITY AGAINST THE WORKED EXAMPLE'
    if (-not (Test-Path $Baseline)) {
      Check 'baseline.json is beside this script' $false $Baseline
    } else {
      $base  = Get-Content $Baseline -Raw | ConvertFrom-Json
      $exDir = if ($ExampleDir) { Resolve-Full $ExampleDir } else { Split-Path $Baseline }
      $exWb  = Join-Path $exDir $base.workbook
      if (-not (Test-Path $exWb)) {
        Check "the example workbook $($base.workbook) is beside the baseline" $false "looked in $exDir"
      } else {
        $b2 = Repair-Snapshot
        $ex = $excel.Workbooks.Open($exWb, 0, $false)
        $script:openedBooks += $ex
        Check 'Excel opened the worked example without a repair log' ((Repair-New $b2).Count -eq 0) ''
        [void](SetProp $excel 'Calculation' -4105)   # xlCalculationAutomatic
        try { $excel.CalculateFullRebuild() } catch { $excel.Calculate() }
        $bad = @()
        foreach ($c in $base.cells) {
          $ws = $null
          try { $ws = $ex.Worksheets($c.sheet) } catch { $bad += "$($c.sheet) missing"; continue }
          if ($c.cell -eq 'TOTAL' -or $c.cell -eq 'NPW') {
            # the alternative sheets differ in length, so find the row by what column A or B says
            $labelCol = if ($c.cell -eq 'TOTAL') { 1 } else { 2 }
            $valueCol = if ($c.cell -eq 'TOTAL') { 7 } else { 5 }
            $want = if ($c.cell -eq 'TOTAL') { 'Total' } else { 'Net Present Worth' }
            $row = 0
            for ($r = 1; $r -le 80; $r++) { if ("$($ws.Cells($r, $labelCol).Value2)" -eq $want) { $row = $r } }
            if ($row -eq 0) { $bad += "$($c.sheet)!$($c.cell) label not found"; continue }
            $got = $ws.Cells($row, $valueCol).Value2
          } else {
            $got = $ws.Range($c.cell).Value2
          }
          if ($null -eq $c.tol) {
            if ("$got" -ne "$($c.expect)") { $bad += "$($c.sheet)!$($c.cell) '$got' vs '$($c.expect)'" }
          } elseif ($null -eq $got) {
            $bad += "$($c.sheet)!$($c.cell) is empty"
          } else {
            if ([Math]::Abs([double]$got - [double]$c.expect) -gt [double]$c.tol) {
              $bad += "$($c.sheet)!$($c.cell) $got vs $($c.expect)"
            }
          }
        }
        Check "Excel reproduces all $($base.cells.Count) baseline values from the $($base.example) run" `
          ($bad.Count -eq 0) (@($bad | Select-Object -First 5) -join '; ')
      }
    }
  }

  # -------------------------------------------------------------------------- 7. the two dialogs
  if ($Interactive) {
    Guard 'the New Study step' {
      Section 'INTERACTIVE: NEW STUDY'
      [void](SetProp $excel 'Visible' $true)
      [void](SetProp $excel 'ScreenUpdating' $true)
      try { $script:wb.Activate() } catch { }
      $giBefore = "$($script:wb.Worksheets('General Information').Range('D9').Value2)"
      $stage = Join-Path $OutDir 'newstudy'
      if (-not (Test-Path $stage)) { New-Item -ItemType Directory -Path $stage -Force | Out-Null }
      Write-Host ''
      Write-Host '  New Study is about to run. When the Save dialog opens, save into:'
      Write-Host "      $stage"
      Read-Host  '  Press Enter to run it'
      try { $excel.Run("'$($script:wb.Name)'!NewStudy") }
      catch { Write-Host "  (New Study returned: $($_.Exception.Message))" }
      $made = @(Get-ChildItem $stage -Filter '*.xlsm' -File -ErrorAction SilentlyContinue)
      Check 'New Study wrote a new workbook where you chose' ($made.Count -ge 1) `
        (@($made | ForEach-Object { $_.Name }) -join ', ')
      Check 'and left the study you had open exactly as it was' `
        ("$($script:wb.Worksheets('General Information').Range('D9').Value2)" -eq $giBefore) "D9 was '$giBefore'"
      if ($made.Count -ge 1) {
        $b3 = Repair-Snapshot
        $ns = $excel.Workbooks.Open($made[0].FullName, 0, $true)
        $script:openedBooks += $ns
        Check 'the new study reopens without a repair log' ((Repair-New $b3).Count -eq 0) ''
        $alts = @(); foreach ($s2 in $ns.Worksheets) { if ($s2.Name -like 'Alt *') { $alts += $s2.Name } }
        Check 'it carries no alternatives from the study you had open' ($alts.Count -eq 0) ($alts -join ', ')
        Check 'and its inputs are blank' `
          ("$($ns.Worksheets('General Information').Range('D9').Value2)" -eq '') ''
      }
    }

    Guard 'the Alternative Setup step' {
      Section 'INTERACTIVE: ALTERNATIVE SETUP'
      [void](SetProp $excel 'Visible' $true)
      try { $script:wb.Activate() } catch { }
      try { $script:wb.Worksheets('General Information').Activate() } catch { }
      $sheetsBefore = @(); foreach ($s2 in $script:wb.Worksheets) { $sheetsBefore += $s2.Name }
      Write-Host ''
      Write-Host '  In Excel, click the Alternative Setup button on General Information,'
      Write-Host '  add one alternative, then close the form. The form is modal, so it has to'
      Write-Host '  be you rather than this script.'
      Read-Host  '  Press Enter when the form is closed'
      $sheetsAfter = @(); foreach ($s2 in $script:wb.Worksheets) { $sheetsAfter += $s2.Name }
      $new = @($sheetsAfter | Where-Object { $sheetsBefore -notcontains $_ })
      Check 'Alternative Setup built a worksheet for the alternative' ($new.Count -ge 1) ($new -join ', ')
      if ($new.Count -ge 1) {
        $as = $script:wb.Worksheets($new[0])
        Check 'the new sheet carries the pay-item pickers' ($as.OLEObjects().Count -gt 0) `
          "$($as.OLEObjects().Count) controls"
        Check 'and the Summary picked it up' `
          ("$($script:wb.Worksheets('Summary').Range('G12').Value2)" -ne '') `
          ("Summary!G12='" + $script:wb.Worksheets('Summary').Range('G12').Value2 + "'")
        $saved = Join-Path $OutDir 'after-alternative-setup.xlsm'
        $script:wb.SaveCopyAs($saved)
        $b4 = Repair-Snapshot
        $reopened = $excel.Workbooks.Open($saved, 0, $true)
        $script:openedBooks += $reopened
        Check 'a workbook saved after the form ran reopens without a repair log' `
          ((Repair-New $b4).Count -eq 0) $saved
      }
    }
  }
}
catch {
  Check 'the run finished without an unexpected error' $false $_.Exception.Message
}
finally {
  Section 'RESULT'
  if ($shippedHash -and (Test-Path $Workbook)) {
    # this run worked on a copy; the original has to come out byte for byte unchanged
    Check 'the workbook you gave it was never written to' `
      ((Get-FileHash $Workbook -Algorithm SHA256).Hash -eq $shippedHash) "SHA256 $shippedHash"
  }
  $json = Join-Path $OutDir 'results.json'
  try {
    @{ workbook = $Workbook
       excel    = if ($excel) { "$(Prop $excel 'Version') build $(Prop $excel 'Build')" } else { 'not started' }
       when     = (Get-Date).ToString('s')
       failed   = $script:Fails
       checks   = $script:Results } | ConvertTo-Json -Depth 5 | Set-Content $json -Encoding UTF8
  } catch { Write-Host "could not write results.json: $($_.Exception.Message)" }
  Write-Host ('{0} checks, {1} failed' -f $script:Results.Count, $script:Fails)
  if ($script:Fails -eq 0) { Write-Host 'ALL PASS' }
  else {
    Write-Host ('FAILED: ' + ((@($script:Results | Where-Object { -not $_.ok } | ForEach-Object { $_.check })) -join '; '))
  }
  Write-Host "wrote $json"

  foreach ($b in $script:openedBooks) { try { $b.Close($false) } catch { } }
  if ($excel) {
    try { $excel.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  try { Stop-Transcript | Out-Null } catch { }
}
exit $(if ($script:Fails -gt 0) { 1 } else { 0 })
