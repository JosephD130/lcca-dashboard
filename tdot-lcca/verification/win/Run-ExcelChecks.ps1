<#
.SYNOPSIS
  Drives real desktop Excel over COM to check the TDOT LCCA workbook.

.DESCRIPTION
  Everything the Linux tooling cannot reach: whether Excel repairs the file on open, whether its
  calculation engine agrees with the baseline, what the ActiveX pay-item pickers actually show,
  how wide each sheet really is in Excel's own metrics, and whether the macros are present and
  compile. Writes results.json, a log, and a PDF of every sheet into -OutDir.

  The automatic suite needs no clicking. -Interactive adds the two steps that open a modal dialog
  (New Study, and Alternative Setup) and verifies what they produced.

.PARAMETER Workbook
  The shipped .xlsm. Default: the file beside the repo root.

.PARAMETER Baseline
  baseline.json plus the example workbook it names, for the calculation-fidelity check.

.PARAMETER OutDir
  Where to write results.json, run.log and the PDFs. Created if missing.

.PARAMETER Interactive
  Also run the two macros that open a dialog, prompting you to click through.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\verification\win\Run-ExcelChecks.ps1 -OutDir .\out

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\verification\win\Run-ExcelChecks.ps1 -OutDir .\out -Interactive
#>
[CmdletBinding()]
param(
  [string] $Workbook  = "$PSScriptRoot\..\..\TDOA_LCCA_Framework_v1.2.0_ARA_09112026.xlsm",
  [string] $Baseline  = "$PSScriptRoot\baseline.json",
  [string] $ExampleDir = '',
  [string] $OutDir    = "$PSScriptRoot\out",
  [switch] $Interactive
)

$ErrorActionPreference = 'Stop'
$script:Results = @()
$script:Fails   = 0

function Resolve-Full([string]$p) { return [IO.Path]::GetFullPath((Join-Path (Get-Location) $p)) }

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutDir    = Resolve-Full $OutDir
$Workbook  = Resolve-Full $Workbook
$logPath   = Join-Path $OutDir 'run.log'
Start-Transcript -Path $logPath -Force | Out-Null

function Check([string]$label, [bool]$ok, $detail = '') {
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

# Excel writes a repair log when it fixes a file on open. The location moves between versions, so
# snapshot every place it is known to land and diff after each Open.
$RepairDirs = @($env:TEMP, (Split-Path $Workbook), (Join-Path $env:USERPROFILE 'Documents')) |
              Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
function Repair-Snapshot {
  $h = @{}
  foreach ($d in $RepairDirs) {
    foreach ($f in Get-ChildItem -Path $d -Filter 'error*.xml' -File -ErrorAction SilentlyContinue) {
      $h[$f.FullName] = $f.LastWriteTimeUtc
    }
  }
  return $h
}
function Repair-New([hashtable]$before) {
  $after = Repair-Snapshot
  return @($after.Keys | Where-Object { -not $before.ContainsKey($_) -or $before[$_] -ne $after[$_] })
}

$excel = $null
$openedBooks = @()
$working = $null
$shippedHash = if (Test-Path $Workbook) { (Get-FileHash $Workbook -Algorithm SHA256).Hash } else { $null }
try {
  Section 'EXCEL'
  try { $excel = New-Object -ComObject Excel.Application }
  catch { Write-Host 'FAIL  no desktop Excel on this machine (New-Object Excel.Application failed)'; throw }
  $excel.Visible       = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.AutomationSecurity = 1        # msoAutomationSecurityLow: let the workbook's macros load
  $excel.Calculation   = -4135         # xlCalculationManual, so nothing recalcs before we ask
  Check "Excel $($excel.Version) ($($excel.Build)) is driving this run" $true "$($excel.Name)"

  # ------------------------------------------------------------------ 1. does it open clean
  Section 'DOES EXCEL REPAIR THE FILE'
  Check 'the workbook exists' (Test-Path $Workbook) $Workbook
  # everything below runs on a copy: the interactive steps add an alternative, and the shipped
  # file must come out of this byte for byte unchanged
  $working = Join-Path $OutDir ([IO.Path]::GetFileName($Workbook))
  Copy-Item $Workbook $working -Force
  $before = Repair-Snapshot
  $wb = $excel.Workbooks.Open($working, 0, $false)    # UpdateLinks = 0, ReadOnly = false
  $openedBooks += $wb
  $logs = Repair-New $before
  Check 'Excel opened it without writing a repair log' ($logs.Count -eq 0) ($logs -join '; ')
  foreach ($l in $logs) {
    Copy-Item $l (Join-Path $OutDir (Split-Path $l -Leaf)) -Force -ErrorAction SilentlyContinue
    Write-Host ('      repair log: ' + (Get-Content $l -Raw))
  }
  Check 'and did not mark it dirty on load, which a silent repair does' $wb.Saved "Saved=$($wb.Saved)"

  # ------------------------------------------------------------------ 2. the macros
  Section 'MACROS'
  $vbeOk = $false
  try { $null = $wb.VBProject.Name; $vbeOk = $true } catch { }
  if (-not $vbeOk) {
    Check 'Trust access to the VBA project object model is on' $false `
      'File > Options > Trust Center > Trust Center Settings > Macro Settings. Without it the macro checks are skipped.'
  } else {
    $comps = @(); $procs = @()
    foreach ($c in $wb.VBProject.VBComponents) {
      $comps += $c.Name
      $cm = $c.CodeModule
      for ($i = 1; $i -le $cm.CountOfLines; $i++) {
        if ($cm.Lines($i, 1) -match '^\s*(?:Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+(\w+)') { $procs += $Matches[2] }
      }
    }
    Check 'the Alternative Setup form and the New Study module are in the project' `
      (($comps -contains 'frmAlternativeSetup') -and ($comps -contains 'LCCA_NewStudy')) ($comps -join ', ')
    foreach ($m in @('NewStudy', 'ClearStudy', 'SetupSummaryWs', 'LoadAlternativeList')) {
      Check "$m is defined" ($procs -contains $m) ''
    }
    Check 'no module mentions the tool that wrote it' `
      (-not ($comps + $procs | Where-Object { $_ -match '(?i)claude|anthropic' })) ''
  }

  # ------------------------------------------------------------------ 3. the ActiveX pay-item pickers
  Section 'THE PAY-ITEM PICKERS (ActiveX, Windows only)'
  $pickerSheets = @()
  foreach ($ws in $wb.Worksheets) { if ($ws.Name -like 'TMP(*') { $pickerSheets += $ws } }
  $checked = 0
  foreach ($ws in $pickerSheets) {
    foreach ($ole in $ws.OLEObjects()) {
      if ($ole.progID -notlike '*ComboBox*') { continue }
      $o = $ole.Object
      if ($checked -eq 0) {
        Check 'the picker shows two columns, the item number and the description' `
          ($o.ColumnCount -ge 2) "ColumnCount=$($o.ColumnCount) ColumnWidths='$($o.ColumnWidths)'"
        Check 'the dropdown is wide enough for both columns' `
          ($o.ListWidth -gt 200) "ListWidth=$($o.ListWidth) pt"
        Check 'it is filled from the Pay_Items list' `
          ("$($o.ListFillRange)" -match 'Pay_Items|Pay_Item') "ListFillRange='$($o.ListFillRange)'"
        if ($o.ListCount -gt 0) {
          $sample = @()
          for ($c = 0; $c -lt [Math]::Min(2, $o.ColumnCount); $c++) { $sample += "$($o.List(0, $c))" }
          Check 'and the first row of the list reads back with both columns populated' `
            (($sample | Where-Object { $_ -ne '' }).Count -ge 2) ($sample -join ' | ')
        }
      }
      $checked++
    }
  }
  Check 'every alternative template carries its pickers' ($checked -ge 10) "$checked combo boxes found"

  # ------------------------------------------------------------------ 4. how wide the sheets really are
  Section 'WIDTH IN EXCEL''S OWN METRICS (target: 1,240 px of content)'
  $bands = @{
    'Summary'                = 'G1:R1'
    'Overview'               = 'B1:M1'
    'Instructions'           = 'B1:M1'
    'General Information'    = 'A1:J1'
    'Pay_Items'              = 'A1:J1'
    'Maintenance Policies'   = 'B1:F1'
    'Typical Values'         = 'A1:G1'
    'Method'                 = 'A1:E1'
  }
  $widths = @{}
  foreach ($k in $bands.Keys) {
    $ws = $null
    try { $ws = $wb.Worksheets($k) } catch { continue }
    $pt = $ws.Range($bands[$k]).Width          # points
    $px = [Math]::Round($pt * 96.0 / 72.0)     # Excel lays out at 96 dpi
    $widths[$k] = $px
    Check "$k fits a 1366-wide laptop" ($px -le 1300) "$px px ($([Math]::Round($px/96.0,2)) in) across $($bands[$k])"
  }

  # ------------------------------------------------------------------ 5. every sheet, as Excel draws it
  Section 'RENDERING'
  $pdf = Join-Path $OutDir 'workbook.pdf'
  try {
    $wb.ExportAsFixedFormat(0, $pdf)           # xlTypePDF
    Check 'exported every sheet to PDF for review' (Test-Path $pdf) $pdf
  } catch { Check 'exported every sheet to PDF for review' $false $_.Exception.Message }

  # ------------------------------------------------------------------ 6. does Excel agree with the baseline
  Section 'CALCULATION FIDELITY AGAINST THE WORKED EXAMPLE'
  if (-not (Test-Path $Baseline)) {
    Check 'baseline.json is present' $false $Baseline
  } else {
    $base = Get-Content $Baseline -Raw | ConvertFrom-Json
    $exDir = if ($ExampleDir) { Resolve-Full $ExampleDir } else { Split-Path $Baseline }
    $exWb  = Join-Path $exDir $base.workbook
    if (-not (Test-Path $exWb)) {
      Check "the example workbook $($base.workbook) is beside the baseline" $false `
        "put it in $exDir, or pass -ExampleDir. It is produced by verification/examples/run_example.py."
    } else {
      $b2 = Repair-Snapshot
      $ex = $excel.Workbooks.Open($exWb, 0, $false)
      $openedBooks += $ex
      Check 'Excel opened the example workbook without a repair log' ((Repair-New $b2).Count -eq 0) ''
      $excel.Calculation = -4105            # xlCalculationAutomatic
      $excel.CalculateFullRebuild()
      $bad = @()
      foreach ($c in $base.cells) {
        $ws = $null
        try { $ws = $ex.Worksheets($c.sheet) } catch { $bad += "$($c.sheet) missing"; continue }
        $ref = $c.cell
        if ($ref -eq 'TOTAL' -or $ref -eq 'NPW') {
          # the alternative sheets differ in length, so find the row by what column A or B says
          $col = if ($ref -eq 'TOTAL') { 1 } else { 2 }
          $want = if ($ref -eq 'TOTAL') { 'Total' } else { 'Net Present Worth' }
          $row = 0
          for ($r = 1; $r -le 80; $r++) { if ("$($ws.Cells($r, $col).Value2)" -eq $want) { $row = $r } }
          if ($row -eq 0) { $bad += "$($c.sheet)!$ref not found"; continue }
          $got = $ws.Cells($row, $(if ($ref -eq 'TOTAL') { 7 } else { 5 })).Value2
        } else {
          $got = $ws.Range($ref).Value2
        }
        if ($null -eq $c.tol) {
          if ("$got" -ne "$($c.expect)") { $bad += "$($c.sheet)!$ref '$got' vs '$($c.expect)'" }
        } else {
          $d = [Math]::Abs([double]$got - [double]$c.expect)
          if ($d -gt [double]$c.tol) { $bad += "$($c.sheet)!$ref $got vs $($c.expect)" }
        }
      }
      Check "Excel reproduces all $($base.cells.Count) baseline values from the $($base.example) run" `
        ($bad.Count -eq 0) (($bad | Select-Object -First 5) -join '; ')
    }
  }

  # ------------------------------------------------------------------ 7. the two dialogs
  if ($Interactive) {
    Section 'INTERACTIVE: THE TWO MACROS THAT OPEN A DIALOG'
    $excel.Visible = $true
    $excel.ScreenUpdating = $true
    $wb.Activate()

    $giBefore = "$($wb.Worksheets('General Information').Range('D9').Value2)"
    $stage = Join-Path $OutDir 'newstudy'
    if (-not (Test-Path $stage)) { New-Item -ItemType Directory -Path $stage | Out-Null }
    Write-Host ''
    Write-Host '  About to run New Study. When the Save dialog opens, save into:'
    Write-Host "      $stage"
    Read-Host '  Press Enter to run it'
    try { $excel.Run("'$($wb.Name)'!NewStudy") } catch { Write-Host "  (New Study returned: $($_.Exception.Message))" }
    $made = @(Get-ChildItem $stage -Filter '*.xlsm' -File -ErrorAction SilentlyContinue)
    Check 'New Study wrote a new workbook where you chose' ($made.Count -ge 1) (($made | ForEach-Object { $_.Name }) -join ', ')
    Check 'and left the study you had open exactly as it was' `
      ("$($wb.Worksheets('General Information').Range('D9').Value2)" -eq $giBefore) "D9 was '$giBefore'"
    if ($made.Count -ge 1) {
      $b3 = Repair-Snapshot
      $ns = $excel.Workbooks.Open($made[0].FullName, 0, $true)
      $openedBooks += $ns
      Check 'the new study opens without a repair log' ((Repair-New $b3).Count -eq 0) ''
      $alts = @(); foreach ($s2 in $ns.Worksheets) { if ($s2.Name -like 'Alt *') { $alts += $s2.Name } }
      Check 'it carries no alternatives from the study you had open' ($alts.Count -eq 0) ($alts -join ', ')
      Check 'and its inputs are blank' `
        ("$($ns.Worksheets('General Information').Range('D9').Value2)" -eq '') ''
    }

    Write-Host ''
    Write-Host '  Now click the Alternative Setup button on General Information, add one alternative,'
    Write-Host '  then close the form. (The form is modal, so it has to be you, not the script.)'
    $sheetsBefore = @(); foreach ($s2 in $wb.Worksheets) { $sheetsBefore += $s2.Name }
    $wb.Activate(); $wb.Worksheets('General Information').Activate()
    Read-Host '  Press Enter when the form is closed'
    $sheetsAfter = @(); foreach ($s2 in $wb.Worksheets) { $sheetsAfter += $s2.Name }
    $new = @($sheetsAfter | Where-Object { $sheetsBefore -notcontains $_ })
    Check 'Alternative Setup built a worksheet for the alternative' ($new.Count -ge 1) ($new -join ', ')
    if ($new.Count -ge 1) {
      $as = $wb.Worksheets($new[0])
      Check 'the new sheet carries the pay-item pickers' ($as.OLEObjects().Count -gt 0) "$($as.OLEObjects().Count) controls"
      Check 'and the Summary picked it up' `
        ("$($wb.Worksheets('Summary').Range('G12').Value2)" -ne '') `
        "Summary!G12='$($wb.Worksheets('Summary').Range('G12').Value2)'"
      $saved = Join-Path $OutDir 'after-alternative-setup.xlsm'
      $wb.SaveCopyAs($saved)
      $b4 = Repair-Snapshot
      $reopened = $excel.Workbooks.Open($saved, 0, $true)
      $openedBooks += $reopened
      Check 'a workbook saved after the form ran reopens without a repair log' `
        ((Repair-New $b4).Count -eq 0) $saved
    }
  }
}
finally {
  Section 'RESULT'
  if ($shippedHash -and (Test-Path $Workbook)) {
    # this run worked on a copy; the file in the repo has to come out byte for byte unchanged
    Check 'the shipped workbook was never written to' `
      ((Get-FileHash $Workbook -Algorithm SHA256).Hash -eq $shippedHash) "SHA256 $shippedHash"
  }
  $json = Join-Path $OutDir 'results.json'
  @{ workbook = $Workbook
     excel    = if ($excel) { "$($excel.Version) build $($excel.Build)" } else { 'not started' }
     when     = (Get-Date).ToString('s')
     failed   = $script:Fails
     checks   = $script:Results } | ConvertTo-Json -Depth 5 | Set-Content $json -Encoding UTF8
  Write-Host ("{0} checks, {1} failed" -f $script:Results.Count, $script:Fails)
  if ($script:Fails -eq 0) { Write-Host 'ALL PASS' }
  else { Write-Host ('FAILED: ' + (($script:Results | Where-Object { -not $_.ok } | ForEach-Object { $_.check }) -join '; ')) }
  Write-Host "wrote $json"

  foreach ($b in $openedBooks) { try { $b.Close($false) } catch { } }
  if ($excel) { try { $excel.Quit() } catch { } ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Stop-Transcript | Out-Null
}
exit $(if ($script:Fails -gt 0) { 1 } else { 0 })
