# GNC Overall Sale Performance Dashboard - daily auto-create (started by GNC_Auto.bat).
# Report date = yesterday (or -ReportDate). Steps, the same as the manual routine:
#   1. Rohan's "GNC Dashboard Data" mail -> Order.xlsx (sale raw + Zecpe) and Allocation.xlsx
#   2. CDR: last day's "GNC <n> CDR ... .xlsb" copied to "GNC <day> CDR ... .xlsb", old CDR rows removed
#      (2 kept), yesterday's outbound CDR pasted, K:N dragged, sorted by Priorty
#   3. Dashboard:
#      - "Date & Camp wise Overall Sale": new orders appended (OrderID already there = skipped), formulas dragged
#      - "Zecpe": new rows appended (Order name already there = skipped)
#      - "Allocation": yesterday's allocation pasted A:Q (Date = yesterday, T-1), R:U = VLOOKUP on the CDR
#        (not found: "-" / "Pending to call "), V = S as value, W dragged; earlier days that got
#        connected in yesterday's CDR take that call (R:U)
#      - "APR-Utilization Raw": GNC APR + PP APR (cleaned like GNCapr_BVA) appended for yesterday
#      - "LOB&Agent Wise Sale Made Perf"!AZ2 = yesterday's day; pivots (Dashboard, Shiprocket Status,
#        Zecpe) pointed at all sale rows and refreshed; saved
# Running it again for the same day replaces that day (no duplicates).
# The dialer files (outbound CDR, GNC APR, PP APR) come from the morning tasks GNCOutCDR / GNCAPR /
# GNCall in -DialerDirs; they are recognised by content (report date inside the file), not by name.
param(
    [string]$Root = "D:\MIS\GNC\Automate",
    [datetime]$ReportDate = (Get-Date).Date.AddDays(-1),
    [string[]]$DialerDirs = @("D:\MIS\GNC\Sep'26\Auto", "D:\MIS\GNC\Oct'26\Auto"),
    [string]$MailDir,                 # testing: use Order.xlsx / Allocation.xlsx already in this folder
    [switch]$KeepCdrIfExists,
    [string]$StopAfter                # testing: save and stop after sale / zecpe / alloc / follow / apr / az2
)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$D = $ReportDate.Date
$MonAbbr = 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'
$tag = "$($MonAbbr[$D.Month - 1])'$($D.ToString('yy'))"
$monthDir = Join-Path $Root $tag
$logDir = Join-Path $Root 'Logs'; New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir 'gnc-auto-log.txt'
function Log($m) { $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $m"; Write-Host $line; Add-Content -LiteralPath $log -Value $line -Encoding utf8 }
# Excel takes a 2-D block only as a 1-based array when written through IDispatch
function To1($a) {
    if (-not ($a -is [array]) -or $a.Rank -ne 2 -or $a.GetLowerBound(0) -eq 1) { return , $a }
    $r = $a.GetLength(0); $c = $a.GetLength(1)
    $b = [Array]::CreateInstance([object], [int[]]($r, $c), [int[]](1, 1))
    for ($i = 0; $i -lt $r; $i++) { for ($j = 0; $j -lt $c; $j++) { $b[($i + 1), ($j + 1)] = $a[$i, $j] } }
    , $b
}function Set-V2($rng, $val) { $val = To1 $val; try { [void]$rng.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $rng, @(, $val)) } catch { throw "writing $($rng.Worksheet.Name)!$($rng.Address(0,0)) failed: $($_.Exception.InnerException.Message)" } }
function New-Excel { $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id); $xl = New-Object -ComObject Excel.Application; $p = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id) | Select-Object -First 1; , @($xl, $p) }
function Stop-Excel($xl, $xpid) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 800; if ($xpid) { Get-Process -Id $xpid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } }
# A row is copied down in pieces that leave out the cells a spilled formula fills (e.g. G of
# =VLOOKUP(..,{5,6},0) in F): copying those cells pastes them as fixed values, which blocks the spill
# and leaves Excel's calc chain damaged (the file then opens only with "repair").
function Is-SpillChild($cell) { try { if (-not $cell.HasSpill) { return $false }; $sp = $cell.SpillParent; if ($null -eq $sp) { return $false }; return ($sp.Row -ne $cell.Row -or $sp.Column -ne $cell.Column) } catch { return $false } }
function Copy-RowDown($ws, [int]$tpl, [string]$c1, [string]$c2, [int]$a, [int]$b) {
    # formulas written (not copied) from the template row; spill cells left for their formula; number formats kept
    $s = $ws.Range("$c1$tpl").Column; $e = $ws.Range("$c2$tpl").Column
    for ($c = $s; $c -le $e; $c++) {
        $cell = $ws.Cells.Item($tpl, $c); $dst = $ws.Range($ws.Cells.Item($a, $c), $ws.Cells.Item($b, $c))
        $dst.NumberFormat = $cell.NumberFormat
        # a spill cell reports its parent's formula text: same text as the cell on its left = part of that spill
        $child = Is-SpillChild $cell
        if (-not $child -and $c -gt 1) { $left = $ws.Cells.Item($tpl, $c - 1); if ($left.HasFormula -and "$($left.Formula2)" -eq "$($cell.Formula2)") { $child = $true } }
        if ($cell.HasFormula -and -not $child) { $dst.Formula2R1C1 = $cell.Formula2R1C1 }
    }
}# Rohan's cells are copied like a manual copy-paste (text such as '245101 keeps its apostrophe):
# source rows $idx (1-based data rows, header = row 1) of columns $sc1:$sc2 -> destination rows from $destRow, column $dc
function Copy-Rows($src, [int[]]$idx, [string]$sc1, [string]$sc2, $dst, [int]$destRow, [string]$dc) {
    $k = 0
    while ($k -lt $idx.Count) {
        $e = $k; while ($e + 1 -lt $idx.Count -and $idx[$e + 1] -eq $idx[$e] + 1) { $e++ }
        $r1 = $idx[$k] + 1; $r2 = $idx[$e] + 1
        [void]$src.Range("$sc1${r1}:$sc2$r2").Copy($dst.Range("$dc$($destRow + $k)"))
        $k = $e + 1
    }
}# New rows get the look of the rows above (font, fill, borders, number formats), like Paste Special > Formats
function Copy-Format($ws, [int]$tpl, [string]$c1, [string]$c2, [int]$a, [int]$b) {
    [void]$ws.Range("$c1${tpl}:$c2$tpl").Copy()
    [void]$ws.Range("$c1${a}:$c2$b").PasteSpecial(-4122)
    try { $ws.Application.CutCopyMode = 1 } catch {}                                          # clear the copy marquee
}function LastRow($ws, [int]$col) { $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }
function Stop-Here([string]$stage) { if ($StopAfter -eq $stage) { $script:x.Calculation = -4105; $script:x.CalculateFullRebuild(); $script:wb.Save(); Log "TEST: stopped after $stage (saved)"; $script:stopped = $true } }
function Arr([int]$r, [int]$c) { New-Object 'object[,]' $r, $c }

Log "==== GNC dashboard auto-create, report date $($D.ToString('dd-MMM-yyyy')) ===="
$dash = Join-Path $monthDir "GNC_Overall_Sale_Performance_Dashboard $tag.xlsb"
if (-not (Test-Path -LiteralPath $dash)) {     # file renamed (e.g. "... Oct'26as.xlsb"): use the one dashboard in the month folder
    $alt = @(Get-ChildItem -LiteralPath $monthDir -Filter "GNC_Overall_Sale_Performance_Dashboard*.xlsb" -ErrorAction SilentlyContinue)
    if ($alt.Count -eq 1) { $dash = $alt[0].FullName }
}
$cdrDir = Join-Path $monthDir "CDR $tag"
if (-not (Test-Path -LiteralPath $dash)) { Log "STOP: dashboard not found: $dash"; exit 1 }
foreach ($f in @($dash) + @(Get-ChildItem -LiteralPath $cdrDir -Filter '*.xlsb' | ForEach-Object FullName)) {
    if (Test-Path -LiteralPath (Join-Path (Split-Path $f) ('~$' + (Split-Path $f -Leaf)))) { Log "STOP: $(Split-Path $f -Leaf) is open in Excel. Close it and run again."; exit 1 }
}

# ---------------------------------------------------------------- 1. Rohan's mail
if (-not $MailDir) {
    $MailDir = Join-Path $monthDir ("Data\Rohan\" + $D.ToString('yyyy-MM-dd'))
    Log "Downloading Rohan's 'GNC Dashboard Data' mail..."
    $out = & "C:\Python314\python.exe" -I (Join-Path $here 'fetch_rohan.py') --date $D.ToString('yyyy-MM-dd') --out $MailDir 2>&1
    $out | ForEach-Object { Log "  $_" }
    if ($LASTEXITCODE -ne 0) { Log "STOP: mail step failed (see above). Nothing was changed."; exit 1 }
}
$orderXlsx = Join-Path $MailDir 'Order.xlsx'; $allocXlsx = Join-Path $MailDir 'Allocation.xlsx'; if (-not (Test-Path -LiteralPath $allocXlsx) -and (Test-Path -LiteralPath (Join-Path $MailDir 'Allocation.csv'))) { $allocXlsx = Join-Path $MailDir 'Allocation.csv' }

# ---------------------------------------------------------------- dialer files of the report date
$dstr = $D.ToString('yyyy-MM-dd')
$recent = Get-ChildItem -LiteralPath ($DialerDirs | Where-Object { Test-Path -LiteralPath $_ }) -File | Where-Object { $_.LastWriteTime -ge $D -and $_.LastWriteTime -lt $D.AddDays(4) } | Sort-Object LastWriteTime -Descending
# outbound CDR straight from the MIS dashboard (http://192.168.11.240/mis_dashboardGNC, "CDR Report OB" = row.php,
# the same CdrReport file as the dialer download); the dialer folders are only the fallback
$cdrHtml = $null
try {
    $cdrDl = Join-Path $root "Data\CDR"; New-Item -ItemType Directory -Force -Path $cdrDl | Out-Null
    $cdrFile = Join-Path $cdrDl "CdrReport OB $dstr.xls"
    $resp = Invoke-WebRequest -UseBasicParsing -Method Post -Uri 'http://192.168.11.240/mis_dashboardGNC/row.php' -TimeoutSec 600 `
        -Body @{ ClientCampName = ''; start_date = $dstr; end_date = $dstr; btnExport = 'Export' }
    [IO.File]::WriteAllBytes($cdrFile, $resp.RawContentStream.ToArray())
    $t = [IO.File]::ReadAllText($cdrFile)
    if ($t.Contains('>Phone Number<') -and $t.Contains('>Call Code<') -and $t.Contains(">$dstr<")) {
        $cdrHtml = Get-Item -LiteralPath $cdrFile
        Log "  outbound CDR downloaded from the MIS dashboard (CDR Report OB): $(([regex]::Matches($t, '<tr')).Count - 1) calls"
    } else { Log "  MIS dashboard CDR download has no calls of $dstr - using the dialer folder" }
} catch { Log "  MIS dashboard CDR download failed ($($_.Exception.Message)) - using the dialer folder" }
if (-not $cdrHtml) {
    $cdrHtml = $recent | Where-Object { $_.Extension -eq '.xls' } | Where-Object {
        $t = [IO.File]::ReadAllText($_.FullName); $t.Contains('>Phone Number<') -and $t.Contains('>Call Code<') -and $t.Contains(">$dstr<") } | Select-Object -First 1
}
# GNC APR + PP APR straight from the dialer http://192.168.11.240 (fetch_apr.py); the dialer folders are the fallback
$aprDir = Join-Path $root 'Data\APR'
$aprOut = & 'C:\Python314\python.exe' (Join-Path $PSScriptRoot 'fetch_apr.py') --date $dstr --out $aprDir 2>&1
$aprOut | ForEach-Object { Log "  APR download: $_" }
$aprDl = @(Get-ChildItem -LiteralPath $aprDir -Filter "*APR $dstr.csv" -ErrorAction SilentlyContinue)
$recent = @($aprDl) + @($recent)
$aprGnc = $recent | Where-Object { $_.Extension -eq '.csv' -and $_.Name -like 'GNCAPR*' -and (Get-Content -LiteralPath $_.FullName -TotalCount 2)[1] -like "*$dstr 00:00:00*" } | Select-Object -First 1
$aprPP = $recent | Where-Object { $_.Extension -eq '.csv' -and $_.Name -like 'PPAPR*' -and (Get-Content -LiteralPath $_.FullName -TotalCount 2)[1] -like "*$dstr 00:00:00*" } | Select-Object -First 1
foreach ($p in @(@('outbound CDR', $cdrHtml), @('GNC APR', $aprGnc), @('PP APR', $aprPP))) {
    if (-not $p[1]) { Log "STOP: no $($p[0]) file for $($D.ToString('dd-MMM')) in $($DialerDirs -join ' / ') (morning download task not run yet?)"; exit 1 }
    Log "  $($p[0]): $($p[1].Name)"
}

# ---------------------------------------------------------------- APR cleaning (GNCapr_BVA)
$aprHeaders = 'USER', 'ID', 'CALLS', 'Login Time', 'WAIT', 'TALK', 'DISPO', 'PAUSE', 'Login', 'Logout', 'ACHT', 'AOC', 'BIO', 'Bre', 'Briefing', 'DOWN', 'Lunch', 'Meet', 'QA', 'SB', 'Tea break', 'Training break', 'Wash', 'Tea'
function TimeVal([string]$s) {
    $s = $s.Trim(); if ($s -match '^(\d+):(\d{1,2}):(\d{1,2})$') { return ([int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]) / 86400.0 }
    $n = 0.0; if ([double]::TryParse($s, [ref]$n)) { return $n }; $s
}
function Clean-Apr([string]$path) {
    $lines = @(Get-Content -LiteralPath $path | Select-Object -Skip 3)                     # 1. first 3 rows
    while ($lines.Count -and -not $lines[-1].Trim()) { $lines = $lines[0..($lines.Count - 2)] }
    $lines = $lines[0..($lines.Count - 2)]                                                    # 2. last (TOTALS) row
    $rows = @($lines | ConvertFrom-Csv -Header (1..200 | ForEach-Object { "c$_" }))
    $hdr = @(); foreach ($p in $rows[0].PSObject.Properties) { if ($null -ne $p.Value -and "$($p.Value)" -ne '') { $hdr += "$($p.Value)".Trim() } }
    # 3-6: keep wanted columns (Excel MATCH = case-insensitive, first hit), add missing as 00:00:00, order as the list
    $out = @()
    foreach ($r in $rows | Select-Object -Skip 1) {
        $vals = @($r.PSObject.Properties | ForEach-Object { "$($_.Value)" })
        $o = [ordered]@{}
        foreach ($h in $aprHeaders) {
            $ix = -1; for ($k = 0; $k -lt $hdr.Count; $k++) { if ($hdr[$k] -ieq $h) { $ix = $k; break } }
            $o[$h] = if ($ix -ge 0) { TimeVal $vals[$ix] } else { 0.0 }
        }
        $out += [pscustomobject]$o
    }
    $out
}
$aprRows = @(Clean-Apr $aprGnc.FullName) + @(Clean-Apr $aprPP.FullName)
Log "  APR rows: $(@(Clean-Apr $aprGnc.FullName).Count) GNC + $(@(Clean-Apr $aprPP.FullName).Count) PP"

# ---------------------------------------------------------------- backup
$bk = Join-Path $Root 'Backups'; New-Item -ItemType Directory -Force -Path $bk | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd HHmm'
Copy-Item -LiteralPath $dash -Destination (Join-Path $bk "Dashboard before $stamp.xlsb")
Get-ChildItem -LiteralPath $bk -Filter 'Dashboard before *.xlsb' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 | Remove-Item -Force

$x = $null
try {
    $nx = New-Excel; $x = $nx[0]; $xpid = $nx[1]
    $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false; $x.EnableEvents = $false; $x.AutomationSecurity = 3

    # ------------------------------------------------------------ 2. CDR workbook for the day
    $cdrFiles = foreach ($f in Get-ChildItem -LiteralPath $cdrDir -Filter 'GNC*CDR*.xlsb') {
        if ($f.Name -match '^GNC\s+(\d+)(?:-(\d+))?\s+CDR(.*)$') { [pscustomobject]@{ F = $f; Day = [int]$(if ($Matches[2]) { $Matches[2] } else { $Matches[1] }); Rest = $Matches[3] } } }
    $prev = $cdrFiles | Where-Object { $_.Day -lt $D.Day } | Sort-Object Day -Descending | Select-Object -First 1
    if (-not $prev) { throw "no earlier 'GNC <day> CDR' file in $cdrDir to start from" }
    $cdrName = "GNC $($D.Day) CDR$($prev.Rest)"; $cdrPath = Join-Path $cdrDir $cdrName
    if (Test-Path -LiteralPath $cdrPath) { Move-Item -LiteralPath $cdrPath -Destination (Join-Path $bk "$([IO.Path]::GetFileNameWithoutExtension($cdrName)) before $stamp.xlsb") -Force }
    Copy-Item -LiteralPath $prev.F.FullName -Destination $cdrPath
    $cw = $x.Workbooks.Open($cdrPath); $x.Calculation = -4105
    $cs = $cw.Worksheets.Item('CDR')
    if ($cs.FilterMode) { $cs.ShowAllData() }
    $ul = $cs.UsedRange.Row + $cs.UsedRange.Rows.Count - 1
    if ($ul -ge 4) { [void]$cs.Range($cs.Rows.Item(4), $cs.Rows.Item($ul)).Delete() }      # keep header + 2 rows
    $hw = $x.Workbooks.Open($cdrHtml.FullName); $hs = $hw.Worksheets.Item(1)
    $hn = LastRow $hs 1
    [void]$hs.Range($hs.Cells.Item(2, 1), $hs.Cells.Item($hn, 10)).Copy($cs.Cells.Item(4, 1))
    $hw.Close($false)
    $cn = 3 + $hn - 1
    # K:N for every row: K = Agent (A), L:N = Calling Status / SUB SCENARIOS 1 / Priorty from this file's
    # 'Short Code' sheet (plain formulas - the old ones were array formulas linked to "GNC 28 CDR 26 Jun.xlsb")
    $sc = $cw.Worksheets.Item('Short Code'); $sl2 = $sc.Cells.Item($sc.Rows.Count, 1).End(-4162).Row
    $scv = $sc.Range("A1:D$sl2").Value2; $caRow = 0
    for ($i = 2; $i -le $sl2; $i++) { if ("$($scv[$i,1])".Trim() -eq 'CA') { $caRow = $i } }
    if (-not $caRow) { $caRow = $sl2 + 1; Set-V2 $sc.Range("A$caRow") 'CA' }
    if (-not "$($sc.Range("B$caRow").Value2)") { Set-V2 $sc.Range("B$caRow") 'Connected' }
    if (-not "$($sc.Range("C$caRow").Value2)") { Set-V2 $sc.Range("C$caRow") 'Auto Rap' }        # same as the Jun list
    [void]$cs.Range("K2:N$cn").ClearContents()
    $cs.Range("K2:K$cn").FormulaR1C1 = '=RC1'
    $cs.Range("L2:L$cn").FormulaR1C1 = "=VLOOKUP(RC4,'Short Code'!C1:C4,2,0)"
    $cs.Range("M2:M$cn").FormulaR1C1 = "=VLOOKUP(RC4,'Short Code'!C1:C4,3,0)"
    $cs.Range("N2:N$cn").FormulaR1C1 = "=VLOOKUP(RC4,'Short Code'!C1:C4,4,0)"
    Copy-Format $cs 3 'A' 'N' 4 $cn                                                     # pasted calls get the format of the row above
    $x.CalculateFull()
    $chk = $cs.Range("N2:N$cn").Value2; $filled = 0; foreach ($v in $chk) { if ($v -is [double]) { $filled++ } }
    if ($filled -lt ($cn - 1) * 0.9) { throw "CDR: Priorty (N) found for only $filled of $($cn - 1) calls (unknown Call Codes?)" }
    $sortRange = $cs.Range("A1:N$cn")
    [void]$sortRange.Sort($cs.Range("N2"), 1, [Type]::Missing, [Type]::Missing, 1, [Type]::Missing, 1, 1)   # A-Z on Priorty, header
    $x.Calculation = -4105; $x.CalculateFullRebuild()
    $cw.Save()
    Log "CDR: $cdrName created from $($prev.F.Name) - $($hn - 1) calls of $($D.ToString('dd-MMM')) + 2 old rows, sorted by Priorty"
    # VLOOKUP(Q, CDR!$B:$N, {10,11,12,2}, 0) done here, so the dashboard gets values and no link to the CDR file:
    # first row (in Priorty order) whose Phone Number equals the allocation number (number = number, text = text, case-insensitive)
    $cv = $cs.Range("B2:N$cn").Value2
    $look = @{}
    for ($i = 1; $i -le $cn - 1; $i++) {
        $p = $cv[$i, 1]; if ($null -eq $p -or "$p" -eq '') { continue }
        $k = if ($p -is [double]) { "n:" + $p.ToString('R') } else { "s:" + "$p".ToLower() }
        if (-not $look.ContainsKey($k)) { $look[$k] = @($cv[$i, 10], $cv[$i, 11], $cv[$i, 12], $cv[$i, 2]) }
    }
    function Look($q) { if ($null -eq $q -or "$q" -eq '') { return $null }; $k = if ($q -is [double]) { "n:" + $q.ToString('R') } else { "s:" + "$q".ToLower() }; $look[$k] }
    $cw.Close($false)                                                                          # CDR saved; the dashboard needs only the values above

    # ------------------------------------------------------------ 3. dashboard
    $wb = $x.Workbooks.Open($dash); $x.Calculation = -4135
    $ow = $x.Workbooks.Open($orderXlsx); $aw = $x.Workbooks.Open($allocXlsx)
    Stop-Here 'open'; if ($stopped) { return }

    # 3a. sale raw -> Date & Camp wise Overall Sale (new OrderIDs only)
    $ss = $wb.Worksheets.Item('Date & Camp wise Overall Sale'); if ($ss.FilterMode) { $ss.ShowAllData() }
    # Rohan's file holds the whole month: all rows of the report date are added (a re-run replaces them), plus
    # late rows of earlier days whose OrderID is not in the sheet at all (OrderIDs repeat, so only for older days)
    $oa = $D.ToOADate()
    $sl = LastRow $ss 13
    $sb = $ss.Range("B1:B$sl").Value2; $old = @(); for ($i = 2; $i -le $sl; $i++) { if ($sb[$i, 1] -is [double] -and $sb[$i, 1] -eq $oa) { $old += $i } }
    for ($k = $old.Count - 1; $k -ge 0; $k--) { [void]$ss.Rows.Item($old[$k]).Delete() }
    if ($old.Count) { Log "Sale: $($old.Count) earlier row(s) of $($D.ToString('dd-MMM')) removed (re-run)" }
    $sl = LastRow $ss 13
    $have = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($e in $ss.Range("M2:M$sl").Value2) { if ($e) { [void]$have.Add("$e".Trim()) } }
    $os = $ow.Worksheets.Item(1); $on = LastRow $os 13
    $ov = $os.Range("A2:S$on").Value2
    $new = @(); $late = 0
    for ($i = 1; $i -le $on - 1; $i++) { if ($ov[$i, 2] -is [double] -and $ov[$i, 2] -eq $oa) { $new += $i } }
    # (late rows of other days are NOT added any more - 12-Oct-26: the team removes them by hand)
    if ($new.Count) {
        $a = $sl + 1; $b = $sl + $new.Count
        Copy-RowDown $ss $sl 'A' 'S' $a $b                         # formulas dragged (A, D:G, S)
        Copy-Rows $os $new 'B' 'C' $ss $a 'B'; Copy-Rows $os $new 'H' 'R' $ss $a 'H'
        Copy-Format $ss $sl 'A' 'S' $a $b
    }
    # (no OrderID de-duplication: Rohan's OrderID column repeats one ID on whole days - 12-Oct-26, all rows stay)
    $saleLast = LastRow $ss 13
    Log "Sale: $($new.Count - $late) row(s) of $($D.ToString('dd-MMM')) + $late late order(s) of earlier days (OrderID not in the sheet yet) appended, sheet rows 2-$saleLast"
    Stop-Here 'sale'; if ($stopped) { return }

    # 3b. Zecpe (new Order names only)
    $zs = $wb.Worksheets.Item('Zecpe'); if ($zs.FilterMode) { $zs.ShowAllData() }
    $zl = LastRow $zs 2
    $zh = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($e in $zs.Range("B2:B$zl").Value2) { if ($e) { [void]$zh.Add("$e".Trim()) } }
    $zsrc = $ow.Worksheets.Item(2); $zn = LastRow $zsrc 2; $zv = $zsrc.Range("A2:G$zn").Value2
    $zadd = @(); for ($i = 1; $i -le $zn - 1; $i++) { $id = "$($zv[$i,2])".Trim(); if ($id -and -not $zh.Contains($id)) { $zadd += $i; [void]$zh.Add($id) } }
    if ($zadd.Count) {
        $a = $zl + 1; $b = $zl + $zadd.Count
        Copy-RowDown $zs $zl 'A' 'G' $a $b
        Copy-Rows $zsrc $zadd 'A' 'G' $zs $a 'A'
        Copy-Format $zs $zl 'A' 'G' $a $b
    }
    Log "Zecpe: $($zadd.Count) new row(s) appended"
    Stop-Here 'zecpe'; if ($stopped) { return }
    $ow.Close($false)

    # 3c. Allocation
    $as = $wb.Worksheets.Item('Allocation'); if ($as.FilterMode) { $as.ShowAllData() }
    $al = LastRow $as 2
    $bv = $as.Range("B1:B$al").Value2; $oa = $D.ToOADate()
    $first = 0; for ($i = 2; $i -le $al; $i++) { if ($bv[$i, 1] -is [double] -and $bv[$i, 1] -eq $oa) { $first = $i; break } }
    if ($first) { [void]$as.Range($as.Rows.Item($first), $as.Rows.Item($al)).Delete(); $al = $first - 1; Log "Allocation: earlier rows of $($D.ToString('dd-MMM')) removed (re-run)" }
    $src = $aw.Worksheets.Item(1); $an = LastRow $src 1; $sv = $src.Range("A2:M$an").Value2; $n = $an - 1
    $a = $al + 1; $b = $al + $n
    Copy-RowDown $as $al 'A' 'W' $a $b                               # formats + A / V / W formulas
    [void]$src.Range("A2:A$an").Copy($as.Range("D$a"))                                         # D Date Type (SKU_Type)
    [void]$src.Range("B2:M$an").Copy($as.Range("F$a"))                                         # F Store .. Q ShippingPhone
    $bce = Arr $n 2; $ec = Arr $n 1
    for ($k = 0; $k -lt $n; $k++) { $bce[$k, 0] = $oa; $bce[$k, 1] = 'T-1'; $ec[$k, 0] = '-' }   # B Date, C Helper, E Time Slot
    Set-V2 $as.Range("B${a}:C$b") $bce; Set-V2 $as.Range("E${a}:E$b") $ec
    Copy-Format $as $al 'A' 'W' $a $b
    $qv = $as.Range("Q${a}:Q$b").Value2; $miss = 0
    $ru = Arr $n 4; $vv = Arr $n 1
    for ($k = 0; $k -lt $n; $k++) {
        $hit = Look $qv[($k + 1), 1]
        if (-not $hit) { $ru[$k, 0] = '-'; $ru[$k, 1] = 'Pending to call '; $ru[$k, 2] = '-'; $ru[$k, 3] = $oa; $miss++ }
        else { for ($j = 0; $j -lt 4; $j++) { $ru[$k, $j] = $hit[$j] } }
        $vv[$k, 0] = $ru[$k, 1]
    }    Set-V2 $as.Range("R${a}:U$b") $ru
    Set-V2 $as.Range("V${a}:V$b") $vv                                                         # Same Day Connect = S (value)
    Log "Allocation: $n row(s) of $($D.ToString('dd-MMM')) pasted (rows $a-$b), $($n - $miss) found in CDR, $miss pending to call"
    Stop-Here 'alloc'; if ($stopped) { return }

    # earlier days: connected in yesterday's CDR -> take that call
    $pl = $a - 1
    $qv = $as.Range("Q2:Q$pl").Value2; $cur = $as.Range("R2:U$pl").Value2; $upd = 0
    for ($i = 1; $i -le $pl - 1; $i++) {
        $hit = Look $qv[$i, 1]
        if ($hit -and "$($hit[1])" -eq 'Connected') { for ($j = 1; $j -le 4; $j++) { $cur[$i, $j] = $hit[$j - 1] }; $upd++ }
    }
    Set-V2 $as.Range("R2:U$pl") $cur    Log "Allocation: $upd earlier allocation(s) connected on $($D.ToString('dd-MMM')) updated (R:U)"
    Stop-Here 'follow'; if ($stopped) { return }
    $aw.Close($false)

    # 3d. APR-Utilization Raw
    $ps = $wb.Worksheets.Item('APR-Utilization Raw'); if ($ps.FilterMode) { $ps.ShowAllData() }
    $pl2 = LastRow $ps 2
    $pb = $ps.Range("B1:B$pl2").Value2; $pf = 0; for ($i = 2; $i -le $pl2; $i++) { if ($pb[$i, 1] -is [double] -and $pb[$i, 1] -eq $oa) { $pf = $i; break } }
    if ($pf) { [void]$ps.Range($ps.Rows.Item($pf), $ps.Rows.Item($pl2)).Delete(); $pl2 = $pf - 1 }
    $hdrRow = @(); for ($c = 1; $c -le 33; $c++) { $hdrRow += "$($ps.Cells.Item(1, $c).Text)".Trim() }
    $m = $aprRows.Count; $a = $pl2 + 1; $b = $pl2 + $m
    Copy-RowDown $ps $pl2 'A' 'AG' $a $b
    $pv = $ps.Range("A${a}:AG$b").Formula                                                     # keep formula cells, overwrite constants
    $out = Arr $m 33
    for ($k = 0; $k -lt $m; $k++) {
        for ($c = 0; $c -lt 33; $c++) { $out[$k, $c] = $pv[($k + 1), ($c + 1)] }
        $out[$k, 1] = $oa
        foreach ($h in $aprHeaders) { $c = [array]::IndexOf(($hdrRow | ForEach-Object { $_.ToLower() }), $h.ToLower()); if ($c -ge 0 -and -not "$($out[$k,$c])".StartsWith('=')) { $out[$k, $c] = $aprRows[$k].$h } }
        for ($c = 0; $c -lt 33; $c++) { if ($hdrRow[$c] -eq 'Downtime') { $out[$k, $c] = 0.0 } }
    }
    # constants as values, formula cells keep their (dragged) formula text
    $rg = $ps.Range("A${a}:AG$b"); [void]$rg.GetType().InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $rg, @(, (To1 $out)))
    Copy-Format $ps $pl2 'A' 'AG' $a $b
    Log "APR-Utilization Raw: $m row(s) of $($D.ToString('dd-MMM')) (rows $a-$b)"
    Stop-Here 'apr'; if ($stopped) { return }

    # 3e. AZ2 = day of the report date
    Set-V2 ($wb.Worksheets.Item('LOB&Agent Wise Sale Made Perf').Range('AZ2')) ([double]$D.Day)
    Stop-Here 'az2'; if ($stopped) { return }

    # 3f. pivots on the sale sheet -> all rows, refreshed
    $x.Calculation = -4105; $x.CalculateFullRebuild()
    $byCache = @{}; $okN = 0; $bad = @()
    foreach ($sn in 'Dashboard', 'Shiprocket Status', 'Zecpe') {
        foreach ($pt in @($wb.Worksheets.Item($sn).PivotTables())) {
            try {
                $sd = "$($pt.PivotCache().SourceData)"
                if ($sd -match "^'Date & Camp wise Overall Sale'!R1C(\d+):R\d+C(\d+)$") {
                    $nsd = "'Date & Camp wise Overall Sale'!R1C$($Matches[1]):R${saleLast}C$($Matches[2])"
                    if ($sd -ne $nsd) {
                        try { $pt.SourceData = $nsd } catch { [void]$pt.ChangePivotCache($wb.PivotCaches().Create(1, $nsd)) }
                    }
                }
                [void]$pt.RefreshTable(); $okN++
                # date fields oldest -> newest (a manual sort puts each new day at the end, e.g. 09-Oct before 05-Oct)
                foreach ($f in @($pt.RowFields()) + @($pt.ColumnFields())) {
                    if ($f.DataType -eq 2) { [void]$f.AutoSort(1, $f.SourceName) }    # xlDate, xlAscending
                }
            } catch { $bad += "$sn/$($pt.Name): $($_.Exception.Message)" }
        }
    }
    Log "Pivots: $okN refreshed on sale rows 1-$saleLast, dates sorted oldest first$(if ($bad) { '; NOT refreshed: ' + ($bad -join '; ') })"
    $wb.Save(); $wb.Close($false)
    Log "DONE: $dash"
}
catch {
    $ie = $_.Exception; $msgs = @(); while ($ie) { $msgs += $ie.Message; $ie = $ie.InnerException }
    Log "FAILED: $($msgs -join ' | ') (line $($_.InvocationInfo.ScriptLineNumber))"
    Log "Nothing was saved to the dashboard; backup: $(Join-Path $bk "Dashboard before $stamp.xlsb")"
    exit 1
}
finally {
    if ($x) {
        foreach ($w in @($x.Workbooks)) { try { $w.Close($false) } catch {} }
        Stop-Excel $x $xpid
    }
    [GC]::Collect()
}
