# GNC Inbound MIS Dashboard - daily raw-data update (started by GNC_Inbound.bat).
# Report date = yesterday (or -ReportDate). Workbook: <Root>\<Mon'yy>\GNC Inbound MIS Dashboard <Mon'yy>.xlsb
#   1. looks at the last date already in each raw sheet and downloads everything from that day to the report date
#      (fetch_inbound.py: Feedback 192.168.11.244, CDR Report IN 192.168.11.240, DialDesk In Call Details) - so missed days are caught up
#   2. pastes, header-checked, as values (text stays text, dates stay dates):
#        "Feedback Raw"        A:L   <- Feedback export         ; formulas M:P dragged   ; rows of the downloaded dates are replaced
#        "CDR Raw"             B:S   <- CDR Report IN           ; formulas A and T:AD dragged ; rows of the downloaded dates are replaced
#        "Inbound DD Tagging"  A:AG  <- In Call Details (All)   ; formulas AH:AI dragged ; only Call Ids not in the sheet are added
#   3. pivots that read those sheets are pointed at all rows and refreshed, everything recalculated, saved
# Every pasted cell is read back and compared; if anything differs nothing is saved. Running it again changes nothing.
param(
    [string]$Root = $PSScriptRoot,
    [datetime]$ReportDate = (Get-Date).Date.AddDays(-1),
    [string]$InboundPath,                 # testing: use this workbook
    [string]$DataDir,                     # testing: use the downloaded files in this folder (no download)
    [switch]$SkipDownload,                # paste the files already in the data folder
    [int]$WaitMinutes = 0,                # wait this long for the workbook to be closed (the scheduled run uses 30)
    [int]$MaxDays = 10,                   # never go back further than this many days
    [switch]$NoBackup,
    [string]$Python
)
$ErrorActionPreference = 'Stop'
$D = $ReportDate.Date
$MonAbbr = 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'
$tag = "$($MonAbbr[$D.Month - 1])'$($D.ToString('yy'))"
$monthDir = Join-Path $Root $tag
$logDir = Join-Path $Root 'Logs'; New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir 'gnc-inbound-log.txt'
function Log($m) { $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $m"; Write-Host $line; Add-Content -LiteralPath $log -Value $line -Encoding utf8 }
function New-Excel { $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id); $xl = New-Object -ComObject Excel.Application; $p = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id) | Select-Object -First 1; , @($xl, $p) }
function Stop-Excel($xl, $xpid) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 800; if ($xpid) { Get-Process -Id $xpid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } }
function LastRow($ws, [int]$col) { $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }
function Norm($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return 'n' + $v.ToString('R') }; if ($v -is [bool]) { return 'b' + $v }; if ($v -is [int]) { return 'n' + ([double]$v).ToString('R') }; 's' + "$v".Trim() }
function NormId($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return $v.ToString('R') }; "$v".Trim() }
function NormH([string]$h) { ("$h" -replace '\s+', ' ').Trim().ToLower() }
# a date cell value (number or text) as an Excel date number (whole day), or $null
function DayOf($v) {
    if ($v -is [double]) { return [Math]::Floor($v) }
    $s = "$v".Trim(); if (-not $s) { return $null }
    $dt = [datetime]::MinValue
    foreach ($f in 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd', 'dd-MM-yyyy HH:mm:ss', 'dd-MM-yyyy', 'dd/MM/yyyy', 'M/d/yyyy h:mm:ss tt') {
        if ([datetime]::TryParseExact($s, $f, [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$dt)) { return [Math]::Floor($dt.ToOADate()) } }
    $null
}
function DayList($ws, [int]$col, [int]$n) { $v = $ws.Range($ws.Cells.Item(2, $col), $ws.Cells.Item($n, $col)).Value2; $out = New-Object 'double[]' ($n - 1); for ($i = 1; $i -le $n - 1; $i++) { $d = DayOf ($v[$i, 1]); $out[$i - 1] = if ($null -eq $d) { -1 } else { $d } }; , $out }

# the three raw sheets: pasted columns start at DestCol1 (heading cells of the sheet define the names), formula columns are dragged
$sheets = @(
    @{ Key = 'Feedback'; Name = 'Feedback Raw';       File = 'Feedback';   DestCol1 = 1; N = 12; Formula = 13..16; KeyCol = 1;  Mode = 'date'; SrcDate = 'CallDate';  LastCol = 16 },
    @{ Key = 'InboundCDR'; Name = 'CDR Raw';          File = 'InboundCDR'; DestCol1 = 2; N = 18; Formula = @(1) + (20..30); KeyCol = 2; Mode = 'date'; SrcDate = 'CallDate'; LastCol = 30 },
    @{ Key = 'DDTagging'; Name = 'Inbound DD Tagging'; File = 'DDTagging'; DestCol1 = 1; N = 33; Formula = 34..35; KeyCol = 19; Mode = 'callid'; SrcDate = 'CallDate'; LastCol = 35 }
)

Log "==== GNC Inbound raw update, report date $($D.ToString('dd-MMM-yyyy')) ===="
if (-not $InboundPath) { $InboundPath = Join-Path $monthDir "GNC Inbound MIS Dashboard $tag.xlsb" }
if (-not (Test-Path -LiteralPath $InboundPath)) { Log "STOP: workbook not found: $InboundPath"; exit 1 }
# really in use? (Excel keeps the file open exclusively; a leftover ~$ lock file alone does not count)
function Test-InUse([string]$f) { try { $fs = [IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $fs.Close(); $false } catch { $true } }
$waited = 0
while (Test-InUse $InboundPath) {
    if ($waited -ge $WaitMinutes) { Log "STOP: $(Split-Path $InboundPath -Leaf) is open in Excel. Save and close it (File > Close), then run again. Nothing was changed."; exit 1 }
    if ($waited -eq 0) { Log "$(Split-Path $InboundPath -Leaf) is open in Excel - waiting up to $WaitMinutes min for it to be closed..." }
    Start-Sleep -Seconds 60; $waited++
}
if (-not $DataDir) { $DataDir = Join-Path $monthDir ("Data\Inbound\" + $D.ToString('yyyy-MM-dd')) }

$x = $null; $wb = $null
try {
    $nx = New-Excel; $x = $nx[0]; $xpid = $nx[1]
    $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false; $x.EnableEvents = $false; $x.AutomationSecurity = 3

    # ---------------------------------------------------------------- 1. last dates -> range to download
    try { $wb = $x.Workbooks.Open($InboundPath, 0, $true) } catch { throw "'$(Split-Path $InboundPath -Leaf)' cannot be opened by Excel automation. Open it once in Excel, File > Save As (same name, .xlsb), close it and run again." }
    $floor = $D.AddDays(-($MaxDays - 1)).ToOADate()
    foreach ($s in $sheets) {
        $ws = $wb.Worksheets.Item($s.Name); $n = LastRow $ws $s.KeyCol
        $last = $null; if ($n -ge 2) { $days = DayList $ws $s.KeyCol $n; $mx = ($days | Measure-Object -Maximum).Maximum; if ($mx -gt 0) { $last = $mx } }
        $from = if ($null -eq $last) { $D.ToOADate() } else { [Math]::Max([Math]::Min($last, $D.ToOADate()), $floor) }
        $s.From = $from
        Log "$($s.Name): rows 2-$n, last date $(if ($last) { [datetime]::FromOADate($last).ToString('dd-MMM') } else { '-' }) -> from $([datetime]::FromOADate($from).ToString('dd-MMM'))"
    }
    $wb.Close($false); $wb = $null
    $dFrom = [datetime]::FromOADate((($sheets | ForEach-Object { $_.From }) | Measure-Object -Minimum).Minimum)

    # ---------------------------------------------------------------- 2. download
    if (-not $SkipDownload -and -not $PSBoundParameters.ContainsKey('DataDir')) {
        if (-not $Python) { foreach ($c in 'C:\Python314\python.exe', "$env:LOCALAPPDATA\Microsoft\WindowsApps\python.exe", 'python.exe') { if (Get-Command $c -ErrorAction SilentlyContinue) { $Python = $c; break } } }
        New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
        Log "Downloading Feedback / CDR Report IN / In Call Details for $($dFrom.ToString('yyyy-MM-dd')) .. $($D.ToString('yyyy-MM-dd')) ..."
        $ErrorActionPreference = 'Continue'; $out = & $Python (Join-Path $PSScriptRoot 'fetch_inbound.py') --from $dFrom.ToString('yyyy-MM-dd') --to $D.ToString('yyyy-MM-dd') --out $DataDir 2>&1 | ForEach-Object { "$_" }; $ErrorActionPreference = 'Stop'
        $out | ForEach-Object { Log "  $_" }
    }
    $res = @{}; $rf = Join-Path $DataDir '_result.json'
    if (Test-Path -LiteralPath $rf) { $j = Get-Content -LiteralPath $rf -Raw | ConvertFrom-Json; foreach ($p in $j.PSObject.Properties) { $res[$p.Name] = $p.Value } }
    foreach ($s in $sheets) {
        $f = $null
        if ($res.ContainsKey($s.File) -and $res[$s.File].ok -and (Test-Path -LiteralPath $res[$s.File].file)) { $f = $res[$s.File].file }
        elseif ($SkipDownload -or $PSBoundParameters.ContainsKey('DataDir')) { $f = Get-ChildItem -LiteralPath $DataDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "$($s.File)*" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName }
        $s.Src = $f
        if (-not $f) { Log "SKIPPED $($s.Name): no downloaded file ($($res[$s.File].error))" }
    }
    if (-not ($sheets | Where-Object { $_.Src })) { Log "STOP: nothing downloaded. Nothing was changed."; exit 1 }

    # ---------------------------------------------------------------- backup + paste
    if (-not $NoBackup) {
        $bk = Join-Path $Root 'Backups'; New-Item -ItemType Directory -Force -Path $bk | Out-Null
        Copy-Item -LiteralPath $InboundPath -Destination (Join-Path $bk "Inbound MIS before $(Get-Date -Format 'yyyy-MM-dd HHmm').xlsb")
        Get-ChildItem -LiteralPath $bk -Filter 'Inbound MIS before *.xlsb' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 | Remove-Item -Force
    }
    $wb = $x.Workbooks.Open($InboundPath); $x.Calculation = -4135
    $failed = @()
    foreach ($s in $sheets) {
        if (-not $s.Src) { $failed += $s.Name; continue }
        $ws = $wb.Worksheets.Item($s.Name); if ($ws.FilterMode) { $ws.ShowAllData() }
        $n = $s.N; $c1 = $s.DestCol1
        $th = $ws.Range($ws.Cells.Item(1, $c1), $ws.Cells.Item(1, $c1 + $n - 1)).Value2
        # --- source file: header row, column map
        $sw = $x.Workbooks.Open($s.Src, 0, $true); $ss = $sw.Worksheets.Item(1)
        $sl = $ss.UsedRange.Row + $ss.UsedRange.Rows.Count - 1; $scn = $ss.UsedRange.Column + $ss.UsedRange.Columns.Count - 1
        $sv = $ss.Range($ss.Cells.Item(1, 1), $ss.Cells.Item($sl, $scn)).Value2
        $hr = 0; $best = -1
        for ($r = 1; $r -le [Math]::Min(6, $sl); $r++) { $hit = 0; for ($c = 1; $c -le $scn; $c++) { $h = NormH $sv[$r, $c]; for ($j = 1; $j -le $n; $j++) { if ($h -and $h -eq (NormH $th[1, $j])) { $hit++; break } } }; if ($hit -gt $best) { $best = $hit; $hr = $r } }
        $map = @(); $miss = @()
        for ($j = 1; $j -le $n; $j++) { $m = 0; for ($c = 1; $c -le $scn; $c++) { if ((NormH $sv[$hr, $c]) -eq (NormH $th[1, $j])) { $m = $c; break } }; $map += $m; if (-not $m) { $miss += $th[1, $j] } }
        if ($miss.Count) {
            if ($scn -eq $n -and $best -lt $n * 0.5) { Log "WARNING $($s.Name): headings of the file do not match - pasted by position"; $map = 1..$n; $hr = 1 }
            else { throw "$($s.Name): column(s) missing in $(Split-Path $s.Src -Leaf): $($miss -join ', ') (file has: $((1..$scn | ForEach-Object { $sv[$hr, $_] }) -join ', '))" }
        }
        $dcol = 0; for ($j = 1; $j -le $n; $j++) { if ((NormH $th[1, $j]) -eq (NormH $s.SrcDate)) { $dcol = $map[$j - 1] } }
        $idcol = 0; if ($s.Mode -eq 'callid') { for ($j = 1; $j -le $n; $j++) { if ((NormH $th[1, $j]) -eq 'call id') { $idcol = $map[$j - 1] } } }
        # --- source rows
        $keepRows = @(); $dropFuture = 0; $dayRows = @{}
        for ($r = $hr + 1; $r -le $sl; $r++) {
            $any = $false; for ($j = 0; $j -lt $n; $j++) { if ($null -ne $sv[$r, $map[$j]] -and "$($sv[$r,$map[$j]])" -ne '') { $any = $true; break } }
            if (-not $any) { continue }
            $dy = if ($dcol) { DayOf ($sv[$r, $dcol]) } else { $null }
            if ($null -ne $dy -and $dy -gt $D.ToOADate()) { $dropFuture++; continue }
            $keepRows += $r; if ($null -ne $dy) { $dayRows[$dy] = 1 + [int]$dayRows[$dy] }
        }
        # --- target rows to replace / skip
        $tn = LastRow $ws $s.KeyCol; if ($tn -lt 1) { $tn = 1 }
        $removed = 0
        if ($s.Mode -eq 'date') {
            $keepRows = @($keepRows | Where-Object { $d2 = DayOf ($sv[$_, $dcol]); $null -eq $d2 -or $d2 -ge $s.From })
            $dset = @{}; foreach ($r in $keepRows) { $d2 = DayOf ($sv[$r, $dcol]); if ($null -ne $d2) { $dset[$d2] = 1 } }
            if ($tn -ge 2 -and $dset.Count) {
                $days = DayList $ws $s.KeyCol $tn
                for ($i = $days.Length - 1; $i -ge 0; $i--) { if ($dset.ContainsKey($days[$i])) { $hi = $i; while ($i -gt 0 -and $dset.ContainsKey($days[$i - 1])) { $i-- }; [void]$ws.Range($ws.Rows.Item($i + 2), $ws.Rows.Item($hi + 2)).Delete(); $removed += $hi - $i + 1 } }
                $tn = LastRow $ws $s.KeyCol; if ($tn -lt 1) { $tn = 1 }
            }
        } else {
            $have = New-Object 'System.Collections.Generic.HashSet[string]'
            if ($tn -ge 2) { foreach ($e in $ws.Range("B2:B$tn").Value2) { if ($null -ne $e) { [void]$have.Add((NormId $e)) } } }
            $keepRows = @($keepRows | Where-Object { $id = NormId ($sv[$_, $idcol]); $id -and $have.Add($id) })
        }
        $m = $keepRows.Count
        if ($m -eq 0) { Log "$($s.Name): file has no new rows$(if ($removed) { " ($removed old row(s) of those days were removed!)" })"; $sw.Close($false); continue }
        if ($tn -lt 2) { throw "$($s.Name): no data row to copy the formulas and formats from" }
        $a = $tn + 1; $b = $tn + $m
        # --- formats + formulas from the last row, values column by column (copy/paste values keeps text as text)
        [void]$ws.Range($ws.Cells.Item($tn, 1), $ws.Cells.Item($tn, $s.LastCol)).Copy()
        [void]$ws.Range($ws.Cells.Item($a, 1), $ws.Cells.Item($b, $s.LastCol)).PasteSpecial(-4122)
        try { $x.CutCopyMode = 0 } catch {}
        foreach ($fc in $s.Formula) { $ws.Range($ws.Cells.Item($a, $fc), $ws.Cells.Item($b, $fc)).Formula2R1C1 = $ws.Cells.Item($tn, $fc).Formula2R1C1 }
        for ($j = 0; $j -lt $n; $j++) {
            $k = 0
            while ($k -lt $m) {
                $e = $k; while ($e + 1 -lt $m -and $keepRows[$e + 1] -eq $keepRows[$e] + 1) { $e++ }
                [void]$ss.Range($ss.Cells.Item($keepRows[$k], $map[$j]), $ss.Cells.Item($keepRows[$e], $map[$j])).Copy()
                [void]$ws.Cells.Item($a + $k, $c1 + $j).PasteSpecial(-4163)
                $k = $e + 1
            }
        }
        try { $x.CutCopyMode = 0 } catch {}
        # --- verify
        $x.Calculate()
        $nv = $ws.Range($ws.Cells.Item($a, $c1), $ws.Cells.Item($b, $c1 + $n - 1)).Value2; $mis = 0
        for ($i = 0; $i -lt $m; $i++) { for ($j = 0; $j -lt $n; $j++) { if ((Norm ($sv[$keepRows[$i], $map[$j]])) -ne (Norm ($nv[($i + 1), ($j + 1)]))) { $mis++; if ($mis -le 5) { Log "  MISMATCH $($s.Name) row $($a + $i) col $($c1 + $j): '$($sv[$keepRows[$i],$map[$j]])' vs '$($nv[($i + 1),($j + 1)])'" } } } }
        if ($mis) { throw "$($s.Name): $mis pasted cell(s) differ from the downloaded file - nothing saved" }
        $fv = $ws.Range($ws.Cells.Item($a, 1), $ws.Cells.Item($b, $s.LastCol)).Value2; $err = @{}
        foreach ($fc in $s.Formula) { for ($i = 1; $i -le $m; $i++) { $q = $fv[$i, $fc]; if ($q -is [int]) { $k2 = "$($ws.Cells.Item(1,$fc).Text)"; $err[$k2] = 1 + [int]$err[$k2] } } }
        $sw.Close($false)
        $rng = if ($dayRows.Count) { "$([datetime]::FromOADate(($dayRows.Keys | Measure-Object -Minimum).Minimum).ToString('dd-MMM'))..$([datetime]::FromOADate(($dayRows.Keys | Measure-Object -Maximum).Maximum).ToString('dd-MMM'))" } else { '' }
        Log "$($s.Name): $m row(s) appended (rows $a-$b)$(if ($rng) { ", dates $rng" })$(if ($removed) { ", $removed old row(s) of the same days replaced" })$(if ($dropFuture) { ", $dropFuture row(s) after the report date left out" }); all cells verified"
        if ($err.Count) { Log "  note: formula errors in new rows: $(($err.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ') (e.g. an agent id missing in 'Team Detail')" }
        $s.NewLast = LastRow $ws $s.KeyCol
    }

    # ---------------------------------------------------------------- "Sales Performance": show the days up to the report date
    # Day columns H:AL (date in row 3) are shown up to the report date, later ones stay hidden; the week columns C:G
    # (Week-1 = days 1-7, Week-2 = 8-14, Week-3 = 15-21, Week-4 = 22-28, Week-5 = 29+) are shown once their first day is reached.
    $sp = $wb.Worksheets.Item('Sales Performance'); $ser = $D.ToOADate(); $shown = 0; $lastShown = $null
    for ($c = 8; $c -le 38; $c++) {
        $dv = $sp.Cells.Item(3, $c).Value2
        $show = ($dv -is [double]) -and $dv -le $ser
        $sp.Columns.Item($c).Hidden = -not $show
        if ($show) { $shown++; $lastShown = $c }
    }
    $firstDay = $sp.Cells.Item(3, 8).Value2
    for ($w = 0; $w -lt 5; $w++) {
        $startDay = if ($firstDay -is [double]) { $firstDay + 7 * $w } else { $null }
        $inMonth = ($null -ne $startDay) -and ([datetime]::FromOADate($startDay).Month -eq [datetime]::FromOADate($firstDay).Month)
        $sp.Columns.Item(3 + $w).Hidden = -not ($inMonth -and $startDay -le $ser)
    }
    Log "Sales Performance: $shown day column(s) shown, through $(if ($lastShown) { $sp.Cells.Item(3, $lastShown).Text + ' (column ' + ($sp.Cells.Item(1, $lastShown).Address($false, $false) -replace '\d', '') + ')' } else { '-' })"
    # ---------------------------------------------------------------- pivots on the three sheets
    $names = $sheets | ForEach-Object { $_.Name }; $okN = 0; $bad = @()
    foreach ($ws2 in $wb.Worksheets) {
        foreach ($pt in @($ws2.PivotTables())) {
            try {
                $sd = "$($pt.PivotCache().SourceData)"
                if ($sd -match "^'(.+)'!R1C(\d+):R(\d+)C(\d+)$" -and $names -contains $Matches[1]) {
                    $sh = $Matches[1]; $c1p = $Matches[2]; $r2 = [int]$Matches[3]; $c2p = $Matches[4]
                    $sk = $sheets | Where-Object { $_.Name -eq $sh }; $nl = LastRow $wb.Worksheets.Item($sh) $sk.KeyCol
                    if ($r2 -ne $nl) { $nsd = "'$sh'!R1C${c1p}:R${nl}C${c2p}"; try { $pt.SourceData = $nsd } catch { [void]$pt.ChangePivotCache($wb.PivotCaches().Create(1, $nsd)) } }
                }
                if ($sd -match "^'(.+)'!" -and $names -contains $Matches[1]) { [void]$pt.RefreshTable(); $okN++ }
            } catch { $bad += "$($ws2.Name)/$($pt.Name): $($_.Exception.Message)" }
        }
    }
    Log "Pivots: $okN refreshed$(if ($bad) { '; NOT refreshed: ' + ($bad -join '; ') })"
    $x.Calculation = -4105; $x.CalculateFull()
    $wb.Save(); $wb.Close($false); $wb = $null
    if ($failed.Count) { Log "DONE with problems - not updated: $($failed -join ', ')"; exit 1 }
    Log "DONE: $InboundPath"
}
catch {
    $ie = $_.Exception; $msgs = @(); while ($ie) { $msgs += $ie.Message; $ie = $ie.InnerException }
    Log "FAILED: $($msgs -join ' | ') (line $($_.InvocationInfo.ScriptLineNumber))"
    Log "Nothing was saved$(if ($bk) { "; backup: $bk" })."
    exit 1
}
finally {
    if ($x) { foreach ($w in @($x.Workbooks)) { try { $w.Close($false) } catch {} }; Stop-Excel $x $xpid }
    [GC]::Collect()
}
