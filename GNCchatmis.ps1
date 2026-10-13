# GNC Chat dashboard - daily chat data update (started by GNCchatmis.bat).
# Report date = yesterday (or -ReportDate). Workbook: <Root>\<Mon'yy>\GNC_Chat_Dashboard_<Mon'yy>.xlsb
#   1. fetch_chat.py: KwikEngage > Analytics > Export Reports (Ticket, date range) > link in the mail > csv download.
#      The range starts at the last chat date already in the sheet (so missed days are caught up), ends at the report date.
#   2. sheet "Chat": the csv's columns are matched to the sheet's headings (A:BY = 77 columns; extra csv columns are left out)
#        - TicketId already in the sheet and unchanged -> left alone
#        - TicketId in the sheet but changed (status, times ...) -> the old row is deleted, the new one appended below
#        - new TicketId                   -> appended below (oldest first), formulas BZ:CE dragged, format of the row above
#   3. "Snap": day columns H:AL shown up to the report date (later ones hidden), week columns C:G shown once the week has started
#   4. everything recalculated, saved
# Every pasted cell is read back and compared; if anything differs nothing is saved. Running it again changes nothing.
param(
    [string]$Root = $PSScriptRoot,
    [datetime]$ReportDate = (Get-Date).Date.AddDays(-1),
    [string]$ChatPath,                    # testing: use this workbook
    [string]$DataDir,                     # testing: use the csv in this folder (no download)
    [switch]$SkipDownload,
    [int]$WaitMinutes = 0,                # wait this long for the workbook to be closed (the scheduled run uses 30)
    [int]$WindowDays = 10,                # export this many days back from the report date (the export is by ticket creation day, queued tickets are assigned days later)
    [switch]$NoBackup,
    [string]$Python
)
$ErrorActionPreference = 'Stop'
$D = $ReportDate.Date
$MonAbbr = 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'
$tag = "$($MonAbbr[$D.Month - 1])'$($D.ToString('yy'))"
$monthDir = Join-Path $Root $tag
$logDir = Join-Path $Root 'Logs'; New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir 'gnc-chatmis-log.txt'
function Log($m) { $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $m"; Write-Host $line; Add-Content -LiteralPath $log -Value $line -Encoding utf8 }
function New-Excel { $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id); $xl = New-Object -ComObject Excel.Application; $p = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id) | Select-Object -First 1; , @($xl, $p) }
function Stop-Excel($xl, $xpid) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 800; if ($xpid) { Get-Process -Id $xpid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } }
function LastRow($ws, [int]$col) { $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }
function Norm($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return 'n' + $v.ToString('R') }; if ($v -is [bool]) { return 'b' + $v }; if ($v -is [int]) { return 'n' + ([double]$v).ToString('R') }; 's' + "$v".Trim() }
function NormId($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return $v.ToString('R') }; "$v".Trim() }
function NormH([string]$h) { ("$h" -replace '\s+', ' ').Trim().ToLower() }
function DayOf($v) {
    if ($v -is [double]) { return [Math]::Floor($v) }
    $s = "$v".Trim(); if (-not $s) { return $null }
    $dt = [datetime]::MinValue
    foreach ($f in 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd', 'dd-MM-yyyy HH:mm:ss', 'dd-MM-yyyy') { if ([datetime]::TryParseExact($s, $f, [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$dt)) { return [Math]::Floor($dt.ToOADate()) } }
    $null
}
function Test-InUse([string]$f) { try { $fs = [IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $fs.Close(); $false } catch { $true } }
$NCOLS = 77; $FORMULA = 78..83; $LASTCOL = 83

Log "==== GNC Chat dashboard update, report date $($D.ToString('dd-MMM-yyyy')) ===="
if (-not $ChatPath) { $ChatPath = Join-Path $monthDir "GNC_Chat_Dashboard_$tag.xlsb" }
if (-not (Test-Path -LiteralPath $ChatPath)) { Log "STOP: workbook not found: $ChatPath"; exit 1 }
$waited = 0
while (Test-InUse $ChatPath) {
    if ($waited -ge $WaitMinutes) { Log "STOP: $(Split-Path $ChatPath -Leaf) is open in Excel. Save and close it (File > Close), then run again. Nothing was changed."; exit 1 }
    if ($waited -eq 0) { Log "$(Split-Path $ChatPath -Leaf) is open in Excel - waiting up to $WaitMinutes min for it to be closed..." }
    Start-Sleep -Seconds 60; $waited++
}
if (-not $DataDir) { $DataDir = Join-Path $monthDir ("Data\Chat\" + $D.ToString('yyyy-MM-dd')) }

$x = $null; $wb = $null
try {
    $nx = New-Excel; $x = $nx[0]; $xpid = $nx[1]
    $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false; $x.EnableEvents = $false; $x.AutomationSecurity = 3

    # ---------------------------------------------------------------- 1. last chat date -> range
    try { $wb = $x.Workbooks.Open($ChatPath, 0, $true) } catch { throw "'$(Split-Path $ChatPath -Leaf)' cannot be opened by Excel automation. Open it once in Excel, File > Save As (same name, .xlsb), close it and run again." }
    $ws = $wb.Worksheets.Item('Chat'); $n = LastRow $ws 1
    $last = $null
    if ($n -ge 2) { $mx = 0; foreach ($e in $ws.Range("J2:J$n").Value2) { if ($e -is [double] -and [Math]::Floor($e) -gt $mx) { $mx = [Math]::Floor($e) } }; if ($mx -gt 0) { $last = $mx } }
    $dFrom = $D.AddDays(-($WindowDays - 1))
    Log "Chat: rows 2-$n, last chat date $(if ($last) { [datetime]::FromOADate($last).ToString('dd-MMM') } else { '-' }) -> exporting the last $WindowDays days, from $($dFrom.ToString('dd-MMM'))"
    $wb.Close($false); $wb = $null

    # ---------------------------------------------------------------- 2. download
    if (-not $SkipDownload -and -not $PSBoundParameters.ContainsKey('DataDir')) {
        if (-not $Python) { foreach ($c in 'C:\Python314\python.exe', "$env:LOCALAPPDATA\Microsoft\WindowsApps\python.exe", 'python.exe') { if (Get-Command $c -ErrorAction SilentlyContinue) { $Python = $c; break } } }
        New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
        Log "Exporting chat tickets $($dFrom.ToString('yyyy-MM-dd')) .. $($D.ToString('yyyy-MM-dd')) (KwikEngage, then the mail link - can take a few minutes) ..."
        $ErrorActionPreference = 'Continue'
        $out = & $Python (Join-Path $PSScriptRoot 'fetch_chat.py') --from $dFrom.ToString('yyyy-MM-dd') --to $D.ToString('yyyy-MM-dd') --out $DataDir 2>&1 | ForEach-Object { "$_" }
        $ErrorActionPreference = 'Stop'
        $out | ForEach-Object { Log "  $_" }
    }
    $csv = $null; $rf = Join-Path $DataDir '_result.json'
    if (Test-Path -LiteralPath $rf) { $j = Get-Content -LiteralPath $rf -Raw | ConvertFrom-Json; if ($j.Chat.ok -and (Test-Path -LiteralPath $j.Chat.file)) { $csv = $j.Chat.file } }
    if (-not $csv -and ($SkipDownload -or $PSBoundParameters.ContainsKey('DataDir'))) { $csv = Get-ChildItem -LiteralPath $DataDir -File -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName }
    if (-not $csv) { Log "STOP: no chat csv downloaded. Nothing was changed."; exit 1 }

    # ---------------------------------------------------------------- backup + paste
    if (-not $NoBackup) {
        $bk = Join-Path $Root 'Backups'; New-Item -ItemType Directory -Force -Path $bk | Out-Null
        Copy-Item -LiteralPath $ChatPath -Destination (Join-Path $bk "Chat MIS before $(Get-Date -Format 'yyyy-MM-dd HHmm').xlsb")
        Get-ChildItem -LiteralPath $bk -Filter 'Chat MIS before *.xlsb' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 | Remove-Item -Force
    }
    $wb = $x.Workbooks.Open($ChatPath); $x.Calculation = -4135
    $ws = $wb.Worksheets.Item('Chat'); if ($ws.FilterMode) { $ws.ShowAllData() }
    $th = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $NCOLS)).Value2

    # csv opened as UTF-8 text (like opening it in Excel: dates, numbers, TRUE/FALSE are converted)
    $x.Workbooks.OpenText($csv, 65001, 1, 1, 1, $false, $false, $false, $true, $false) | Out-Null
    $sw = $x.ActiveWorkbook; $ss = $sw.Worksheets.Item(1)
    $sl = LastRow $ss 1; $scn = $ss.Cells.Item(1, $ss.Columns.Count).End(-4159).Column
    $hdr = $ss.Range($ss.Cells.Item(1, 1), $ss.Cells.Item(1, $scn)).Value2
    $map = @(); $miss = @()
    for ($j = 1; $j -le $NCOLS; $j++) { $m = 0; for ($c = 1; $c -le $scn; $c++) { if ((NormH $hdr[1, $c]) -eq (NormH $th[1, $j])) { $m = $c; break } }; $map += $m; if (-not $m) { $miss += $th[1, $j] } }
    if ($miss.Count) { throw "Chat: column(s) of the sheet missing in $(Split-Path $csv -Leaf): $($miss -join ', ')" }
    $identity = $true; for ($j = 0; $j -lt $NCOLS; $j++) { if ($map[$j] -ne $j + 1) { $identity = $false } }
    $dcol = 0; for ($j = 1; $j -le $NCOLS; $j++) { if ((NormH $th[1, $j]) -eq 'firstassignedat') { $dcol = $map[$j - 1] } }
    if ($sl -lt 2) { Log "Chat: the csv has no rows"; $sw.Close($false) }
    else {
        if ($dcol) { [void]$ss.Range($ss.Cells.Item(2, 1), $ss.Cells.Item($sl, $scn)).Sort($ss.Cells.Item(2, $dcol), 1, [Type]::Missing, [Type]::Missing, 1, [Type]::Missing, 1, 2) }   # oldest first
        $sv = $ss.Range($ss.Cells.Item(1, 1), $ss.Cells.Item($sl, $scn)).Value2
        $tn = LastRow $ws 1
        $tv = $null; if ($tn -ge 2) { $tv = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($tn, $NCOLS)).Value2 }
        $rowOf = @{}; for ($r = 2; $r -le $tn; $r++) { $id = NormId ($tv[$r, 1]); if ($id -and -not $rowOf.ContainsKey($id)) { $rowOf[$id] = $r } }
        $monthStart = (Get-Date -Year $D.Year -Month $D.Month -Day 1).Date.ToOADate()
        $delRows = @(); $addS = @(); $same = 0; $seen = @{}; $before = 0
        for ($r = 2; $r -le $sl; $r++) {
            $id = NormId ($sv[$r, $map[0]]); if (-not $id) { continue }
            if ($seen.ContainsKey($id)) { continue }; $seen[$id] = 1
            if ($dcol) { $dy = DayOf ($sv[$r, $dcol]); if ($null -ne $dy -and $dy -lt $monthStart) { $before++; continue } }
            if ($rowOf.ContainsKey($id)) {
                $tr = $rowOf[$id]; $diff = $false
                for ($j = 0; $j -lt $NCOLS; $j++) { if ((Norm ($sv[$r, $map[$j]])) -ne (Norm ($tv[$tr, ($j + 1)]))) { $diff = $true; break } }
                if ($diff) { $delRows += $tr; $addS += $r } else { $same++ }
            } else { $addS += $r }
        }
        # changed tickets: old row out (runs of consecutive rows, from the bottom up), the new version goes below with the new ones
        $delRows = @($delRows | Sort-Object -Descending); $k = 0
        while ($k -lt $delRows.Count) { $e = $k; while ($e + 1 -lt $delRows.Count -and $delRows[$e + 1] -eq $delRows[$e] - 1) { $e++ }; [void]$ws.Range($ws.Rows.Item($delRows[$e]), $ws.Rows.Item($delRows[$k])).Delete(); $k = $e + 1 }
        if ($delRows.Count) { $tn = LastRow $ws 1 }
        function Paste-Block([int]$s1, [int]$s2, [int]$d1) {
            if ($identity) { [void]$ss.Range($ss.Cells.Item($s1, 1), $ss.Cells.Item($s2, $NCOLS)).Copy(); [void]$ws.Cells.Item($d1, 1).PasteSpecial(-4163) }
            else { for ($j = 0; $j -lt $NCOLS; $j++) { [void]$ss.Range($ss.Cells.Item($s1, $map[$j]), $ss.Cells.Item($s2, $map[$j])).Copy(); [void]$ws.Cells.Item($d1, $j + 1).PasteSpecial(-4163) } }
        }
        $m = $addS.Count; $a = $tn + 1; $b = $tn + $m
        if ($m) {
            if ($tn -lt 2) { throw "Chat: no data row to copy the formulas and formats from" }
            [void]$ws.Range($ws.Cells.Item($tn, 1), $ws.Cells.Item($tn, $LASTCOL)).Copy()
            [void]$ws.Range($ws.Cells.Item($a, 1), $ws.Cells.Item($b, $LASTCOL)).PasteSpecial(-4122)
            foreach ($fc in $FORMULA) { $ws.Range($ws.Cells.Item($a, $fc), $ws.Cells.Item($b, $fc)).Formula2R1C1 = $ws.Cells.Item($tn, $fc).Formula2R1C1 }
            $k = 0
            while ($k -lt $m) { $e = $k; while ($e + 1 -lt $m -and $addS[$e + 1] -eq $addS[$e] + 1) { $e++ }; Paste-Block $addS[$k] $addS[$e] ($a + $k); $k = $e + 1 }
        }
        try { $x.CutCopyMode = 0 } catch {}
        $x.Calculate()
        # verify every touched row
        $dstRows = @{}; for ($i = 0; $i -lt $m; $i++) { $dstRows[$addS[$i]] = $a + $i }
        $nl = LastRow $ws 1; $nv = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nl, $LASTCOL)).Value2; $mis = 0
        foreach ($sr in $dstRows.Keys) { $dr = $dstRows[$sr]; for ($j = 0; $j -lt $NCOLS; $j++) { if ((Norm ($sv[$sr, $map[$j]])) -ne (Norm ($nv[$dr, ($j + 1)]))) { $mis++; if ($mis -le 5) { Log "  MISMATCH row $dr col $($j + 1): '$($sv[$sr,$map[$j]])' vs '$($nv[$dr,($j + 1)])'" } } } }
        if ($mis) { throw "Chat: $mis pasted cell(s) differ from the csv - nothing saved" }
        $err = 0; for ($i = 0; $i -lt $m; $i++) { foreach ($fc in $FORMULA) { if ($nv[($a + $i), $fc] -is [int]) { $err++ } } }
        $sw.Close($false)
        Log "Chat: csv $($seen.Count) ticket(s): $same already correct, $($delRows.Count) changed (old row removed, new one appended), $($m - $delRows.Count) new, $m appended in all$(if ($m) { " (rows $a-$b)" }); all cells verified$(if ($err) { "; $err formula error cell(s) in new rows" })$(if ($before) { "; $before ticket(s) assigned before the month left out" })"
    }

    # ---------------------------------------------------------------- Snap: show the days up to the report date
    $sp = $wb.Worksheets.Item('Snap'); $ser = $D.ToOADate(); $shown = 0; $lastShown = $null
    for ($c = 8; $c -le 38; $c++) { $dv = $sp.Cells.Item(2, $c).Value2; $show = ($dv -is [double]) -and $dv -le $ser; $sp.Columns.Item($c).Hidden = -not $show; if ($show) { $shown++; $lastShown = $c } }
    $firstDay = $sp.Cells.Item(2, 8).Value2
    for ($w = 0; $w -lt 5; $w++) {
        $startDay = if ($firstDay -is [double]) { $firstDay + 7 * $w } else { $null }
        $inMonth = ($null -ne $startDay) -and ([datetime]::FromOADate($startDay).Month -eq [datetime]::FromOADate($firstDay).Month)
        $sp.Columns.Item(3 + $w).Hidden = -not ($inMonth -and $startDay -le $ser)
    }
    Log "Snap: $shown day column(s) shown, through $(if ($lastShown) { $sp.Cells.Item(2, $lastShown).Text } else { '-' })"

    $x.Calculation = -4105; $x.CalculateFull()
    $wb.Save(); $wb.Close($false); $wb = $null
    Log "DONE: $ChatPath"
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
