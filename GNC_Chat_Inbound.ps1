# GNC Chat + Inbound sale sync (started by GNC_Chat_Inbound.bat).
# Source : "Date & Camp wise Overall Sale" in GNC_Overall_Sale_Performance_Dashboard <Mon'yy>.xlsb (opened read-only)
# Targets: GNC_Chat_Dashboard_<Mon'yy>.xlsb            sheet "Sale Raw"   <- rows with Campaign = Chat    (columns A:S)
#          GNC Inbound MIS Dashboard <Mon'yy>.xlsb     sheet "Sales Raw"  <- rows with Campaign = Inbound (columns A:N)
# Only the columns the target sheet has are pasted, as values (text stays text, dates stay dates), after the
# column headings were checked against the source. For every source row:
#   - same OrderID already in the target and all columns equal  -> left alone
#   - same OrderID, a column differs                            -> that row is corrected
#   - not found by OrderID, but same Date + EMP ID + CustomerNumber on a row whose OrderID is not a sale
#     (a wrong / placeholder OrderID)                           -> that row is corrected
#   - otherwise                                                 -> appended below the last row (format of the row above)
# After the paste every synced row is read back and compared with the source; if anything differs nothing is saved.
# APR: the day's GNC APR + PP APR (dialer files GNCAPR_*.csv / PPAPR_*.csv, found by the report date inside the file) are
# cleaned like GNC_Auto does and the rows of the Inbound agents (Inbound 'Team Detail') / Chat agents (Chat 'Agent Details')
# are appended to the sheet "APR Utilization Raw" of that dashboard (created in the Chat dashboard if missing). A re-run
# replaces that day's rows.
# Running it again changes nothing (no duplicates). Rows in the target that are not in the source are kept (reported).
param(
    [string]$Root = $PSScriptRoot,
    [datetime]$ReportDate = (Get-Date).Date.AddDays(-1),    # only picks the month folder, e.g. Oct'26
    [string]$DashPath, [string]$ChatPath, [string]$InboundPath,    # testing: use these files instead of the month folder
    [switch]$NoBackup,
    [string[]]$DialerDirs = @("D:\MIS\GNC\Sep'26\Auto", "D:\MIS\GNC\Oct'26\Auto"),
    [switch]$SkipApr                                         # sales only, no APR
)
$ErrorActionPreference = 'Stop'
$D = $ReportDate.Date
$MonAbbr = 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'
$tag = "$($MonAbbr[$D.Month - 1])'$($D.ToString('yy'))"
$monthDir = Join-Path $Root $tag
$logDir = Join-Path $Root 'Logs'; New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir 'gnc-chat-inbound-log.txt'
function Log($m) { $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $m"; Write-Host $line; Add-Content -LiteralPath $log -Value $line -Encoding utf8 }
function New-Excel { $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id); $xl = New-Object -ComObject Excel.Application; $p = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id) | Select-Object -First 1; , @($xl, $p) }
function Stop-Excel($xl, $xpid) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 800; if ($xpid) { Get-Process -Id $xpid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } }
function LastRow($ws, [int]$col) { $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }
# a cell value as a comparable string: numbers (dates too) in full precision, text trimmed
function Norm($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return 'n' + $v.ToString('R') }; if ($v -is [bool]) { return 'b' + $v }; if ($v -is [int]) { return 'n' + ([double]$v).ToString('R') }; 's' + "$v".Trim() }
function NormId($v) { if ($null -eq $v) { return '' }; if ($v -is [double]) { return $v.ToString('R') }; "$v".Trim() }
function RowKey($v, [int]$i) { (NormId ($v[$i, 2])) + '|' + (NormId ($v[$i, 3])) + '|' + (NormId ($v[$i, 8])) }     # Date | EMP ID | CustomerNumber
# copy the source rows (1-based sheet rows) to the target rows, values only; runs of consecutive rows in one go
function Paste-Rows($src, [int[]]$srcRows, $dst, [int[]]$dstRows, [int]$n) {
    $k = 0
    while ($k -lt $srcRows.Count) {
        $e = $k; while ($e + 1 -lt $srcRows.Count -and $srcRows[$e + 1] -eq $srcRows[$e] + 1 -and $dstRows[$e + 1] -eq $dstRows[$e] + 1) { $e++ }
        [void]$src.Range($src.Cells.Item($srcRows[$k], 1), $src.Cells.Item($srcRows[$e], $n)).Copy()
        [void]$dst.Cells.Item($dstRows[$k], 1).PasteSpecial(-4163)                           # xlPasteValues
        $k = $e + 1
    }
    try { $dst.Application.CutCopyMode = 0 } catch {}
}

# Excel takes a 2-D block only as a 1-based array when written through IDispatch
function To1($a) {
    if (-not ($a -is [array]) -or $a.Rank -ne 2 -or $a.GetLowerBound(0) -eq 1) { return , $a }
    $r = $a.GetLength(0); $c = $a.GetLength(1)
    $b = [Array]::CreateInstance([object], [int[]]($r, $c), [int[]](1, 1))
    for ($i = 0; $i -lt $r; $i++) { for ($j = 0; $j -lt $c; $j++) { $b[($i + 1), ($j + 1)] = $a[$i, $j] } }
    , $b
}
function Set-V2($rng, $val) { $val = To1 $val; try { [void]$rng.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $rng, @(, $val)) } catch { throw "writing $($rng.Worksheet.Name)!$($rng.Address(0,0)) failed: $($_.Exception.InnerException.Message)" } }
# ---- APR cleaning (same as GNC_Auto's GNCapr_BVA; USER / ID kept, the dialer's 'Teabr' / 'Tra' columns = Tea break / Training break)
$aprHeaders = 'USER', 'ID', 'CALLS', 'Login Time', 'WAIT', 'TALK', 'DISPO', 'PAUSE', 'Login', 'Logout', 'ACHT', 'AOC', 'BIO', 'Bre', 'Briefing', 'DOWN', 'Lunch', 'Meet', 'QA', 'SB', 'Tea break', 'Training break', 'Wash'
$aprAlias = @{ 'Tea break' = 'Teabr'; 'Training break' = 'Tra' }
function TimeVal([string]$s) {
    $s = $s.Trim(); if ($s -match '^(\d+):(\d{1,2}):(\d{1,2})$') { return ([int]$Matches[1] * 3600 + [int]$Matches[2] * 60 + [int]$Matches[3]) / 86400.0 }
    $n = 0.0; if ([double]::TryParse($s, [ref]$n)) { return $n }; $s
}
function Clean-Apr([string]$path) {
    $lines = @(Get-Content -LiteralPath $path | Select-Object -Skip 3)                     # first 3 rows
    while ($lines.Count -and -not $lines[-1].Trim()) { $lines = $lines[0..($lines.Count - 2)] }
    $lines = $lines[0..($lines.Count - 2)]                                                    # last (TOTALS) row
    $rows = @($lines | ConvertFrom-Csv -Header (1..200 | ForEach-Object { "c$_" }))
    $hdr = @(); foreach ($p in $rows[0].PSObject.Properties) { if ($null -ne $p.Value -and "$($p.Value)" -ne '') { $hdr += "$($p.Value)".Trim() } }
    $out = @()
    foreach ($r in $rows | Select-Object -Skip 1) {
        $vals = @($r.PSObject.Properties | ForEach-Object { "$($_.Value)" })
        $o = [ordered]@{}
        foreach ($h in $aprHeaders) {
            $ix = -1; for ($k = 0; $k -lt $hdr.Count; $k++) { if ($hdr[$k] -ieq $h) { $ix = $k; break } }
            if ($ix -lt 0 -and $aprAlias.ContainsKey($h)) { for ($k = 0; $k -lt $hdr.Count; $k++) { if ($hdr[$k] -ieq $aprAlias[$h]) { $ix = $k; break } } }
            $o[$h] = if ($ix -ge 0) { TimeVal $vals[$ix] } else { 0.0 }
        }
        $out += [pscustomobject]$o
    }
    $out
}
# the 33 columns of "APR Utilization Raw" (UID .. Capping) for the agents in $team (ID -> @(TL, LOB)), as in the sale dashboard's sheet
$aprCols = 'UID','Date','USER','ID','TL Name','CALLS','Process Type','Login Time','WAIT','TALK','DISPO','PAUSE','Login','Logout','ACHT','AOC','BIO','Bre','Briefing','DOWN','Lunch','Meet','QA','SB','Tea break','Training break','Wash','Net Login','Break','TRA+QA','Downtime','Atten','Capping'
$aprFmt = @('General','dd-mmm-yy','General','General','#,##0','#,##0','#,##0'); $aprFmt += @('hh:mm:ss') * 7; $aprFmt += '0'; $aprFmt += @('hh:mm:ss') * 16; $aprFmt += 'General'; $aprFmt += 'hh:mm:ss'   # 33 formats
function Build-Apr($rows, $team, [double]$ser) {
    $keep = @($rows | Where-Object { $team.ContainsKey("$($_.ID)".Trim()) })
    $m = New-Object 'object[,]' $keep.Count, 33
    for ($k = 0; $k -lt $keep.Count; $k++) {
        $r = $keep[$k]; $id = "$($r.ID)".Trim(); $tm = $team[$id]
        $m[$k, 0] = "$id$([int]$ser)"; $m[$k, 1] = $ser; $m[$k, 2] = "$($r.USER)".Trim(); $m[$k, 3] = $id; $m[$k, 4] = $tm[0]
        $m[$k, 5] = $r.CALLS; $m[$k, 6] = $tm[1]
        $j = 7; foreach ($h in 'Login Time', 'WAIT', 'TALK', 'DISPO', 'PAUSE', 'Login', 'Logout', 'ACHT', 'AOC', 'BIO', 'Bre', 'Briefing', 'DOWN', 'Lunch', 'Meet', 'QA', 'SB', 'Tea break', 'Training break', 'Wash') { $m[$k, $j] = $r.$h; $j++ }
        $brk = [double]$r.BIO + [double]$r.Lunch + [double]$r.'Tea break' + [double]$r.SB + [double]$r.Wash
        $tra = [double]$r.Briefing + [double]$r.Meet + [double]$r.QA + [double]$r.'Training break'
        $net = ([double]$r.'Login Time' - $brk) + $tra + 0.0
        $m[$k, 27] = $net; $m[$k, 28] = $brk; $m[$k, 29] = $tra; $m[$k, 30] = 0.0
        $m[$k, 31] = $(if ($net -ge 0.322916666666667) { 1.0 } elseif ($net -ge 0.159722222222222) { 0.5 } else { 0.0 })
        $m[$k, 32] = $(if ($net -gt 8 / 24) { 8 / 24 } else { $net })
    }
    , $m
}

Log "==== GNC Chat + Inbound sale sync, month $tag ===="
if (-not $DashPath)    { $DashPath    = Join-Path $monthDir "GNC_Overall_Sale_Performance_Dashboard $tag.xlsb" }
if (-not $ChatPath)    { $ChatPath    = Join-Path $monthDir "GNC_Chat_Dashboard_$tag.xlsb" }
if (-not $InboundPath) { $InboundPath = Join-Path $monthDir "GNC Inbound MIS Dashboard $tag.xlsb" }
foreach ($f in @($DashPath, $ChatPath, $InboundPath)) { if (-not (Test-Path -LiteralPath $f)) { Log "STOP: file not found: $f"; exit 1 } }
foreach ($f in @($ChatPath, $InboundPath)) {
    if (Test-Path -LiteralPath (Join-Path (Split-Path $f) ('~$' + (Split-Path $f -Leaf)))) { Log "STOP: $(Split-Path $f -Leaf) is open in Excel. Save and close it, then run again. Nothing was changed."; exit 1 }
}
if (Test-Path -LiteralPath (Join-Path (Split-Path $DashPath) ('~$' + (Split-Path $DashPath -Leaf)))) { Log "WARNING: $(Split-Path $DashPath -Leaf) is open in Excel - the saved file is read, unsaved changes are not seen." }

$aprRows = @()
if (-not $SkipApr) {
    $dstr = $D.ToString('yyyy-MM-dd')
    $dirs = @($DialerDirs + (Join-Path $monthDir 'Auto')) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique
    $csvs = @(Get-ChildItem -LiteralPath $dirs -File -Filter '*.csv' | Sort-Object LastWriteTime -Descending)
    $aprGnc = $csvs | Where-Object { $_.Name -like 'GNCAPR*' -and (Get-Content -LiteralPath $_.FullName -TotalCount 2)[1] -like "*$dstr 00:00:00*" } | Select-Object -First 1
    $aprPP  = $csvs | Where-Object { $_.Name -like 'PPAPR*'  -and (Get-Content -LiteralPath $_.FullName -TotalCount 2)[1] -like "*$dstr 00:00:00*" } | Select-Object -First 1
    foreach ($p in @(@('GNC APR', $aprGnc), @('PP APR', $aprPP))) {
        if (-not $p[1]) { Log "STOP: no $($p[0]) file for $($D.ToString('dd-MMM')) in $($dirs -join ' / ') (morning download not run yet? use -SkipApr for sales only). Nothing was changed."; exit 1 }
        Log "  $($p[0]): $($p[1].Name)"
    }
    $g = @(Clean-Apr $aprGnc.FullName); $pp = @(Clean-Apr $aprPP.FullName); $aprRows = $g + $pp
    Log "APR files read: $($g.Count) GNC + $($pp.Count) PP rows of $($D.ToString('dd-MMM'))"
}

$targets = @(
    @{ Label = 'Chat';    Path = $ChatPath;    Sheet = 'Sale Raw';  Campaign = 'Chat';    Cols = 19; Prefix = 'Chat Dashboard'; TeamSheet = 'Agent Details'; TeamId = 3; TeamLob = 5; TeamTl = 11 },
    @{ Label = 'Inbound'; Path = $InboundPath; Sheet = 'Sales Raw'; Campaign = 'Inbound'; Cols = 14; Prefix = 'Inbound Dashboard'; TeamSheet = 'Team Detail'; TeamId = 2; TeamLob = 4; TeamTl = 11 }
)
if (-not $NoBackup) {
    $bk = Join-Path $Root 'Backups'; New-Item -ItemType Directory -Force -Path $bk | Out-Null
    $stamp = Get-Date -Format 'yyyy-MM-dd HHmm'
    foreach ($t in $targets) {
        Copy-Item -LiteralPath $t.Path -Destination (Join-Path $bk "$($t.Prefix) before $stamp.xlsb")
        Get-ChildItem -LiteralPath $bk -Filter "$($t.Prefix) before *.xlsb" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 | Remove-Item -Force
    }
}

$x = $null; $open = @(); $skipped = @()
try {
    $nx = New-Excel; $x = $nx[0]; $xpid = $nx[1]
    $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false; $x.EnableEvents = $false; $x.AutomationSecurity = 3

    $dw = $x.Workbooks.Open($DashPath, 0, $true)                                              # read-only: the dashboard is never changed
    $ds = $dw.Worksheets.Item('Date & Camp wise Overall Sale'); if ($ds.FilterMode) { $ds.ShowAllData() }
    $dl = LastRow $ds 13                                                                      # OrderID column
    $sv = $ds.Range("A1:S$dl").Value2
    Log "Dashboard: 'Date & Camp wise Overall Sale' rows 2-$dl read"

    foreach ($t in $targets) {
        $lab = $t.Label; $n = $t.Cols
        try { $tw = $x.Workbooks.Open($t.Path) }
        catch { Log "SKIPPED ${lab}: '$(Split-Path $t.Path -Leaf)' cannot be opened by Excel automation. Open it once in Excel, File > Save As (same name, .xlsb), close it and run again."; $skipped += $lab; continue }
        $open += $tw
        $ts = $tw.Worksheets.Item($t.Sheet); if ($ts.FilterMode) { $ts.ShowAllData() }
        $x.Calculation = -4135

        # headings: every pasted column must have the same heading in source and target
        $th = $ts.Range($ts.Cells.Item(1, 1), $ts.Cells.Item(1, $n)).Value2
        $bad = @(); for ($c = 1; $c -le $n; $c++) { if ("$($sv[1,$c])".Trim().ToLower() -ne "$($th[1,$c])".Trim().ToLower()) { $bad += "col $c '$($sv[1,$c])' vs '$($th[1,$c])'" } }
        if ($bad) { throw "$lab '$($t.Sheet)': column headings differ from the dashboard sheet ($($bad -join '; ')) - nothing pasted" }

        # source rows of this campaign
        $src = @(); for ($i = 2; $i -le $dl; $i++) { if ("$($sv[$i,14])".Trim() -ieq $t.Campaign) { $src += $i } }

        # target rows
        $tl = LastRow $ts 2; if ($tl -lt 1) { $tl = 1 }
        $tv = $null; if ($tl -ge 2) { $tv = $ts.Range($ts.Cells.Item(1, 1), $ts.Cells.Item($tl, $n)).Value2 }       # (not 'if' as an expression: it flattens the 2-D array)
        $byOrder = @{}; $byKey = @{}; $used = @{}
        for ($r = 2; $r -le $tl; $r++) {
            $o = NormId ($tv[$r, 13]); if ($o) { if (-not $byOrder.ContainsKey($o)) { $byOrder[$o] = New-Object System.Collections.Generic.List[int] }; $byOrder[$o].Add($r) }
            $k = RowKey $tv $r; if (-not $byKey.ContainsKey($k)) { $byKey[$k] = New-Object System.Collections.Generic.List[int] }; $byKey[$k].Add($r)
        }
        $srcOrders = @{}; foreach ($i in $src) { $srcOrders[(NormId ($sv[$i, 13]))] = $true }
        $plan = @{}                                                                           # source row -> target row (0 = append)
        foreach ($i in $src) {                                                                # pass 1: same OrderID
            $o = NormId ($sv[$i, 13])
            if ($o -and $byOrder.ContainsKey($o)) { foreach ($r in $byOrder[$o]) { if (-not $used.ContainsKey($r)) { $used[$r] = $true; $plan[$i] = $r; break } } }
        }
        foreach ($i in $src) {                                                                # pass 2: same Date+EMP+Customer on a row with a non-sale OrderID
            if ($plan.ContainsKey($i)) { continue }
            $k = RowKey $sv $i
            if ($byKey.ContainsKey($k)) { foreach ($r in $byKey[$k]) { if (-not $used.ContainsKey($r) -and -not $srcOrders.ContainsKey((NormId ($tv[$r, 13])))) { $used[$r] = $true; $plan[$i] = $r; break } } }
        }
        $fixSrc = @(); $fixDst = @(); $addSrc = @(); $same = 0; $fixLog = @()
        foreach ($i in $src) {
            if ($plan.ContainsKey($i)) {
                $r = $plan[$i]; $diff = @()
                for ($c = 1; $c -le $n; $c++) { if ((Norm ($sv[$i, $c])) -ne (Norm ($tv[$r, $c]))) { $diff += $sv[1, $c] } }
                if ($diff) { $fixSrc += $i; $fixDst += $r; $fixLog += "row $r ($($diff -join ', '))" } else { $same++ }
            } else { $addSrc += $i }
        }
        $extra = @(); for ($r = 2; $r -le $tl; $r++) { if (-not $used.ContainsKey($r)) { $extra += $r } }

        if ($fixSrc.Count) { Paste-Rows $ds $fixSrc $ts $fixDst $n }
        $addDst = @(); $a = $tl + 1
        if ($addSrc.Count) {
            for ($k = 0; $k -lt $addSrc.Count; $k++) { $addDst += ($a + $k) }
            if ($tl -ge 2) {                                                                  # new rows look like the row above (number formats, font, fill, borders)
                [void]$ts.Range($ts.Cells.Item($tl, 1), $ts.Cells.Item($tl, $n)).Copy()
                [void]$ts.Range($ts.Cells.Item($a, 1), $ts.Cells.Item($a + $addSrc.Count - 1, $n)).PasteSpecial(-4122)
                try { $x.CutCopyMode = 0 } catch {}
            }
            Paste-Rows $ds $addSrc $ts $addDst $n
        }

        # read back every synced row and compare with the source, cell by cell
        $dstOf = @{}; for ($q = 0; $q -lt $fixSrc.Count; $q++) { $dstOf[$fixSrc[$q]] = $fixDst[$q] }
        for ($q = 0; $q -lt $addSrc.Count; $q++) { $dstOf[$addSrc[$q]] = $addDst[$q] }
        foreach ($i in $src) { if (-not $dstOf.ContainsKey($i)) { $dstOf[$i] = $plan[$i] } }
        $nl = LastRow $ts 2
        $nv = $ts.Range($ts.Cells.Item(1, 1), $ts.Cells.Item($nl, $n)).Value2
        $mis = 0
        foreach ($i in $src) { $r = $dstOf[$i]; for ($c = 1; $c -le $n; $c++) { if ((Norm ($sv[$i, $c])) -ne (Norm ($nv[$r, $c]))) { $mis++; Log "  MISMATCH $lab row $r col $c : '$($sv[$i,$c])' vs '$($nv[$r,$c])'" } } }
        if ($mis) { throw "${lab}: $mis cell(s) differ after the paste - nothing saved" }
        Log "${lab}: $($src.Count) '$($t.Campaign)' row(s) in the dashboard -> '$($t.Sheet)': $same already correct, $($fixSrc.Count) corrected, $($addSrc.Count) appended (rows $(if ($addSrc.Count) { "$($addDst[0])-$($addDst[-1])" } else { '-' })); sheet now rows 2-$nl, all cells verified"
        foreach ($f in $fixLog) { Log "  corrected $f" }
        if ($extra.Count) { Log "  note: $($extra.Count) row(s) of '$($t.Sheet)' are not in the dashboard (kept): rows $($extra -join ', ')" }

        # ---- APR Utilization Raw: this day's rows of the agents of this dashboard
        if (-not $SkipApr) {
            $tms = $tw.Worksheets.Item($t.TeamSheet); $tml = LastRow $tms $t.TeamId
            $tmv = $tms.Range($tms.Cells.Item(1, 1), $tms.Cells.Item($tml, 11)).Value2
            $team = @{}; for ($i = 2; $i -le $tml; $i++) { $id = "$($tmv[$i,$t.TeamId])".Trim(); if ($id -and -not $team.ContainsKey($id)) { $team[$id] = @("$($tmv[$i,$t.TeamTl])".Trim(), "$($tmv[$i,$t.TeamLob])".Trim()) } }
            $ser = $D.ToOADate()
            $am = Build-Apr $aprRows $team $ser; $m = $am.GetLength(0)
            $newApr = $false; $aps = $null; foreach ($w2 in $tw.Worksheets) { if ($w2.Name -eq 'APR Utilization Raw') { $aps = $w2 } }
            if (-not $aps) {
                $aps = $tw.Worksheets.Add([Type]::Missing, $tw.Worksheets.Item($tw.Worksheets.Count)); $aps.Name = 'APR Utilization Raw'
                $dapr = $dw.Worksheets.Item('APR-Utilization Raw'); [void]$dapr.Range("A1:AG1").Copy(); [void]$aps.Range("A1").PasteSpecial(-4104)
                try { $x.CutCopyMode = 0 } catch {}; $newApr = $true
                $ts.Activate(); Log "${lab}: sheet 'APR Utilization Raw' created (headings of the sale dashboard's APR sheet)"
            }
            if ($aps.FilterMode) { $aps.ShowAllData() }
            $ah = $aps.Range("A1:AG1").Value2; $bad = @(); for ($c = 1; $c -le 33; $c++) { if ("$($ah[1,$c])".Trim().ToLower() -ne $aprCols[$c - 1].ToLower()) { $bad += "col $c '$($ah[1,$c])'" } }
            if ($bad) { throw "${lab} 'APR Utilization Raw': headings differ from expected ($($bad -join '; '))" }
            $al = LastRow $aps 2; if ($al -lt 1) { $al = 1 }; $gone = 0
            if ($al -ge 2) { $bv = $aps.Range("B1:B$al").Value2; for ($i = $al; $i -ge 2; $i--) { if ($bv[$i, 1] -is [double] -and $bv[$i, 1] -eq $ser) { [void]$aps.Rows.Item($i).Delete(); $gone++ } }; $al = LastRow $aps 2; if ($al -lt 1) { $al = 1 } }
            if ($m -gt 0) {
                $a = $al + 1; $b = $al + $m
                if ($al -ge 2) { [void]$aps.Range($aps.Cells.Item($al, 1), $aps.Cells.Item($al, 33)).Copy(); [void]$aps.Range($aps.Cells.Item($a, 1), $aps.Cells.Item($b, 33)).PasteSpecial(-4122); try { $x.CutCopyMode = 0 } catch {} }
                for ($c = 1; $c -le 33; $c++) { $aps.Range($aps.Cells.Item($a, $c), $aps.Cells.Item($b, $c)).NumberFormat = $aprFmt[$c - 1] }
                Set-V2 $aps.Range($aps.Cells.Item($a, 1), $aps.Cells.Item($b, 33)) $am
                if ($newApr) { [void]$aps.Columns.Item("A:AG").AutoFit() }
                $rb = $aps.Range($aps.Cells.Item($a, 1), $aps.Cells.Item($b, 33)).Value2; $mis = 0
                for ($i = 0; $i -lt $m; $i++) { for ($c = 0; $c -lt 33; $c++) { if ((Norm ($am[$i, $c])) -ne (Norm ($rb[($i + 1), ($c + 1)]))) { $mis++; Log "  MISMATCH ${lab} APR row $($a + $i) col $($c + 1): '$($am[$i,$c])' vs '$($rb[($i + 1),($c + 1)])'" } } }
                if ($mis) { throw "${lab}: $mis APR cell(s) differ after the paste - nothing saved" }
                $names = (@(for ($i = 0; $i -lt $m; $i++) { $am[$i, 2] })) -join ', '
                Log "${lab}: APR Utilization Raw: $m row(s) of $($D.ToString('dd-MMM')) appended (rows $a-$b)$(if ($gone) { ", $gone earlier row(s) of that day replaced" }): $names"
            } else { Log "${lab}: APR: none of the $($team.Count) agents of '$($t.TeamSheet)' is in the GNC / PP APR of $($D.ToString('dd-MMM')) - nothing appended$(if ($gone) { " ($gone earlier row(s) of that day removed)" })" }
        }
    }

    $dw.Close($false)
    $x.Calculation = -4105
    foreach ($tw in $open) { $x.CalculateFull(); $tw.Save(); Log "Saved: $($tw.Name)" }
    foreach ($tw in $open) { $tw.Close($false) }; $open = @()
    if ($skipped.Count) { Log "DONE with problems - not updated: $($skipped -join ', ')"; exit 1 }
    Log "DONE"
}
catch {
    $ie = $_.Exception; $msgs = @(); while ($ie) { $msgs += $ie.Message; $ie = $ie.InnerException }
    Log "FAILED: $($msgs -join ' | ') (line $($_.InvocationInfo.ScriptLineNumber))"
    Log "Nothing was saved$(if (-not $NoBackup) { "; backups are in $bk" })."
    exit 1
}
finally {
    if ($x) {
        foreach ($w in @($x.Workbooks)) { try { $w.Close($false) } catch {} }
        Stop-Excel $x $xpid
    }
    [GC]::Collect()
}
