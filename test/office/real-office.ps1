# Real-Office acceptance suite for the dav-service POC (Windows + desktop Office).
# Drives Word/Excel/PowerPoint through COM against a running demo-site + dav-service
# and checks the outcome on the server side (versions, locks), since Office's COM
# return values can lie (e.g. Save As "succeeds" when the server refused the PUT).
#
#   powershell -File test/office/real-office.ps1 -Demo http://localhost:3443 -Jars <dir with alice.jar, bob.jar>
#   -Phase main     : open/edit/save per app, second user, crash, Save As, force unlock, revoke, web upload
#   -Phase expiry   : run with LOCK_MAX_TIMEOUT_S=60 on both services; idle > timeout then save (refresh/revival)
#   -Phase big      : 100 MB xlsx open/edit/save
#
# Every Office step runs in a separate job with a timeout, because a hidden modal
# dialog (e.g. "File in use") blocks COM forever. On timeout all Office processes are killed.
param(
  [string]$Demo = 'http://localhost:3443',
  [string]$Jars = $PSScriptRoot,
  [ValidateSet('main', 'expiry', 'big')][string]$Phase = 'main',
  [string]$BigFile = '',
  [int]$IdleSeconds = 90,
  # Show Office windows and real dialogs, slowed down, so a person can watch (dialogs then wait for a click).
  [switch]$Visible,
  # Run only these case numbers (default: all in the phase).
  [int[]]$Only = @()
)
function Want([int]$n) { $Only.Count -eq 0 -or $Only -contains $n }
# Every case starts here. KillOffice first: a Word still quitting from the previous case
# is reused by the next COM New-Object and then exits, closing (and UNLOCKing) the new document.
function Announce([string]$t) { KillOffice; if ($Visible) { Write-Host ''; Write-Host "=== WATCH: $t ===" -ForegroundColor Cyan } }
$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
$flagDir = Join-Path $env:TEMP "office-suite-flags"
New-Item -ItemType Directory -Force $flagDir | Out-Null
# A flag left by an aborted run would release the next run's wait step at once
# (case 2 then closed Alice's document before Bob opened it, a false FAIL).
Get-ChildItem $flagDir -File | Remove-Item -Force

function Api([string]$user, [string]$method, [string]$path, [string]$extra = '') {
  $jar = Join-Path $Jars "$user.jar"
  # X-Requested-With is the demo site's CSRF guard for /api.
  $out = & curl.exe -s -b $jar -X $method -H 'X-Requested-With: fetch' -w "`n%{http_code}" $extra.Split(' ', [StringSplitOptions]::RemoveEmptyEntries) "$Demo$path"
  $lines = @($out)
  $code = [int]$lines[-1]
  $body = ($lines[0..([Math]::Max(0, $lines.Count - 2))] -join "`n")
  [pscustomobject]@{ Code = $code; Json = $(if ($body.Trim().StartsWith('{') -or $body.Trim().StartsWith('[')) { $body | ConvertFrom-Json } else { $null }) }
}
function Files { (Api alice GET '/api/files').Json }
function Detail([string]$id) { (Api alice GET "/api/files/$id").Json }
function NewSession([string]$user, [string]$id) { (Api $user POST "/api/files/$id/edit-session").Json }
function KillOffice { Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep 2 }

function Record([string]$case, [string]$expect, [bool]$pass, [string]$observed) {
  $results.Add([pscustomobject]@{ Case = $case; Pass = $(if ($pass) { 'PASS' } else { 'FAIL' }); Expected = $expect; Observed = $observed })
  Write-Host ("[{0}] {1} :: {2}" -f $(if ($pass) { 'PASS' } else { 'FAIL' }), $case, $observed)
}

# One Office session as a job. Steps: open, edit, save, close, saveas:<url>, wait:<flagname>, hold:<sec>, noclose.
# Prints "step=result" lines; exceptions are caught per step so later steps still run.
$officeJob = {
  param([string]$app, [string]$url, [string[]]$steps, [string]$flagDir, [bool]$visible = $false)
  $ErrorActionPreference = 'Stop'
  function Say($s) { Write-Output $s }
  try {
    switch ($app) {
      'word' { $o = New-Object -ComObject Word.Application; $o.Visible = $visible; $o.DisplayAlerts = $(if ($visible) { -1 } else { 0 }) }
      'excel' { $o = New-Object -ComObject Excel.Application; $o.Visible = $visible; $o.DisplayAlerts = $visible }
      'powerpoint' { $o = New-Object -ComObject PowerPoint.Application; $o.DisplayAlerts = 1 }
    }
  } catch { Say "start=ERR $($_.Exception.Message)"; return }
  $d = $null; $closed = $false; $n = 0
  foreach ($step in $steps) {
    $n++
    if ($visible) { Start-Sleep 3 } # slow enough to follow on screen
    try {
      switch -Regex ($step) {
        '^open$' {
          $d = switch ($app) {
            'word' { $o.Documents.Open($url, $false, $false) }
            'excel' { $o.Workbooks.Open($url, 0, $false) }
            'powerpoint' { $o.Presentations.Open($url, 0, 0, $(if ($visible) { -1 } else { 0 })) }
          }
          $ro = [bool]($d.ReadOnly -ne 0 -and $d.ReadOnly -ne $false)
          Say "open=ok readonly=$ro"
        }
        '^edit$' {
          switch ($app) {
            'word' { $d.Content.InsertAfter(" edit$n") | Out-Null }
            'excel' { $d.Worksheets.Item(1).Range('A1').Value2 = "edit$n" }
            'powerpoint' { $d.Slides.Item(1).Shapes.AddTextbox(1, 10, 10 + $n, 200, 40).TextFrame.TextRange.Text = "edit$n" }
          }
          Say "edit=ok"
        }
        '^save$' { $d.Save(); $s = if ($app -eq 'powerpoint') { $d.Saved -ne 0 } else { [bool]$d.Saved }; Say "save=returned saved=$s" }
        '^saveas:(.+)$' {
          $target = $Matches[1]
          switch ($app) {
            'word' { $d.SaveAs2($target, 16) }
            'excel' { $d.SaveAs($target, 51) }
            'powerpoint' { $d.SaveAs($target, 24) }
          }
          $name = if ($app -eq 'word') { $d.FullName } else { $d.FullName }
          Say "saveas=returned fullname=$name"
        }
        '^wait:(.+)$' {
          $f = Join-Path $flagDir $Matches[1]
          Say "waiting=$($Matches[1])"
          while (-not (Test-Path $f)) { Start-Sleep -Milliseconds 300 }
          Remove-Item $f -Force
        }
        '^hold:(\d+)$' { Start-Sleep ([int]$Matches[1]) ; Say "hold=done" }
        '^close$' { if ($app -eq 'powerpoint') { $d.Close() } else { $d.Close(0) }; $closed = $true; Say "close=ok" }
        '^noclose$' { Say "noclose"; return }
      }
    } catch { Say "$($step.Split(':')[0])=ERR $($_.Exception.Message.Split([char]10)[0])" }
  }
  try { if (-not $closed -and $d) { if ($app -eq 'powerpoint') { $d.Close() } else { $d.Close(0) } } } catch {}
  try { $o.Quit() } catch {}
}

function RunOffice([string]$app, [string]$url, [string[]]$steps, [int]$timeout = 90, [scriptblock]$during = $null) {
  $job = Start-Job -ScriptBlock $officeJob -ArgumentList $app, $url, $steps, $flagDir, $Visible.IsPresent
  if ($during) { & $during $job }
  $done = Wait-Job $job -Timeout $timeout
  $out = @(Receive-Job $job -ErrorAction SilentlyContinue)
  if (-not $done) { $out += 'TIMEOUT (blocked, likely a hidden Office dialog)'; Stop-Job $job; KillOffice }
  Remove-Job $job -Force
  return ($out -join ' | ')
}

function WaitFor([string]$job, [string]$marker, [int]$sec = 60) {
  $deadline = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $deadline) {
    $o = @(Receive-Job -Id $job -Keep -ErrorAction SilentlyContinue) -join ' '
    if ($o -match [regex]::Escape($marker)) { return $true }
    Start-Sleep -Milliseconds 300
  }
  return $false
}
function Flag([string]$name) { New-Item -ItemType File -Force (Join-Path $flagDir $name) | Out-Null }

KillOffice
$all = Files
$byExt = @{}
foreach ($f in $all) { $byExt[$f.name] = $f }
$word = $all | Where-Object { $_.name -eq 'Project Plan.docx' }
$xl = $all | Where-Object { $_.name -like '*.xlsx' -and $_.name -notlike 'big*' } | Select-Object -First 1
$ppt = $all | Where-Object { $_.name -like '*.pptx' } | Select-Object -First 1
$heb = $all | Where-Object { $_.name -match '[֐-׿]' } | Select-Object -First 1
$appOf = @{ docx = 'word'; xlsx = 'excel'; pptx = 'powerpoint' }

if ($Phase -eq 'main') {
  # 1. Basic round trip per app: open editable, two saves become two versions, close releases the lock.
  if (Want 1) { Announce "1. Basic round trip per app: open editable, two saves become two versions, close releases the lock."
  foreach ($f in @($word, $xl, $ppt, $heb)) {
    $app = $appOf[$f.name.Split('.')[-1]]
    $before = (Detail $f.id).versionCount
    $s = NewSession alice $f.id
    $o = RunOffice $app $s.davUrl @('open', 'edit', 'save', 'edit', 'save', 'close')
    Start-Sleep 2
    $after = Detail $f.id
    $ok = ($o -match 'readonly=False') -and ($after.versionCount -eq $before + 2) -and ($null -eq $after.lock)
    Record "1 round trip $app ($($f.name))" 'editable, +2 versions, lock released' $ok "$o || versions $before->$($after.versionCount), lock=$([bool]$after.lock)"
  }

  }
  # 2. Second user while Alice holds the lock.
  if (Want 2) { Announce "2. Second user while Alice holds the lock."
  $sa = NewSession alice $word.id
  $ja = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $sa.davUrl, @('open', 'wait:release2', 'close'), $flagDir, $Visible.IsPresent
  WaitFor $ja.Id 'waiting=release2' | Out-Null
  $sb = NewSession bob $word.id
  $ob = RunOffice word $sb.davUrl @('open', 'edit', 'save', 'close') 45
  $v = (Detail $word.id).versionCount
  Record '2 second user while locked' 'Bob read-only or blocked; no version from Bob' ($ob -match 'readonly=True' -or $ob -match 'TIMEOUT' -or $ob -match 'open=ERR') "$ob || lockedBy in session API: $($sb.lockedBy)"
  Flag 'release2'; Wait-Job $ja -Timeout 30 | Out-Null; Receive-Job $ja | Out-Null; Remove-Job $ja -Force; KillOffice

  }
  # 3. Crash: kill Word while it holds the lock, then the same user reopens (new session, like clicking Edit again).
  if (Want 3) { Announce "3. Crash: kill Word while it holds the lock, then the same user reopens (new session, like clicking Edit again)."
  $s1 = NewSession alice $word.id
  $jc = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $s1.davUrl, @('open', 'wait:never3'), $flagDir, $Visible.IsPresent
  WaitFor $jc.Id 'waiting=never3' | Out-Null
  KillOffice; Stop-Job $jc; Remove-Job $jc -Force
  $lockAfterCrash = (Detail $word.id).lock
  $s2 = NewSession alice $word.id
  $oc = RunOffice word $s2.davUrl @('open', 'edit', 'save', 'close') 45
  Record '3 crash then same-user reopen' 'reopens editable (plan: same-user takeover)' ($oc -match 'readonly=False' -and $oc -notmatch 'save=ERR') "lock still held after crash=$([bool]$lockAfterCrash) || $oc"
  KillOffice
  Api alice POST "/api/files/$($word.id)/force-unlock" | Out-Null

  }
  # 4. Save As to a new name inside the session folder.
  if (Want 4) { Announce "4. Save As to a new name inside the session folder."
  foreach ($f in @($word, $xl)) {
    $app = $appOf[$f.name.Split('.')[-1]]
    $before = (Detail $f.id).versionCount
    $filesBefore = @(Files).Count
    $s = NewSession alice $f.id
    $target = $s.davUrl -replace '[^/]+$', ("SaveAsCopy." + $f.name.Split('.')[-1])
    $o = RunOffice $app $s.davUrl @('open', 'edit', "saveas:$target", 'edit', 'save', 'close') 60
    Start-Sleep 2
    $after = Detail $f.id
    $filesAfter = @(Files).Count
    # The save after Save As must land on the new file (Word re-points the document there).
    $copy = @(Files) | Where-Object { $_.name -eq ("SaveAsCopy." + $f.name.Split('.')[-1]) } | Select-Object -First 1
    $copyVersions = if ($copy) { (Detail $copy.id).versionCount } else { 0 }
    $ok = ($o -match 'saveas=returned') -and $copy -and ($copyVersions -ge 2) -and ($after.versionCount -eq $before)
    Record "4 Save As $app" 'new file created, the later save lands on it, original untouched' $ok "$o || files $filesBefore->$filesAfter, copy versions=$copyVersions, original versions $before->$($after.versionCount), lock=$([bool]$after.lock)"
    KillOffice
    if ($after.lock) { Api alice POST "/api/files/$($f.id)/force-unlock" | Out-Null }
  }

  }
  # 5. Admin force unlock while Word is open with unsaved edits, then the user saves.
  if (Want 5) { Announce "5. Admin force unlock while Word is open with unsaved edits, then the user saves."
  $before = (Detail $word.id).versionCount
  $s = NewSession alice $word.id
  $j5 = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $s.davUrl, @('open', 'edit', 'wait:go5', 'save', 'close'), $flagDir, $Visible.IsPresent
  WaitFor $j5.Id 'waiting=go5' | Out-Null
  $fu = Api alice POST "/api/files/$($word.id)/force-unlock"
  Flag 'go5'; $done = Wait-Job $j5 -Timeout 60; $o5 = (@(Receive-Job $j5) -join ' | '); if (-not $done) { $o5 += ' | TIMEOUT'; KillOffice }; Remove-Job $j5 -Force
  $d5 = Detail $word.id; $v5 = $d5.versionCount
  Record '5 force unlock then save' 'save kept as a conflict version' (($v5 -gt $before) -and ($d5.versions[0].source -eq 'conflict')) "force-unlock=$($fu.Code) || $o5 || versions $before->$v5, newest source=$($d5.versions[0].source)"
  KillOffice

  }
  # 6. Revoke the session mid-edit, then save.
  if (Want 6) { Announce "6. Revoke the session mid-edit, then save."
  $before = (Detail $word.id).versionCount
  $s = NewSession alice $word.id
  $j6 = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $s.davUrl, @('open', 'edit', 'wait:go6', 'save', 'close'), $flagDir, $Visible.IsPresent
  WaitFor $j6.Id 'waiting=go6' | Out-Null
  $rv = Api alice POST "/api/sessions/$($s.sessionId)/revoke"
  Start-Sleep 6 # past RESOLVE_CACHE_MS
  Flag 'go6'; $done = Wait-Job $j6 -Timeout 60; $o6 = (@(Receive-Job $j6) -join ' | '); if (-not $done) { $o6 += ' | TIMEOUT'; KillOffice }; Remove-Job $j6 -Force
  $v6 = (Detail $word.id).versionCount
  $q6 = @((Api alice GET '/api/admin/quarantine').Json) | Where-Object { $_.sessionId -eq $s.sessionId }
  Record '6 revoke then save' 'save quarantined (not a version), no prompt or hang' ((@($q6).Count -ge 1) -and ($v6 -eq $before) -and ($o6 -notmatch 'TIMEOUT')) "revoke=$($rv.Code) || $o6 || versions $before->$v6, quarantined=$(@($q6).Count) ($(@($q6 | ForEach-Object { $_.reason }) -join ','))"
  KillOffice

  }
  # 7. Web upload while Office holds the lock.
  if (Want 7) { Announce "7. Web upload while Office holds the lock."
  $s = NewSession alice $word.id
  $j7 = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $s.davUrl, @('open', 'wait:go7', 'close'), $flagDir, $Visible.IsPresent
  WaitFor $j7.Id 'waiting=go7' | Out-Null
  # A real docx body (the current version, fetched over DAV) so the upload is refused for the lock, not the format.
  $tmp = Join-Path $env:TEMP 'upload7.docx'; & curl.exe -s -o $tmp $s.davUrl
  $up = Api alice POST "/api/files/$($word.id)/versions" "-F file=@$tmp"
  Record '7 web upload while locked' '409' ($up.Code -eq 409) "upload=$($up.Code)"
  Flag 'go7'; Wait-Job $j7 -Timeout 30 | Out-Null; Remove-Job $j7 -Force; KillOffice
  }
}

if ($Phase -eq 'expiry') {
  # 8. Long idle: services run with LOCK_MAX_TIMEOUT_S=60. Office asks for 600 s, gets 60 s.
  if (Want 8) { Announce "8. Long idle: services run with LOCK_MAX_TIMEOUT_S=60. Office asks for 600 s, gets 60 s."
  # Expect either a LOCK refresh during the idle, or a revival on the next save.
  $before = (Detail $word.id).versionCount
  $s = NewSession alice $word.id
  $o = RunOffice word $s.davUrl @('open', 'edit', "hold:$IdleSeconds", 'save', 'close') ($IdleSeconds + 90)
  $v = (Detail $word.id).versionCount
  Record "8 idle ${IdleSeconds}s past a 60s lock, then save" 'save stored (refresh or revival)' ($v -eq $before + 1) "$o || versions $before->$v"
  KillOffice

  }
  # 9. Expired lock overtaken: Alice idles past expiry, Bob opens, edits, saves and closes, then Alice saves.
  if (Want 9) { Announce "9. Expired lock overtaken: Alice idles past expiry, Bob opens, edits, saves and closes, then Alice saves."
  $before = (Detail $word.id).versionCount
  $sa = NewSession alice $word.id
  $ja = Start-Job -ScriptBlock $officeJob -ArgumentList 'word', $sa.davUrl, @('open', 'edit', 'wait:go9', 'save', 'close'), $flagDir, $Visible.IsPresent
  WaitFor $ja.Id 'waiting=go9' | Out-Null
  Start-Sleep ($IdleSeconds)
  $sb = NewSession bob $word.id
  $ob = RunOffice word $sb.davUrl @('open', 'edit', 'save', 'close') 60
  Flag 'go9'; $done = Wait-Job $ja -Timeout 60; $oa = (@(Receive-Job $ja) -join ' | '); if (-not $done) { $oa += ' | TIMEOUT'; KillOffice }; Remove-Job $ja -Force
  $d9 = Detail $word.id; $v = $d9.versionCount
  $top = @($d9.versions | Select-Object -First 2 | ForEach-Object { "$($_.source)/$($_.createdBy.id)" }) -join ','
  # Newest first: Alice's late save as a conflict on top of Bob's current version.
  Record '9 expired lock overtaken by another user' "Bob's save current; Alice's save kept as a conflict" ($top -eq 'conflict/alice,dav/bob') "bob: $ob || alice: $oa || versions $before->$v, newest=$top"
  KillOffice
  }
}

if ($Phase -eq 'big') {
  # 10. 100 MB workbook: open, edit, save. Uploaded beforehand through the web API.
  if (Want 10) { Announce "10. 100 MB workbook: open, edit, save. Uploaded beforehand through the web API."
  $f = $all | Where-Object { $_.name -like 'big*' } | Select-Object -First 1
  $before = (Detail $f.id).versionCount
  $s = NewSession alice $f.id
  $t0 = Get-Date
  $o = RunOffice excel $s.davUrl @('open', 'edit', 'save', 'close') 900
  # Excel uploads big files in the background: COM Save() returns before the PUT ends, so poll.
  $deadline = (Get-Date).AddSeconds(120)
  do { Start-Sleep 2; $v = (Detail $f.id).versionCount } while ($v -eq $before -and (Get-Date) -lt $deadline)
  $o += " | lock after close=$([bool](Detail $f.id).lock)"
  Record "10 big xlsx ($([math]::Round($f.size/1MB)) MB)" 'editable, +1 version' (($o -match 'readonly=False') -and $v -eq $before + 1) "$o || versions $before->$v || $([int]((Get-Date)-$t0).TotalSeconds)s"
  }
}

$results | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 (Join-Path $Jars "results-$Phase.json")
$results | Format-Table Case, Pass, Observed -Wrap -AutoSize | Out-String -Width 220
