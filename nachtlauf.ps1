# Arbeitet die Tagesrotation am Stueck ab (statt stuendlich): laeuft, bis NICHTS_ZU_TUN
# oder max. 8 Durchgaenge. Pausiert waehrenddessen den geplanten Task (kein Parallel-Scan).
$ErrorActionPreference = "Continue"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$log = Join-Path $dir "lauf.log"
try { Disable-ScheduledTask -TaskName "cron-urlaubspreise" | Out-Null } catch {}
Set-Location $dir
for ($i = 1; $i -le 8; $i++) {
    python flugpreise.py
    $letzte = Get-Content $log -Tail 1
    if ($letzte -match "NICHTS_ZU_TUN|CONFIG_GESPERRT") { break }
}
try { Enable-ScheduledTask -TaskName "cron-urlaubspreise" | Out-Null } catch {}
