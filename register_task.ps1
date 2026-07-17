# Registriert (oder aktualisiert) den geplanten Task fuer cron-urlaubspreise.
# Liest den Zeitplan aus der Konfiguration (config.xlsx): frequenz_stunden, verpasste_nachholen.
# Laeuft rund um die Uhr (kein Zeitfenster).
# Aufruf:  powershell -ExecutionPolicy Bypass -File register_task.ps1
$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
# Zeitplan aus der Konfiguration (config.xlsx, Ort via config-pfad.json) - Laden ueber
# die Python-Ladefunktion, damit die Parse-Logik nur an einer Stelle lebt.
$py = "import sys; sys.path.insert(0, r'$dir'); import flugpreise; zp = flugpreise.lade_config(flugpreise.CONFIG)['job']['zeitplan']; print(str(zp['frequenz_stunden']) + ';' + str(zp['verpasste_nachholen']))"
$vals = (& python -c $py | Select-Object -Last 1).Trim() -split ';'

$freq = [int]$vals[0]
$nachholen = $vals[1] -eq 'True'

# pythonw.exe = lautlos (kein Konsolenfenster)
$pyExe = (Get-Command python).Source
$pyw = Join-Path (Split-Path $pyExe) "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $pyExe }

$TaskName = "cron-urlaubspreise"
$action = New-ScheduledTaskAction -Execute $pyw -Argument "flugpreise.py" -WorkingDirectory $dir

# Taeglicher Trigger ab Mitternacht, Wiederholung alle freq Stunden ueber 24h (kein Zeitfenster)
$trigger = New-ScheduledTaskTrigger -Daily -At "00:00"
$rep = (New-ScheduledTaskTrigger -Once -At "00:00" `
        -RepetitionInterval (New-TimeSpan -Hours $freq) `
        -RepetitionDuration (New-TimeSpan -Hours 24)).Repetition
$trigger.Repetition = $rep

# StartWhenAvailable:$false = verpasste Laeufe NICHT nachholen; IgnoreNew = keine Ueberlappung
# AllowStartIfOnBatteries/DontStopIfGoingOnBatteries = laeuft auch im Akkubetrieb (Laptop)
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable:$nachholen `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -DontStopOnIdleEnd `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Description "Flugpreis-Abfrage cron-urlaubspreise" -RunLevel Limited | Out-Null

Write-Host "Task '$TaskName' registriert."
Write-Host ("  Rund um die Uhr, alle {0} h" -f $freq)
Write-Host ("  Verpasste nachholen: {0}" -f $nachholen)
Write-Host ("  Programm: {0} flugpreise.py" -f $pyw)
