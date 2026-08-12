<#
.SYNOPSIS
    Register (or remove) a Windows scheduled task that renders pending case
    study slide images.

.DESCRIPTION
    The job is deliberately timid. It runs only while you are logged in, does
    nothing at all if PowerPoint is already open (automating a window you are
    using steals focus, and a modal dialog in it blocks every COM call), and
    exits quietly when no case study is waiting -- which is most days. It
    never opens a window: pythonw.exe is used when available, and everything
    it did is appended to the log.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\schedule_render_pending.ps1 -Register
    powershell -ExecutionPolicy Bypass -File .\schedule_render_pending.ps1 -Register -At 17:30
    powershell -ExecutionPolicy Bypass -File .\schedule_render_pending.ps1 -Status
    powershell -ExecutionPolicy Bypass -File .\schedule_render_pending.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [switch]$Register,
    [switch]$Unregister,
    [switch]$Status,
    [string]$TaskName = "Premion Proposal Builder - render case study images",
    [string]$At = "08:30"
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-PythonPath {
    # pythonw runs without a console window, which is what makes a scheduled
    # run invisible. Fall back to python if it isn't there.
    foreach ($candidate in @(".venv\Scripts\pythonw.exe", "venv\Scripts\pythonw.exe")) {
        $full = Join-Path $repo $candidate
        if (Test-Path $full) { return $full }
    }
    $pythonw = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($pythonw) { return $pythonw.Source }
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return $python.Source }
    throw "No python.exe or pythonw.exe found on PATH."
}

if ($Unregister) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task: $TaskName"
    } else {
        Write-Host "No such scheduled task: $TaskName"
    }
    return
}

if ($Status) {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) { Write-Host "Not registered."; return }
    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    Write-Host "State        : $($task.State)"
    Write-Host "Last run     : $($info.LastRunTime)"
    Write-Host "Last result  : $($info.LastTaskResult)  (0 = ok, 1 = something failed, 2 = couldn't run)"
    Write-Host "Next run     : $($info.NextRunTime)"
    return
}

if (-not $Register) {
    Write-Host "Pass -Register, -Unregister or -Status. See -? for details."
    return
}

$python = Get-PythonPath
$arguments = "render_case_study_images.py --pending --quiet --skip-if-busy"

$action = New-ScheduledTaskAction -Execute $python -Argument $arguments -WorkingDirectory $repo
$trigger = New-ScheduledTaskTrigger -Weekly -At $At `
    -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday
# Interactive: it needs your desktop session, because it drives PowerPoint.
# A task set to "run whether logged on or not" would launch PowerPoint in a
# session with no display and hang.
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force | Out-Null

Write-Host "Registered: $TaskName"
Write-Host "  runs     : weekdays at $At, while you're logged in"
Write-Host "  command  : $python $arguments"
Write-Host "  folder   : $repo"
Write-Host "  skips    : silently when PowerPoint is open, or nothing is pending"
Write-Host ""
Write-Host "Run it now to check:  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Remove it later    :  .\schedule_render_pending.ps1 -Unregister"
