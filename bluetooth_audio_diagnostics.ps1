# Deep Audio Diagnostics Script
# Run this in PowerShell as Administrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Deep Audio Diagnostics" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Import module
Import-Module AudioDeviceCmdlets -ErrorAction SilentlyContinue

# Check current device with full details
Write-Host "[1] Current Default Device - Full Details:" -ForegroundColor Yellow
$currentDevice = Get-AudioDevice -Playback
$currentDevice | Format-List *
Write-Host ""

# Check ALL playback devices with their state
Write-Host "[2] All Playback Devices with Status:" -ForegroundColor Yellow
Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback"} | Format-Table Index, Default, Name, DeviceState -AutoSize
Write-Host ""

# Check volume level of S6520
Write-Host "[3] S6520 Volume Level:" -ForegroundColor Yellow
try {
    $s6520Volume = Get-AudioDevice -Index 1
    Write-Host "   Device: $($s6520Volume.Name)" -ForegroundColor White
    Write-Host "   Volume: $($s6520Volume.Volume)%" -ForegroundColor White
    Write-Host "   Muted: $($s6520Volume.Mute)" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Host "   Unable to get volume info.`n" -ForegroundColor Red
}

# Check if device is actually disabled
Write-Host "[4] S6520 PnP Device Status:" -ForegroundColor Yellow
$s6520Pnp = Get-PnpDevice | Where-Object {$_.FriendlyName -like "*S6520*"}
if ($s6520Pnp) {
    $s6520Pnp | Format-Table FriendlyName, Status, Class, ProblemCode -AutoSize
    Write-Host ""
} else {
    Write-Host "   S6520 PnP device not found!`n" -ForegroundColor Red
}

# Check Bluetooth audio driver status
Write-Host "[5] Bluetooth Audio Driver Status:" -ForegroundColor Yellow
Get-PnpDevice -Class "MEDIA" | Where-Object {$_.FriendlyName -like "*Bluetooth*" -or $_.FriendlyName -like "*S6520*"} | 
    Format-Table FriendlyName, Status, Class -AutoSize
Write-Host ""

# Check for Bluetooth A2DP profile
Write-Host "[6] Bluetooth Audio Profiles (A2DP):" -ForegroundColor Yellow
Get-PnpDevice | Where-Object {$_.FriendlyName -like "*A2dp*" -or $_.InstanceId -like "*A2DP*"} | 
    Select-Object FriendlyName, Status, InstanceId | Format-Table -AutoSize
Write-Host ""

# Check Windows audio troubleshooter recommendations
Write-Host "[7] Audio Device Driver Info:" -ForegroundColor Yellow
Get-WmiObject Win32_SoundDevice | Where-Object {$_.Name -like "*Bluetooth*" -or $_.Name -like "*S6520*"} | 
    Select-Object Name, Status, StatusInfo, Manufacturer | Format-Table -AutoSize
Write-Host ""

# Check for audio format issues
Write-Host "[8] Checking System Audio Configuration:" -ForegroundColor Yellow
$audioKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
if (Test-Path $audioKey) {
    Write-Host "   Audio registry key exists.`n" -ForegroundColor Green
} else {
    Write-Host "   WARNING: Audio registry key not found!`n" -ForegroundColor Red
}

# Check for exclusive mode issues
Write-Host "[9] Getting S6520 Device ID for manual check:" -ForegroundColor Yellow
$s6520Device = Get-AudioDevice -Index 1
Write-Host "   Device ID: $($s6520Device.ID)" -ForegroundColor White
Write-Host "   Copy this ID if needed for manual registry checks.`n" -ForegroundColor Gray

# Test if we can actually switch devices
Write-Host "[10] Testing Device Switching:" -ForegroundColor Yellow
Write-Host "   Current device: $(Get-AudioDevice -Playback | Select-Object -ExpandProperty Name)" -ForegroundColor White
Write-Host "   Attempting to switch to device Index 2..." -ForegroundColor White
try {
    Set-AudioDevice -Index 2
    Start-Sleep -Seconds 1
    Write-Host "   Switched to: $(Get-AudioDevice -Playback | Select-Object -ExpandProperty Name)" -ForegroundColor White
    Write-Host "   Switching back to S6520..." -ForegroundColor White
    Set-AudioDevice -Index 1
    Start-Sleep -Seconds 1
    Write-Host "   Current device: $(Get-AudioDevice -Playback | Select-Object -ExpandProperty Name)" -ForegroundColor White
    Write-Host "   Device switching works!`n" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Device switching failed!`n" -ForegroundColor Red
}

# Check for any error codes
Write-Host "[11] Checking for Device Problem Codes:" -ForegroundColor Yellow
$allAudioDevices = Get-PnpDevice -Class "AudioEndpoint","MEDIA" | Where-Object {$_.FriendlyName -like "*S6520*" -or $_.FriendlyName -like "*Soundbar*"}
if ($allAudioDevices) {
    foreach ($device in $allAudioDevices) {
        Write-Host "   Device: $($device.FriendlyName)" -ForegroundColor White
        Write-Host "   Status: $($device.Status)" -ForegroundColor White
        Write-Host "   Problem: $($device.Problem)" -ForegroundColor White
        if ($device.Status -ne "OK") {
            Write-Host "   ⚠️  ISSUE DETECTED!" -ForegroundColor Red
        }
        Write-Host ""
    }
} else {
    Write-Host "   No S6520 devices found in audio endpoints.`n" -ForegroundColor Yellow
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Diagnostics Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nIMPORTANT: While this script is still open, please:" -ForegroundColor Yellow
Write-Host "1. Play a test sound (YouTube, etc.)" -ForegroundColor White
Write-Host "2. Right-click the volume icon in taskbar" -ForegroundColor White
Write-Host "3. Select 'Open Volume Mixer'" -ForegroundColor White
Write-Host "4. Check if S6520 shows up and what its volume level is" -ForegroundColor White
Write-Host "5. Also check if individual apps have volume set to 0`n" -ForegroundColor White