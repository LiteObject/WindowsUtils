# Advanced Bluetooth Audio Fix Script
# Run this in PowerShell as Administrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Advanced Bluetooth Audio Fix" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Import-Module AudioDeviceCmdlets -ErrorAction SilentlyContinue

Write-Host "ISSUE DETECTED: S6520 volume is unreadable, indicating audio stream problem.`n" -ForegroundColor Red

# Option 1: Force disable and re-enable the Bluetooth audio driver
Write-Host "[Fix 1] Resetting Bluetooth Audio Driver (A2DP)..." -ForegroundColor Yellow
$s6520A2DP = Get-PnpDevice | Where-Object {$_.FriendlyName -eq "S6520 A2DP SNK"}

if ($s6520A2DP) {
    Write-Host "   Found S6520 A2DP device. Resetting..." -ForegroundColor White
    try {
        Disable-PnpDevice -InstanceId $s6520A2DP.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "   Disabled A2DP driver." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        
        Enable-PnpDevice -InstanceId $s6520A2DP.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "   Re-enabled A2DP driver." -ForegroundColor Green
        Start-Sleep -Seconds 3
    } catch {
        Write-Host "   ERROR: Could not reset A2DP driver. Try running as Administrator.`n" -ForegroundColor Red
    }
} else {
    Write-Host "   S6520 A2DP device not found.`n" -ForegroundColor Red
}

# Option 2: Reset the audio endpoint itself
Write-Host "[Fix 2] Resetting Audio Endpoint..." -ForegroundColor Yellow
$s6520Endpoint = Get-PnpDevice | Where-Object {$_.FriendlyName -eq "Soundbar (S6520)" -and $_.Class -eq "AudioEndpoint"}

if ($s6520Endpoint) {
    Write-Host "   Found S6520 audio endpoint. Resetting..." -ForegroundColor White
    try {
        Disable-PnpDevice -InstanceId $s6520Endpoint.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "   Disabled audio endpoint." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        Enable-PnpDevice -InstanceId $s6520Endpoint.InstanceId -Confirm:$false -ErrorAction Stop
        Write-Host "   Re-enabled audio endpoint." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "   ERROR: Could not reset audio endpoint.`n" -ForegroundColor Red
    }
} else {
    Write-Host "   Audio endpoint not found.`n" -ForegroundColor Red
}

# Option 3: Restart Bluetooth services
Write-Host "[Fix 3] Restarting Bluetooth Services..." -ForegroundColor Yellow
try {
    Restart-Service bthserv -Force -ErrorAction Stop
    Write-Host "   Bluetooth Support Service restarted." -ForegroundColor Green
    
    $btAudio = Get-Service BTAGService -ErrorAction SilentlyContinue
    if ($btAudio -and $btAudio.Status -eq "Running") {
        Restart-Service BTAGService -Force
        Write-Host "   Bluetooth Audio Gateway Service restarted." -ForegroundColor Green
    }
    Start-Sleep -Seconds 3
} catch {
    Write-Host "   Could not restart all Bluetooth services.`n" -ForegroundColor Yellow
}

# Option 4: Restart audio services again
Write-Host "[Fix 4] Restarting Audio Services..." -ForegroundColor Yellow
Restart-Service Audiosrv, AudioEndpointBuilder -Force
Write-Host "   Audio services restarted.`n" -ForegroundColor Green
Start-Sleep -Seconds 3

# Option 5: Force set as default again
Write-Host "[Fix 5] Setting S6520 as default device..." -ForegroundColor Yellow
Set-AudioDevice -Index 1
Write-Host "   S6520 set as default.`n" -ForegroundColor Green
Start-Sleep -Seconds 2

# Check the result
Write-Host "[Verification] Checking S6520 status now:" -ForegroundColor Yellow
$currentDevice = Get-AudioDevice -Playback
Write-Host "   Current Device: $($currentDevice.Name)" -ForegroundColor White
Write-Host "   Volume: $($currentDevice.Volume)%" -ForegroundColor White
Write-Host "   Muted: $($currentDevice.Mute)" -ForegroundColor White

if ([string]::IsNullOrEmpty($currentDevice.Volume)) {
    Write-Host "`n   ⚠️  WARNING: Volume is still blank!" -ForegroundColor Red
    Write-Host "   This suggests the S6520 audio stream isn't working properly.`n" -ForegroundColor Red
} else {
    Write-Host "`n   ✓ Volume is now readable! Audio stream should work.`n" -ForegroundColor Green
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fixes Applied!" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Test audio NOW (play YouTube video, etc.)" -ForegroundColor White
Write-Host "2. If still not working, the Bluetooth connection may need to be removed:" -ForegroundColor White
Write-Host "   - Go to Settings > Bluetooth & devices" -ForegroundColor Gray
Write-Host "   - Find S6520 and click 'Remove device'" -ForegroundColor Gray
Write-Host "   - Turn off/on the speaker and re-pair it" -ForegroundColor Gray
Write-Host "3. Check if your soundbar has multiple audio modes (stereo/surround)" -ForegroundColor White
Write-Host "   - Some modes don't work properly with Windows Bluetooth`n" -ForegroundColor Gray