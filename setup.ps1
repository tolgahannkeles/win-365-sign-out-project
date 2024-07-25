# RESET OFFICE ACTIVATION STATE
# ... (rest of the header is the same)

# Determine the script's location dynamically
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Construct relative paths to the files
$OLicenseCleanupPath = Join-Path $ScriptPath "OLicenseCleanup.vbs"
$SignoutOfWamAccountsPath = Join-Path $ScriptPath "signoutofwamaccounts.ps1"
$WPJCleanUpPath = Join-Path $ScriptPath "WPJCleanUp\WPJCleanUp\WPJCleanUp.cmd"

# Check existence of files with relative paths
Write-Host "Test OLicenseCleanup.vbs: " -NoNewLine; Test-Path $OLicenseCleanupPath
Write-Host "Test signoutofwamaccounts.ps1: " -NoNewLine; Test-Path $SignoutOfWamAccountsPath
Write-Host "Test WPJCleanUp.cmd: " -NoNewLine; Test-Path $WPJCleanUpPath

# Execute scripts using relative paths
Write-Host "Execute OLicenseCleanup.vbs: "
Start-Process -FilePath cscript.exe -ArgumentList $OLicenseCleanupPath -NoNewWindow -PassThru
Start-Sleep 3

Write-Host "Execute signoutofwamaccounts.ps1: "
Start-Process -FilePath powershell.exe -ArgumentList $SignoutOfWamAccountsPath -NoNewWindow -PassThru
Start-Sleep 3

Write-Host "Execute WPJCleanUp.cmd: "
Start-Process $WPJCleanUpPath -NoNewWindow -PassThru
Start-Sleep 3
