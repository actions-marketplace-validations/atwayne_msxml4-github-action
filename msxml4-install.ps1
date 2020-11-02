# Copied from https://raw.githubusercontent.com/dblock/msiext/master/appveyor.yml
# MSXML 4.0 Service Pack 3 (Microsoft XML Core Services)
Write-Host "Downloading MSXML 4.0 Service Pack 3 (Microsoft XML Core Services)..." -ForegroundColor Green
$msiFilePath = "$($env:USERPROFILE)\msxml.msi"
$logFilePath = "$($env:TEMP)\msxml.txt"
(New-Object Net.WebClient).DownloadFile('http://download.microsoft.com/download/A/2/D/A2D8587D-0027-4217-9DAD-38AFDB0A177E/msxml.msi', $msiFilePath)
$process = (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $msiFilePath /quiet /l*v $logFilePath" -Wait -Passthru)
$exitCode = $process.ExitCode
if ($exitCode -ne 0) {
    Get-Content $logFilePath
    throw "Command failed with exit code $exitCode."
}

del $msiFilePath
del $logFilePath
Write-Host "MSXML 4.0 Service Pack 3 (Microsoft XML Core Services) installed successfully" -ForegroundColor Green