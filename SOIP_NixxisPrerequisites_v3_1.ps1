# Enable running PowerShell scripts
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

# Function to check if an application is installed
function Test-SoftwareInstalled {
    param (
        [string]$name
    )
    $uninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $uninstallKeyWOW = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    $keys = @($uninstallKey, $uninstallKeyWOW)
    foreach ($key in $keys) {
        if (Get-ItemProperty -Path "$key\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$name*" }) {
            return $true
        }
    }
    return $false
}

# Check and install Microsoft Edge WebView2 Runtime if not installed
if (-not (Test-SoftwareInstalled -name "Microsoft Edge WebView2 Runtime")) {
    $webview2Url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
    $webview2Installer = "$env:TEMP\MicrosoftEdgeWebView2RuntimeInstallerX64.exe"
    Invoke-WebRequest -Uri $webview2Url -OutFile $webview2Installer
    Start-Process -FilePath $webview2Installer -ArgumentList "/silent /install" -Wait
}

# Check and install Microsoft .Net Framework 4.8 if not installed
if (-not (Test-SoftwareInstalled -name ".NET Framework 4.8")) {
    $dotNetUrl = "https://go.microsoft.com/fwlink/?linkid=2088631"
    $dotNetInstaller = "$env:TEMP\ndp48-x86-x64-allos-enu.exe"
    Invoke-WebRequest -Uri $dotNetUrl -OutFile $dotNetInstaller
    Start-Process -FilePath $dotNetInstaller -ArgumentList "/q /norestart" -Wait
}

# Check and install Visual C++ Redistributable if not installed
if (-not (Test-SoftwareInstalled -name "Microsoft Visual C++ 2015-2022 Redistributable")) {
    $vcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    $vcRedistInstaller = "$env:TEMP\vc_redist.x64.exe"
    Invoke-WebRequest -Uri $vcRedistUrl -OutFile $vcRedistInstaller
    Start-Process -FilePath $vcRedistInstaller -ArgumentList "/install /quiet /norestart" -Wait
}

# Ensure C:\Nixxis exists
$nixxisPath = "C:\Nixxis"
if (-Not (Test-Path -Path $nixxisPath)) {
    New-Item -Path $nixxisPath -ItemType Directory
}

# Download clientsoftware.zip to C:\Nixxis
$nixxisUrl = "http://soip.cloud.soip.info:8088/provisioning/Client.3.1.3.zip"
$nixxisZip = "$nixxisPath\clientsoftware.zip"
try {
    Invoke-WebRequest -Uri $nixxisUrl -OutFile $nixxisZip -ErrorAction Stop
    Write-Output "Download successful: $nixxisZip"
} catch {
    Write-Error "Failed to download file from $nixxisUrl"
    return
}

# Rename clientsoftware.zip to clientsoftwareV3.zip
$nixxisZipV3 = "$nixxisPath\clientsoftwareV3.zip"
try {
    Rename-Item -Path $nixxisZip -NewName $nixxisZipV3 -ErrorAction Stop
    Write-Output "File renamed to: $nixxisZipV3"
} catch {
    Write-Error "Failed to rename file to $nixxisZipV3"
    return
}

# Unzip clientsoftwareV3.zip to folder C:\Nixxis\clientsoftwareV3
$nixxisUnzipPath = "$nixxisPath\clientsoftwareV3"
try {
    Expand-Archive -Path $nixxisZipV3 -DestinationPath $nixxisUnzipPath -ErrorAction Stop
    Write-Output "Unzipped to: $nixxisUnzipPath"
} catch {
    Write-Error "Failed to unzip file to $nixxisUnzipPath"
    return
}

# Create desktop shortcut for NixxisClientDesktop.exe
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut("$env:USERPROFILE\Desktop\Nixxis v3.lnk")
$shortcut.TargetPath = "C:\Nixxis\clientsoftwareV3\NixxisClientDesktop.exe"
$shortcut.Save()

# Cleanup installers
$installers = @($webview2Installer, $dotNetInstaller, $vcRedistInstaller)
foreach ($installer in $installers) {
    if (Test-Path -Path $installer) {
        Remove-Item -Path $installer
        Write-Output "Removed installer: $installer"
    } else {
        Write-Output "Installer not found: $installer"
    }
}
