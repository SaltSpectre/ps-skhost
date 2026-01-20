# app_handler.ps1
# Handles installation and uninstallation logic for skHost

. "$PSScriptRoot\icon_handler.ps1"

Function Test-RegistryValue ($regkey, $name) {
    if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
        $true
    }
    else {
        $false
    }
    # https://adamtheautomator.com/powershell-get-registry-value/
}

# Install logic. Does not start after install, even if autostart is specified.
Function Install-skHost {
    param (
        [Parameter(Mandatory = $true)] [string] $ParentPath,
        [Parameter(Mandatory = $false)] [bool] $AutoStart = $false
    )
    
    $InstallPath = "$env:LOCALAPPDATA\SaltSpectre\ps-skhost"
    $ManifestPath = Join-Path $ParentPath "manifest.txt"
    $installLog = "install.log"
    $SESSION_LOG = Join-Path $InstallPath $installLog
    $installLogHeader = $sessionLogHeader -replace "skHost Session Log", "skHost Installation Log"
    
    # Create installation directory
    if (-not (Test-Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
        $createdInstallDir = $true
    }

    if ($createdInstallDir) {
        Set-Content -Path $SESSION_LOG -Value $installLogHeader -Encoding utf8
        Write-skSessionLog -Message "✔️ Created installation directory: $InstallPath" -Type "SUCCESS" -Color Green
    } else {
        Set-Content -Path $SESSION_LOG -Value $installLogHeader -Encoding utf8
        Write-skSessionLog -Message "ℹ️ Installation directory already exists: $InstallPath" -Type "INFO" -Color Cyan
    }
    
    # Copy files listed in manifest.txt
    if (Test-Path $ManifestPath) {
        $ManifestFiles = Get-Content $ManifestPath | Where-Object { $_.Trim() -and -not $_.StartsWith('#') }
        foreach ($file in $ManifestFiles) {
            $SourceFile = Join-Path $ParentPath $file.Trim()
            $DestFile = Join-Path $InstallPath $file.Trim()
            
            if (Test-Path $SourceFile) {
                # Create destination directory if needed
                $DestDir = Split-Path $DestFile -Parent
                if ($DestDir -and -not (Test-Path $DestDir)) {
                    New-Item -Path $DestDir -ItemType Directory -Force | Out-Null
                }
                Copy-Item -Path $SourceFile -Destination $DestFile -Force -Confirm:$false
                Write-skSessionLog -Message "✔️ Copied: $file" -Type "SUCCESS" -Color Green
            } else {
                Write-skSessionLog -Message "❌ File not found: $file" -Type "ERROR" -Color Red
                return
            }
        }
    } else {
        Write-skSessionLog -Message "manifest.txt not found. Installation cannot proceed without manifest file." -Type "ERROR" -Color Red
        return
    }
    
    if ($AutoStart) {
        if (Test-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" "skHost") {
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -Force -Confirm:$false
        }
        # Specify to use conhost in case Windows Terminal is set as default as it does not support WindowStyle Hidden like conhost
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -PropertyType String -Value "conhost powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallPath\skhost.ps1`"" -Force -Confirm:$false
    }

    # Create Start Menu Shortcut
    $WScript = New-Object -ComObject ("WScript.Shell")
    $Shortcut = $Wscript.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\skHost.lnk")
    $Shortcut.TargetPath = "$env:SystemRoot\System32\Conhost.exe" 
    $Shortcut.Arguments = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallPath\skhost.ps1`""
    
    # Set shortcut icon
    $installedIcon = Find-IconFile -BasePath $InstallPath -IconName "skHost"
    if ($installedIcon) {
        $Shortcut.IconLocation = $installedIcon
    }
    else {
        $Shortcut.IconLocation = "$env:SystemRoot\system32\shell32.dll,43"
    }
    $Shortcut.Save()
}

Function Uninstall-skHost {
    # Remove new installation directory structure
    $NewInstallPath = "$env:LOCALAPPDATA\SaltSpectre\ps-skhost"
    if (Test-Path $NewInstallPath) {
        Remove-Item -Path $NewInstallPath -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-skSessionLog -Message "Removed: $NewInstallPath" -Type "SUCCESS" -Color Green
    }
    
    # Clean up old installation files (backward compatibility)
    Remove-Item -Path "$env:LOCALAPPDATA\skHost.ps1" -Force -Confirm:$false -ErrorAction SilentlyContinue
    foreach ($format in @('ico', 'png', 'bmp')) {
        Remove-Item -Path "$env:LOCALAPPDATA\skHost.$format" -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    
    # Remove shared components
    Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\skHost.lnk" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -Force -Confirm:$false -ErrorAction SilentlyContinue
    
    Write-skSessionLog -Message "Uninstall completed." -Type "SUCCESS" -Color Green
}
