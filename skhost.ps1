<#
.PARAMETER Install
    Copies the script to %LocalAppData%\skHost.ps1 and creates a Start Menu shortcut.
    Does not automatically start the application after installation.
.PARAMETER AutoStart
    Implies -Install and additionally creates a registry entry in HKCU Run to automatically 
    start skHost when Windows starts. Uses conhost to ensure compatibility with hidden 
    window execution.
.PARAMETER Uninstall
    Removes the installed script, Start Menu shortcut, custom icon, and auto-start registry 
    entry. Does not terminate any currently running instances.
.NOTES
    - Additional files are required to use this script.
    - See the GitHub repository for more information and documentation.
.LINK
    https://github.com/SaltSpectre/ps-skhost
#>

Param( 
    [Parameter(Mandatory = $false)] [Switch] $Install,  
    [Parameter(Mandatory = $false)] [Switch] $AutoStart,
    [Parameter(Mandatory = $false)] [Switch] $Uninstall 
)

if ($AutoStart) { $Install = $true }
if ($Uninstall) { $Install = $false }

$ParentPath = Split-Path ($MyInvocation.MyCommand.Path) -Parent

Add-Type -AssemblyName System.Drawing, System.Windows.Forms

Function Test-RegistryValue ($regkey, $name) {
    if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
        $true
    }
    else {
        $false
    }
    # https://adamtheautomator.com/powershell-get-registry-value/
}

Function Find-IconFile {
    param([string]$BasePath, [string]$IconName)
    
    foreach ($format in @('ico', 'png', 'bmp')) {
        $iconPath = Join-Path $BasePath "$($IconName -replace '\.[^.]*$', '').$format"
        if (Test-Path $iconPath) {
            return $iconPath
        }
    }
    return $null
}

Function New-IconFromFile {
    param([string]$IconPath)
    
    if (-not (Test-Path $IconPath)) { return $null }
    
    $extension = [System.IO.Path]::GetExtension($IconPath).ToLower()
    
    switch ($extension) {
        '.ico' { return [System.Drawing.Icon]::new($IconPath) }
        { $_ -in '.png', '.bmp' } {
            $bitmap = [System.Drawing.Bitmap]::new($IconPath)
            $hIcon = $bitmap.GetHicon()
            $icon = [System.Drawing.Icon]::FromHandle($hIcon)
            $bitmap.Dispose()
            return $icon
        }
        default {
            Write-Warning "Unsupported icon format: $extension"
            return $null
        }
    }
}

# Install logic. Does not start after install, even if autostart is specified.
Function Install-skHost {
    $InstallPath = "$env:LOCALAPPDATA\SaltSpectre\ps-skhost"
    $ManifestPath = Join-Path $ParentPath "manifest.txt"
    
    # Create installation directory
    if (-not (Test-Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
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
                Write-Host "Copied: $file" -ForegroundColor Green
            } else {
                Write-Error "File not found: $file"
                return
            }
        }
    } else {
        Write-Error "manifest.txt not found. Installation cannot proceed without manifest file."
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

if ($Install) {
    Install-skHost
    Break
}

# Uninstall logic. Does not stop currently running script.
if ($Uninstall) {
    # Remove new installation directory structure
    $NewInstallPath = "$env:LOCALAPPDATA\SaltSpectre\ps-skhost"
    if (Test-Path $NewInstallPath) {
        Remove-Item -Path $NewInstallPath -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Removed: $NewInstallPath" -ForegroundColor Green
    }
    
    # Clean up old installation files (backward compatibility)
    Remove-Item -Path "$env:LOCALAPPDATA\skHost.ps1" -Force -Confirm:$false -ErrorAction SilentlyContinue
    foreach ($format in @('ico', 'png', 'bmp')) {
        Remove-Item -Path "$env:LOCALAPPDATA\skHost.$format" -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    
    # Remove shared components
    Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\skHost.lnk" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -Force -Confirm:$false -ErrorAction SilentlyContinue
    
    Write-Host "Uninstall completed." -ForegroundColor Green
    Break
}

# Create system tray icon
$iconFile = Find-IconFile -BasePath $PSScriptRoot -IconName "skHost"
$SysTrayIconImage = New-IconFromFile -IconPath $iconFile

# Create the system tray icon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Text = "skHost ($(Get-Content "$PSScriptRoot\version.txt" -TotalCount 1))"
$notifyIcon.Icon = $SysTrayIconImage
$notifyIcon.ContextMenu = New-Object System.Windows.Forms.ContextMenu
$menuItem = New-Object System.Windows.Forms.MenuItem
$menuItem.Text = "Exit"
$menuItem.add_click({
        # Exit logic
        $skTimer.Stop()
        $skTimer.Dispose()
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
        Start-Sleep 1
        [System.Windows.Forms.Application]::Exit()
    })
$notifyIcon.ContextMenu.MenuItems.AddRange($menuItem)
$notifyIcon.Visible = $true

#region CoreLogic and Helper Functions
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $(Get-Content "$PSScriptRoot\user32.cs" -Raw)
Add-Type -TypeDefinition $(Get-Content "$PSScriptRoot\mouse.cs" -Raw)

# Load configuration file
$Config = Get-Content "$PSScriptRoot\config.json" | ConvertFrom-Json

# Read keystroke from config file or default to Ctrl+Shift+F15
if ($Config.skHostKeystroke -and $Config.skHostKeystroke.Trim() -ne "") {
    $skHostKeystroke = $Config.skHostKeystroke.Trim()
} else {
    $skHostKeystroke = "^+{F15}"  # Ctrl+Shift+F15
}   

# Validate keystroke and use Ctrl+Shift+F15 if invalid
Try {
    [System.Windows.Forms.SendKeys]::SendWait($skHostKeystroke)
} Catch {
    Write-Warning "Invalid or unconfigured keystroke format in config.json. Defaulting to Ctrl+Shift+F15."
    $skHostKeystroke = "^+{F15}"
}

# Read loop interval from config file or default to 4 minutes
if ($Config.loopIntervalSeconds -and $Config.loopIntervalSeconds -is [int] -and $Config.loopIntervalSeconds -gt 0) {
    $skHostInterval = $Config.loopIntervalSeconds * 1000  # Convert to milliseconds
} else {
    Write-Warning "Invalid or unconfigured loop interval format in config.json. Defaulting to 4 minutes."
    $skHostInterval = 240000  # Default to 4 minutes
}

# Update the tooltip to show the configured keystroke and interval
$notifyIcon.Text = "skHost ($(Get-Content "$PSScriptRoot\version.txt" -TotalCount 1))`nKeystroke: $skHostKeystroke`nInterval: $($skHostInterval / 1000) seconds"

Function Move-MouseCursor {
    [Mouse]::MoveMouse()
}

Function Get-WindowInformation {
    [CmdletBinding()] 
    param ( [Parameter(Mandatory = $true)] $Handle )
    
    $title = New-Object System.Text.StringBuilder 256
    $class = New-Object System.Text.StringBuilder 256
    
    [User32]::GetWindowText($Handle, $title, $title.Capacity) | Out-Null
    [User32]::GetClassName($Handle, $class, $class.Capacity) | Out-Null
    
    return New-Object PSObject -Property @{
        Handle = $Handle
        Class  = $class.ToString()
        Title  = $title.ToString()
    }
}

Function Get-WindowsByClass {
    # Use this function to find all active windows of a certain Window class by specifying the -ClassName parameter
    [CmdletBinding()] 
    param ( [Parameter(Mandatory = $true)] $ClassName )

    $EnumeratedWindows = New-Object System.Collections.ArrayList
    $enumWindows = {
        param($hWnd, $lParam)
        $windowInfo = Get-WindowInformation -Handle $hWnd
        if ($windowInfo.Title -and ($ClassName -contains $windowInfo.Class)) {
            $EnumeratedWindows.Add($windowInfo) | Out-Null
        }
        return $true
    }
    [User32]::EnumWindows($enumWindows, [IntPtr]::Zero) | Out-Null
    return $EnumeratedWindows
}

# Things get weird if every window in a RemoteApp session is called to the foreground. Define specific windows.
$RemoteAppWhiteList = $Config.remoteAppWhitelist

$RdpWindowClasses = @(
    "TscShellContainerClass", # MSTSC/MSRDC
    "RAIL_WINDOW", # RemoteApp
    "WindowsForms10.Window.8.app.0.aa0c13_r6_ad1" #Hyper-V Console
)

Function Invoke-skLogic {
    Write-Host "`nDEBUG: Entering main loop at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan

    # Create temporary window ("skSink"), activate it, send keystroke via SendKeys, then destroy it
    Write-Host "DEBUG: Creating temporary window for keystroke activity" -ForegroundColor Yellow
    $hInstance = [User32]::GetModuleHandle($null)
    $skSink = [User32]::CreateWindowEx(0, "Static", "", 0x10000000, -10000, -10000, 1, 1, [IntPtr]::Zero, [IntPtr]::Zero, $hInstance, [IntPtr]::Zero)
    
    if ($skSink -ne [IntPtr]::Zero) {
        Write-Host "DEBUG: Temporary window created with handle: $skSink" -ForegroundColor Green
        
        # Activate the temporary window
        [User32]::SetForegroundWindow($skSink) | Out-Null
        Write-Host "DEBUG: skSink temporary window activated" -ForegroundColor Green
        
        # Send keystroke using SendKeys (goes to active window)
        [System.Windows.Forms.SendKeys]::SendWait("$skHostKeystroke")
        Write-Host "DEBUG: Sent keystroke to skSink: $skHostKeystroke" -ForegroundColor Green
        
        # Destroy the temporary window
        [User32]::DestroyWindow($skSink) | Out-Null
        Write-Host "DEBUG: skSink temporary window destroyed" -ForegroundColor Green
    } else {
        Write-Host "DEBUG: Failed to create skSink temporary window" -ForegroundColor Red
    }

    Write-Host "DEBUG: Searching for RDP, RemoteApp, and Hyper-V windows across all desktops." -ForegroundColor Yellow
    $ActiveRdpWindows = Get-WindowsByClass -ClassName $RdpWindowClasses
    if ($ActiveRdpWindows -ne $null) {
        Write-Host "DEBUG: Active sessions found. Caching foreground window information:" -ForegroundColor Magenta
        $ForegroundWindowHandle = [User32]::GetForegroundWindow()
        $ForegroundWindow = Get-WindowInformation -Handle $ForegroundWindowHandle
        Write-Host "DEBUG: [$ForegroundWindowHandle]: $($ForegroundWindow.Class) - $($ForegroundWindow.Title)" -ForegroundColor Magenta

        foreach ($RdpWindow in $ActiveRdpWindows) {
            $WindowHandle = $RdpWindow.Handle
            Write-Host "DEBUG: [$WindowHandle] Found: $($RdpWindow.Class) - $($RdpWindow.Title)" -ForegroundColor Yellow
            if ($RdpWindow.Class -like "*RAIL*" -and -not ($RemoteAppWhiteList | Where-Object { $RdpWindow.Title -like "*$_*" })) {
                Write-Host "DEBUG: RemoteApp not in the whitelist. Ignoring." -ForegroundColor DarkGray
            }
            else {
                [User32]::SetForegroundWindow($WindowHandle) | Out-Null
                Write-Host "DEBUG: [$WindowHandle] Activated window in foreground." -ForegroundColor Yellow
                Move-MouseCursor
                Write-Host "DEBUG: [$WindowHandle] Sent simulated mouse input." -ForegroundColor Green
            }
        }
        Write-Host "DEBUG: [$ForegroundWindowHandle] Restoring foreground window: $($ForegroundWindow.Class) - $($ForegroundWindow.Title)" -ForegroundColor Magenta
        [User32]::SetForegroundWindow($ForegroundWindowHandle) | Out-Null

    }
    else {
        Write-Host "DEBUG: No matching windows found--nothing to do." -ForegroundColor Yellow
    }
    Write-Host "DEBUG: Main loop completed at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
}

# Create a Windows Forms Timer for the main logic
$skTimer = New-Object System.Windows.Forms.Timer
$skTimer.Interval = $skHostInterval  # Use the configured interval
$skTimer.Add_Tick({
        Invoke-skLogic
    })
Invoke-skLogic # Immediately run on script execution
$skTimer.Start() # Start the Timer

# Create an application context to keep the system tray icon interactive
$AppContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($AppContext)
