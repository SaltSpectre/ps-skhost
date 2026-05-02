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

# Dot source logging and icon handler functions
. "$PSScriptRoot\log.ps1"
. "$PSScriptRoot\icon_handler.ps1"

Add-Type -AssemblyName System.Drawing, System.Windows.Forms

if ($Install) {
    . "$PSScriptRoot\app_handler.ps1"
    Install-skHost -ParentPath $ParentPath -AutoStart $AutoStart
    Break
}

# Uninstall logic. Does not stop currently running script.
if ($Uninstall) {
    . "$PSScriptRoot\app_handler.ps1"
    Uninstall-skHost
    Break
}

$ConfigPath = Join-Path $PSScriptRoot "config.json"
$Config = Get-Content $ConfigPath | ConvertFrom-Json
$script:localEnabled = if ($null -ne $Config.localEnabled) { [bool]$Config.localEnabled } else { $true }
$script:rdpEnabled = if ($null -ne $Config.rdpEnabled) { [bool]$Config.rdpEnabled } else { $true }

# Create system tray icon
$iconFile = Find-IconFile -BasePath $PSScriptRoot -IconName "skHost"
$SysTrayIconImage = New-IconFromFile -IconPath $iconFile

# Create the system tray icon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Text = "skHost ($(Get-Content "$PSScriptRoot\version.txt" -TotalCount 1))"
$notifyIcon.Icon = $SysTrayIconImage
$notifyIcon.ContextMenu = New-Object System.Windows.Forms.ContextMenu
$notifyIcon.add_DoubleClick({
    Toggle-LogViewer
})
$menuItem_OpenLog = New-Object System.Windows.Forms.MenuItem
$menuItem_OpenLog.Text = "Log Viewer"
$menuItem_OpenLog.add_click({
        Toggle-LogViewer
    })
$menuItem_Local = New-Object System.Windows.Forms.MenuItem
$menuItem_Local.add_click({
        $script:localEnabled = -not $script:localEnabled
        Save-skConfig
        Update-skMenuLabels
        Write-skSessionLog -Message "Local activity $($(if ($script:localEnabled) { 'enabled' } else { 'disabled' }))" -Type "INFO" -Color Cyan
    })
$menuItem_Rdp = New-Object System.Windows.Forms.MenuItem
$menuItem_Rdp.add_click({
        $script:rdpEnabled = -not $script:rdpEnabled
        Save-skConfig
        Update-skMenuLabels
        Write-skSessionLog -Message "RDP activity $($(if ($script:rdpEnabled) { 'enabled' } else { 'disabled' }))" -Type "INFO" -Color Cyan
    })
$menuItem_Exit = New-Object System.Windows.Forms.MenuItem
$menuItem_Exit.Text = "Exit"
$menuItem_Exit.add_click({
        # Exit logic
        $skTimer.Stop()
        $skTimer.Dispose()
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
        Start-Sleep 1
        [System.Windows.Forms.Application]::Exit()
    })
$notifyIcon.ContextMenu.MenuItems.AddRange(@($menuItem_OpenLog, $menuItem_Local, $menuItem_Rdp, $menuItem_Exit))
$notifyIcon.Visible = $true

#region CoreLogic and Helper Functions
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $(Get-Content "$PSScriptRoot\user32.cs" -Raw)
Add-Type -TypeDefinition $(Get-Content "$PSScriptRoot\mouse.cs" -Raw)

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
    Write-skSessionLog -Message "Invalid keystroke format in config.json. Defaulting to Ctrl+Shift+F15." -Type "WARNING" -Color Yellow
    $skHostKeystroke = "^+{F15}"
}

# Read loop interval from config file or default to 4 minutes
if ($Config.loopIntervalSeconds -and $Config.loopIntervalSeconds -is [int] -and $Config.loopIntervalSeconds -gt 0) {
    $skHostInterval = $Config.loopIntervalSeconds * 1000  # Convert to milliseconds
} else {
    Write-skSessionLog -Message "Invalid loop interval in config.json. Defaulting to 4 minutes." -Type "WARNING" -Color Yellow
    $skHostInterval = 240000  # Default to 4 minutes
}

# Update the tooltip to show the configured keystroke and interval
Function Save-skConfig {
    $configOutput = [ordered]@{
        skHostKeystroke     = if ($Config.skHostKeystroke) { $Config.skHostKeystroke } else { $skHostKeystroke }
        loopIntervalSeconds = if ($Config.loopIntervalSeconds) { $Config.loopIntervalSeconds } else { [int]($skHostInterval / 1000) }
        localEnabled        = $script:localEnabled
        rdpEnabled          = $script:rdpEnabled
    }

    $configOutput | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding utf8
    $script:Config = Get-Content $ConfigPath | ConvertFrom-Json
}

Function Update-skMenuLabels {
    $menuItem_Local.Text = if ($script:localEnabled) { "Disable Local" } else { "Enable Local" }
    $menuItem_Rdp.Text = if ($script:rdpEnabled) { "Disable RDP" } else { "Enable RDP" }
    $notifyIcon.Text = "skHost ($(Get-Content "$PSScriptRoot\version.txt" -TotalCount 1))`nKeystroke: $skHostKeystroke`nInterval: $($skHostInterval / 1000) seconds"
}

Update-skMenuLabels

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

Function Get-ChildWindows {
    [CmdletBinding()]
    param ( [Parameter(Mandatory = $true)] [IntPtr] $ParentHandle )

    $EnumeratedWindows = New-Object System.Collections.ArrayList
    $enumWindows = {
        param($hWnd, $lParam)
        $windowInfo = Get-WindowInformation -Handle $hWnd
        $EnumeratedWindows.Add($windowInfo) | Out-Null
        return $true
    }
    [User32]::EnumChildWindows($ParentHandle, $enumWindows, [IntPtr]::Zero) | Out-Null
    return $EnumeratedWindows
}

Function New-MousePointLParam {
    param (
        [Parameter(Mandatory = $true)] [int] $X,
        [Parameter(Mandatory = $true)] [int] $Y
    )

    return [IntPtr](($Y -shl 16) -bor ($X -band 0xFFFF))
}

Function Send-BackgroundMouseMove {
    param (
        [Parameter(Mandatory = $true)] $Handle,
        [Parameter(Mandatory = $true)] [int] $X,
        [Parameter(Mandatory = $true)] [int] $Y
    )

    [User32]::PostMessage($Handle, 0x0200, [IntPtr]::Zero, (New-MousePointLParam -X $X -Y $Y)) | Out-Null
}

Function Send-BackgroundMouseJiggle {
    param ( [Parameter(Mandatory = $true)] $Handle )

    Send-BackgroundMouseMove -Handle $Handle -X 10 -Y 10
    Start-Sleep -Milliseconds 100
    Send-BackgroundMouseMove -Handle $Handle -X 11 -Y 10
    Start-Sleep -Milliseconds 100
    Send-BackgroundMouseMove -Handle $Handle -X 10 -Y 10
}

Function Test-WindowIsForegroundRoot {
    param (
        [Parameter(Mandatory = $true)] $WindowHandle,
        [Parameter(Mandatory = $true)] $ForegroundWindowHandle
    )

    $ForegroundRoot = [User32]::GetAncestor($ForegroundWindowHandle, 2)
    return $ForegroundRoot -eq $WindowHandle
}

Function Invoke-BackgroundSessionMouseInput {
    param ( [Parameter(Mandatory = $true)] $Window )

    if ($Window.Class -eq "TscShellContainerClass") {
        Send-BackgroundMouseJiggle -Handle $Window.Handle
        $ChildWindows = @(Get-ChildWindows -ParentHandle $Window.Handle)
        $InputWindows = @($ChildWindows | Where-Object { $_.Class.Trim() -eq "IHWindowClass_rdclientax" })

        if ($InputWindows -and $InputWindows.Count -gt 0) {
            foreach ($InputWindow in $InputWindows) {
                Send-BackgroundMouseJiggle -Handle $InputWindow.Handle
            }
            Write-skSessionLog -Message "✔️ Sent background mouse input to RDP top-level and $($InputWindows.Count) child input target(s): [$($Window.Handle)] '$($Window.Title)'" -Type "SUCCESS" -Color Green
        }
        else {
            $ChildClasses = ($ChildWindows | Select-Object -ExpandProperty Class -Unique) -join ", "
            Write-skSessionLog -Message "⚠️ Sent background mouse input to RDP top-level only; no child input target found: [$($Window.Handle)] '$($Window.Title)'. Child classes: $ChildClasses" -Type "WARNING" -Color Yellow
        }

        return $true
    }

    if ($Window.Class -eq "RAIL_WINDOW") {
        Send-BackgroundMouseJiggle -Handle $Window.Handle
        Write-skSessionLog -Message "✔️ Sent background mouse input to RemoteApp window: [$($Window.Handle)] '$($Window.Title)'" -Type "SUCCESS" -Color Green
        return $true
    }

    return $false
}

$RdpWindowClasses = @(
    "TscShellContainerClass", # MSTSC/MSRDC
    "RAIL_WINDOW" # RemoteApp
    # "WindowsForms10.Window.8.app.0.aa0c13_r6_ad1" # Hyper-V Console
)

Function Invoke-skLogic {
    Write-skSessionLog -Message "========== Main Loop Started ==========" -Type "INFO" -Color Cyan

    if ($script:localEnabled) {
        # Create temporary window ("skSink"), activate it, send keystroke via SendKeys, then destroy it
        Write-skSessionLog -Message "🪄 Creating temporary skSink window for keystroke activity" -Type "DEBUG" -Color Yellow
        $hInstance = [User32]::GetModuleHandle($null)
        $skSink = [User32]::CreateWindowEx(0, "Static", "", 0x10000000, -10000, -10000, 1, 1, [IntPtr]::Zero, [IntPtr]::Zero, $hInstance, [IntPtr]::Zero)
        
        if ($skSink -ne [IntPtr]::Zero) {
            Write-skSessionLog -Message "✔️ Temporary skSink window created successfully [Handle: $skSink]" -Type "SUCCESS" -Color Green
            
            # Activate the temporary window
            [User32]::SetForegroundWindow($skSink) | Out-Null
            Write-skSessionLog -Message "✔️ Temporary skSink window activated" -Type "SUCCESS" -Color Green
            
            # Send keystroke using SendKeys (goes to active window)
            [System.Windows.Forms.SendKeys]::SendWait("$skHostKeystroke")
            Write-skSessionLog -Message "✔️ Keystroke sent to skSink: $skHostKeystroke" -Type "SUCCESS" -Color Green
            
            # Destroy the temporary window
            [User32]::DestroyWindow($skSink) | Out-Null
            Write-skSessionLog -Message "✔️ Temporary skSink window destroyed" -Type "SUCCESS" -Color Green
        } else {
            Write-skSessionLog -Message "❌ Failed to create temporary skSink window" -Type "ERROR" -Color Red
        }
    } else {
        Write-skSessionLog -Message "⏭️ Local activity disabled--skipping skSink keystroke." -Type "INFO" -Color Cyan
    }

    if ($script:rdpEnabled) {
        Write-skSessionLog -Message "🔍 Searching for RDP and RemoteApp windows across all desktops." -Type "DEBUG" -Color Yellow
        $ActiveRdpWindows = @(Get-WindowsByClass -ClassName $RdpWindowClasses)
        if ($ActiveRdpWindows -and $ActiveRdpWindows.Count -gt 0) {
            Write-skSessionLog -Message "✔️ Found $($ActiveRdpWindows.Count) active session(s). Checking foreground window..." -Type "INFO" -Color Magenta
            $ForegroundWindowHandle = [User32]::GetForegroundWindow()
            $ForegroundWindow = Get-WindowInformation -Handle $ForegroundWindowHandle
            Write-skSessionLog -Message "🎯 Current foreground: [$ForegroundWindowHandle] $($ForegroundWindow.Class) - '$($ForegroundWindow.Title)'" -Type "DEBUG" -Color Magenta

            foreach ($RdpWindow in $ActiveRdpWindows) {
                $WindowHandle = $RdpWindow.Handle
                Write-skSessionLog -Message "⌛ Processing session: [$WindowHandle] $($RdpWindow.Class) - '$($RdpWindow.Title)'" -Type "DEBUG" -Color Yellow

                if ([User32]::IsIconic($WindowHandle)) {
                    Write-skSessionLog -Message "⏭️ Skipping minimized session: [$WindowHandle] '$($RdpWindow.Title)'" -Type "DEBUG" -Color DarkGray
                }
                elseif (Test-WindowIsForegroundRoot -WindowHandle $WindowHandle -ForegroundWindowHandle $ForegroundWindowHandle) {
                    Move-MouseCursor
                    Write-skSessionLog -Message "✔️ Sent foreground mouse input to active session: [$WindowHandle] '$($RdpWindow.Title)'" -Type "SUCCESS" -Color Green
                }
                else {
                    $backgroundInputSent = Invoke-BackgroundSessionMouseInput -Window $RdpWindow
                    if (-not $backgroundInputSent) {
                        Write-skSessionLog -Message "⏭️ Skipping: no background input target found for [$WindowHandle] '$($RdpWindow.Title)'" -Type "WARNING" -Color Yellow
                    }
                }
            }
        } else {
            Write-skSessionLog -Message "✔️ No matching windows found--nothing to do." -Type "DEBUG" -Color Yellow
        }
    } else {
        Write-skSessionLog -Message "⏭️ RDP activity disabled--skipping RDP and RemoteApp windows." -Type "INFO" -Color Cyan
    }
    Write-skSessionLog -Message "========== Main Loop Completed ==========" -Type "INFO" -Color Cyan
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
