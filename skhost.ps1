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
$notifyIcon.ContextMenu.MenuItems.AddRange(@($menuItem_OpenLog, $menuItem_Exit))
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
    Write-skSessionLog -Message "========== Main Loop Started ==========" -Type "INFO" -Color Cyan

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

    Write-skSessionLog -Message "🔍 Searching for RDP, RemoteApp, and Hyper-V windows across all desktops." -Type "DEBUG" -Color Yellow
    $ActiveRdpWindows = Get-WindowsByClass -ClassName $RdpWindowClasses
    if ($ActiveRdpWindows -ne $null) {
        Write-skSessionLog -Message "✔️ Found $($ActiveRdpWindows.Count) active session(s). Caching current foreground window..." -Type "INFO" -Color Magenta
        $ForegroundWindowHandle = [User32]::GetForegroundWindow()
        $ForegroundWindow = Get-WindowInformation -Handle $ForegroundWindowHandle
        Write-skSessionLog -Message "🎯 Current foreground: [$ForegroundWindowHandle] $($ForegroundWindow.Class) - '$($ForegroundWindow.Title)'" -Type "DEBUG" -Color Magenta

        foreach ($RdpWindow in $ActiveRdpWindows) {
            $WindowHandle = $RdpWindow.Handle
            Write-skSessionLog -Message "⌛ Processing session: [$WindowHandle] $($RdpWindow.Class) - '$($RdpWindow.Title)'" -Type "DEBUG" -Color Yellow
            if ($RdpWindow.Class -like "*RAIL*" -and -not ($RemoteAppWhiteList | Where-Object { $RdpWindow.Title -like "*$_*" })) {
                Write-skSessionLog -Message "⏭️ Skipping: RemoteApp not in whitelist" -Type "DEBUG" -Color DarkGray
            }
            else {
                [User32]::SetForegroundWindow($WindowHandle) | Out-Null
                Write-skSessionLog -Message "🪄 Activated window in foreground: [$WindowHandle] '$($RdpWindow.Title)'" -Type "DEBUG" -Color Yellow
                Move-MouseCursor
                Write-skSessionLog -Message "⌨️ Sent simulated mouse input to window: [$WindowHandle] '$($RdpWindow.Title)'" -Type "SUCCESS" -Color Green
            }
        }
        Write-skSessionLog -Message "🔄 Restoring original foreground window: [$ForegroundWindowHandle] '$($ForegroundWindow.Title)'" -Type "INFO" -Color Magenta
        [User32]::SetForegroundWindow($ForegroundWindowHandle) | Out-Null
    } else {
        Write-skSessionLog -Message "✔️ No matching windows found--nothing to do." -Type "DEBUG" -Color Yellow
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
