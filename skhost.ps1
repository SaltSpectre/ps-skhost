# skHost.ps1
# -----------
# 
# New in version 3.2
#    - Refactored to use a Windows Forms Timer instead of a PowerShell Job
#       - Reducing to a single PowerShell process instead of a job reduces resource usage and improves responsiveness.
#    - Changed the key used for simulated input from F15 to F16 to avoid conflicts with Linux terminals including SSH and WSL in Windows Terminal.
# New in version 3.1:
#    -Added support for RemoteApp windows. Due to the volume of windows per RemoteApp session, limited to a whitelist only.
#    -Enhanced debug logging
# New in version 3.0.1:
#    -Added support for Hyper-V keep-alive in Basic and Enhanced Session modes
#    -Added PS Job debugging mode for development
# New in version 3.0:
#    -Stability improvements
#    -Added mouse movement to keep all active RDP windows, including Azure Virtual Desktop, from going idle. RDP window cannot be minimized.


Param( 
    [Parameter(Mandatory = $false)] [Switch] $Install, # Copies the script to AppData and creates a Start Menu shortcut
    [Parameter(Mandatory = $false)] [Switch] $AutoStart, # $Install + creates an auto-run entry in HKCU
    [Parameter(Mandatory = $false)] [Switch] $Uninstall   # Deletes the script from AppData and removes shortcut + auto-run
)

if ($AutoStart) { $Install = $true }
if ($Uninstall) { $Install = $false }

$ParentPath = Split-Path ($MyInvocation.MyCommand.Path) -Parent

Add-Type -AssemblyName System.Drawing, System.Windows.Forms

# https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-extracticonexw
# https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-destroyicon
# This function allows for the extraction of an icon at an index from a system icon library. Also helps us dispose of handles.
Add-Type -TypeDefinition '
    using System;
    using System.Runtime.InteropServices;

    public class Shell32_Extract {

        [DllImport(
            "Shell32.dll",
            EntryPoint = "ExtractIconExW",
            CharSet  = CharSet.Unicode,
            ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall
        )]

        public static extern int ExtractIconEx(
            string lpszFile,
            int iconIndex,
            out IntPtr phiconLarge,
            out IntPtr phiconSmall,
            int nIcons
        );
    }

    public class User32_DestroyIcon {
        
        [DllImport(
            "User32.dll",
            EntryPoint = "DestroyIcon"
        )]

        public static extern int DestroyIcon(IntPtr hIcon);
    }
'

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
    Copy-Item -Path $PSCommandPath -Destination "$env:LOCALAPPDATA\skHost.ps1" -Force -Confirm:$false
    if ($AutoStart) {
        if (Test-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" "skHost") {
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -Force -Confirm:$false
        }
        # Specify to use conhost in case Windows Terminal is set as default as it does not support WindowStyle Hidden like conhost
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -PropertyType String -Value "conhost powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File %LocalAppData%\skHost.ps1" -Force -Confirm:$false
    }

    # Create Start Menu Shortcut
    $WScript = New-Object -ComObject ("WScript.Shell")
    $Shortcut = $Wscript.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\skHost.lnk")
    $Shortcut.TargetPath = "$env:SystemRoot\System32\Conhost.exe" 
    $Shortcut.Arguments = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File %LocalAppData%\skHost.ps1'
    # Use my fancy icon if it is located in AppData or the script's parent path, otherwise fallback to star icon from Shell32.dll
    if ((Test-Path "$ParentPath\skHost.ico") -and !(Test-path "$env:LOCALAPPDATA\skhost.ico")) { Copy-Item "$ParentPath\skHost.ico" "$env:LOCALAPPDATA\skHost.ico" }
    if (Test-path "$env:LOCALAPPDATA\skhost.ico") {
        $Shortcut.IconLocation = "$env:LOCALAPPDATA\skhost.ico"
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
    Remove-Item -Path "$env:LOCALAPPDATA\skHost.ps1" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:LOCALAPPDATA\skHost.ico" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\skHost.lnk" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "skHost" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Break
}

# Use my fancy icon if it is located in AppData, otherwise fallback to star icon from Shell32.dll.
if (Test-Path "$env:LOCALAPPDATA\skHost.ico") {
    $SysTrayIconImage = New-Object System.Drawing.Icon ("$env:LOCALAPPDATA\skHost.ico")
}
else {
    [System.IntPtr] $icoHandleSm = 0
    [System.IntPtr] $icoHandleLg = 0
    [void] [Shell32_Extract]::ExtractIconEx("%systemroot%\system32\shell32.dll", 43, [ref] $icoHandleLg, [ref] $icoHandleSm, 1)
    $SysTrayIconImage = [System.Drawing.Icon]::FromHandle($icoHandleSm)
}

# Create the system tray icon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Text = "skHost"
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
        [System.Windows.Forms.Application]::Exit()
    })
$notifyIcon.ContextMenu.MenuItems.AddRange($menuItem)
$notifyIcon.Visible = $true

#region CoreLogic and Helper Functions
Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition '
    using System;
    using System.Runtime.InteropServices;
    public static class User32 {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc enumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
'

Function Move-MouseCursor {
    Add-Type -TypeDefinition '
    using System;
    using System.Runtime.InteropServices;
    public static class Mouse {
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT {
        public int type;
        public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll")]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT {
        public int x;
        public int y;
        }

        const int INPUT_MOUSE = 0;
        const int MOUSEEVENTF_MOVE = 0x0001;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        const int SM_CXSCREEN = 0;
        const int SM_CYSCREEN = 1;

        public static void MoveMouse() {
        INPUT Input = new INPUT();
        POINT CurrentPosition;

        GetCursorPos(out CurrentPosition);

        Input.type = INPUT_MOUSE;
        Input.mi.dwFlags = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE;

        Input.mi.dx = (int)(CurrentPosition.x * (65536.0f / GetSystemMetrics(SM_CXSCREEN)));
        Input.mi.dy = (int)(CurrentPosition.y * (65536.0f / GetSystemMetrics(SM_CYSCREEN)));

        SendInput(1, new INPUT[] { Input }, Marshal.SizeOf(typeof(INPUT)));
        }
    }
'
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

# Things get weird if every window in a RemoteApp session is called to the foreground. Define specific windows. Alternatively use skRAHelper mode.
# TODO: Actually implement a skRAHelper mode
$RemoteAppWhiteList = @(
    "Teams",
    "skRAHelper"
)

$RdpWindowClasses = @(
    "TscShellContainerClass", # MSTSC/MSRDC
    "RAIL_WINDOW", # RemoteApp
    "WindowsForms10.Window.8.app.0.aa0c13_r6_ad1" #Hyper-V Console
)

Function Invoke-skLogic {
    Write-Host "`nDEBUG: Executing logic at $(Get-Date -Format "yyyy-MM-dd HH:mm")" -ForegroundColor Cyan

    [System.Windows.Forms.SendKeys]::SendWait("{F16}")    # Presses the F16 key
    Write-Host "DEBUG: Sent simulated keypress to the host: {F16}" -ForegroundColor Green

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
}

# Create a Windows Forms Timer for the main logic
$skTimer = New-Object System.Windows.Forms.Timer
$skTimer.Interval = 120000  # 120,000 milliseconds = 2 minutes
$skTimer.Add_Tick({
        Invoke-skLogic
    })
Invoke-skLogic # First run immediately
$skTimer.Start() # Start the Timer

# Create an application context to keep the system tray icon interactive
$AppContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($AppContext)
