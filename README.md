# skHost: An idle prevention tool for local, RDP, RemoteApp, and Hyper-V sessions written entirely in PowerShell for maximum portability.

## Description
skHost is a Windows tool that runs in the background to keep a system alive. It does this by simulating keyboard input (Default: F15) and mouse movement every 4 minutes (240 seconds). The mouse cursor does not visibly move but the input is registered by the system. *New: Simulated keyboard input is now sent to an ephemeral window ("skSink") to reduce interference with other applications, especially WSL & SSH sessions to Linux hosts.*

In addition to keeping a physical machine session active, skHost will also bring forward all non-minimized RDP, (whitelisted) RemoteApp, and Hyper-V VM windows to simulate mouse movement at the same interval. This keeps these sessions from timing out. For RemoteApp windows, skHost uses a configurable whitelist to determine which windows to activate.
 
skHost is designed to run on the **client side** and is **transparent to the RDP or Hyper-V session**, requiring no installation or configuration on the remote host.

skHost runs in a background PowerShell process that exists in a hidden window. It creates a system tray icon that is used for a visual cue that the process is running and provides an interaction point in order to terminate the process without having to resort to the Task Manager. Ideally, skHost never gets in the way of your work and is never accidentally terminated when closing extraneous windows on your desktop.

### Important note
**You are solely responsible for deciding whether and how to use this tool.** This was not created to help people avoid work while appearing active when away from their computer. Rather, I developed it because I work across multiple systems simultaneously throughout the day and wanted to maintain visibility for my teams to facilitate communication by keeping my Teams status active.

I make no claim that this tool is transparent to your IT administration teams, and I make no guarantee that you will not be fired, sanctioned, or otherwise disciplined for it's use in a work environment. **The responsibility is yours to determine whether use of this tool is appropriate for you and whether its use is ethical.** When in doubt, you should either seek permission or just not use it.

## System Requirements
- Windows 10+
- Windows PowerShell 5.0 or 5.1
- PowerShell execution policy must be set to Unrestricted or Bypass. Alternatively you can sign the script yourself.
- PowerShell Constrained Language Mode is not supported and will prevent the script from running.

*Note: This version is incompatible with newer versions of PowerShell (aka PowerShell Core), however Windows 10 and Windows 11 both ship with compatible versions of Windows PowerShell that are retained even if one of the newer versions is installed. Ensure you are using `PowerShell.exe` and not `pwsh.exe`.*

## Usage Instructions
skHost can operate in stand-alone or installed modes. The default is stand-alone.

Clone this repository using `git clone https://github.com/SaltSpectre/ps-skhost.git` or download the latest source code zip from the Releases page.

### Preparing your environment
Prior to running the script, you must update your PowerShell execution policy. Learn more about [PowerShell Execution Policies](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1).

Supported execution policies are Unrestricted or Bypass. Please read the link above to make an informed decision about which to use, ensuring you understand the implications of your choice.

Example:
```powershell
Set-ExecutionPolicy Bypass
```

### Stand-alone Mode
Stand-alone mode just requires you to run the script from the command line like so

```powershell
.\skhost.ps1
```

Your indicator that the process is running is a system tray icon. This icon may be hidden in the overflow area by default. I recommend dragging it out of the overflow and onto the system tray so that it will always be visible.

### Installed Mode

#### Installation
Installed mode copies the script and all required files as specified in `manifest.txt` to `%LOCALAPPDATA%\SaltSpectre\ps-skhost` and creates a Start Menu shortcut. The installation will fail if `manifest.txt` is missing or if any files listed in it cannot be found.

```powershell
.\skhost.ps1 -Install
```

Once installed, the script will not launch automatically. Simply navigate to your Start Menu, search for skHost, and launch it like you would any other app. Look for the custom icon in the System Tray or overflow area to ensure it is running. See the Stand-alone mode section for more info.

#### Auto-start Mode
In installed mode, you must manually execute the script each time you want to use it. Alternatively, you can use the `autostart` parameter to create an entry for skHost in `HKCU:\Software\Microsoft\Windows\CurrentVersion\Run`.

Using the `autostart` parameter implies `install`, so only one is necessary.

```powershell
.\skhost.ps1 -AutoStart
```

The only difference in this mode is the auto start entry in the registry which will take effect on the next login. A restart *is not required* to simply run the app. You can find it in your Start Menu under skHost.

#### Uninstallation
Since skhost is only "installed" by copying files to `%LOCALAPPDATA%\SaltSpectre\ps-skhost` and making Start Menu entries, it cannot be uninstalled via traditional methods. To uninstall, execute the script with the `-Uninstall` parameter either from your installed copy or from the original copy in your working directory.

This parameter will remove the shortcut and autostart entries as applicable and delete all files from the installation directory. For backward compatibility, it will also clean up files from the old installation location.

```powershell
.\skhost.ps1 -Uninstall
```

### Changing the keystroke
By default it will use Ctrl+Shift+F15 as the keystroke that is sent to the system. I have experimented with several different keystrokes and found many to be problematic. It is impossible to identify a universally safe keystroke, so if you find this to be problematic, you can change the keystroke by modifying the `skHostKeystroke` variable in `config.json`. The script will validate the keystroke value and revert back to Ctrl+Shift+F15 if the configured value is invalid.

For a list of all possible keystrokes, refer to the [SendKeys Class documentation on Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys).

### Changing the execution interval
By default the script will invoke the keep-alive logic every 240 seconds (4 minutes). This can be configured in `config.json` via the `loopIntervalSeconds` property. The value must be in seconds and must be a positive integer value. The script will validate the value and revert back to 240 seconds if the configured value is invalid.

### Required Files and Configuration

skHost requires the following files to operate properly:
- `skhost.ps1` - The main script
- `config.json` - Contains the RemoteApp whitelist configuration
- `user32.cs` - C# definitions for Windows API functions
- `mouse.cs` - C# definitions for mouse movement simulation
- `manifest.txt` - List of files to be included in installation
- `version.txt` - Used for the versioning string for the systray icon tooltip

#### RemoteApp Whitelist
By default, skHost will not activate every RemoteApp window to prevent interference with specific applications. You can modify the whitelist in `config.json` to include the RemoteApp windows you want to keep active. `skRAHelper` is included in the default configuration and is reserved for future use.

#### Custom Icon
A custom icon file can be included in the same directory as the script to display in the system tray. The icon must be in `.ico` format.

### Known Issues
- This script will not work with [Constrained Language Mode](https://devblogs.microsoft.com/powershell/powershell-constrained-language-mode/). CLM and this script are wholly incompatible, and there are no plans to support it in the future.
- When using this script to keep an RDP, RemoteApp window, and/or Hyper-V session active, the window ***cannot be minimized***. It does not have to be the focused window as the script will grab focus momentarily, but it will not restore a remote session window for focus. If minimizing the window is something you feel is a must, consider using a virtual desktop (`Win + Tab`, not a VM) instead. When windows are running on a separate virtual desktop, the logic still works, but the window will not flash or take up space on the current desktop.
- For RDP, RemoteApp, and Hyper-V, the script **MUST** be running on the host (aka your computer). The session will not remain active if the script is running in the remote session.
- RDP, RemoteApp, and Hyper-V windows will flash when focus is grabbed, however this will be a momentary disruption.

---

## License

This project is proudly open source under the MIT License. Please refer to `LICENSE` for licensing information.

## Contributing

Got an idea? Open an Issue or fork me and make a pull request!
