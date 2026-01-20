$LOG_NAME = "session.log"
$SESSION_LOG = Join-Path $PSScriptRoot $LOG_NAME
$script:LogViewerProcess = $null # Track log viewer process

# Initialize session log with header (overwrites on each launch)
$sessionLogHeader = @"
================================================================================
skHost Session Log
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
================================================================================

"@
Set-Content -Path $SESSION_LOG -Value $sessionLogHeader -Encoding utf8

Function Write-skSessionLog {
    param (
        [Parameter(Mandatory = $true)] [String] $Message,
        [Parameter(Mandatory = $false)] [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS")] [String] $Type = "INFO",
        [Parameter(Mandatory = $false)] [ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "DarkGray", "Gray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White")] [String] $Color = "Gray"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logEntry = "[$timestamp] [$Type] $Message"
    
    # Write to session log file with retry logic for file locking
    $maxRetries = 10
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        if ($retryCount -gt 0) {
            $logText = "[$timestamp] [$Type] $Message (Retry Success #$retryCount // Last Exception: $logException)"
        } else {
            $logText = $logEntry
        }
        try {
            Add-Content -Path $SESSION_LOG -Value $logText -Encoding utf8 -ErrorAction Stop
            $success = $true
        }
        catch {
            $logException = $_.Exception.Message
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Start-Sleep -Milliseconds 50
            }
        }
    }
    
    # Write to console
    Write-Host $logText -ForegroundColor $Color
    
    # Log error message to console if logging to file failed
    if (-not $success) {
        Write-Host "[$timestamp] [ERROR] Failed to write previous message to session log after $maxRetries attempts. Last Exception: $logException" -ForegroundColor Red
    }

}

Function Toggle-LogViewer {
    if ($script:LogViewerProcess -and -not $script:LogViewerProcess.HasExited) {
        # Close existing log viewer
        $script:LogViewerProcess.Kill()
        $script:LogViewerProcess = $null
        Write-skSessionLog -Message "Log viewer closed" -Type "DEBUG" -Color Gray
    } else {
        # Open new log viewer window
        $script:LogViewerProcess = Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-NoExit", "-Command", "Get-Content -Path '$SESSION_LOG' -Wait -Tail 1000" -PassThru
        Write-skSessionLog -Message "Log viewer opened" -Type "DEBUG" -Color Gray
    }
}