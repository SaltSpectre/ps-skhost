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
            Write-skSessionLog -Message "Unsupported icon format: $extension" -Type "WARNING" -Color Yellow
            return $null
        }
    }
}
