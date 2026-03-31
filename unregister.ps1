$appPath = "$PSScriptRoot\OpenCode Office Add-in.exe"
$registrationName = "OpenCodeOfficeAddin"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

function Get-OfficeManifestId {
    param(
        [string]$Path
    )

    if (!(Test-Path $Path)) {
        return $null
    }

    try {
        [xml]$manifest = Get-Content -LiteralPath $Path -Raw
        return [string]$manifest.OfficeApp.Id
    } catch {
        return $null
    }
}

function Get-DeveloperRegistrationNames {
    if (!(Test-Path $regPath)) {
        return @()
    }

    return (Get-Item -Path $regPath).Property | Where-Object { $_ -notlike "PS*" }
}

function Remove-OpenCodeRegistrations {
    param(
        [string]$ManifestPath,
        [string]$ManifestId
    )

    if (!(Test-Path $regPath)) {
        return 0
    }

    $removed = 0
    foreach ($name in Get-DeveloperRegistrationNames) {
        $value = (Get-ItemProperty -Path $regPath -Name $name -ErrorAction SilentlyContinue).$name
        $shouldRemove = $name -eq $registrationName

        if (!$shouldRemove -and $ManifestPath -and $value -eq $ManifestPath) {
            $shouldRemove = $true
        }

        if (!$shouldRemove -and $ManifestId -and $value -and (Test-Path $value)) {
            $existingManifestId = Get-OfficeManifestId -Path $value
            if ($existingManifestId -and $existingManifestId -eq $ManifestId) {
                $shouldRemove = $true
            }
        }

        if ($shouldRemove) {
            Remove-ItemProperty -Path $regPath -Name $name -ErrorAction SilentlyContinue
            $removed++
        }
    }

    if ((Get-DeveloperRegistrationNames).Count -eq 0) {
        Remove-Item -Path $regPath -Force -ErrorAction SilentlyContinue
    }

    return $removed
}

# Check if running from release (exe exists) or dev (manifest in root)
if (Test-Path $appPath) {
    $manifestPath = "$PSScriptRoot\resources\manifest.xml"
} else {
    $manifestPath = "$PSScriptRoot\manifest.xml"
}

$manifestFullPath = if (Test-Path $manifestPath) { (Resolve-Path $manifestPath).Path } else { $null }
$manifestId = if ($manifestFullPath) { Get-OfficeManifestId -Path $manifestFullPath } else { $null }

Write-Host "Removing OpenCode Office add-in registration..." -ForegroundColor Cyan

if (Test-Path $regPath) {
    $removedCount = Remove-OpenCodeRegistrations -ManifestPath $manifestFullPath -ManifestId $manifestId
    if ($removedCount -gt 0) {
        Write-Host "Removed $removedCount OpenCode registration(s)" -ForegroundColor Green
    } else {
        Write-Host "OpenCode add-in registration was not found" -ForegroundColor Gray
    }
} else {
    Write-Host "No sideloaded add-ins were registered" -ForegroundColor Gray
}

Write-Host ""
Write-Host "To re-register, run: .\register.ps1" -ForegroundColor Gray
