$appPath = "$PSScriptRoot\OpenCode Office Add-in.exe"
$registrationName = "OpenCodeOfficeAddin"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$userDataRegPath = "HKCU:\Software\OpenCode\OfficeAddin"
$userDataRegValue = "UserDataDir"

function Get-InstalledUserDataDir {
    $configured = (Get-ItemProperty -Path $userDataRegPath -Name $userDataRegValue -ErrorAction SilentlyContinue).$userDataRegValue
    if ($configured) {
        return $configured
    }

    return Join-Path ([Environment]::GetFolderPath("ApplicationData")) "OpenCode Office Add-in"
}

function Remove-TrustedLocalhostCertificate {
    param(
        [string]$UserDataDir
    )

    $thumbprintPath = Join-Path $UserDataDir "certs\thumbprint.txt"
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $store.Open("ReadWrite")

    if (Test-Path $thumbprintPath) {
        $thumbprint = (Get-Content -LiteralPath $thumbprintPath -Raw).Trim()
        $cert = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
        if ($cert) {
            $store.Remove($cert)
        }
    }

    $store.Close()
}

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

# Prefer packaged resources when the desktop bundle is present; otherwise use repo-local assets.
if (Test-Path $appPath) {
    $manifestPath = "$PSScriptRoot\resources\manifest.xml"
} else {
    $manifestPath = "$PSScriptRoot\manifest.xml"
}

$manifestFullPath = if (Test-Path $manifestPath) { (Resolve-Path $manifestPath).Path } else { $null }
$manifestId = if ($manifestFullPath) { Get-OfficeManifestId -Path $manifestFullPath } else { $null }

Write-Host "Removing OpenCode Office sideload registration..." -ForegroundColor Cyan

if (Test-Path $regPath) {
    $removedCount = Remove-OpenCodeRegistrations -ManifestPath $manifestFullPath -ManifestId $manifestId
    if ($removedCount -gt 0) {
        Write-Host "Removed $removedCount OpenCode registration(s)" -ForegroundColor Green
    } else {
        Write-Host "No OpenCode registration was found in the Office developer key" -ForegroundColor Gray
    }
} else {
    Write-Host "The Office developer sideload key is not present" -ForegroundColor Gray
}

if (Test-Path $appPath) {
    $userDataDir = Get-InstalledUserDataDir
    Remove-TrustedLocalhostCertificate -UserDataDir $userDataDir
    if (Test-Path $userDataDir) {
        Remove-Item -LiteralPath $userDataDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -Path $userDataRegPath -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "To register again later, run: .\register.ps1" -ForegroundColor Gray
