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

    return $removed
}

# Check if running from release (exe exists) or dev (manifest in root)
if (Test-Path $appPath) {
    # Release mode: certs and manifest are in resources subfolder
    $manifestPath = "$PSScriptRoot\resources\manifest.xml"
    $certPath = "$PSScriptRoot\resources\certs\localhost.pem"
} else {
    # Dev mode: certs and manifest are in the repo root
    $manifestPath = "$PSScriptRoot\manifest.xml"
    $certPath = "$PSScriptRoot\certs\localhost.pem"
}

$manifestFullPath = (Resolve-Path $manifestPath).Path
$manifestId = Get-OfficeManifestId -Path $manifestFullPath

Write-Host "Setting up Office Add-in for Windows..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Trust the SSL certificate
Write-Host "Step 1: Trusting development SSL certificate..." -ForegroundColor Yellow

if (!(Test-Path $certPath)) {
    Write-Host "Error: Certificate not found at $certPath" -ForegroundColor Red
    Write-Host "Certificates are required for HTTPS. Please ensure certs are in the repository." -ForegroundColor Red
    exit 1
}

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")

# Check if certificate is already trusted
$existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($existing) {
    Write-Host "  ✓ Certificate already trusted" -ForegroundColor Green
} else {
    $store.Add($cert)
    Write-Host "  ✓ Certificate trusted" -ForegroundColor Green
}

$store.Close()

Write-Host ""

# Step 2: Register manifest
Write-Host "Step 2: Registering add-in manifest..." -ForegroundColor Yellow
Write-Host "  Manifest: $manifestFullPath"

if (!(Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

$removedCount = Remove-OpenCodeRegistrations -ManifestPath $manifestFullPath -ManifestId $manifestId

New-ItemProperty -Path $regPath -Name $registrationName -Value $manifestFullPath -PropertyType String -Force | Out-Null

if ($removedCount -gt 0) {
    Write-Host "  ✓ Removed $removedCount previous OpenCode registration(s)" -ForegroundColor Green
}

Write-Host "  ✓ Add-in registered" -ForegroundColor Green
Write-Host ""

Write-Host "Setup complete! Next steps:" -ForegroundColor Cyan
Write-Host "1. Close Word, PowerPoint, Excel, and OneNote if they are open"
Write-Host "2. Start the tray app: bun run start:tray"
Write-Host "3. Open Word, PowerPoint, Excel, or OneNote"
Write-Host "4. Go to Insert > Add-ins > My Add-ins and look for 'OpenCode'"
Write-Host ""
Write-Host "To unregister, run: .\unregister.ps1" -ForegroundColor Gray
