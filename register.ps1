$appPath = "$PSScriptRoot\OpenCode Office Add-in.exe"
$registrationName = "OpenCodeOfficeAddin"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$userDataRegPath = "HKCU:\Software\OpenCode\OfficeAddin"
$userDataRegValue = "UserDataDir"
$certPassphrase = "OpenCodeOfficeLocalCert"

function Get-InstalledUserDataDir {
    $configured = (Get-ItemProperty -Path $userDataRegPath -Name $userDataRegValue -ErrorAction SilentlyContinue).$userDataRegValue
    if ($configured) {
        return $configured
    }

    return Join-Path ([Environment]::GetFolderPath("ApplicationData")) "OpenCode Office Add-in"
}

function Ensure-LocalhostCertificate {
    param(
        [string]$UserDataDir,
        [switch]$Packaged
    )

    if (!$Packaged) {
        return "$PSScriptRoot\certs\localhost.pem"
    }

    $certDir = Join-Path $UserDataDir "certs"
    $pemPath = Join-Path $certDir "localhost.pem"
    $pfxPath = Join-Path $certDir "localhost.pfx"
    $thumbprintPath = Join-Path $certDir "thumbprint.txt"
    New-Item -ItemType Directory -Path $certDir -Force | Out-Null

    if (!(Test-Path $pfxPath) -or !(Test-Path $pemPath)) {
        $san = "dns=localhost&dns=*.localhost&ipaddress=127.0.0.1&ipaddress=::1"
        $cert = New-SelfSignedCertificate -Subject "CN=localhost" -FriendlyName "OpenCode Office Add-in localhost" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -CertStoreLocation "Cert:\CurrentUser\My" -TextExtension @("2.5.29.17={text}$san") -NotAfter (Get-Date).AddYears(5) -DnsName @("localhost", "*.localhost")
        $password = ConvertTo-SecureString $certPassphrase -Force -AsPlainText
        Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath $pfxPath -Password $password | Out-Null
        Export-Certificate -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath $pemPath -Type CERT | Out-Null
        $cert.Thumbprint | Out-File -FilePath $thumbprintPath -NoNewline

        $leafStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
        $leafStore.Open("ReadWrite")
        $leaf = $leafStore.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
        if ($leaf) {
            $leafStore.Remove($leaf)
        }
        $leafStore.Close()
    }

    New-Item -Path $userDataRegPath -Force | Out-Null
    New-ItemProperty -Path $userDataRegPath -Name $userDataRegValue -Value $UserDataDir -PropertyType String -Force | Out-Null
    return $pemPath
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

    return $removed
}

# Prefer packaged resources when the desktop bundle is present; otherwise use repo-local assets.
if (Test-Path $appPath) {
    $manifestPath = "$PSScriptRoot\resources\manifest.xml"
    $userDataDir = Get-InstalledUserDataDir
    $certPath = Ensure-LocalhostCertificate -UserDataDir $userDataDir -Packaged
} else {
    $manifestPath = "$PSScriptRoot\manifest.xml"
    $certPath = "$PSScriptRoot\certs\localhost.pem"
}

$manifestFullPath = (Resolve-Path $manifestPath).Path
$manifestId = Get-OfficeManifestId -Path $manifestFullPath

Write-Host "Preparing OpenCode Office Add-in on Windows..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Trust the localhost HTTPS certificate used by the local add-in server.
Write-Host "Step 1: Trusting the localhost HTTPS certificate..." -ForegroundColor Yellow

if (!(Test-Path $certPath)) {
    Write-Host "Error: Missing certificate at $certPath" -ForegroundColor Red
    Write-Host "The local HTTPS endpoint cannot start without it." -ForegroundColor Red
    exit 1
}

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")

# Reuse trust if the same certificate thumbprint is already installed.
$existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($existing) {
    Write-Host "  ✓ Certificate trust is already in place" -ForegroundColor Green
} else {
    $store.Add($cert)
    Write-Host "  ✓ Certificate trust installed" -ForegroundColor Green
}

$store.Close()

Write-Host ""

# Step 2: Refresh the Office sideload manifest registration.
Write-Host "Step 2: Refreshing Office sideload manifest..." -ForegroundColor Yellow
Write-Host "  Manifest source: $manifestFullPath"

if (!(Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

$removedCount = Remove-OpenCodeRegistrations -ManifestPath $manifestFullPath -ManifestId $manifestId

New-ItemProperty -Path $regPath -Name $registrationName -Value $manifestFullPath -PropertyType String -Force | Out-Null

if ($removedCount -gt 0) {
    Write-Host "  ✓ Removed $removedCount earlier OpenCode registration(s)" -ForegroundColor Green
}

Write-Host "  ✓ Office sideload registration updated" -ForegroundColor Green
Write-Host ""

Write-Host "Setup complete. Next steps:" -ForegroundColor Cyan
Write-Host "1. Close Word, PowerPoint, and Excel if they are open"
Write-Host "2. Launch the installed tray app"
Write-Host "3. Open Word, PowerPoint, or Excel"
Write-Host "4. Go to Insert > Add-ins > My Add-ins and look for 'OpenCode'"
Write-Host ""
Write-Host "To remove the sideload registration later, run: .\unregister.ps1" -ForegroundColor Gray
