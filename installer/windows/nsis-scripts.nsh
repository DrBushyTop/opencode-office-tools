; OpenCode Office Add-in NSIS hooks

!define OfficeDevRegPath "Software\Microsoft\Office\16.0\WEF\Developer"
!define StartupRunPath "Software\Microsoft\Windows\CurrentVersion\Run"
!define StartupValueName "OpenCodeOfficeAddin"
!define UserDataRegPath "Software\OpenCode\OfficeAddin"
!define UserDataValueName "UserDataDir"
!define CertPassphrase "OpenCodeOfficeLocalCert"

!macro OpenCodeRunPowerShell SCRIPT
  nsExec::ExecToLog 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "${SCRIPT}"'
!macroend

!macro OpenCodeInstallCertificate
  DetailPrint "Trusting localhost HTTPS certificate..."
  !insertmacro OpenCodeRunPowerShell '
    $$userDataDir = [Environment]::GetFolderPath("ApplicationData") + "\OpenCode Office Add-in"; \
    $$certDir = Join-Path $$userDataDir "certs"; \
    $$thumbprintPath = Join-Path $$certDir "thumbprint.txt"; \
    $$pfxPath = Join-Path $$certDir "localhost.pfx"; \
    $$pemPath = Join-Path $$certDir "localhost.pem"; \
    $$subject = "CN=localhost"; \
    New-Item -ItemType Directory -Path $$certDir -Force | Out-Null; \
    $$dnsNames = @("localhost", "*.localhost"); \
    $$san = "dns=localhost&dns=*.localhost&ipaddress=127.0.0.1&ipaddress=::1"; \
    $$cert = New-SelfSignedCertificate -Subject $$subject -FriendlyName "OpenCode Office Add-in localhost" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -CertStoreLocation "Cert:\CurrentUser\My" -TextExtension @("2.5.29.17={text}$$san") -NotAfter (Get-Date).AddYears(5) -DnsName $$dnsNames; \
    $$password = ConvertTo-SecureString "${CertPassphrase}" -Force -AsPlainText; \
    Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$$($$cert.Thumbprint)" -FilePath $$pfxPath -Password $$password | Out-Null; \
    Export-Certificate -Cert "Cert:\CurrentUser\My\$$($$cert.Thumbprint)" -FilePath $$pemPath -Type CERT | Out-Null; \
    $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser"); \
    $$store.Open("ReadWrite"); \
    $$existing = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$cert.Thumbprint }; \
    if (-not $$existing) { $$store.Add($$cert); } \
    $$store.Close(); \
    $$cert.Thumbprint | Out-File -FilePath $$thumbprintPath -NoNewline; \
    New-Item -Path "HKCU:\${UserDataRegPath}" -Force | Out-Null; \
    Set-ItemProperty -Path "HKCU:\${UserDataRegPath}" -Name "${UserDataValueName}" -Value $$userDataDir; \
    $$leafStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser"); \
    $$leafStore.Open("ReadWrite"); \
    $$leaf = $$leafStore.Certificates | Where-Object { $$_.Thumbprint -eq $$cert.Thumbprint }; \
    if ($$leaf) { $$leafStore.Remove($$leaf); } \
    $$leafStore.Close(); \
    if (Test-Path $$pemPath) { \
      $$trusted = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($$pemPath); \
      $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser"); \
      $$store.Open("ReadWrite"); \
      $$existing = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$trusted.Thumbprint }; \
      if (-not $$existing) { $$store.Add($$trusted); } \
      $$store.Close(); \
    }'
!macroend

!macro OpenCodeSyncManifest
  DetailPrint "Refreshing Office sideload registration..."
  !insertmacro OpenCodeRunPowerShell '
    $$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"; \
    $$manifestPath = "$INSTDIR\resources\manifest.xml"; \
    function Get-ManifestId([string]$$path) { \
      if (-not (Test-Path $$path)) { return $$null; } \
      try { [xml]$$xml = Get-Content -LiteralPath $$path -Raw; return [string]$$xml.OfficeApp.Id; } catch { return $$null; } \
    }; \
    if (-not (Test-Path $$regPath)) { New-Item -Path $$regPath -Force | Out-Null; } \
    $$manifestId = Get-ManifestId $$manifestPath; \
    foreach ($$name in (Get-Item -Path $$regPath).Property | Where-Object { $$_ -notlike "PS*" }) { \
      $$value = (Get-ItemProperty -Path $$regPath -Name $$name -ErrorAction SilentlyContinue).$$name; \
      $$shouldRemove = $$name -eq "${StartupValueName}" -or $$value -eq $$manifestPath; \
      if (-not $$shouldRemove -and $$manifestId -and $$value -and (Test-Path $$value)) { \
        $$existingManifestId = Get-ManifestId $$value; \
        if ($$existingManifestId -and $$existingManifestId -eq $$manifestId) { $$shouldRemove = $$true; } \
      } \
      if ($$shouldRemove) { Remove-ItemProperty -Path $$regPath -Name $$name -ErrorAction SilentlyContinue; } \
    }'
  WriteRegStr HKCU "${OfficeDevRegPath}" "${StartupValueName}" "$INSTDIR\resources\manifest.xml"
!macroend

!macro OpenCodeRegisterStartup
  ReadRegStr $0 HKCU "${UserDataRegPath}" "${UserDataValueName}"
  ${If} $0 != ""
    WriteRegStr HKCU "${UserDataRegPath}" "${UserDataValueName}" "$0"
  ${EndIf}
  WriteRegStr HKCU "${StartupRunPath}" "${StartupValueName}" '"$INSTDIR\OpenCode Office Add-in.exe"'
!macroend

!macro OpenCodeRemoveCertificate
  DetailPrint "Removing localhost HTTPS certificate..."
  !insertmacro OpenCodeRunPowerShell '
    $$userDataDir = (Get-ItemProperty -Path "HKCU:\${UserDataRegPath}" -Name "${UserDataValueName}" -ErrorAction SilentlyContinue).${UserDataValueName}; \
    if (-not $$userDataDir) { $$userDataDir = [Environment]::GetFolderPath("ApplicationData") + "\OpenCode Office Add-in"; } \
    $$thumbprintFile = Join-Path $$userDataDir "certs\thumbprint.txt"; \
    $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser"); \
    $$store.Open("ReadWrite"); \
    if (Test-Path $$thumbprintFile) { \
      $$thumbprint = (Get-Content $$thumbprintFile -Raw).Trim(); \
      $$cert = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$thumbprint }; \
      if ($$cert) { $$store.Remove($$cert); } \
    } else { \
      $$certs = $$store.Certificates | Where-Object { $$_.Subject -eq "CN=localhost" }; \
      foreach ($$c in $$certs) { $$store.Remove($$c); } \
    } \
    $$store.Close(); \
    if (Test-Path $$userDataDir) { Remove-Item $$userDataDir -Recurse -Force -ErrorAction SilentlyContinue; } \
    Remove-Item "HKCU:\${UserDataRegPath}" -Recurse -Force -ErrorAction SilentlyContinue;'
!macroend

!macro OpenCodeRemoveManifest
  !insertmacro OpenCodeRunPowerShell '
    $$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"; \
    $$manifestPath = "$INSTDIR\resources\manifest.xml"; \
    function Get-ManifestId([string]$$path) { \
      if (-not (Test-Path $$path)) { return $$null; } \
      try { [xml]$$xml = Get-Content -LiteralPath $$path -Raw; return [string]$$xml.OfficeApp.Id; } catch { return $$null; } \
    }; \
    if (Test-Path $$regPath) { \
      $$manifestId = Get-ManifestId $$manifestPath; \
      foreach ($$name in (Get-Item -Path $$regPath).Property | Where-Object { $$_ -notlike "PS*" }) { \
        $$value = (Get-ItemProperty -Path $$regPath -Name $$name -ErrorAction SilentlyContinue).$$name; \
        $$shouldRemove = $$name -eq "${StartupValueName}" -or $$value -eq $$manifestPath; \
        if (-not $$shouldRemove -and $$manifestId -and $$value -and (Test-Path $$value)) { \
          $$existingManifestId = Get-ManifestId $$value; \
          if ($$existingManifestId -and $$existingManifestId -eq $$manifestId) { $$shouldRemove = $$true; } \
        } \
        if ($$shouldRemove) { Remove-ItemProperty -Path $$regPath -Name $$name -ErrorAction SilentlyContinue; } \
      } \
      $$remaining = (Get-Item -Path $$regPath -ErrorAction SilentlyContinue).Property | Where-Object { $$_ -notlike "PS*" }; \
      if (-not $$remaining) { Remove-Item $$regPath -Force -ErrorAction SilentlyContinue; } \
    }'
!macroend

!macro OpenCodeRemoveStartup
  DeleteRegValue HKCU "${StartupRunPath}" "${StartupValueName}"
!macroend

!macro customInstall
  !insertmacro OpenCodeInstallCertificate
  !insertmacro OpenCodeSyncManifest
  !insertmacro OpenCodeRegisterStartup
!macroend

!macro customUnInstall
  nsExec::ExecToLog 'taskkill /F /IM "OpenCode Office Add-in.exe"'
  !insertmacro OpenCodeRemoveCertificate
  !insertmacro OpenCodeRemoveManifest
  !insertmacro OpenCodeRemoveStartup
!macroend
