; Custom NSIS hooks for OpenCode Office Add-in

!define OfficeDevRegPath "Software\Microsoft\Office\16.0\WEF\Developer"
!define StartupRunPath "Software\Microsoft\Windows\CurrentVersion\Run"
!define StartupValueName "OpenCodeOfficeAddin"

!macro RunLoggedPowerShell SCRIPT
  nsExec::ExecToLog 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "${SCRIPT}"'
!macroend

!macro customInstall
  DetailPrint "Installing SSL certificate..."
  !insertmacro RunLoggedPowerShell '
    $$certPath = "$INSTDIR\resources\certs\localhost.pem"; \
    if (Test-Path $$certPath) { \
      $$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($$certPath); \
      $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser"); \
      $$store.Open("ReadWrite"); \
      $$existing = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$cert.Thumbprint }; \
      if (-not $$existing) { $$store.Add($$cert); } \
      $$store.Close(); \
      $$cert.Thumbprint | Out-File -FilePath "$INSTDIR\resources\certs\.thumbprint" -NoNewline; \
    }'

  DetailPrint "Registering Office Add-in..."
  !insertmacro RunLoggedPowerShell '
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
  WriteRegStr HKCU "${StartupRunPath}" "${StartupValueName}" '"$INSTDIR\OpenCode Office Add-in.exe"'
!macroend

!macro customUnInstall
  nsExec::ExecToLog 'taskkill /F /IM "OpenCode Office Add-in.exe"'

  DetailPrint "Removing SSL certificate..."
  !insertmacro RunLoggedPowerShell '
    $$thumbprintFile = "$INSTDIR\resources\certs\.thumbprint"; \
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
    $$store.Close();'

  !insertmacro RunLoggedPowerShell '
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

  DeleteRegValue HKCU "${StartupRunPath}" "${StartupValueName}"
!macroend
