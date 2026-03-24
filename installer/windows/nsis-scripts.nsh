; Custom NSIS scripts for OpenCode Office Add-in
; These run during installation to set up SSL certificates and Office integration

!macro customInstall
  ; Trust the SSL certificate using PowerShell (handles PEM format correctly)
  DetailPrint "Installing SSL certificate..."
  nsExec::ExecToLog 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\
    $$certPath = \"$INSTDIR\resources\certs\localhost.pem\"; \
    if (Test-Path $$certPath) { \
      $$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($$certPath); \
      $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store(\"Root\", \"CurrentUser\"); \
      $$store.Open(\"ReadWrite\"); \
      $$existing = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$cert.Thumbprint }; \
      if (-not $$existing) { $$store.Add($$cert); } \
      $$store.Close(); \
      $$cert.Thumbprint | Out-File -FilePath \"$INSTDIR\resources\certs\.thumbprint\" -NoNewline; \
    }"'
  
  ; Register the add-in manifest with Office
  DetailPrint "Registering Office Add-in..."
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\WEF\Developer" "OpenCodeOfficeAddin" "$INSTDIR\resources\manifest.xml"
  
  ; Create startup registry entry to run on login
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "OpenCodeOfficeAddin" '"$INSTDIR\OpenCode Office Add-in.exe"'
!macroend

!macro customUnInstall
  ; Kill the running process first
  nsExec::ExecToLog 'taskkill /F /IM "OpenCode Office Add-in.exe"'
  
  ; Remove SSL certificate using thumbprint (more reliable than subject name)
  DetailPrint "Removing SSL certificate..."
  nsExec::ExecToLog 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\
    $$thumbprintFile = \"$INSTDIR\resources\certs\.thumbprint\"; \
    $$store = New-Object System.Security.Cryptography.X509Certificates.X509Store(\"Root\", \"CurrentUser\"); \
    $$store.Open(\"ReadWrite\"); \
    if (Test-Path $$thumbprintFile) { \
      $$thumbprint = (Get-Content $$thumbprintFile -Raw).Trim(); \
      $$cert = $$store.Certificates | Where-Object { $$_.Thumbprint -eq $$thumbprint }; \
      if ($$cert) { $$store.Remove($$cert); } \
    } else { \
      $$certs = $$store.Certificates | Where-Object { $$_.Subject -eq \"CN=localhost\" }; \
      foreach ($$c in $$certs) { $$store.Remove($$c); } \
    } \
    $$store.Close();"'
  
  ; Remove Office add-in registration
  DeleteRegValue HKCU "Software\Microsoft\Office\16.0\WEF\Developer" "OpenCodeOfficeAddin"
  
  ; Remove startup entry
  DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "OpenCodeOfficeAddin"
!macroend
