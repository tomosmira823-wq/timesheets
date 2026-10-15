; Timesheets installer (NSIS) - per-user install, no admin needed
Unicode true
!include "MUI2.nsh"
!include "x64.nsh"

!define APP "Timesheets"
!define VER "1.0.0"
!define PUB "Kivo"
!define UNINST "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}"

Name "${APP}"
OutFile "Timesheets-Setup-${VER}.exe"
InstallDir "$LOCALAPPDATA\Programs\${APP}"
InstallDirRegKey HKCU "Software\${PUB}\${APP}" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma
BrandingText "${PUB}"

VIProductVersion "${VER}.0"
VIAddVersionKey "ProductName" "${APP}"
VIAddVersionKey "CompanyName" "${PUB}"
VIAddVersionKey "FileDescription" "${APP} installer"
VIAddVersionKey "FileVersion" "${VER}"
VIAddVersionKey "LegalCopyright" "${PUB}"

!define MUI_ICON "launcher\timesheets.ico"
!define MUI_UNICON "launcher\timesheets.ico"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "${APP}"
!define MUI_WELCOMEPAGE_TEXT "This installs the ${APP} app. Your company account, sign-in and data are online, so you can use it from any PC.$\r$\n$\r$\nNew company? Click Create your company on the sign-in screen.$\r$\n$\r$\nClick Next to continue."
!define MUI_FINISHPAGE_RUN "$INSTDIR\Timesheets.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Open ${APP} now"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Function .onInit
  ${IfNot} ${RunningX64}
    MessageBox MB_OK|MB_ICONSTOP "${APP} needs 64-bit Windows."
    Abort
  ${EndIf}
FunctionEnd

Section "Install"
  SetOutPath "$INSTDIR"
  ; close a running copy so files can be replaced on upgrade
  nsExec::ExecToLog 'taskkill /F /IM Timesheets.exe'
  nsExec::ExecToLog 'taskkill /F /IM pythonw.exe /FI "WINDOWTITLE eq Timesheets*"'
  File "stage\Timesheets.exe"
  SetOutPath "$INSTDIR\app"
  File /r "stage\app\*.*"
  SetOutPath "$INSTDIR\runtime"
  File /r "stage\runtime\*.*"
  SetOutPath "$INSTDIR"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateDirectory "$SMPROGRAMS\${APP}"
  CreateShortcut "$SMPROGRAMS\${APP}\${APP}.lnk" "$INSTDIR\Timesheets.exe"
  CreateShortcut "$SMPROGRAMS\${APP}\Uninstall ${APP}.lnk" "$INSTDIR\Uninstall.exe"
  CreateShortcut "$DESKTOP\${APP}.lnk" "$INSTDIR\Timesheets.exe"
  WriteRegStr HKCU "Software\${PUB}\${APP}" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "${UNINST}" "DisplayName" "${APP}"
  WriteRegStr HKCU "${UNINST}" "DisplayVersion" "${VER}"
  WriteRegStr HKCU "${UNINST}" "Publisher" "${PUB}"
  WriteRegStr HKCU "${UNINST}" "DisplayIcon" "$INSTDIR\Timesheets.exe"
  WriteRegStr HKCU "${UNINST}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINST}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegDWORD HKCU "${UNINST}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINST}" "NoRepair" 1
SectionEnd

Section "Uninstall"
  nsExec::ExecToLog 'taskkill /F /IM Timesheets.exe'
  Delete "$DESKTOP\${APP}.lnk"
  Delete "$SMPROGRAMS\${APP}\*.lnk"
  RMDir "$SMPROGRAMS\${APP}"
  RMDir /r "$INSTDIR\app"
  RMDir /r "$INSTDIR\runtime"
  Delete "$INSTDIR\Timesheets.exe"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINST}"
  DeleteRegKey HKCU "Software\${PUB}\${APP}"
  ; %APPDATA%\Timesheets (database, receipts, backups) is deliberately kept.
SectionEnd
