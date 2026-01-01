# Agent Performance Processor - Windows Installer Script
# This creates a professional Windows installer for the application

!define APP_NAME "Agent Performance Processor"
!define APP_VERSION "2.3"
!define APP_PUBLISHER "Agent Performance Solutions"
!define APP_URL "https://github.com/anilsunil97/agent-performance-processor"
!define APP_DESCRIPTION "Professional Data Analysis & Reporting Tool"

# Installer settings
Name "${APP_NAME}"
OutFile "AgentPerformanceProcessor_Setup.exe"
InstallDir "$PROGRAMFILES\${APP_NAME}"
InstallDirRegKey HKLM "Software\${APP_NAME}" "InstallDir"
RequestExecutionLevel admin

# Modern UI
!include "MUI2.nsh"

# Interface Settings
!define MUI_ABORTWARNING
!define MUI_ICON "app_icon_hd.ico"
!define MUI_UNICON "app_icon_hd.ico"

# Welcome page
!insertmacro MUI_PAGE_WELCOME

# License page
!insertmacro MUI_PAGE_LICENSE "LICENSE"

# Directory page
!insertmacro MUI_PAGE_DIRECTORY

# Start menu page
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKLM"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\${APP_NAME}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"
Var StartMenuFolder
!insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder

# Installation page
!insertmacro MUI_PAGE_INSTFILES

# Finish page
!define MUI_FINISHPAGE_RUN "$INSTDIR\AgentPerformanceProcessor_Offline.exe"
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\README_OFFLINE.txt"
!insertmacro MUI_PAGE_FINISH

# Uninstaller pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

# Languages
!insertmacro MUI_LANGUAGE "English"

# Version Information
VIProductVersion "${APP_VERSION}.0.0"
VIAddVersionKey "ProductName" "${APP_NAME}"
VIAddVersionKey "ProductVersion" "${APP_VERSION}"
VIAddVersionKey "CompanyName" "${APP_PUBLISHER}"
VIAddVersionKey "FileDescription" "${APP_DESCRIPTION}"
VIAddVersionKey "FileVersion" "${APP_VERSION}"
VIAddVersionKey "LegalCopyright" "© 2025 ${APP_PUBLISHER}"

# Installation section
Section "Main Application" SecMain
    SectionIn RO
    
    # Set output path to the installation directory
    SetOutPath $INSTDIR
    
    # Install files
    File "AgentPerformanceProcessor_Offline\AgentPerformanceProcessor_Offline.exe"
    File "AgentPerformanceProcessor_Offline\README_OFFLINE.txt"
    File "app_icon_hd.ico"
    
    # Store installation folder
    WriteRegStr HKLM "Software\${APP_NAME}" "InstallDir" $INSTDIR
    
    # Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    
    # Add to Add/Remove Programs
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\app_icon_hd.ico"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${APP_PUBLISHER}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${APP_VERSION}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "URLInfoAbout" "${APP_URL}"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoRepair" 1
    
    # Calculate installed size (approximate - 50MB)
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "EstimatedSize" 51200
    
SectionEnd

# Start Menu shortcuts
Section "Start Menu Shortcuts" SecStartMenu
    !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
    CreateShortcut "$SMPROGRAMS\$StartMenuFolder\${APP_NAME}.lnk" "$INSTDIR\AgentPerformanceProcessor_Offline.exe" "" "$INSTDIR\app_icon_hd.ico"
    CreateShortcut "$SMPROGRAMS\$StartMenuFolder\README.lnk" "$INSTDIR\README_OFFLINE.txt"
    CreateShortcut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
    
    !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

# Desktop shortcut
Section "Desktop Shortcut" SecDesktop
    CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\AgentPerformanceProcessor_Offline.exe" "" "$INSTDIR\app_icon_hd.ico"
SectionEnd

# File associations
Section "File Associations" SecFileAssoc
    # Associate .csv files with the application (optional)
    WriteRegStr HKCR ".csv\OpenWithProgids" "AgentPerformanceProcessor.csv" ""
    WriteRegStr HKCR "AgentPerformanceProcessor.csv" "" "Agent Performance CSV File"
    WriteRegStr HKCR "AgentPerformanceProcessor.csv\DefaultIcon" "" "$INSTDIR\app_icon_hd.ico"
    WriteRegStr HKCR "AgentPerformanceProcessor.csv\shell\open\command" "" '"$INSTDIR\AgentPerformanceProcessor_Offline.exe" "%1"'
SectionEnd

# Section descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
!insertmacro MUI_DESCRIPTION_TEXT ${SecMain} "Main application files (required)"
!insertmacro MUI_DESCRIPTION_TEXT ${SecStartMenu} "Create shortcuts in Start Menu"
!insertmacro MUI_DESCRIPTION_TEXT ${SecDesktop} "Create shortcut on Desktop"
!insertmacro MUI_DESCRIPTION_TEXT ${SecFileAssoc} "Associate CSV files with Agent Performance Processor"
!insertmacro MUI_FUNCTION_DESCRIPTION_END

# Uninstaller section
Section "Uninstall"
    # Remove files
    Delete "$INSTDIR\AgentPerformanceProcessor_Offline.exe"
    Delete "$INSTDIR\README_OFFLINE.txt"
    Delete "$INSTDIR\app_icon_hd.ico"
    Delete "$INSTDIR\Uninstall.exe"
    
    # Remove directories
    RMDir "$INSTDIR"
    
    # Remove Start Menu shortcuts
    !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
    Delete "$SMPROGRAMS\$StartMenuFolder\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\$StartMenuFolder\README.lnk"
    Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk"
    RMDir "$SMPROGRAMS\$StartMenuFolder"
    
    # Remove Desktop shortcut
    Delete "$DESKTOP\${APP_NAME}.lnk"
    
    # Remove registry keys
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
    DeleteRegKey HKLM "Software\${APP_NAME}"
    
    # Remove file associations
    DeleteRegKey HKCR "AgentPerformanceProcessor.csv"
    DeleteRegValue HKCR ".csv\OpenWithProgids" "AgentPerformanceProcessor.csv"
    
SectionEnd

# Include required plugins (removed GetSize as it's not needed)