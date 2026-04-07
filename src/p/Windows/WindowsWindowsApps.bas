Attribute VB_Name = "WindowsWindowsApps"
'@Folder Windows
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'AUMID = Application User Model ID (App ID or AppUserModelID)

'Powershell command to get all app package names:
'    powershell Get-AppxPackage | Select Name, PackageFamilyName

'Code removed due to triggering false Malware alert in OneDrive
