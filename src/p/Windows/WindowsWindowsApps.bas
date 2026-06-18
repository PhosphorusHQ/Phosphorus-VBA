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
' Add: '!App' to Package Family Name to get AppID

'Return the Official Name of the app

Public Const DuckDuckGo_OfficialName = "DuckDuckGo.DesktopBrowser"
Public Const MicrosoftEdge_OfficialName = "Microsoft.MicrosoftEdge.Stable"
Public Const MicrosoftWindowsCalculator_OfficialName = "Microsoft.WindowsCalculator"

Public Function GetAppID(OfficialName As String) As String

  Dim WShell As Object
  Dim Exec As Object
  Dim Output As String
  Dim PackageFamilyName As String
  Dim Cmd As String

  Cmd = "Get-AppxPackage -Name '*" & OfficialName & "*' | Select-Object -ExpandProperty PackageFamilyName"
  Output = Phosphorus.WindowsPowerShell.Execute(Cmd, "WindowsWindowsApps.GetAppID(" & OfficialName & ")")

  ' Clean up the output (remove extra whitespace, newlines)
  Output = VBA.Strings.Trim(Output)

  ' Check if output is empty (app not found)
  If Output = "" Then
    'Raise an exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverWindowsAppNotFound, OfficialName
  Else
    ' Construct and return the App ID
    PackageFamilyName = Output
    'Remove any line feeds from the PS command re
    PackageFamilyName = VBA.Strings.Replace(PackageFamilyName, vbCrLf, "")
    'Append !App to th epackage name!
    GetAppID = PackageFamilyName & "!App"
  End If
    
End Function

'Other code removed due to triggering false Malware alert in OneDrive
