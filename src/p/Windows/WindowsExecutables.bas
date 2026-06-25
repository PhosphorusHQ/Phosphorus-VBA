Attribute VB_Name = "WindowsExecutables"
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

Public Type Executable
  Name As String
  ExeFile As String
  FullPath As String
End Type

Public Function AdobeAcrobat() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Acrobat"
  myExecutable.ExeFile = "Acrobat.exe"
  AdobeAcrobat = myExecutable
End Function

Public Function Brave() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Brave"
  myExecutable.ExeFile = "brave.exe"
  Brave = myExecutable
End Function

Public Function Chrome() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Chrome"
  myExecutable.ExeFile = "chrome.exe"
  myExecutable.FullPath = GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_ProgramFiles) & "\Google\Chrome\Application\" & myExecutable.ExeFile
  Chrome = myExecutable
End Function

Public Function Chromium() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Chromium"
  myExecutable.ExeFile = "chrome.exe"
  myExecutable.FullPath = GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_ProgramFiles) & "\Chromium\Application\" & myExecutable.ExeFile
  Chromium = myExecutable
End Function

Public Function Epic() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Epic"
  myExecutable.ExeFile = "epic.exe"
  Epic = myExecutable
End Function

Public Function Firefox() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Firefox"
  myExecutable.ExeFile = "firefox.exe"
  Firefox = myExecutable
End Function
 
Public Function MicrosoftEdge() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Microsoft Edge"
  myExecutable.ExeFile = "msedge.exe"
  MicrosoftEdge = myExecutable
End Function

Public Function MicrosoftWord() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Microsoft Word"
  myExecutable.ExeFile = "winword.exe"
  MicrosoftWord = myExecutable
End Function

Public Function Opera() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Opera"
  myExecutable.ExeFile = "opera.exe"
  Opera = myExecutable
End Function

Public Function NotepadPlusPlus() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Notepad++"
  myExecutable.ExeFile = "notepad++.exe"
  NotepadPlusPlus = myExecutable
End Function

Public Function PowerShell() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "PowerShell"
  myExecutable.ExeFile = "powershell.exe"
  'PowerShell will always be on the PATH so can be called just by it's name
  myExecutable.FullPath = myExecutable.ExeFile
  PowerShell = myExecutable
End Function

Public Function SamsungBrowser() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "SamsungBrowser"
  myExecutable.ExeFile = "samsunginternet.exe"
  SamsungBrowser = myExecutable
End Function

Public Function WindowsCharacterMap() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Character Map"
  myExecutable.ExeFile = "charmap.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsCharacterMap = myExecutable
End Function

Public Function WindowsDiskCleanUp() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Disk Clean Up"
  myExecutable.ExeFile = "cleanmgr.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsDiskCleanUp = myExecutable
End Function

Public Function WindowsExplorer() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Explorer"
  myExecutable.ExeFile = "explorer.exe"
  'Explorer will always be on the PATH so can be called just by it's name
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsExplorer = myExecutable
End Function

Public Function WindowsPCHealthCheck() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows PC Health Check"
  myExecutable.ExeFile = "PCHealthCheck.exe"
  'Explorer will always installed here?
  myExecutable.FullPath = "C:\Program Files\PCHealthCheck\" & myExecutable.ExeFile
  WindowsPCHealthCheck = myExecutable
End Function

Public Function WindowsMediaPlayer() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Media Player"
  myExecutable.ExeFile = "wmplayer.exe"
  WindowsMediaPlayer = myExecutable
End Function

Public Function WindowsNotepad() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Notepad"
  myExecutable.ExeFile = "notepad.exe"
  'Explorer will always be on the PATH so can be called just by it's name
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsNotepad = myExecutable
End Function

Public Function WindowsTaskManager() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Task Manager"
  myExecutable.ExeFile = "taskmgr.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsTaskManager = myExecutable
End Function

Public Function YandexWebBrowser() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Yandex Web Browser"
  myExecutable.ExeFile = "browser.exe"
  Dim InstalledEXEPath As String
  InstalledEXEPath = "C:\Program Files\Yandex\YandexBrowser\Application\browser.exe"
  If Dir(InstalledEXEPath) <> "" Then
    myExecutable.FullPath = InstalledEXEPath
  Else
    myExecutable.FullPath = GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_LocalAppData) & "\Yandex\YandexBrowser\Application\" & myExecutable.ExeFile
  End If
  YandexWebBrowser = myExecutable
End Function

' ************************
' * WINDOWS SYSTEM TOOLS *
' ************************

'C:\WINDOWS\system32\*.exe
'These will always be on the PATH so can be called just by their file name
