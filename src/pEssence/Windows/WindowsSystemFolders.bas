Attribute VB_Name = "WindowsSystemFolders"
'@Folder Windows
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================

'Taken from: https://www.devhut.net/vba-getting-the-path-of-system-folders/

Public Enum KNOWNFOLDERIDS
  FOLDERID_AddNewPrograms
  FOLDERID_AdminTools
  FOLDERID_AppUpdates
  FOLDERID_CDBurning
  FOLDERID_ChangeRemovePrograms
  FOLDERID_CommonAdminTools
  FOLDERID_CommonOEMLinks
  FOLDERID_CommonPrograms
  FOLDERID_CommonStartMenu
  FOLDERID_CommonStartup
  FOLDERID_CommonTemplates
  FOLDERID_ComputerFolder
  FOLDERID_ConflictFolder
  FOLDERID_ConnectionsFolder
  FOLDERID_Contacts
  FOLDERID_ControlPanelFolder
  FOLDERID_Cookies
  FOLDERID_Desktop
  FOLDERID_Documents
  FOLDERID_Downloads
  FOLDERID_Favorites
  FOLDERID_Fonts
  FOLDERID_Games
  FOLDERID_GameTasks
  FOLDERID_History
  FOLDERID_InternetCache
  FOLDERID_InternetFolder
  FOLDERID_Links
  FOLDERID_LocalAppData
  FOLDERID_LocalAppDataLow
  FOLDERID_LocalizedResourcesDir
  FOLDERID_Music
  FOLDERID_NetHood
  FOLDERID_NetworkFolder
  FOLDERID_OriginalImages
  FOLDERID_PhotoAlbums
  FOLDERID_Pictures
  FOLDERID_Playlists
  FOLDERID_PrintersFolder
  FOLDERID_PrintHood
  FOLDERID_Profile
  FOLDERID_ProgramData
  FOLDERID_ProgramFiles
  FOLDERID_ProgramFilesCommon
  FOLDERID_ProgramFilesCommonX64
  FOLDERID_ProgramFilesCommonX86
  FOLDERID_ProgramFilesX64
  FOLDERID_ProgramFilesX86
  FOLDERID_Programs
  FOLDERID_Public
  FOLDERID_PublicDesktop
  FOLDERID_PublicDocuments
  FOLDERID_PublicDownloads
  FOLDERID_PublicGameTasks
  FOLDERID_PublicMusic
  FOLDERID_PublicPictures
  FOLDERID_PublicVideos
  FOLDERID_QuickLaunch
  FOLDERID_Recent
  FOLDERID_RecordedTV
  FOLDERID_RecycleBinFolder
  FOLDERID_ResourceDir
  FOLDERID_RoamingAppData
  FOLDERID_SampleMusic
  FOLDERID_SamplePictures
  FOLDERID_SamplePlaylists
  FOLDERID_SampleVideos
  FOLDERID_SavedGames
  FOLDERID_SavedSearches
  FOLDERID_SEARCH_CSC
  FOLDERID_SEARCH_MAPI
  FOLDERID_SearchHome
  FOLDERID_SendTo
  FOLDERID_SidebarDefaultParts
  FOLDERID_SidebarParts
  FOLDERID_StartMenu
  FOLDERID_Startup
  FOLDERID_SyncManagerFolder
  FOLDERID_SyncResultsFolder
  FOLDERID_SyncSetupFolder
  FOLDERID_System
  FOLDERID_SystemX86
  FOLDERID_Templates
  FOLDERID_TreeProperties
  FOLDERID_UserProfiles
  FOLDERID_UsersFiles
  FOLDERID_Videos
  FOLDERID_Windows
End Enum

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte   ' Array of 8 bytes
End Type

#If VBA7 Then
  
  ' -----------------------------------------------
  ' Modern declarations (Office 2010+ / VBA7+)
  ' Works in both 32-bit and 64-bit Office
  ' -----------------------------------------------

  Private Declare PtrSafe Function SHGetKnownFolderPath Lib "shell32.dll" ( _
    ByRef rfid As Any, _
    ByVal dwFlags As Long, _
    ByVal hToken As LongPtr, _
    ByRef ppszPath As LongPtr) As Long
    
  Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" ( _
    ByVal lpszGuid As LongPtr, _
    ByRef pGuid As Any) As Long

  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef pDest As Any, _
    ByRef pSrc As Any, _
    ByVal ByteLen As LongPtr)

  Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" ( _
    ByVal hMem As LongPtr)

  Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal ptr As LongPtr) As Long

#Else
    
    ' -----------------------------------------------
    ' Legacy 32-bit only declarations (Office 2007 and older)
    ' -----------------------------------------------

  Private Declare Function SHGetKnownFolderPath Lib "shell32.dll" ( _
    rfid As Any, _
    ByVal dwFlags As Long, _
    ByVal hToken As Long, _
    ppszPath As Long) As Long

  Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
    ByVal lpszGuid As Long, _
    pGuid As Any) As Long

  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDest As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long)

  Private Declare Sub CoTaskMemFree Lib "ole32.dll" ( _
    ByVal hMem As Long)

  Private Declare Function lstrlenW Lib "kernel32" ( _
    ByVal ptr As Long) As Long

#End If

Public Function GetFolderBySHGetKnownFolderPath(lRfid As KNOWNFOLDERIDS) As String
#If VBA7 Then
  Dim ppszPath              As LongPtr
#Else
  Dim ppszPath              As Long
#End If
  Dim gGuid                 As GUID
  Dim sRfid                 As String

  sRfid = Get_sRfid(lRfid)

  'attempt to convert the knownfolder ID to a GUID
  If CLSIDFromString(StrPtr(sRfid), gGuid) = 0 Then
    'pass gGuid to SHGetKnownFolderPath, which returns
    'the path as a pointer to a Unicode string.
    If SHGetKnownFolderPath(gGuid, 0, 0, ppszPath) = 0 Then
      'if no error on return get the path
      'if present and release the pointer
       GetFolderBySHGetKnownFolderPath = GetPointerToByteStringW(ppszPath)
       Call CoTaskMemFree(ppszPath)
    End If
  End If  'CLSIDFromString
  
End Function

#If VBA7 Then
Private Function GetPointerToByteStringW(ByVal dwData As LongPtr) As String
#Else
Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
#End If
  Dim tmp()                 As Byte
  Dim tmplen                As Long

  If dwData <> 0 Then
    'determine the size of the returned data
    tmplen = lstrlenW(dwData) * 2

    If tmplen <> 0 Then
      'create a byte buffer for the string
      'then assign it to the return value
      'of the function
      ReDim tmp(0 To (tmplen - 1)) As Byte
      CopyMemory tmp(0), ByVal dwData, tmplen
      GetPointerToByteStringW = tmp
    End If
  End If

End Function

Private Function Get_sRfid(lREFKNOWNFOLDERID As Long) As String
  'xref REFKNOWNFOLDERID Enum value to GUID string value
  Select Case lREFKNOWNFOLDERID
    Case FOLDERID_AddNewPrograms: Get_sRfid = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
    Case FOLDERID_AdminTools: Get_sRfid = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
    Case FOLDERID_AppUpdates: Get_sRfid = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
    Case FOLDERID_CDBurning: Get_sRfid = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
    Case FOLDERID_ChangeRemovePrograms: Get_sRfid = "{df7266ac-9274-4867-8d55-3bd661de872d}"
    Case FOLDERID_CommonAdminTools: Get_sRfid = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
    Case FOLDERID_CommonOEMLinks: Get_sRfid = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
    Case FOLDERID_CommonPrograms: Get_sRfid = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
    Case FOLDERID_CommonStartMenu: Get_sRfid = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
    Case FOLDERID_CommonStartup: Get_sRfid = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
    Case FOLDERID_CommonTemplates: Get_sRfid = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
    Case FOLDERID_ComputerFolder: Get_sRfid = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
    Case FOLDERID_ConflictFolder: Get_sRfid = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
    Case FOLDERID_ConnectionsFolder: Get_sRfid = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
    Case FOLDERID_Contacts: Get_sRfid = "{56784854-C6CB-462b-8169-88E350ACB882}"
    Case FOLDERID_ControlPanelFolder: Get_sRfid = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
    Case FOLDERID_Cookies: Get_sRfid = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
    Case FOLDERID_Desktop: Get_sRfid = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
    Case FOLDERID_Documents: Get_sRfid = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
    Case FOLDERID_Downloads: Get_sRfid = "{374DE290-123F-4565-9164-39C4925E467B}"
    Case FOLDERID_Favorites: Get_sRfid = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
    Case FOLDERID_Fonts: Get_sRfid = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
    Case FOLDERID_Games: Get_sRfid = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
    Case FOLDERID_GameTasks: Get_sRfid = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
    Case FOLDERID_History: Get_sRfid = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
    Case FOLDERID_InternetCache: Get_sRfid = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
    Case FOLDERID_InternetFolder: Get_sRfid = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
    Case FOLDERID_Links: Get_sRfid = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
    Case FOLDERID_LocalAppData: Get_sRfid = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
    Case FOLDERID_LocalAppDataLow: Get_sRfid = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
    Case FOLDERID_LocalizedResourcesDir: Get_sRfid = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
    Case FOLDERID_Music: Get_sRfid = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
    Case FOLDERID_NetHood: Get_sRfid = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
    Case FOLDERID_OriginalImages: Get_sRfid = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
    Case FOLDERID_PhotoAlbums: Get_sRfid = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
    Case FOLDERID_Pictures: Get_sRfid = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
    Case FOLDERID_Playlists: Get_sRfid = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
    Case FOLDERID_PrintersFolder: Get_sRfid = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
    Case FOLDERID_PrintHood: Get_sRfid = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
    Case FOLDERID_Profile: Get_sRfid = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
    Case FOLDERID_ProgramData: Get_sRfid = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
    Case FOLDERID_ProgramFiles: Get_sRfid = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
    Case FOLDERID_ProgramFilesCommon: Get_sRfid = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
    Case FOLDERID_ProgramFilesCommonX64: Get_sRfid = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
    Case FOLDERID_ProgramFilesCommonX86: Get_sRfid = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
    Case FOLDERID_ProgramFilesX64: Get_sRfid = "{6D809377-6AF0-444b-8957-A3773F02200E}"
    Case FOLDERID_ProgramFilesX86: Get_sRfid = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
    Case FOLDERID_Programs: Get_sRfid = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
    Case FOLDERID_Public: Get_sRfid = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
    Case FOLDERID_PublicDesktop: Get_sRfid = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
    Case FOLDERID_PublicDocuments: Get_sRfid = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
    Case FOLDERID_PublicDownloads: Get_sRfid = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
    Case FOLDERID_PublicGameTasks: Get_sRfid = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
    Case FOLDERID_PublicMusic: Get_sRfid = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
    Case FOLDERID_PublicPictures: Get_sRfid = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
    Case FOLDERID_PublicVideos: Get_sRfid = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
    Case FOLDERID_QuickLaunch: Get_sRfid = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
    Case FOLDERID_Recent: Get_sRfid = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
    Case FOLDERID_RecordedTV: Get_sRfid = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
    Case FOLDERID_RecycleBinFolder: Get_sRfid = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
    Case FOLDERID_ResourceDir: Get_sRfid = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
    Case FOLDERID_RoamingAppData: Get_sRfid = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
    Case FOLDERID_SampleMusic: Get_sRfid = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
    Case FOLDERID_SamplePictures: Get_sRfid = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
    Case FOLDERID_SamplePlaylists: Get_sRfid = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
    Case FOLDERID_SampleVideos: Get_sRfid = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
    Case FOLDERID_SavedGames: Get_sRfid = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
    Case FOLDERID_SavedSearches: Get_sRfid = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
    Case FOLDERID_SEARCH_CSC: Get_sRfid = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
    Case FOLDERID_SEARCH_MAPI: Get_sRfid = "{98ec0e18-2098-4d44-8644-66979315a281}"
    Case FOLDERID_SearchHome: Get_sRfid = "{190337d1-b8ca-4121-a639-6d472d16972a}"
    Case FOLDERID_SendTo: Get_sRfid = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
    Case FOLDERID_SidebarDefaultParts: Get_sRfid = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
    Case FOLDERID_SidebarParts: Get_sRfid = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
    Case FOLDERID_StartMenu: Get_sRfid = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
    Case FOLDERID_Startup: Get_sRfid = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
    Case FOLDERID_SyncManagerFolder: Get_sRfid = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
    Case FOLDERID_SyncResultsFolder: Get_sRfid = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
    Case FOLDERID_SyncSetupFolder: Get_sRfid = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
    Case FOLDERID_System: Get_sRfid = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
    Case FOLDERID_SystemX86: Get_sRfid = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
    Case FOLDERID_Templates: Get_sRfid = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
    Case FOLDERID_TreeProperties: Get_sRfid = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
    Case FOLDERID_UserProfiles: Get_sRfid = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
    Case FOLDERID_UsersFiles: Get_sRfid = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
    Case FOLDERID_Videos: Get_sRfid = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
    Case FOLDERID_Windows: Get_sRfid = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
    Case FOLDERID_NetworkFolder: Get_sRfid = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
  End Select
End Function

Private Sub Test()
  Debug.Print GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_AdminTools)
  Debug.Print GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_Documents)
  Debug.Print GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_Pictures)
  Debug.Print GetFolderBySHGetKnownFolderPath(KNOWNFOLDERIDS.FOLDERID_Downloads)
End Sub

