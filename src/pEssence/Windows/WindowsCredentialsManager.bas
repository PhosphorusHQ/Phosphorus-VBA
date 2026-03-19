Attribute VB_Name = "WindowsCredentialsManager"
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

' --------------------------------------------------------------
' Windows Credential Manager API Declarations
' Works in both 32-bit and 64-bit Office (VBA7+ = Office 2010+)
' --------------------------------------------------------------

#If VBA7 Then   ' Office 2010 and newer

  #If Win64 Then
    Private Declare PtrSafe Function CredWriteW Lib "advapi32.dll" ( _
      ByRef CREDENTIAL As CREDENTIAL, _
      ByVal Flags As Long) As Long

    Private Declare PtrSafe Function CredReadW Lib "advapi32.dll" ( _
      ByVal TargetName As LongPtr, _
      ByVal Type_ As Long, _
      ByVal ReservedFlag As Long, _
      ByRef CredentialPtr As LongPtr) As Long

    Private Declare PtrSafe Sub CredFree Lib "advapi32.dll" ( _
      ByVal cred As LongPtr)

    Private Declare PtrSafe Function CredDeleteW Lib "advapi32.dll" ( _
      ByVal TargetName As LongPtr, _
      ByVal Type_ As Long, _
      ByVal Flags As Long) As Long
      
  #Else
    
    ' 32-bit Office on 64-bit Windows (still uses 32-bit pointers)
    Private Declare PtrSafe Function CredWriteW Lib "advapi32.dll" ( _
      ByRef CREDENTIAL As CREDENTIAL, _
      ByVal Flags As Long) As Long

    Private Declare PtrSafe Function CredReadW Lib "advapi32.dll" ( _
      ByVal TargetName As String, _
      ByVal Type_ As Long, _
      ByVal ReservedFlag As Long, _
      ByRef CredentialPtr As LongPtr) As Long

    Private Declare PtrSafe Sub CredFree Lib "advapi32.dll" ( _
      ByVal cred As LongPtr)

    Private Declare PtrSafe Function CredDeleteW Lib "advapi32.dll" ( _
      ByVal TargetName As String, _
      ByVal Type_ As Long, _
      ByVal Flags As Long) As Long
  
  #End If

#Else

  ' Office 2007 or older (very rare in 2026)
  Private Declare Function CredWriteW Lib "advapi32.dll" ( _
    ByRef Credential As CREDENTIAL, _
    ByVal Flags As Long) As Long

  Private Declare Function CredReadW Lib "advapi32.dll" ( _
    ByVal TargetName As String, _
    ByVal Type_ As Long, _
    ByVal ReservedFlag As Long, _
    ByRef CredentialPtr As Long) As Long

  Private Declare Sub CredFree Lib "advapi32.dll" ( _
    ByVal Cred As Long)

  Private Declare Function CredDeleteW Lib "advapi32.dll" ( _
    ByVal TargetName As String, _
    ByVal Type_ As Long, _
    ByVal Flags As Long) As Long

#End If

#If VBA7 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#Else
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)
  Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
#End If

' --------------------------------------------------------------
' Required structures
' --------------------------------------------------------------

Private Type CREDENTIAL
  Flags              As Long
  Type               As Long
  TargetName         As LongPtr   ' LPCWSTR
  Comment            As LongPtr   ' LPCWSTR
  LastWritten_low    As Long      ' FILETIME low part
  LastWritten_high   As Long      ' FILETIME high part
  CredentialBlobSize As Long
  CredentialBlob     As LongPtr   ' LPBYTE
  Persist            As Long
  AttributeCount     As Long
  Attributes         As LongPtr   ' PCREDENTIAL_ATTRIBUTE
  TargetAlias        As LongPtr   ' LPCWSTR
  UserName           As LongPtr   ' LPCWSTR
End Type

Private Type CREDENTIAL_ATTRIBUTE
    Keyword   As LongPtr   ' LPWSTR — pointer to Unicode string (keyword/name)
    Flags     As Long      ' DWORD — reserved; MUST be 0
    ValueSize As Long      ' DWORD — size of Value in bytes (= 256)
    Value     As LongPtr   ' LPBYTE — pointer to the raw byte data
End Type

Private Const CRED_TYPE_GENERIC     As Long = 1
Private Const CRED_PERSIST_LOCAL_MACHINE As Long = 1
Private Const CRED_PERSIST_SESSION   As Long = 2
Private Const CRED_PERSIST_ENTERPRISE   As Long = 3
Private Const CRED_MAX_ATTRIBUTES   As Long = 64
Private Const CRED_MAX_VALUE_SIZE   As Long = 256

Public Enum CredentialType
  Generic = 1
End Enum

Public Enum CredentialPersistence
  LocalMachine = 1
  Session = 2
  Enterprise = 3
End Enum

Public Type GenericCredential
  Type As CredentialType
  Target  As String
  User    As String
  Secret  As String
  Persistence As CredentialPersistence
  Comment As String
  Attributes As Scripting.Dictionary
End Type

Private Sub Test_NoAttributes()

  Dim myCred As GenericCredential
  myCred.Type = CredentialType.Generic
  myCred.Target = "MyApp:SecretToken"
  myCred.User = "anystring"
  myCred.Secret = "my-very-long-secret-2026"
  myCred.Persistence = CRED_PERSIST_ENTERPRISE
  myCred.Comment = "My App Secret Token"
  
  If WriteCredential(myCred) Then

    Dim myCred2 As GenericCredential
    myCred2 = ReadCredentialNoAttributes(myCred.Target)
    
    Debug.Print "Type Read: " & (myCred2.Type = myCred.Type)
    Debug.Print "Target Read: " & (myCred2.Target = myCred.Target)
    Debug.Print "User Read: " & (myCred2.User = myCred.User)
    Debug.Print "Secret Read: " & (myCred2.Secret = myCred.Secret)
    Debug.Print "Persistence Read: " & (myCred2.Persistence = myCred.Persistence)
    
    DeleteCredential myCred.Target, myCred.User
    
  End If
  
End Sub

Public Function WriteCredential(gc As GenericCredential) As Boolean

  Dim cred As CREDENTIAL
  With cred
    .Type = VBA.Conversion.CLng(gc.Type)
    .TargetName = StrPtr(gc.Target)
    .UserName = StrPtr(gc.User)
    .Comment = StrPtr(gc.Comment)
    .CredentialBlobSize = Len(gc.Secret) * 2                    ' bytes (Unicode)
    .CredentialBlob = StrPtr(gc.Secret)
    .Persist = VBA.Conversion.CLng(gc.Persistence)
    .AttributeCount = 0
    .Attributes = 0
  End With

  If CredWriteW(cred, 0) = 0 Then
    WriteCredential = False
  Else
    WriteCredential = True
  End If

End Function

Public Function ReadCredentialNoAttributes(Target As String) As GenericCredential
  
  Dim ReturnCredentials As GenericCredential

  Dim CredPtr     As LongPtr
  Dim cred        As CREDENTIAL
  Dim Password    As String

  If CredReadW(StrPtr(Target), CRED_TYPE_GENERIC, 0, CredPtr) = 0 Then
    MsgBox "Not found or error: " & Err.LastDllError
    Exit Function
  End If

  ' Copy structure from pointer
  CopyMemory cred, ByVal CredPtr, LenB(cred)

  ' Extract password (Unicode string)
  If cred.CredentialBlobSize > 0 Then
    Password = String$(cred.CredentialBlobSize \ 2, vbNullChar)
    CopyMemory ByVal StrPtr(Password), ByVal cred.CredentialBlob, cred.CredentialBlobSize
  End If

  With ReturnCredentials
    .Type = cred.Type
    .Target = PtrToString(cred.TargetName)
    .User = PtrToString(cred.UserName)
    .Secret = Password
    .Persistence = cred.Persist
    .Comment = PtrToString(cred.Comment)
  End With
  
  CredFree CredPtr

  ReadCredentialNoAttributes = ReturnCredentials
  
End Function

' ----------------------------------------------------------------------------
' Delete credential by Target (and optional User match)
' ----------------------------------------------------------------------------
Public Sub DeleteCredential(ByVal Target As String, Optional ByVal User As String = vbNullString)

  On Error Resume Next
  If Err.Number <> 0 Then
    Err.Clear
    Exit Sub   ' already doesn't exist ? silent success
  End If
  On Error GoTo 0

  Dim hr As Long
  hr = CredDeleteW(StrPtr(Target), CRED_TYPE_GENERIC, 0)
  If hr = 0 Then
    Err.Raise vbObjectError + 4, , "CredDeleteW failed: " & Err.LastDllError
  End If
    
End Sub

Private Sub Test_Attributes()

  Dim myCred As GenericCredential
  myCred.Type = CredentialType.Generic
  myCred.Target = "MyCompany_API_Token_v1"
  myCred.User = "token-user"
  myCred.Secret = "abc123-super-secret-token-2026"
  myCred.Persistence = CRED_PERSIST_ENTERPRISE

  Dim Attributes As Scripting.Dictionary
  Set myCred.Attributes = New Scripting.Dictionary
  myCred.Attributes.Add "MyCompany_CreatedBy", "VBA-Script-1.0"
  myCred.Attributes.Add "MyCompany_Environment", "Production"

  If WriteCredentialWithAttributes(myCred) Then
        
    Dim myCred2 As GenericCredential
    myCred2 = ReadCredentialWithAttributes(myCred.Target)
    
    Debug.Print "Type Read: " & (myCred2.Type = myCred.Type)
    Debug.Print "Target Read: " & (myCred2.Target = myCred.Target)
    Debug.Print "User Read: " & (myCred2.User = myCred.User)
    Debug.Print "Secret Read: " & (myCred2.Secret = myCred.Secret)
    Debug.Print "Persistence Read: " & (myCred2.Persistence = myCred.Persistence)
    Debug.Print "Persistence Read: " & (myCred2.Attributes("MyCompany_CreatedBy") = myCred.Attributes("MyCompany_CreatedBy"))

    DeleteCredential myCred.Target, myCred.User
    
  End If

End Sub

Public Function WriteCredentialWithAttributes(gc As GenericCredential) As Boolean

  Dim AttributeCount     As Long
  AttributeCount = gc.Attributes.Count
  Dim attrs() As CREDENTIAL_ATTRIBUTE
  ReDim attrs(0 To AttributeCount - 1) As CREDENTIAL_ATTRIBUTE ' Example: 2 attributes

  Dim attrData() As Byte
  Dim i As Long
  Dim k As Variant
  i = -1
  For Each k In gc.Attributes.Keys
    i = i + 1
    attrData = gc.Attributes(k)
    With attrs(i)
      .Keyword = StrPtr(k)
      .Flags = 0
      .ValueSize = UBound(attrData) + 1   ' bytes
      .Value = VarPtr(attrData(0))
    End With
  Next

  Dim cred As CREDENTIAL
  With cred
    .Flags = 0
    .Type = VBA.Conversion.CLng(gc.Type) 'CRED_TYPE_GENERIC
    .TargetName = StrPtr(gc.Target)
    .UserName = StrPtr(gc.User)
    .Comment = StrPtr("My secure API token")
    .CredentialBlobSize = Len(gc.Secret) * 2                    ' bytes (Unicode)
    .CredentialBlob = StrPtr(gc.Secret)
    .Persist = gc.Persistence ' CRED_PERSIST_ENTERPRISE
    .AttributeCount = AttributeCount
    .Attributes = VarPtr(attrs(0))           ' Pointer to first attribute
  End With

  If CredWriteW(cred, 0) <> 0 Then
'    MsgBox "Credential + attributes saved successfully."
    WriteCredentialWithAttributes = True
  Else
'    MsgBox "CredWriteW failed: " & Err.LastDllError
    WriteCredentialWithAttributes = False
  End If

  ' Arrays go out of scope ? memory freed automatically (no CredFree needed here)

End Function

Public Function ReadCredentialWithAttributes(Target As String) As GenericCredential

  Dim ReturnCredential As GenericCredential
  Dim CredPtr As LongPtr
  Dim cred As CREDENTIAL
  Dim i As Long
  Dim attr As CREDENTIAL_ATTRIBUTE
  Dim AttrArray() As Byte

  If CredReadW(StrPtr(Target), CRED_TYPE_GENERIC, 0, CredPtr) = 0 Then
     MsgBox "CredReadW failed or not found: " & Err.LastDllError
     Exit Function
  End If

  CopyMemory cred, ByVal CredPtr, LenB(cred)

  ' Extract password
  Dim Password As String
  If cred.CredentialBlobSize > 0 Then
    Password = String$(cred.CredentialBlobSize \ 2, vbNullChar)
    CopyMemory ByVal StrPtr(Password), ByVal cred.CredentialBlob, cred.CredentialBlobSize
  End If

  ' Read attributes
  Set ReturnCredential.Attributes = New Scripting.Dictionary

  If cred.AttributeCount > 0 Then
    For i = 0 To cred.AttributeCount - 1
      CopyMemory attr, ByVal cred.Attributes + (i * LenB(attr)), LenB(attr)

      If attr.ValueSize > 0 Then
    
        Dim valBytes() As Byte
        ReDim valBytes(0 To attr.ValueSize - 1)
        CopyMemory valBytes(0), ByVal attr.Value, attr.ValueSize

        'Get value
        ' ------------------------------------------------
        ' Treat bytes as UTF-16LE (no StrConv needed!)
        ' Just assign directly via pointer or manual string creation
        ' ------------------------------------------------
        Dim valStr As String
        valStr = String$(attr.ValueSize \ 2, vbNullChar)
        CopyMemory ByVal StrPtr(valStr), valBytes(0), attr.ValueSize

        ' Optional: trim any accidental trailing null char (shouldn't be there, but safe)
        If Right$(valStr, 1) = vbNullChar Then
          valStr = Left$(valStr, Len(valStr) - 1)
        End If
          
        Dim AttrKey As String
        AttrKey = PtrToString(attr.Keyword)
        ReturnCredential.Attributes.Add AttrKey, valStr
      
      End If
    Next i
  End If

  With ReturnCredential
    .Type = cred.Type
    .Target = PtrToString(cred.TargetName)
    .User = PtrToString(cred.UserName)
    .Comment = PtrToString(cred.Comment)
    .Secret = Password
    .Persistence = cred.Persist
  End With

  CredFree CredPtr
  
  ReadCredentialWithAttributes = ReturnCredential

End Function

' Helper: Convert LPWSTR pointer to VBA string
Private Function PtrToString(ByVal lpwstr As LongPtr) As String
  If lpwstr = 0 Then
    Exit Function
  End If
  Dim lenStr As Long
  lenStr = lstrlenW(lpwstr)
  PtrToString = String$(lenStr, vbNullChar)
  CopyMemory ByVal StrPtr(PtrToString), ByVal lpwstr, lenStr * 2
End Function

Public Sub OpenCredentialManager()
  WindowsProcesses.RunShellExecuteToStartNewProcess _
    "Credential Manager", _
    "open", _
    "control.exe", _
    "/name Microsoft.CredentialManager", _
    vbNullString, _
    WindowStyle.Normal
End Sub
