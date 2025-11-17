Attribute VB_Name = "RuntimeIDs"
'@Folder pPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

' Declare the AccessibleObjectFromWindow API function
#If VBA7 Then
  Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
    ByVal hwnd As LongPtr, _
    ByVal dwObjectID As LongPtr, _
    ByRef riid As GUID, _
    ByRef ppvObject As Object _
    ) As LongPtr
#Else
  Private Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
   ByVal hwnd As Long, _
   ByVal dwObjectID As Long, _
   ByRef riid As GUID, _
   ByRef ppvObject As Object _
   ) As Long
#End If

#If VBA7 Then ' Check if using VBA7 or later (64-bit compatible)
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
  ) As LongPtr ' Use LongPtr for 64-bit compatibility
#Else ' For older VBA versions (32-bit only)
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
  ) As Long ' Use Long
#End If


'http://www.cpearson.com/Excel/CreateGUID.aspx
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'This declaration of the AccessibleObjectFromWindow API function is compatible with both 32-bit and 64-bit versions of Windows. The key is the use of Long for the hwnd parameter.
'Here 's why:
' * Long Data Type: In VBA, the Long data type is 32 bits wide. However, on 64-bit Windows, even though pointers are 64 bits, window handles (HWNDs) are still 32-bit values. So, using Long for the hwnd parameter works correctly on both 32-bit and 64-bit Windows.
' * Pointers: The ppvObject parameter (which receives the interface pointer) is declared as Object. In VBA, Object can hold pointers of any size, so it correctly handles the 64-bit pointer returned on 64-bit Windows.
' * No Need for Separate Declarations: You do not need separate declarations for 32-bit and 64-bit versions of AccessibleObjectFromWindow when using VBA. The single declaration provided works correctly in both environments.
'
'key Point:
'The crucial part is the use of Long for the hwnd and Object for the pointer. This combination ensures that the declaration is compatible with both 32-bit and 64-bit Windows. You don't need to use LongPtr or other 64-bit specific types for this particular API function.

' Define the IID_IAccessible GUID
Private Const IID_IAccessible As String = "{618736E0-3E94-11CF-A842-0080C7381A2E}"

' Define the object IDs
Private Const OBJID_CLIENT As LongPtr = &HFFFFFFF4
Private Const OBJID_WINDOW As Long = &H0

Public Function GetElementRuntimeID(ByRef MatchingElement As UIAutomationClient.IUIAutomationElement) As String
     
  Dim strRuntimeIDString As String
  strRuntimeIDString = ""
  
  'The runtime ID is a unique array of elements (SAFEARRAY)
  Dim varRuntimeIDArray As Variant
  varRuntimeIDArray = MatchingElement.GetRuntimeId
  
  'Check if the RuntimeID was retrived successfully
  If IsArray(varRuntimeIDArray) Then
    'Iterate through the array
    Dim i As Long
    
    For i = LBound(varRuntimeIDArray) To UBound(varRuntimeIDArray)
      If i > LBound(varRuntimeIDArray) Then
        strRuntimeIDString = strRuntimeIDString & " "
      End If
      strRuntimeIDString = strRuntimeIDString & varRuntimeIDArray(i)
    Next i
       
  Else
  
    MsgBox "Failed to get runtime id!"
       
  End If
  
  If strRuntimeIDString = "" Then
    
'    Debug.Print _
'      "ControlTypeID: " & MatchingElement.CurrentControlType, _
'      "Name: " & MatchingElement.CurrentName
      'MatchingElement.CurrentClassName
      
    Dim hwnd As LongPtr
    'hwnd = GetWindowFromElement(ParentElement)
'    hwnd = FindWindow(vbNullString, MatchingElement.CurrentName)
    hwnd = FindWindow("Excel", MatchingElement.CurrentName)
    'Get the IAccessible interface
    Dim iAccessible As iAccessible
    Set iAccessible = GetAccessibleObjectFromWindow(hwnd, OBJID_CLIENT)
    Dim childId As Long

    If Not iAccessible Is Nothing Then
      ' 3. Use the IAccessible interface
'      Debug.Print iAccessible.accName ' Example: Get the name of the accessible object
'      Debug.Print iAccessible.accRole ' Example: Get the role
      ' ... other IAccessible properties and methods ...
    
      'Get child ID
      childId = iAccessible.accChild(0)
      'Some applications don't return a valid ID for 0, try -1
      If childId = 0 Then
        childId = iAccessible.accChild(-1)
      End If
    
      ' 4. (Optional) Convert to UIAutomation element (if needed)
      Dim UIAutomation As IUIAutomation
      Dim element As IUIAutomationElement
      Set UIAutomation = New CUIAutomation
      Set element = UIAutomation.ElementFromIAccessible(iAccessible, childId)

      If Not element Is Nothing Then
        MsgBox "Converted to UIAutomation element = - no we need to get th eruntime ID!"
        'Now you can use the UIAutomation Element
      Else
'        Debug.Print "Failed to convert to UIAutomation element"
      End If
    
    End If
    
  End If
  
  GetElementRuntimeID = strRuntimeIDString

End Function

'If we can't get a RuntimeID for an element it may be that it is an older MSAA type element.
'This code gets the corresponding UIA element.
'We may need to call it on the parent elements window handle, or use Find Window WinAPI call with the element name?
'Function to get the IAccessible interface
Private Function GetAccessibleObjectFromWindow(hwnd As LongPtr, objectID As LongPtr) As iAccessible

  Dim iid As GUID
  Dim ppvObject As Object
  Dim hr As LongPtr

  ' Convert the IID string to a GUID
  Call StringToGUID(IID_IAccessible, iid)

  ' Call the AccessibleObjectFromWindow API
  hr = AccessibleObjectFromWindow(hwnd, objectID, iid, ppvObject)

  ' Check for success and return the IAccessible interface
  If hr = 0 And Not ppvObject Is Nothing Then
    Set GetAccessibleObjectFromWindow = ppvObject
  Else
    Set GetAccessibleObjectFromWindow = Nothing
  End If

End Function

' Helper function to convert a string to a GUID (you might already have this)
Private Function StringToGUID(strGuid As String, GUID As GUID)

  Dim bytes() As Byte
  bytes = VBA.Strings.StrConv(strGuid, vbFromUnicode)

  With GUID
    .Data1 = VBA.Conversion.CLng("&H" & VBA.Strings.Mid(strGuid, 2, 8))
    .Data2 = VBA.Conversion.CInt("&H" & VBA.Strings.Mid(strGuid, 11, 4))
    .Data3 = VBA.Conversion.CInt("&H" & VBA.Strings.Mid(strGuid, 16, 4))
    .Data4(0) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 21, 2))
    .Data4(1) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 23, 2))
    .Data4(2) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 26, 2))
    .Data4(3) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 27, 2))
    .Data4(4) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 29, 2))
    .Data4(5) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 32, 2))
    .Data4(6) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 33, 2))
    .Data4(7) = VBA.Conversion.CByte("&H" & VBA.Strings.Mid(strGuid, 35, 2))
  End With

End Function


' Example usage:
Private Sub TestAccessibleObjectFromWindow()

  Dim hwnd As LongPtr
  Dim iAccessible As iAccessible
  Dim element As IUIAutomationElement
  Dim UIAutomation As IUIAutomation
'  Dim accessible As Object 'Used to receive the IAccessible interface

  ' *** 1. Get the HWND of the window you want to interact with ***
  ' You'll need to find a way to get the HWND. This is highly application-specific.
  ' Here's a placeholder -- replace with your actual HWND retrieval method.
  hwnd = 123456 'FindWindow("Notepad", "Untitled - Notepad") ' Example: Find Notepad window

  If hwnd = 0 Then
    MsgBox "Window not found!"
    Exit Sub
  End If

  ' 2. Get the IAccessible interface
  Set iAccessible = GetAccessibleObjectFromWindow(hwnd, OBJID_CLIENT)

  If Not iAccessible Is Nothing Then
    ' 3. Use the IAccessible interface
'    Debug.Print iAccessible.accName ' Example: Get the name of the accessible object
'    Debug.Print iAccessible.accRole ' Example: Get the role
    ' ... other IAccessible properties and methods ...

    ' 4. (Optional) Convert to UIAutomation element (if needed)
    Set UIAutomation = New CUIAutomation
    Set element = UIAutomation.ElementFromIAccessible(iAccessible, 0)

    If Not element Is Nothing Then
'      Debug.Print "Converted to UIAutomation element"
      'Now you can use the UIAutomation Element
    Else
 '     Debug.Print "Failed to convert to UIAutomation element"
    End If

    Set element = Nothing
    Set UIAutomation = Nothing

  Else
'    Debug.Print "Failed to get IAccessible interface."
  End If

  Set iAccessible = Nothing

End Sub

'Helper function to get the window from an element
Private Function GetWindowFromElement(element As IUIAutomationElement) As LongPtr

  On Error Resume Next 'Handle errors

  Dim tempElement As IUIAutomationElement
  Set tempElement = element

  Do While Not tempElement Is Nothing
    GetWindowFromElement = tempElement.GetCurrentPropertyValue(UIAutomationClient.UIA_PropertyIds.UIA_NativeWindowHandlePropertyId)
    If GetWindowFromElement <> 0 Then
      Exit Do 'Found it!
    End If
    Dim UIAutomation As CUIAutomation
    Set UIAutomation = New CUIAutomation
'    Set tempElement = tempElement.FindFirst(UIAutomationClient.TreeScope.TreeScope_Parent, uiAutomation.CreateTrueCondition())
    Set tempElement = tempElement.GetCachedParent
  Loop

  Set tempElement = Nothing

End Function

'' Example usage (works in both 32-bit and 64-bit VBA):
'Sub TestFindWindow()
'
'Dim hwnd As LongPtr ' Use LongPtr
'
'hwnd = FindWindow("Notepad", "Untitled - Notepad") ' Example: Find Notepad
'
'If hwnd <> 0 Then
'Debug.Print "Window found! HWND: " & hwnd
'Else
'Debug.Print "Window not found."
'End If
'
'End Sub




