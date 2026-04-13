Attribute VB_Name = "pPathTestsCommon"
'@Folder pPath
'@TestModule
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'Public Assert As Object
'Public Fakes As Object

Public Sub OutputActualXPaths(strActualXPaths() As String)
  Dim strPath As String
  strPath = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "") & "\Logs\Phosphorus Ouput " & VBA.Strings.Replace(VBA.Strings.Replace(Now(), "/", "-"), ":", "") & ".txt"
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Dim oFile As Object
  Set oFile = FSO.CreateTextFile(strPath)
  Dim i As Integer
  Dim upper As Integer
  upper = 0
  On Error Resume Next
  upper = UBound(strActualXPaths)
  On Error GoTo 0
  For i = 1 To upper
    oFile.WriteLine strActualXPaths(i)
  Next
  oFile.Close
  Set FSO = Nothing
  Set oFile = Nothing
End Sub

Public Sub TestExpectedAndActualAllElementsPPath(strExpectedPPaths As String, EvaluatedXpath As pPath.ReturnClass)
  
  Dim strExpectedXPathsArray() As String
  strExpectedXPathsArray = VBA.Split(strExpectedPPaths, vbCrLf)
    
  Dim strActualXPathsArray() As String
'  strActualXPathsArray = VBA.Split(strActualXPaths, vbCrLf)
  strActualXPathsArray = EvaluatedXpath.GetMatchingNavigationalPPaths
  
  Dim intExpectedNumberOfXPaths As Integer
  intExpectedNumberOfXPaths = UBound(strExpectedXPathsArray) + 1
  
  Dim intActualNumberOfXPaths As Integer
  intActualNumberOfXPaths = 0
  On Error Resume Next
  intActualNumberOfXPaths = UBound(strActualXPathsArray)
  On Error GoTo 0
  
  Dim intLeastNumberOfXPaths As Integer
  intLeastNumberOfXPaths = VBA.Interaction.IIf(intExpectedNumberOfXPaths <= intActualNumberOfXPaths, intExpectedNumberOfXPaths, intActualNumberOfXPaths)
  
  Dim intXPathLoopCounter As Integer
  For intXPathLoopCounter = 1 To intLeastNumberOfXPaths
'    If intXPathLoopCounter = 999 Then
      If strExpectedXPathsArray(intXPathLoopCounter - 1) <> strActualXPathsArray(intXPathLoopCounter) Then
        Phosphorus.Log4PStatic.Logger.Info "Number: " & intXPathLoopCounter
        Phosphorus.Log4PStatic.Logger.Info "Expected: " & strExpectedXPathsArray(intXPathLoopCounter - 1), VBA.Strings.Len(strExpectedXPathsArray(intXPathLoopCounter - 1))
        Phosphorus.Log4PStatic.Logger.Info "Actual  : " & strActualXPathsArray(intXPathLoopCounter), VBA.Strings.Len(strActualXPathsArray(intXPathLoopCounter))
        Dim i As Integer
        For i = 1 To VBA.Strings.Len(strExpectedXPathsArray(intXPathLoopCounter - 1))
          If VBA.Strings.Mid(strExpectedXPathsArray(intXPathLoopCounter - 1), i, 1) <> VBA.Strings.Mid(strActualXPathsArray(intXPathLoopCounter), i, 1) Then
            Dim lngExpectedUnicode As Long
            Dim lngActualUnicode As Long
            Dim strExpectedCharacter As String
            Dim strActualCharacter As String
            strExpectedCharacter = VBA.Strings.Mid(strExpectedXPathsArray(intXPathLoopCounter - 1), i, 1)
            strActualCharacter = VBA.Strings.Mid(strActualXPathsArray(intXPathLoopCounter), i, 1)
            lngExpectedUnicode = Excel.WorksheetFunction.Unicode(strExpectedCharacter)
            If strActualCharacter = "" Then
              lngActualUnicode = 0
            Else
              lngActualUnicode = Excel.WorksheetFunction.Unicode(strActualCharacter)
            End If
            'lngActualUnicode = Application.WorksheetFunction.Unicode(strActualCharacter)
            'lngActualUnicode = VBA.Strings.AscW(strActualCharacter)
            Phosphorus.Log4PStatic.Logger.Info "Error Character #" & i & " (" & strExpectedCharacter & ")" & "(" & strActualCharacter & ")" & " Expected Unicode: " & lngExpectedUnicode & " Actual Unicode: " & lngActualUnicode
          End If
        Next i
      End If
'    End If
    Phosphorus.AssertionsStatic.pAssert.Equal strExpectedXPathsArray(intXPathLoopCounter - 1), strActualXPathsArray(intXPathLoopCounter), "GetAllElementsXPaths (" & intXPathLoopCounter & ")", isCritical:=True
  Next intXPathLoopCounter
  Phosphorus.AssertionsStatic.pAssert.Equal intExpectedNumberOfXPaths, intActualNumberOfXPaths, "GetAllElementsXpath_NumberOfXPaths", isCritical:=True
End Sub


