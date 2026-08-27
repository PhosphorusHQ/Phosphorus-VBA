Attribute VB_Name = "ExcelTopLevelFunctions"
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

Public TopLevelFunctions As Collection

Public Sub ClearDownTopLevelFunctions()
  Set TopLevelFunctions = Nothing
End Sub

Public Sub InitialiseTopLevelFunctions()
  If TopLevelFunctions Is Nothing Then
    Set TopLevelFunctions = New Collection
    AddATopLevelFunction "sum"
    AddATopLevelFunction "count"
    AddATopLevelFunction "not"
    AddATopLevelFunction "average"
    AddATopLevelFunction "stdeva"
    AddATopLevelFunction "min"
    AddATopLevelFunction "max"
    AddATopLevelFunction "product"
    AddATopLevelFunction "round", 1
    AddATopLevelFunction "roundup", 1
    AddATopLevelFunction "rounddown", 1
  End If
End Sub

Private Sub AddATopLevelFunction(strFunctionName As String, Optional intNumberOfAddtionalParameters As Integer)
  Dim TopLevelFunction As Collection
  Set TopLevelFunction = New Collection
  TopLevelFunction.Add strFunctionName, "FunctionName"
  TopLevelFunction.Add intNumberOfAddtionalParameters, "NumberOfAddtionalParameters"
  TopLevelFunctions.Add TopLevelFunction
End Sub

Private Sub test()
  pPath.ExcelTopLevelFunctions.ClearDownTopLevelFunctions
  pPath.ExcelTopLevelFunctions.InitialiseTopLevelFunctions
  Dim TopLevelFunction As Collection
  For Each TopLevelFunction In pPath.ExcelTopLevelFunctions.TopLevelFunctions
    Dim strFunctionName As String
    Dim intNumberOfAddtionalParameters As Integer
    strFunctionName = TopLevelFunction("FunctionName")
    intNumberOfAddtionalParameters = TopLevelFunction("NumberOfAddtionalParameters")
    'Debug.Print strFunctionName, intNumberOfAddtionalParameters
  Next TopLevelFunction
End Sub

