Attribute VB_Name = "PPathExceTopLevelFunctions"
'@Folder PPath
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
  PPathExceTopLevelFunctions.ClearDownTopLevelFunctions
  PPathExceTopLevelFunctions.InitialiseTopLevelFunctions
  Dim TopLevelFunction As Collection
  For Each TopLevelFunction In PPathExceTopLevelFunctions.TopLevelFunctions
    Dim strFunctionName As String
    Dim intNumberOfAddtionalParameters As Integer
    strFunctionName = TopLevelFunction("FunctionName")
    intNumberOfAddtionalParameters = TopLevelFunction("NumberOfAddtionalParameters")
    Debug.Print strFunctionName, intNumberOfAddtionalParameters
  Next TopLevelFunction
End Sub
