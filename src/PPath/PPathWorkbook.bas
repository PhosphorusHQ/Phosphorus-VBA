Attribute VB_Name = "PPathWorkbook"
'@Folder PPath
Option Explicit

Public gPPathWB As Excel.Workbook

Public Sub OpenWB()
  If gPPathWB Is Nothing Then
    Set gPPathWB = Excel.Workbooks.Add
    'gPPathWB.Windows(1).Visible = False : this seems to make the tests fail??
  End If
  Excel.Application.ScreenUpdating = False
End Sub

Public Sub CloseWB()
  'Auotmatically closed on Module Cleanup?
  On Error Resume Next
  gPPathWB.Close savechanges:=False
  Set gPPathWB = Nothing
  On Error GoTo 0
End Sub
