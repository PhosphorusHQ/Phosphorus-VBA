Attribute VB_Name = "OfficeRibbon"
'@Folder OfficeRibbon
Option Explicit

Private Rib As IRibbonUI

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
  Set Rib = ribbon
End Sub

Public Sub ActivateHomeTab()
  If Not Rib Is Nothing Then
    Rib.ActivateTabMso "TabHome"
  Else
    MsgBox "Ribbon reference not initialized!", vbExclamation
  End If
End Sub

Public Sub RunMacro1(control As IRibbonControl)
  ActivateHomeTab
  MsgBox "Macro 1 executed!", vbInformation
End Sub

Public Sub RunMacro2(control As IRibbonControl)
  ActivateHomeTab
  MsgBox "Macro 2 executed!", vbInformation
End Sub



