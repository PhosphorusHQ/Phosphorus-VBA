Attribute VB_Name = "WindowsMMCSnapIns"
'@Folder Windows
Option Explicit

Public Type WindowsMMCSnapIn
  Name As String
  FileName As String
End Type

'C:\WINDOWS\system32\*.msc

'01/04/2024  08:22            41,587 azman.msc
'01/04/2024  08:22            63,081 certlm.msc
'01/04/2024  08:22            63,070 certmgr.msc
'01/04/2024  08:22           124,118 comexp.msc

Public Function ComputerManagement() As Phosphorus.WindowsMMCSnapIn
  Dim myMMCSnapIn As Phosphorus.WindowsMMCSnapIn
  myMMCSnapIn.Name = "Computer Management"
  myMMCSnapIn.FileName = "compmgmt.msc"
  ComputerManagement = myMMCSnapIn
End Function

'01/04/2024  08:22           145,622 devmgmt.msc
'01/04/2024  08:22            13,091 DevModeRunAsUserConfig.msc

'01/04/2024  08:22            47,682 diskmgmt.msc

Public Function EventViewer() As Phosphorus.WindowsMMCSnapIn
  Dim myMMCSnapIn As Phosphorus.WindowsMMCSnapIn
  myMMCSnapIn.Name = "Event Viewer"
  myMMCSnapIn.FileName = "eventvwr.msc"
  EventViewer = myMMCSnapIn
End Function

'01/04/2024  08:22           144,909 fsmgmt.msc
'01/04/2024  08:22           144,998 lusrmgr.msc

Public Function PerformanceMonitor() As Phosphorus.WindowsMMCSnapIn
  Dim myMMCSnapIn As Phosphorus.WindowsMMCSnapIn
  myMMCSnapIn.Name = "Performance Monitor"
  myMMCSnapIn.FileName = "perfmon.msc"
  PerformanceMonitor = myMMCSnapIn
End Function

Public Function WindowsServices() As Phosphorus.WindowsMMCSnapIn
  Dim myMMCSnapIn As Phosphorus.WindowsMMCSnapIn
  myMMCSnapIn.Name = "Windows Services"
  myMMCSnapIn.FileName = "services.msc"
  WindowsServices = myMMCSnapIn
End Function

'01/04/2024  08:22           145,059 taskschd.msc
'01/04/2024  08:22           144,862 tpm.msc
'01/04/2024  08:22           115,109 WF.msc

Public Function WindowsManagementInstrumentation() As Phosphorus.WindowsMMCSnapIn
  Dim myMMCSnapIn As Phosphorus.WindowsMMCSnapIn
  myMMCSnapIn.Name = "Windows Management Instrumentation"
  myMMCSnapIn.FileName = "WmiMgmt.msc"
  WindowsManagementInstrumentation = myMMCSnapIn
End Function


