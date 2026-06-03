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
  Chrome = myExecutable
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

' ************************
' * WINDOWS SYSTEM TOOLS *
' ************************

'C:\WINDOWS\system32\*.exe
'These will always be on the PATH so can be called just by their file name

'26/01/2025  00:12            40,960 agentactivationruntimestarter.exe
'10/09/2025  15:05           315,392 AggregatorHost.exe
'10/09/2025  15:02         3,258,736 aitstatic.exe
'10/09/2025  15:02           212,992 alg.exe
'30/04/2021  09:53           517,536 amdfendrsr.exe
'13/07/2021  05:45           474,936 amdlogum.exe
'11/06/2025  13:15           139,264 AppHostRegistrationVerifier.exe
'12/02/2025  17:20            49,152 appidcertstorecheck.exe
'12/02/2025  17:20           155,648 appidpolicyconverter.exe
'12/02/2025  17:20            49,152 appidtel.exe
'10/09/2025  15:02            79,232 AppInstallerBackgroundUpdate.exe
'14/05/2025  17:50            96,504 ApplicationFrameHost.exe
'10/09/2025  15:02         1,348,488 ApplyTrustOffline.exe
'26/01/2025  00:12           258,048 ApproveChildRequest.exe
'24/01/2025  19:47           162,760 appverif.exe
'26/01/2025  00:12            49,152 ARP.EXE
'26/01/2025  00:11            49,152 at.exe
'10/09/2025  15:02           151,552 AtBroker.exe
'13/07/2021  05:46           464,720 atieah64.exe
'13/07/2021  05:47           829,776 atieclxx.exe
'01/04/2024  08:22            45,056 attrib.exe
'10/09/2025  15:02           881,592 audiodg.exe
'12/02/2025  17:20            61,440 auditpol.exe
'26/01/2025  00:12           158,520 AuthHost.exe
'01/04/2024  08:22         1,044,480 autochk.exe
'10/09/2025  15:02            94,208 autofstx.exe
'26/01/2025  00:12            86,016 AxInstUI.exe
'01/04/2024  08:22            50,496 backgroundTaskHost.exe
'01/04/2024  08:22            61,440 BackgroundTransferHost.exe
'09/07/2025  13:44           282,624 bcdboot.exe
'10/09/2025  15:02           521,632 bcdedit.exe
'26/01/2025  00:12            81,920 BdeUISrv.exe
'09/07/2025  13:45           286,560 bdeunlock.exe
'10/09/2025  15:05           679,120 BioIso.exe
'10/09/2025  15:04           192,512 BitLockerDeviceEncryption.exe
'26/01/2025  00:12           126,976 BitLockerWizardElev.exe
'26/01/2025  00:11           241,664 bitsadmin.exe
'10/09/2025  15:02            53,248 bootim.exe
'11/06/2025  13:14           112,032 bootsect.exe
'26/01/2025  00:12            49,152 bridgeunattend.exe
'10/09/2025  15:02           159,744 browserexport.exe
'26/01/2025  00:11            67,016 browser_broker.exe
'11/06/2025  13:14            69,632 bthudtask.exe
'26/01/2025  00:11           114,688 ByteCodeGenerator.exe
'01/04/2024  08:22            61,440 cacls.exe
'26/01/2025  00:12            49,152 calc.exe

'26/01/2025  00:12            63,336 CameraSettingsUIHost.exe
'26/01/2025  00:12            92,400 CastSrv.exe
'10/09/2025  15:06           118,784 CertEnrollCtrl.exe
'10/09/2025  15:02           532,480 certreq.exe
'10/09/2025  15:02         1,585,152 certutil.exe
'26/01/2025  00:12           129,464 changepk.exe

'10/09/2025  15:02           286,720 charmap.exe
Public Function WindowsCharacterMap() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Character Map"
  myExecutable.ExeFile = "charmap.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsCharacterMap = myExecutable
End Function

'26/01/2025  00:12            69,632 CheckNetIsolation.exe
'26/01/2025  00:11            53,248 chkdsk.exe
'01/04/2024  08:22            45,056 chkntfs.exe
'01/04/2024  08:22            57,344 choice.exe
'10/09/2025  15:02            65,536 CIDiag.exe
'26/01/2025  00:11            77,824 cipher.exe
'26/01/2025  00:11           382,368 CiTool.exe

'10/09/2025  15:02           307,200 cleanmgr.exe
Public Function WindowsDiskCleanUp() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Disk Clean Up"
  myExecutable.ExeFile = "cleanmgr.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsDiskCleanUp = myExecutable
End Function

'26/01/2025  00:12            53,248 cliconfg.exe
'13/07/2021  05:47           355,120 clinfo.exe
'01/04/2024  08:22            57,344 clip.exe
'10/09/2025  15:02           144,792 ClipRenew.exe
'10/09/2025  15:02         1,129,488 ClipUp.exe
'09/07/2025  13:44           189,864 ClipUpgrade.exe
'26/01/2025  00:11            95,648 CloudExperienceHostBroker.exe
'09/04/2025  18:01           113,072 CloudNotifications.exe
'10/09/2025  15:02           344,064 cmd.exe
'26/01/2025  00:11            49,152 cmdkey.exe
'26/01/2025  00:12            77,824 cmdl32.exe
'10/09/2025  15:04           106,496 cmmon32.exe
'26/01/2025  00:12           122,880 cmstp.exe
'26/01/2025  00:12            45,056 cofire.exe
'26/01/2025  00:12           110,592 colorcpl.exe
'01/04/2024  08:22            49,152 comp.exe
'01/04/2024  08:22            73,728 compact.exe
'10/09/2025  15:02           186,736 CompatTelRunner.exe
'26/01/2025  00:12           114,688 CompMgmtLauncher.exe
'13/08/2025  17:23           118,784 CompPkgSrv.exe
'26/01/2025  00:12            73,728 ComputerDefaults.exe
'10/09/2025  15:02         1,007,616 conhost.exe
'10/09/2025  15:04           259,496 consent.exe
'11/06/2025  13:14           122,880 control.exe
'26/01/2025  00:12            45,056 convert.exe
'10/09/2025  15:05           238,992 convertvhd.exe
'10/09/2025  15:03            81,920 coredpussvr.exe
'10/09/2025  15:03           451,848 CredentialEnrollmentManager.exe
'10/09/2025  15:02           220,456 CredentialUIBroker.exe
'12/03/2025  14:20            65,536 credwiz.exe
'10/09/2025  15:04           192,512 cscript.exe
'01/04/2024  08:22            38,576 csrss.exe
'09/04/2025  18:02            69,632 ctfmon.exe
'01/04/2024  08:22           114,688 cttune.exe
'10/09/2025  15:03            65,536 cttunesvr.exe
'13/08/2025  17:23           727,088 curl.exe
'10/09/2025  15:02           163,840 CustomInstallExec.exe
'10/09/2025  15:04           200,704 dasHost.exe
'11/06/2025  13:14           267,720 DataExchangeHost.exe
'13/08/2025  17:24           200,704 DataStoreCacheDumpTool.exe
'14/05/2025  17:49           151,552 dccw.exe
'01/04/2024  08:22            32,768 dcomcnfg.exe
'01/04/2024  08:22            69,632 ddodiag.exe
'10/09/2025  15:03           241,664 Defrag.exe
'01/04/2024  08:22            61,440 deploymentcsphelper.exe
'10/09/2025  15:04           155,648 desktopimgdownldr.exe
'10/09/2025  15:03           148,888 DeviceCensus.exe
'26/01/2025  00:11           122,880 DeviceCredentialDeployment.exe
'01/04/2024  08:22            49,152 DeviceEject.exe
'10/09/2025  15:03           565,248 DeviceEnroller.exe
'26/01/2025  00:12           122,880 DevicePairingWizard.exe
'01/04/2024  08:22            94,208 DeviceProperties.exe
'26/01/2025  00:12            77,824 DFDWiz.exe
'10/09/2025  15:03           143,360 dfrgui.exe
'13/07/2021  05:47           501,592 dgtrayicon.exe
'10/09/2025  15:05            69,632 dialer.exe
'09/04/2025  18:01           176,128 directxdatabaseupdater.exe
'26/01/2025  00:11           176,128 diskpart.exe
'26/01/2025  00:12            49,152 diskperf.exe
'01/04/2024  08:22           356,352 diskraid.exe
'26/01/2025  00:12            98,304 DiskSnapshot.exe
'26/01/2025  00:12            73,728 diskusage.exe
'10/09/2025  15:03           329,128 Dism.exe
'10/09/2025  15:03           204,800 dispdiag.exe
'26/01/2025  00:12         1,877,392 DisplaySwitch.exe
'01/04/2024  08:22           106,496 djoin.exe
'26/01/2025  00:12            50,504 dllhost.exe
'26/01/2025  00:11            40,960 dllhst3g.exe
'10/09/2025  15:03           204,800 dmcertinst.exe
'01/04/2024  08:22            65,536 dmcfghost.exe
'09/04/2025  18:01           172,032 dmclient.exe
'01/04/2024  08:22            53,248 DmNotificationBroker.exe
'10/09/2025  15:03            61,440 DmOmaCpMo.exe
'12/08/2015  17:03            96,528 dns-sd.exe
'01/04/2024  08:22            61,440 dnscacheugc.exe
'01/04/2024  08:22            45,056 doskey.exe
'01/04/2024  08:22           102,400 dpapimig.exe
'01/04/2024  08:22           102,400 DpiScaling.exe
'11/06/2025  13:14            36,864 dpnsvr.exe
'11/06/2025  13:14           106,496 driverquery.exe
'10/09/2025  15:02           438,272 drvinst.exe
'11/06/2025  13:14            57,344 DsmUserTask.exe
'10/09/2025  15:05           524,288 dsregcmd.exe
'26/01/2025  00:12            40,960 dstokenclean.exe
'10/09/2025  15:05           159,744 dtdump.exe
'12/02/2025  17:20            77,824 dusmtask.exe
'26/01/2025  00:12            40,960 dvdplay.exe
'09/04/2025  18:01           131,072 dwm.exe
'10/09/2025  15:03           262,144 DWWIN.EXE
'10/09/2025  15:03           327,680 dxdiag.exe
'09/04/2025  18:01           131,072 dxgiadaptercache.exe
'10/09/2025  15:03           311,296 Dxpserver.exe
'14/05/2025  17:49            40,960 Eap3Host.exe
'10/09/2025  15:05           319,488 EaseOfAccessDialog.exe
'10/09/2025  15:05           100,608 easinvoker.exe
'13/08/2025  17:23            94,208 EASPolicyManagerBrokerHost.exe
'10/09/2025  15:02           167,936 EDPCleanup.exe
'26/01/2025  00:11            98,304 edpnotify.exe
'10/09/2025  15:04           143,360 EduPrintProv.exe
'13/07/2021  05:47           441,144 EEURestart.exe
'01/04/2024  08:22            40,960 efsui.exe
'26/01/2025  00:12           151,552 EhStorAuthn.exe
'11/06/2025  13:14           176,128 EoAExperiences.exe
'09/04/2025  18:01           847,872 esentutl.exe
'26/01/2025  00:12            62,768 esimtool.exe
'26/01/2025  00:12           380,928 eudcedit.exe
'11/06/2025  13:14            69,632 eventcreate.exe
'10/09/2025  15:02           106,496 eventvwr.exe
'01/04/2024  08:22            73,728 expand.exe
'26/01/2025  00:11            61,440 extrac32.exe
'01/04/2024  08:22            49,152 fc.exe
'26/01/2025  00:12           163,840 fhmanagew.exe
'10/09/2025  15:05           159,744 FileDialogBroker.exe
'10/09/2025  15:03           258,048 FileHistory.exe
'01/04/2024  08:22            40,960 find.exe
'26/01/2025  00:12            65,536 findstr.exe
'26/01/2025  00:12            40,960 finger.exe
'26/01/2025  00:12            49,152 fixmapi.exe
'26/01/2025  00:12            53,248 fltMC.exe
'26/01/2025  00:12            73,728 fodhelper.exe
'26/01/2025  00:12           139,264 Fondue.exe
'14/05/2025  17:49           832,016 fontdrvhost.exe
'26/01/2025  00:11           151,552 fontview.exe
'26/01/2025  00:12            73,728 forfiles.exe
'01/04/2024  08:22            40,960 fsavailux.exe
'10/09/2025  15:04           137,832 FsIso.exe
'10/09/2025  15:02           172,032 fsquirt.exe
'10/09/2025  15:03           279,936 fsutil.exe
'26/01/2025  00:12            86,016 ftp.exe
'26/01/2025  00:12           217,088 fvenotify.exe
'10/09/2025  15:03           278,528 FXSCOVER.exe
'11/06/2025  13:14           716,800 FXSSVC.exe
'09/04/2025  18:01            77,824 FXSUNATD.exe
'26/01/2025  00:12           372,736 GameBarPresenceWriter.exe
'10/09/2025  15:02            80,232 GameInputSvc.exe
'10/09/2025  15:06         1,335,296 GamePanel.exe
'11/06/2025  13:14           658,480 GenValObj.exe
'01/04/2024  08:22           106,496 getmac.exe
'10/09/2025  15:03           290,816 gpresult.exe
'12/03/2025  14:20            57,344 gpupdate.exe
'01/04/2024  08:22            45,056 grpconv.exe
'01/04/2024  08:22            81,920 hdwwiz.exe
'01/04/2024  08:22            32,768 help.exe
'26/01/2025  00:12            40,960 HOSTNAME.EXE
'25/07/2020  11:28           144,680 HPMUIDir.exe
'10/09/2025  15:02         1,971,600 hvax64.exe
'10/09/2025  15:02         2,053,536 hvix64.exe
'26/01/2025  00:12            61,440 icacls.exe
'23/08/2025  23:41            56,128 icarus_rvrt.exe
'11/06/2025  13:14            51,064 icsunattend.exe
'10/09/2025  15:03           290,816 ie4uinit.exe
'10/09/2025  15:03           135,168 ie4ushowIE.exe
'10/09/2025  15:03           561,152 IESettingSync.exe
'26/01/2025  00:12           102,400 ieUnatt.exe
'10/09/2025  15:03           196,608 iexpress.exe
'12/03/2025  14:20           147,456 immersivetpmvscmgrsvr.exe
'01/04/2024  08:22            36,864 InfDefaultInstall.exe
'11/06/2025  13:14           135,168 InputSwitchToastHandler.exe
'10/09/2025  15:03            61,440 ipconfig.exe
'12/02/2025  17:20            73,728 iscsicli.exe
'10/09/2025  15:03            32,768 iscsicpl.exe
'14/05/2025  17:49            61,440 ISM.exe
'10/09/2025  15:03           147,456 isoburn.exe
'01/04/2024  08:22            65,536 klist.exe
'26/01/2025  00:12            65,536 ksetup.exe
'26/01/2025  00:12            45,056 ktmutil.exe
'26/01/2025  00:11            40,960 la57setup.exe
'01/04/2024  08:22            40,960 label.exe
'12/02/2025  17:20            77,824 LanguageComponentsInstallerComHandler.exe
'10/09/2025  15:02            32,768 LaunchTM.exe
'10/09/2025  15:05            86,016 LaunchWinApp.exe
'10/09/2025  15:04           225,280 LegacyNetUXHost.exe
'26/01/2025  00:11            77,824 LicenseManagerShellext.exe
'10/09/2025  15:02           647,168 licensingdiag.exe
'26/01/2025  00:12           174,904 LicensingUI.exe
'10/09/2025  15:04           188,416 LiveCaptions.exe
'01/04/2024  08:22            32,768 Locator.exe
'10/09/2025  15:04           100,600 LockAppHost.exe
'09/07/2025  13:44            75,832 LockScreenContentServer.exe
'26/01/2025  00:12           106,496 lodctr.exe
'10/09/2025  15:04           131,072 logagent.exe
'26/01/2025  00:12           118,784 logman.exe
'10/09/2025  15:02            86,016 LogonUI.exe
'10/09/2025  15:04            61,440 lpkinstall.exe
'10/09/2025  15:04           745,472 lpksetup.exe
'10/09/2025  15:04           131,072 lpremove.exe
'10/09/2025  15:05           394,040 LsaIso.exe
'10/09/2025  15:04            84,096 lsass.exe
'10/09/2025  15:04           872,448 Magnify.exe
'01/04/2024  08:22           110,592 makecab.exe
'10/09/2025  15:04           286,720 manage-bde.exe
'26/01/2025  00:12           843,776 mblctr.exe
'10/09/2025  15:05           360,448 MBR2GPT.EXE
'12/02/2025  17:20           131,072 mcbuilder.exe
'10/09/2025  15:03           487,424 MDEServer.exe
'10/09/2025  15:04           176,128 MDMAgent.exe
'10/09/2025  15:04           196,608 MDMAppInstaller.exe
'10/09/2025  15:03            90,112 MdmDiagnosticsTool.exe
'10/09/2025  15:04           110,592 MdRes.exe
'10/09/2025  15:04           155,648 MdSched.exe
'11/06/2025  13:14            75,864 mfpmp.exe
'24/01/2025  19:46            50,768 microsoft.windows.softwarelogo.showdesktop.exe
'11/06/2025  13:14           114,688 MicrosoftEdgeBCHost.exe
'26/01/2025  00:11           114,688 MicrosoftEdgeCP.exe
'11/06/2025  13:14           114,688 MicrosoftEdgeDevTools.exe
'26/01/2025  00:11            81,920 MicrosoftEdgeSH.exe
'10/09/2025  15:05            86,016 MLEngineStub.exe
'10/09/2025  15:04         1,863,680 mmc.exe
'26/01/2025  00:12         1,327,104 mmgaserver.exe
'26/01/2025  00:12           131,072 mobsync.exe
'10/09/2025  15:05            86,016 MoNotificationUxStub.exe
'01/04/2024  08:22            45,056 mountvol.exe
'01/04/2024  08:22            45,056 mpnotify.exe
'12/01/2024  19:09           918,944 MpSigStub.exe
'10/09/2025  15:04           162,624 MptfGenericService.exe
'26/01/2025  00:12            40,960 MRINFO.EXE
'16/08/2025  21:39       223,939,376 MRT.exe
'01/04/2024  08:22           106,496 MSchedExe.exe
'10/09/2025  15:04           253,952 msconfig.exe
'10/09/2025  15:04           557,056 msdt.exe
'11/06/2025  13:14           208,896 msdtc.exe
'26/01/2025  00:12            36,864 msfeedssync.exe
'26/01/2025  00:12            36,864 mshta.exe
'10/09/2025  15:03           180,224 msiexec.exe
'10/09/2025  15:04           380,928 msinfo32.exe
'10/09/2025  15:04           614,400 msra.exe
'09/07/2025  13:45           135,168 MsSpellCheckingHost.exe
'10/09/2025  15:05         1,388,544 mstsc.exe
'26/01/2025  00:12           155,648 mtstocom.exe
'09/04/2025  18:01           118,784 MuiUnattend.exe
'26/01/2025  00:12            81,920 MultiDigiMon.exe
'10/09/2025  15:04         1,060,864 Narrator.exe
'26/01/2025  00:11            45,056 nbtstat.exe
'26/01/2025  00:12            94,208 ndadmin.exe
'26/01/2025  00:12            58,800 NDKPerfCmd.exe
'26/01/2025  00:12            58,800 NDKPing.exe
'11/06/2025  13:14            81,920 net.exe
'26/01/2025  00:12           184,320 net1.exe
'01/04/2024  08:22            49,152 netbtugc.exe
'10/09/2025  15:04           114,688 netcfg.exe
'26/01/2025  00:12            98,304 NetCfgNotifyObjectHost.exe
'26/01/2025  00:12            57,344 NetEvtFwdr.exe
'01/04/2024  08:22            36,864 NetHost.exe
'26/01/2025  00:12            57,344 netiougc.exe
'26/01/2025  00:11            69,632 Netplwiz.exe
'10/09/2025  15:04           126,976 netsh.exe
'13/08/2025  17:24            69,632 NETSTAT.EXE
'26/01/2025  00:12            94,208 newdev.exe
'10/09/2025  15:05           736,952 NgcIso.exe
'30/09/2025  13:08           322,216 nllBoot.exe
'10/09/2025  15:05           602,112 nltest.exe
'10/09/2025  15:04           360,448 notepad.exe
'10/09/2025  15:04           114,688 nslookup.exe
'10/09/2025  15:04        12,952,992 ntoskrnl.exe
'10/09/2025  15:04            94,208 ntprint.exe
'10/09/2025  15:04           102,400 odbcad32.exe
'01/04/2024  08:22            49,152 odbcconf.exe
'10/09/2025  15:04           102,400 ofdeploy.exe
'10/09/2025  15:03           589,824 omadmclient.exe
'10/09/2025  15:03           167,936 omadmprc.exe
'10/09/2025  15:04        89,771,848 OneDriveSetup.exe
'26/01/2025  00:11            40,960 OOBEFodSetup.exe
'14/05/2025  17:50            98,304 OobeShellHost.exe
'11/06/2025  13:14            94,208 openfiles.exe
'10/09/2025  15:04           170,848 OpenWith.exe
'26/01/2025  00:12           139,264 OptionalFeatures.exe
'10/09/2025  15:04           589,824 osk.exe
'10/09/2025  15:06            46,912 pacjsworker.exe
'26/01/2025  00:11            69,632 PackagedCWALauncher.exe
'09/04/2025  18:01           483,328 pairtool.exe
'09/04/2025  18:02            92,392 PasswordOnWakeSettingFlyout.exe
'10/09/2025  15:04            45,056 PATHPING.EXE
'10/09/2025  15:02           126,976 pcalua.exe
'10/09/2025  15:02           253,952 pcaui.exe
'26/01/2025  00:12            40,960 pcwrun.exe
'10/09/2025  15:04           180,224 perfmon.exe
'26/01/2025  00:12           133,688 phoneactivate.exe
'10/09/2025  15:04           162,608 PickerHost.exe
'09/04/2025  18:02           135,168 PinEnrollmentBroker.exe
'26/01/2025  00:12            45,056 PING.EXE
'10/09/2025  15:04           294,912 PkgMgr.exe
'10/09/2025  15:04           705,920 PktMon.exe
'26/01/2025  00:12            32,768 plasrv.exe
'01/04/2024  08:22            77,824 PnPUnattend.exe
'10/09/2025  15:04           286,720 pnputil.exe
'27/08/2025  03:14           651,264 poqexec.exe
'10/09/2025  15:04           192,512 powercfg.exe
'09/04/2025  18:02           278,528 PresentationHost.exe
'10/09/2025  15:05            61,440 prevhost.exe
'01/04/2024  08:22            40,960 print.exe
'10/09/2025  15:04           585,728 printfilterpipelinesvc.exe
'10/09/2025  15:04           122,880 PrintIsolationHost.exe
'10/09/2025  15:04            94,208 printui.exe
'26/01/2025  00:11            73,728 proquota.exe
'12/03/2025  14:20            86,016 provlaunch.exe
'12/03/2025  14:20           114,688 provtool.exe
'10/09/2025  15:04           286,568 ProximityUxHost.exe
'26/01/2025  00:11            46,936 prproc.exe
'10/09/2025  15:02           389,120 psr.exe
'26/01/2025  00:12            57,344 pwlauncher.exe
'10/09/2025  15:04            45,056 rasautou.exe
'10/09/2025  15:04            49,152 rasdial.exe
'10/09/2025  15:04           159,744 raserver.exe
'10/09/2025  15:04            65,536 rasphone.exe
'09/07/2025  13:45           552,960 rdpclip.exe
'10/09/2025  15:05           258,048 rdpinput.exe
'26/01/2025  00:12            81,920 RdpSa.exe
'26/01/2025  00:12            65,536 RdpSaProxy.exe
'26/01/2025  00:12            61,440 RdpSaUacHelper.exe
'10/09/2025  15:04            77,824 rdrleakdiag.exe
'11/06/2025  13:14           106,496 readCloudDataSettings.exe
'10/09/2025  15:06           114,688 ReAgentc.exe
'10/09/2025  15:04           217,088 recdisc.exe
'01/04/2024  08:22            40,960 recover.exe
'10/09/2025  15:04           356,352 RecoveryDrive.exe
'10/09/2025  15:05         2,191,360 ReFsDedupSvc.exe
'10/09/2025  15:04         1,945,600 refsutil.exe
'10/09/2025  15:04           110,592 reg.exe
'10/09/2025  15:04            32,768 regedt32.exe
'10/09/2025  15:04            69,632 regini.exe
'01/04/2024  08:22            53,248 Register-CimProvider.exe
'10/09/2025  15:04            94,208 regsvr32.exe
'26/01/2025  00:11           147,456 rekeywiz.exe
'26/01/2025  00:12            77,824 relog.exe
'10/09/2025  15:04           258,048 RelPost.exe
'01/04/2024  08:22            40,960 RemotePosWorker.exe
'13/08/2025  17:24           155,648 repair-bde.exe
'01/04/2024  08:22            45,056 replace.exe
'10/09/2025  15:05            42,368 ResetEngine.exe
'10/09/2025  15:04           139,264 resmon.exe
'10/09/2025  15:05           150,200 rgnupdt.exe
'11/06/2025  13:14           593,920 RMActivate.exe
'11/06/2025  13:14           622,592 RMActivate_isv.exe
'11/06/2025  13:14           524,288 RMActivate_ssp.exe
'11/06/2025  13:14           524,288 RMActivate_ssp_isv.exe
'09/04/2025  18:01            45,056 RmClient.exe
'12/03/2025  14:20           143,360 rmttpmvscmgrsvr.exe
'10/09/2025  15:04           208,896 Robocopy.exe
'26/01/2025  00:12            49,152 ROUTE.EXE
'26/01/2025  00:11            57,344 RpcPing.exe
'26/01/2025  00:12            77,824 rrinstaller.exe
'11/06/2025  13:14           335,872 rstrui.exe
'13/10/2020  09:07         1,085,224 RtkAudUService64.exe
'01/04/2024  08:22            45,056 runas.exe
'10/09/2025  15:04            98,304 rundll32.exe
'10/09/2025  15:05           110,592 runexehelper.exe
'26/01/2025  00:11            53,248 RunLegacyCPLElevated.exe
'10/09/2025  15:04           122,880 runonce.exe
'10/09/2025  15:02           133,656 RuntimeBroker.exe
'09/07/2025  13:45           102,400 sc.exe
'11/06/2025  13:14           290,816 schtasks.exe
'10/09/2025  15:02           299,008 sdbinst.exe
'26/01/2025  00:12            98,304 sdchange.exe
'10/09/2025  15:04         1,097,728 sdclt.exe
'26/01/2025  00:12            77,824 sdiagnhost.exe
'10/09/2025  15:06           491,520 SearchFilterHost.exe
'10/09/2025  15:06         1,015,808 SearchIndexer.exe
'10/09/2025  15:06           700,416 SearchProtocolHost.exe
'26/01/2025  00:11            73,728 SecEdit.exe
'01/04/2024  08:22            32,768 secinit.exe
'13/10/2020  09:07           161,096 SECOMN64.exe
'10/09/2025  15:02         1,320,320 securekernel.exe
'09/04/2025  18:02           120,240 SecurityHealthHost.exe
'09/04/2025  18:02           146,128 SecurityHealthService.exe
'26/01/2025  00:12           270,336 SecurityHealthSystray.exe
'13/08/2025  17:24         1,191,936 SensorDataService.exe
'10/09/2025  15:05           106,496 SensorRuntimeBroker.exe
'10/09/2025  15:05           906,400 services.exe
'10/09/2025  15:05           108,864 sessionmsg.exe
'10/09/2025  15:05           159,744 sethc.exe
'01/04/2024  08:22            57,344 setspn.exe
'01/04/2024  08:22            28,672 setupcl.exe
'10/09/2025  15:03           249,856 setupugc.exe
'01/04/2024  08:22            86,016 setx.exe
'13/08/2025  17:24           110,592 sfc.exe
'10/09/2025  15:05         1,650,104 ShellAppRuntime.exe
'10/09/2025  15:05           544,768 ShellHost.exe
'11/06/2025  13:14            77,824 shrpubw.exe
'01/04/2024  08:22            57,344 shutdown.exe
'01/04/2024  08:22           102,400 sigverif.exe
'10/09/2025  15:06           549,104 SIHClient.exe
'10/09/2025  15:05           139,264 sihost.exe
'01/04/2024  08:22            42,696 SlideToShutDown.exe
'10/09/2025  15:05           716,800 slui.exe
'10/09/2025  15:05           610,304 smartscreen.exe
'10/09/2025  15:05           228,768 smss.exe
'10/09/2025  15:02           311,360 SndVol.exe
'26/01/2025  00:11            40,960 snmptrap.exe
'26/01/2025  00:12            49,152 sort.exe
'10/09/2025  15:05           229,376 SpaceAgent.exe
'10/09/2025  15:05           112,024 spaceman.exe
'10/09/2025  15:05           536,576 spaceutil.exe
'13/08/2025  17:23           176,128 SpatialAudioLicenseSrv.exe
'10/09/2025  15:04           987,136 spoolsv.exe
'10/09/2025  15:04           958,464 spoolsvworker.exe
'10/09/2025  15:05           577,536 SppExtComObj.Exe
'10/09/2025  15:05         4,823,520 sppsvc.exe
'01/04/2024  08:22            36,864 srdelayed.exe
'01/04/2024  08:22            81,920 SrTasks.exe
'10/09/2025  15:05           196,608 stordiag.exe
'01/04/2024  08:22            40,960 subst.exe
'26/01/2025  00:12           936,368 sudo.exe
'10/09/2025  15:05            88,232 svchost.exe
'26/01/2025  00:12            65,536 sxstrace.exe
'10/09/2025  15:05            69,632 SyncHost.exe
'26/08/2020  12:03         4,305,240 SynTPEnh.exe
'26/08/2020  12:02           375,640 SynTPEnhService.exe
'10/09/2025  15:05            75,160 SysResetErr.exe
'11/06/2025  13:14           126,976 systeminfo.exe
'01/04/2024  08:22           106,496 SystemPropertiesAdvanced.exe
'01/04/2024  08:22           106,496 SystemPropertiesComputerName.exe
'01/04/2024  08:22           106,496 SystemPropertiesDataExecutionPrevention.exe
'01/04/2024  08:22           106,496 SystemPropertiesHardware.exe
'01/04/2024  08:22           106,496 SystemPropertiesPerformance.exe
'01/04/2024  08:22           106,496 SystemPropertiesProtection.exe
'01/04/2024  08:22           106,496 SystemPropertiesRemote.exe
'10/09/2025  15:05           790,688 SystemSettingsAdminFlows.exe
'10/09/2025  15:05           232,824 SystemSettingsBroker.exe
'10/09/2025  15:05            67,544 SystemSettingsRemoveDevice.exe
'14/05/2025  17:49           118,784 SystemUWPLauncher.exe
'01/04/2024  08:22            32,768 systray.exe
'26/01/2025  00:12           114,688 tabcal.exe
'01/04/2024  08:22            90,112 takeown.exe
'10/09/2025  15:05            40,960 TapiUnattend.exe
'10/09/2025  15:01            92,192 tar.exe
'10/09/2025  15:05           117,160 taskhostw.exe
'01/04/2024  08:22           118,784 taskkill.exe
'01/04/2024  08:22           122,880 tasklist.exe

Public Function WindowsTaskManager() As Phosphorus.Executable
  Dim myExecutable As Phosphorus.Executable
  myExecutable.Name = "Windows Task Manager"
  myExecutable.ExeFile = "taskmgr.exe"
  myExecutable.FullPath = myExecutable.ExeFile
  WindowsTaskManager = myExecutable
End Function

'10/09/2025  15:02           974,608 tcblaunch.exe
'10/09/2025  15:04            45,056 tcmsetup.exe
'26/01/2025  00:12            36,864 TCPSVCS.EXE
'10/09/2025  15:05            90,112 ThumbnailExtractionHost.exe
'10/09/2025  15:05           348,160 TieringEngineService.exe
'01/04/2024  08:22            57,344 timeout.exe
'26/01/2025  00:12            77,824 TokenBrokerCookies.exe
'01/04/2024  08:22            94,208 TpmInit.exe
'10/09/2025  15:05           385,024 TpmTool.exe
'12/03/2025  14:20           126,976 tpmvscmgr.exe
'12/03/2025  14:20           143,360 tpmvscmgrsvr.exe
'10/09/2025  15:04           430,080 tracerpt.exe
'26/01/2025  00:12            40,960 TRACERT.EXE
'26/01/2025  00:12            94,208 TSTheme.exe
'10/09/2025  15:05           118,784 TSWbPrxy.exe
'26/01/2025  00:08            37,864 TsWpfWrp.exe
'10/09/2025  15:05           311,336 ttdinject.exe
'10/09/2025  15:05           137,808 tttracer.exe
'26/01/2025  00:12            81,920 typeperf.exe
'10/09/2025  15:05           236,032 tzsync.exe
'01/04/2024  08:22            65,536 tzutil.exe
'11/06/2025  13:14            57,344 UCConfigTask.exe
'10/09/2025  15:05           185,344 UCPDMgr.exe
'26/01/2025  00:12            75,168 ucsvc.exe
'10/09/2025  15:05           443,808 UIEOrchestrator.exe
'11/06/2025  13:14            69,632 UIMgrBroker.exe
'10/09/2025  15:05            81,920 UndockedFlightingUpdateTask.exe
'26/01/2025  00:12            86,016 unlodctr.exe
'26/01/2025  00:12           262,144 unregmp2.exe
'10/09/2025  15:05           203,952 upfc.exe
'14/05/2025  17:50            77,824 UpgradeResultsUI.exe
'10/09/2025  15:05            73,728 upnpcont.exe
'10/09/2025  15:04           114,688 UPPrinterInstaller.exe
'26/01/2025  00:11            75,840 UserAccountBroker.exe
'10/09/2025  15:05           155,648 UserAccountControlSettings.exe
'09/07/2025  13:44            69,632 UserDataSource.exe
'10/09/2025  15:05           143,360 userinit.exe
'10/09/2025  15:05            86,016 UsoClient.exe
'10/09/2025  15:05           176,128 UtcDecoderHost.exe
'10/09/2025  15:05           315,392 Utilman.exe
'26/01/2025  00:11            57,344 VaultCmd.exe
'10/09/2025  15:05           741,376 vds.exe
'26/01/2025  00:12            77,824 vdsldr.exe
'01/04/2024  08:22            45,056 verclsid.exe
'26/01/2025  00:11           214,448 verifier.exe
'26/01/2025  00:11           204,800 verifiergui.exe
'10/09/2025  15:05           262,144 VoiceAccess.exe
'10/09/2025  15:05           167,936 vssadmin.exe
'10/09/2025  15:05         1,474,560 VSSVC.exe
'13/07/2021  05:47         1,866,552 vulkaninfo-1-999-0-0-0.exe
'13/07/2021  05:47         1,866,552 vulkaninfo.exe
'11/06/2025  13:14           274,432 w32tm.exe
'01/04/2024  08:22            65,536 waitfor.exe
'26/01/2025  00:11            49,152 WallpaperHost.exe
'11/06/2025  13:14           380,928 wbadmin.exe
'09/07/2025  13:44         1,540,096 wbengine.exe
'01/04/2024  08:22           131,072 wecutil.exe
'10/09/2025  15:03           640,424 WerFault.exe
'10/09/2025  15:03           220,488 WerFaultSecure.exe
'10/09/2025  15:03           300,440 wermgr.exe
'26/01/2025  00:11           299,008 wevtutil.exe
'10/09/2025  15:03           172,032 wextract.exe
'10/09/2025  15:03           978,944 WFS.exe
'01/04/2024  08:22            65,536 where.exe
'26/01/2025  00:12            98,304 whoami.exe
'10/09/2025  15:05           126,976 wiaacmgr.exe
'10/09/2025  15:05            65,536 wiawow64.exe
'26/01/2025  00:12           185,776 wifitask.exe
'10/09/2025  15:05           640,384 wimserv.exe
'26/01/2025  00:12           110,592 WinBioDataModelOOBE.exe
'26/01/2025  00:12            40,960 Windows.Media.BackgroundPlayback.exe
'26/01/2025  00:11            94,208 Windows.WARP.JITService.exe
'10/09/2025  15:03           106,496 WindowsActionDialog.exe
'12/02/2025  17:21            69,632 WindowsUpdateElevatedInstaller.exe
'10/09/2025  15:06           790,680 wininit.exe
'10/09/2025  15:02         1,864,600 winload.exe
'10/09/2025  15:06           954,368 winlogon.exe
'10/09/2025  15:02         1,460,184 winresume.exe
'26/01/2025  00:12            77,824 winrs.exe
'26/01/2025  00:12            57,344 winrshost.exe
'26/01/2025  00:11            45,056 WinRTNetMUAHostServer.exe
'09/04/2025  18:02         2,768,896 WinSAT.exe
'01/04/2024  08:22            32,768 winver.exe
'10/09/2025  15:05           323,728 wkspbroker.exe
'10/09/2025  15:05           425,984 wksprt.exe
'11/06/2025  13:14           151,552 wlanext.exe
'10/09/2025  15:06           154,344 wlrmdr.exe
'10/09/2025  15:06         1,519,616 WMPDMC.exe
'10/09/2025  15:03           114,688 WorkFolders.exe
'10/09/2025  15:05            49,152 wowreg32.exe
'10/09/2025  15:04         1,224,552 WpcMon.exe
'10/09/2025  15:03           299,008 WpcTok.exe
'26/01/2025  00:12            57,344 WPDShextAutoplay.exe
'10/09/2025  15:04            49,152 wpnpinst.exe
'09/07/2025  13:44           401,408 wpr.exe
'26/01/2025  00:12            32,768 wscadminui.exe
'01/04/2024  08:22           102,400 WSCollect.exe
'10/09/2025  15:04           204,800 wscript.exe
'10/09/2025  15:04           258,048 wsl.exe
'11/06/2025  13:14            65,536 WSManHTTPConfig.exe
'26/01/2025  00:12            69,632 wsmprovhost.exe
'01/04/2024  08:22            86,016 wsqmcons.exe
'09/04/2025  18:02           139,264 WSReset.exe
'12/02/2025  17:21            45,984 wuapihost.exe
'10/09/2025  15:06           156,072 wuauclt.exe
'10/09/2025  15:03           232,880 WUDFCompanionHost.exe
'10/09/2025  15:03           357,776 WUDFHost.exe
'10/09/2025  15:06           241,664 wusa.exe
'10/09/2025  15:05           923,008 WWAHost.exe
'11/06/2025  13:15            61,440 XblGameSaveTask.exe
'01/04/2024  08:22            73,728 xcopy.exe
'10/09/2025  15:06         3,690,496 xpsrchvw.exe
'26/01/2025  00:12            94,208 xwizard.exe


