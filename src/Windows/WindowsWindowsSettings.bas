Attribute VB_Name = "WindowsWindowsSettings"
'@Folder Windows
Option Explicit

'https://learn.microsoft.com/en-us/windows/apps/develop/launch/launch-settings
'This topic describes the ms-settings: URI scheme.
'Use this URI scheme to launch Windows Settings to specific settings pages.

'https://gist.github.com/dbilanoski/8e931a0072b61ea215996a18fd215fb3

Public Type WindowsSetting
  Category As String
  SettingsPage As String
  URISuffix As String
End Type

Public Function Accounts_AccessWorkOrSchool() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Accounts"
  myWindowsSetting.SettingsPage = "Access work or school"
  myWindowsSetting.URISuffix = "workplace"
  Accounts_AccessWorkOrSchool = myWindowsSetting
End Function

'Accounts
'Settings page URI
'Email & app accounts  ms-settings:emailandaccounts
'Family & other people ms-settings:otherusers
'Provisioning  ms-settings:provisioning (only available on mobile and if the enterprise has deployed a provisioning package)
'ms-settings:workplace-provisioning (only available if enterprise has deployed a provisioning package)
'Repair token  ms-settings:workplace-repairtoken
'Set up a kiosk  ms-settings:assignedaccess
'Sign-in options ms-settings:signinoptions
'ms -Settings: signinoptions -dynamiclock
'Sync your settings  ms-settings:sync
'ms-settings:backup (Backup page deprecated in Windows 11)
'Windows Anywhere  ms-settings:windowsanywhere (device must be Windows Anywhere-capable)
'Windows Hello setup ms-settings:signinoptions-launchfaceenrollment
'ms -Settings: signinoptions -launchfingerprintenrollment
'Your info ms-settings:yourinfo

Public Function Accounts_YourInfo() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Accounts"
  myWindowsSetting.SettingsPage = "Your info"
  myWindowsSetting.URISuffix = "yourinfo"
  Accounts_YourInfo = myWindowsSetting
End Function

'Apps
'Settings page URI

Public Function Apps_AppsAndFeatures() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Apps"
  myWindowsSetting.SettingsPage = "Apps & Features"
  myWindowsSetting.URISuffix = "appsfeatures"
  Apps_AppsAndFeatures = myWindowsSetting
End Function

'App features  ms-settings:appsfeatures-app (Reset, manage add-on & downloadable content, etc. for the app)
Public Function Apps_AppsFeatures() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Apps"
  myWindowsSetting.SettingsPage = "App features"
  myWindowsSetting.URISuffix = "appsfeatures-app"
  Apps_AppsFeatures = myWindowsSetting
End Function

'
'To access this page with a URI, use the ms-settings:appsfeatures-app URI and pass an optional parameter of the package family name of the app.
'
'Example: ms-settings:appsfeatures-app?<PFN>
'Apps for websites ms-settings:appsforwebsites
'Default apps  ms-settings:defaultapps (Behavior introduced in Windows 11, version 21H2 (with 2023-04 Cumulative Update) or 22H2 (with 2023-04 Cumulative Update), or later.)
'Append the query string parameter in the following formats using the Uri-escaped name of an app to directly launch the default settings page for that app:
'
'- registeredAppMachine=<Uri-escaped per machine installed name of app>
'- registeredAppUser=<Uri-escaped per user installed name of app>
'- registeredAUMID=<Uri-escaped Application User Model ID>
'
'For more information, see Launch the Default Apps settings page.
'Default browser settings  ms-settings:defaultbrowsersettings (Deprecated in Windows 11)
'Manage optional features  ms-settings:optionalfeatures
'Offline Maps  ms-settings:maps
'ms-settings:maps-downloadmaps (Download maps)
'Startup apps  ms-settings:startupapps
'Video playback  ms-settings:videoplayback

'Control Center
Public Function ControlCenter_ControlCenter() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Control Center"
  myWindowsSetting.SettingsPage = "Control Center"
  myWindowsSetting.URISuffix = "controlcenter"
  ControlCenter_ControlCenter = myWindowsSetting
End Function

'Cortana
'Settings page URI
'Cortana across my devices ms-settings:cortana-notifications
'More details  ms-settings:cortana-moredetails
'Permissions & History ms-settings:cortana-permissions
'Searching Windows ms-settings:cortana-windowssearch
'Talk to Cortana ms-settings:cortana-language
'ms -Settings: Cortana
'ms -Settings: Cortana -talktocortana
' Important
'
'Cortana voice assistance in Windows as a standalone app was retired in the spring of 2023. For more information, see End of support for Cortana.
'
' Note
'
'This Settings section on desktop will be called Search when the PC is set to regions where Cortana is not currently available or Cortana has been disabled. Cortana-specific pages (Cortana across my devices, and Talk to Cortana) will not be listed in this case.
'
'devices
'Settings page URI

Public Function Devices_Autoplay() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Devices"
  myWindowsSetting.SettingsPage = "Autoplay"
  myWindowsSetting.URISuffix = "autoplay"
  Devices_Autoplay = myWindowsSetting
End Function

Public Function Devices_Bluetooth() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Devices"
  myWindowsSetting.SettingsPage = "Bluetooth"
  myWindowsSetting.URISuffix = "bluetooth"
  Devices_Bluetooth = myWindowsSetting
End Function

'Connected Devices ms-settings:connecteddevices
'Default camera  ms-settings:camera (Behavior deprecated in Windows 10, version 1809 and later)
'Camera settings ms-settings:camera (Behavior introduced in Windows 11, build 22000 and later) Append the query string parameter cameraId set to the Uri-escaped symbolic link name of a camera device to directly launch the settings for that camera. For more information, see Launch the camera settings page.
'Mouse & touchpad  ms-settings:mousetouchpad (touchpad settings only available on devices that have a touchpad)
'Pen & Windows Ink ms-settings:pen
'Printers & scanners ms-settings:printers
'Touch ms - Settings: devices -Touch
'Touchpad  ms-settings:devices-touchpad (only available if touchpad hardware is present)
'Text Suggestions  ms-settings:devicestyping-hwkbtextsuggestions
'typing ms - Settings: typing
'usb ms - Settings: usb
'Wheel ms-settings:wheel (only available if a Surface Dial device is paired)

Public Function Devices_YourPhone() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Devices"
  myWindowsSetting.SettingsPage = "Your Phone"
  myWindowsSetting.URISuffix = "mobile-devices"
  Devices_YourPhone = myWindowsSetting
End Function

'Ease of access
'Settings page URI
'Audio ms - Settings: easeofaccess -Audio
'Closed captions ms-settings:easeofaccess-closedcaptioning
'Color filters ms-settings:easeofaccess-colorfilter
'ms -Settings: easeofaccess -colorfilter - adaptivecolorlink
'ms -Settings: easeofaccess -colorfilter - bluelightlink
'display ms - Settings: easeofaccess -display
'Eye control ms-settings:easeofaccess-eyecontrol
'Hearing devices ms-settings:easeofaccess-hearingaids (Added in Windows 11, Version 24H2)
'High contrast ms-settings:easeofaccess-highcontrast
'keyboard ms - Settings: easeofaccess -keyboard
'magnifier ms - Settings: easeofaccess -magnifier
'Mouse ms - Settings: easeofaccess -Mouse
'Mouse pointer & touch ms-settings:easeofaccess-mousepointer
'narrator ms - Settings: easeofaccess -narrator
'ms -Settings: easeofaccess -narrator - isautostartenabled
'Speech ms - Settings: easeofaccess -speechrecognition
'Text cursor ms-settings:easeofaccess-cursor
'Visual Effects  ms-settings:easeofaccess-visualeffects

'extras
'Settings page URI
'Extras  ms-settings:extras (only available if apps with these settings have been installed, for example, by a 3rd party)

'Family Group
'Settings page URI
'Family Group  ms-settings:family-group

'gaming
'Settings page URI
'Game bar  ms-settings:gaming-gamebar
'Game DVR  ms-settings:gaming-gamedvr
'Game Mode ms-settings:gaming-gamemode
'Playing a game full screen  ms-settings:quietmomentsgame
'TruePlay  ms-settings:gaming-trueplay (As of Windows 10, version 1809 (10.0; Build 17763), this feature is removed from Windows)

'Mixed Reality
' Note
'
'These settings are only available if the Mixed Reality Portal app is installed.
'
' Important
'
'Windows Mixed Reality devices are not supported with Windows 11, version 24H2 and newer.
'
'Windows Mixed Reality support is limited to Windows 10, version 20H2 through Windows 11, version 23H2.
'
'Settings page URI
'Audio and speech  ms-settings:holographic-audio
'Environment ms - Settings: privacy -holographic - Environment
'Headset display ms-settings:holographic-headset
'Uninstall ms - Settings: holographic -management
'Startup and desktop ms-settings:holographic-startupandesktop

'Network and internet
'Settings page URI
'Network & internet  ms-settings:network-status
Public Function NetworkAndInternet_NetworkAndInternet() As Phosphorus.WindowsSetting
  Dim myWindowsSetting As Phosphorus.WindowsSetting
  myWindowsSetting.Category = "Network and internet"
  myWindowsSetting.SettingsPage = "Network And Internet"
  myWindowsSetting.URISuffix = "network-status"
  NetworkAndInternet_NetworkAndInternet = myWindowsSetting
End Function

'Advanced settings ms-settings:network-advancedsettings
'Airplane mode ms-settings:network-airplanemode
'              ms -Settings: proximity
'Cellular & SIM  ms-settings:network-cellular
'Dial-up ms-settings:network-dialup
'DirectAccess  ms-settings:network-directaccess (only available if DirectAccess is enabled)
'ethernet ms - Settings: network -ethernet
'Manage known networks ms-settings:network-wifisettings
'Mobile hotspot  ms-settings:network-mobilehotspot
'proxy ms - Settings: network -proxy
'vpn ms - Settings: network -vpn
'Wi-Fi ms-settings:network-wifi (only available if the device has a wifi adapter)
'Wi-Fi provisioning  ms-settings:wifi-provisioning

'personalization
'Settings page URI
'Background ms - Settings: personalization -Background
'Choose which folders appear on Start  ms-settings:personalization-start-places
'Colors ms - Settings: personalization -Colors
'ms -Settings: Colors
'Customize Copilot key on keyboard ms-settings:personalization-textinput-copilot-hardwarekey
'Dynamic Lighting  ms-settings:personalization-lighting
'Fonts ms - Settings: Fonts
'Glance  ms-settings:personalization-glance (Deprecated in Windows 10, version 1809 and later)
'Lock screen ms-settings:lockscreen
'Navigation bar  ms-settings:personalization-navbar (Deprecated in Windows 10, version 1809 and later)
'Personalization (category)  ms-settings:personalization
'start ms - Settings: personalization -start
'taskbar ms - Settings: taskbar
'Text input  ms-settings:personalization-textinput
'Touch Keyboard  ms-settings:personalization-touchkeyboard
'themes ms - Settings: themes
'phone
'Settings page URI
'Your phone  ms-settings:mobile-devices
'ms -Settings: mobile -devices - addphone
'ms-settings:mobile-devices-addphone-direct (Opens Your Phone app)
'Device Usage  ms-settings:deviceusage
'privacy
'Settings page URI
'Accessory apps  ms-settings:privacy-accessoryapps (Deprecated in Windows 10, version 1809 and later)
'Account info  ms-settings:privacy-accountinfo
'Activity history  ms-settings:privacy-activityhistory
'Advertising ID  ms-settings:privacy-advertisingid (Deprecated in Windows 10, version 1809 and later)
'App diagnostics ms-settings:privacy-appdiagnostics
'Automatic file downloads  ms-settings:privacy-automaticfiledownloads
'Background Apps ms-settings:privacy-backgroundapps (Deprecated in Windows 11, 21H2 and later)
'
'Note: In Windows 11, the background app permissions are accessed individually. To view the permissions, go to Apps->Installed apps and then select "..." on a modern app and choose Advanced options. The advanced page is present for modern apps, and the Background apps permissions section will be present unless a group policy has been set or the user’s global toggle value (the deprecated setting from Windows 10) is set. To access this page with a URI, use the ms-settings:appsfeatures-app URI and pass an optional parameter of the package family name of the app.
'Background Spatial Perception ms-settings:privacy-backgroundspatialperception
'Calendar ms - Settings: privacy -Calendar
'Call history  ms-settings:privacy-callhistory
'Camera ms - Settings: privacy -webcam
'contacts ms - Settings: privacy -contacts
'documents ms - Settings: privacy -documents
'Downloads folder  ms-settings:privacy-downloadsfolder
'Email ms - Settings: privacy -Email
'Eye tracker ms-settings:privacy-eyetracker (requires eyetracker hardware)
'Feedback & diagnostics  ms-settings:privacy-feedback
'File system ms-settings:privacy-broadfilesystemaccess
'General ms-settings:privacy or ms-settings:privacy-general
'Graphics ms - Settings: privacy -graphicscaptureprogrammatic
'ms -Settings: privacy -graphicscapturewithoutborder
'Inking & typing ms-settings:privacy-speechtyping
'Location ms - Settings: privacy -Location
'messaging ms - Settings: privacy -messaging
'microphone ms - Settings: privacy -microphone
'motion ms - Settings: privacy -motion
'Music Library ms-settings:privacy-musiclibrary
'Notifications ms - Settings: privacy -Notifications
'Other devices ms-settings:privacy-customdevices
'Phone calls ms-settings:privacy-phonecalls
'Pictures ms - Settings: privacy -Pictures
'radios ms - Settings: privacy -radios
'Speech ms - Settings: privacy -Speech
'Tasks ms - Settings: privacy -Tasks
'videos ms - Settings: privacy -videos
'Voice activation  ms-settings:privacy-voiceactivation
'Search
'Settings page URI
'Search ms - Settings: Search
'Search more details ms-settings:search-moredetails
'Search Permissions  ms-settings:search-permissions
'Sound
'Settings page URI
'Volume mixer  ms-settings:apps-volume
'Sound ms - Settings: Sound
'Sound devices ms-settings:sound-devices
'Default microphone  ms-settings:sound-defaultinputproperties
'Default audio output  ms-settings:sound-defaultoutputproperties
'Audio device properties
'(specific device) ms-settings:sound-properties?endpointId={0.0.0.00000000}.{aaaaaaaa-0000-1111-2222-bbbbbbbbbbbb}
'
'Note: User of URI must know the endpointId string to use.
'Audio device properties
'(specific device) ms-settings:sound-properties?interfaceId=\\?\SWD#MMDEVAPI#{3.0.0.00000003}.{bbbbbbbb-1111-2222-3333-cccccccccccc}#{cccccccc-2222-3333-4444-dddddddddddd}
'
'Note: User of URI must know the interfaceId string to use and the string must be escaped correctly before sending.
'Surface Hub
'Settings page URI
'accounts ms - Settings: surfacehub -accounts
'Session cleanup ms-settings:surfacehub-sessioncleanup
'Team Conferencing ms-settings:surfacehub-calling
'Team device management  ms-settings:surfacehub-devicemanagement
'Welcome screen  ms-settings:surfacehub-welcome
'System
'Settings page URI
'about ms - Settings: about
'Advanced display settings ms-settings:display-advanced (only available on devices that support advanced display options)
'Battery Saver ms-settings:batterysaver (only available on devices that have a battery, such as a tablet)
'Battery Saver settings  ms-settings:batterysaver-settings (only available on devices that have a battery, such as a tablet)
'Battery use ms-settings:batterysaver-usagedetails (only available on devices that have a battery, such as a tablet)
'clipboard ms - Settings: clipboard
'Default Save Locations  ms-settings:savelocations
'display ms - Settings: display
'ms -Settings: screenrotation
'Duplicating my display  ms-settings:quietmomentspresentation
'During these hours  ms-settings:quietmomentsscheduled
'Encryption ms - Settings: deviceencryption
'Energy recommendations  ms-settings:energyrecommendations (Available in February Moment update for Windows 11, Version 22H2, Build 22624 or later)
'Focus assist  ms-settings:quiethours
'Graphics Settings ms-settings:display-advancedgraphics (only available on devices that support advanced graphics options)
'Graphics Default Settings ms-settings:display-advancedgraphics-default
'multitasking ms - Settings: multitasking
'ms -Settings: multitasking -sgupdate
'Night light settings  ms-settings:nightlight
'Projecting to this PC ms-settings:project
'Shared experiences  ms-settings:crossdevice
'Tablet mode ms-settings:tabletmode (Removed in Windows 11)
'taskbar ms - Settings: taskbar
'Notifications & actions ms-settings:notifications
'Remote Desktop  ms-settings:remotedesktop
'Phone ms-settings:phone (Deprecated in Windows 10, version 1809 and later)
'Power & sleep ms-settings:powersleep
'Presence sensing  ms-settings:presence (Added in May Moment update for Windows 11, Version 22H2, Build 22624)
'Storage ms - Settings: storagesense
'Storage Sense ms-settings:storagepolicies
'Storage recommendations ms-settings:storagerecommendations
'Disks & volumes ms-settings:disksandvolumes
'Time and language
'Settings page URI
'Date & time ms-settings:dateandtime
'Japan IME settings  ms-settings:regionlanguage-jpnime (available if the Microsoft Japan input method editor is installed)
'Region ms - Settings: regionformatting
'Language ms - Settings: keyboard
'ms -Settings: keyboard -Advanced
'ms -Settings: regionlanguage
'ms -Settings: regionlanguage -bpmfime
'ms -Settings: regionlanguage -cangjieime
'ms -Settings: regionlanguage -chsime - Wubi - udp
'ms -Settings: regionlanguage -quickime
'ms -Settings: regionlanguage -korime
'Pinyin IME settings ms-settings:regionlanguage-chsime-pinyin (available if the Microsoft Pinyin input method editor is installed)
'ms -Settings: regionlanguage -chsime - pinyin - domainlexicon
'ms -Settings: regionlanguage -chsime - pinyin - keyconfig
'ms -Settings: regionlanguage -chsime - pinyin - udp
'Speech ms - Settings: Speech
'Wubi IME settings ms-settings:regionlanguage-chsime-wubi (available if the Microsoft Wubi input method editor is installed)
'Update and security
'Settings page URI
'activation ms - Settings: activation
'Backup  ms-settings:backup (page removed in Windows 11; opens Sync)
'Delivery Optimization ms-settings:delivery-optimization
'ms -Settings: Delivery -optimization - activity
'ms -Settings: Delivery -optimization - Advanced
'Find My Device  ms-settings:findmydevice
'For developers  ms-settings:developers
'recovery ms - Settings: recovery
'Launch Security Key Enrollment  ms-settings:signinoptions-launchsecuritykeyenrollment
'troubleshoot ms - Settings: troubleshoot
'Windows Security  ms-settings:windowsdefender
'Windows Insider Program ms-settings:windowsinsider (only present if user is enrolled in WIP)
'ms -Settings: windowsinsider -optin
'Windows Update  ms-settings:windowsupdate
'ms -Settings: windowsupdate -Action
'Windows Update-Active hours ms-settings:windowsupdate-activehours
'Windows Update-Advanced options ms-settings:windowsupdate-options
'Windows Update-Optional updates ms-settings:windowsupdate-optionalupdates
'Windows Update-Restart options  ms-settings:windowsupdate-restartoptions
'Windows Update-Seeker on demand ms-settings:windowsupdate-seekerondemand
'Windows Update-View update history  ms-settings:windowsupdate-history

