Attribute VB_Name = "HTML_RadioButtonsAndCheckboxes"
'@Folder HTML
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2026 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function RadioButtonsAndCheckboxes() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False

  Dim WebBrowser As Object
  Dim SearchBox As pLocator
  Dim SearchButton As pLocator

  Dim SourceFileURL As String
  Dim DocumentRoot As pLocator
  
  Dim Header1 As String
  Dim YesText As pElement
  Dim YesRadioButton As pElement
  Dim NoText As pElement
  Dim NoRadioButton As pElement
  
  Dim Header2 As String
  Dim RocksRadioText As pElement
  Dim RocksRadioButton As pElement
  Dim ReallyRocksRadioText As pElement
  Dim ReallyRocksRadioButton As pElement
  Dim MaybeRocksRadioText As pElement
  Dim MaybeRocksRadioButton As pElement

  Dim Header3 As String
  Dim PhosphorusRocksText As pElement
  Dim PhosphorusRocksCheckbox As pElement

  On Error GoTo ErrorHandler
  
  If Not RunningAllExamples Then
    Factory.CurrentWebBrowserType = WebBrowserType.Brave
    Window.HighlightElements = True
  End If
  
  SourceFileURL = HTMLCommon.GetULREncodedFilepath("Tests\WebPages\HTML\RadioButtonsAndCheckboxes.html")
  
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start "HTML Radio Buttons And Checkboxes", SourceFileURL, "Radio Buttons And Checkboxes"

  Set DocumentRoot = WebBrowser.GetRootWebArea

  'Check the 1st header
  Header1 = "I Love Phosphorus ..."
  Dim Header1Text As pElement

  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    'Chromium rolls all the text elements into one!
    Debug.Assert VBA.Strings.Trim(DocumentRoot.Element.Text) = "I Love Phosphorus ...  Yes  No Confirm last field is disabled. Phosphorus ...  Rocks  Really Rocks  Maybe Rocks Tick the check:  Phosphorus rocks!"
  Else
    Set Header1Text = DocumentRoot.FindElement("ILovePhosphorusHeader", By.NameIs, Header1)
    Debug.Assert Header1Text.ControlType = UIAControlTypeIDs.Text
  End If

  'Check the Yes text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set YesText = DocumentRoot.FindElement("Yes", By.NameIs, "Yes")
  Else
    Set YesText = Header1Text.FindSibling("Yes", UIAControlTypeIDs.Text)
  End If
  Debug.Assert YesText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert YesText.Name = "Yes"
  
  'Check the Yes radio button
  Set YesRadioButton = YesText.FindSibling("yes", UIAControlTypeIDs.RadioButton)
  Debug.Assert YesRadioButton.ControlType = UIAControlTypeIDs.RadioButton
  Debug.Assert YesRadioButton.IsSelected
 
  'Check the No text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set NoText = DocumentRoot.FindElement("No", By.NameIs, "No")
  Else
    Set NoText = Header1Text.FindSibling("No", UIAControlTypeIDs.Text)
  End If
  Debug.Assert NoText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert NoText.Name = "No"
    
  'Check the No radio button & click it
  Set NoRadioButton = NoText.FindSibling("no", UIAControlTypeIDs.RadioButton)
  Debug.Assert NoRadioButton.ControlType = UIAControlTypeIDs.RadioButton
  NoRadioButton.Click().IsSelected

  'Click the Yes radio button & check it again
  YesRadioButton.Click().IsSelected
    
  'Check the header
  Header2 = "Confirm last field is disabled. Phosphorus ..."
  Dim Header2Text As pElement
  
  If Not Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set Header2Text = DocumentRoot.FindElement("ConfirmLastFieldDisabled", By.NameIs, Header2)
    Debug.Assert Header2Text.ControlType = UIAControlTypeIDs.Text
  End If

  'Check the Rocks text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set RocksRadioText = DocumentRoot.FindElement("Rocks", By.NameIs, "Rocks")
  Else
    Set RocksRadioText = Header2Text.FindSibling("Rocks", UIAControlTypeIDs.Text)
  End If
  Debug.Assert RocksRadioText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert RocksRadioText.Name = "Rocks"
  
  'Check the Rocks radio button
  Set RocksRadioButton = RocksRadioText.FindSibling("rocks", UIAControlTypeIDs.RadioButton)
  Debug.Assert RocksRadioButton.ControlType = UIAControlTypeIDs.RadioButton
  Debug.Assert RocksRadioButton.IsNotSelected
  
  'Check the Really Rocks text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set ReallyRocksRadioText = DocumentRoot.FindElement("Really Rocks", By.NameIs, "Really Rocks")
  Else
    Set ReallyRocksRadioText = Header2Text.FindSibling("Really Rocks", UIAControlTypeIDs.Text)
  End If
  Debug.Assert ReallyRocksRadioText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert ReallyRocksRadioText.Name = "Really Rocks"
    
  'Check the Really Rocks radio button
  Set ReallyRocksRadioButton = ReallyRocksRadioText.FindSibling("reallyrocks", UIAControlTypeIDs.RadioButton)
  Debug.Assert ReallyRocksRadioButton.ControlType = UIAControlTypeIDs.RadioButton
  Debug.Assert ReallyRocksRadioButton.IsNotSelected
    
  'Check the Maybe Rocks text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set MaybeRocksRadioText = DocumentRoot.FindElement("Maybe Rocks", By.NameIs, "Maybe Rocks")
  Else
    Set MaybeRocksRadioText = Header2Text.FindSibling("Maybe Rocks", UIAControlTypeIDs.Text)
  End If
  Debug.Assert MaybeRocksRadioText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert MaybeRocksRadioText.Name = "Maybe Rocks"

  'Check the Maybe Rocks radio button
  Set MaybeRocksRadioButton = MaybeRocksRadioText.FindSibling("mayberocks", UIAControlTypeIDs.RadioButton)
  Debug.Assert MaybeRocksRadioButton.ControlType = UIAControlTypeIDs.RadioButton
  Debug.Assert MaybeRocksRadioButton.IsDisabled
  Debug.Assert MaybeRocksRadioButton.IsNotSelected
    
  'Click the Rocks radio button & check it again
  Debug.Assert RocksRadioButton.Click().IsSelected
  
  'Click the Really Rocks radio button & check it again
  Debug.Assert ReallyRocksRadioButton.Click().IsSelected
    
  'Check the Maybe Rocks radio button cannot be clicked
  With MaybeRocksRadioButton
    On Error GoTo ErrorTrap
    .Click
    Debug.Assert Err.Number = 0
    On Error GoTo ErrorHandler
    GoTo NextTest
ErrorTrap:
    Debug.Assert Err.Number = -2147218503
    Debug.Assert Err.Description = "Element 'mayberocks' is not alive!"
    Resume Next
  End With

NextTest:

  'Check the 3rd header
  Header3 = "Tick the check:"
  Dim Header3Text As pElement
  
  If Not Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set Header3Text = DocumentRoot.FindElement("TickTheCheckHeader", By.NameIs, Header3)
    Debug.Assert Header3Text.ControlType = UIAControlTypeIDs.Text
  End If

  'Check the Phosphorus Rocks! text
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set PhosphorusRocksText = DocumentRoot.FindElement("Phosphorus rocks!", By.NameIs, "Phosphorus rocks!")
  Else
    Set PhosphorusRocksText = Header3Text.FindSibling("Phosphorus rocks!", UIAControlTypeIDs.Text)
  End If
  Debug.Assert PhosphorusRocksText.ControlType = UIAControlTypeIDs.Text
  Debug.Assert PhosphorusRocksText.Name = "Phosphorus rocks!"
  
  'Check the Phosphorus Rocks! checkbox & click it
  'Set PhosphorusRocksCheckbox = DocumentRoot.FindElement("PhosphorusRocksCheckbox", By.NameIs, Header3, RelativeElementNumber:=1)
  Set PhosphorusRocksCheckbox = PhosphorusRocksText.FindSibling("phosphorusrocks", UIAControlTypeIDs.CheckBox)
  Debug.Assert PhosphorusRocksCheckbox.ControlType = UIAControlTypeIDs.CheckBox
  Debug.Assert PhosphorusRocksCheckbox.IsNotSelected
  Debug.Assert PhosphorusRocksCheckbox.Click().IsSelected
  Debug.Assert PhosphorusRocksCheckbox.Click().IsNotSelected
    
  If RunningAllExamples Then
    Debug.Print "Radio Buttons & Checkboxed Clicked!"
  Else
    MsgBox "Radio Buttons & Checkboxes Checked!"
    Phosphorus.Log4PStatic.CloseLogger
  End If
  
  Succeeded = True
  GoTo ExitSub
   
ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'RadioButtonsAndCheckboxes' Example!"
  Else
    MsgBox Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'RadioButtonsAndCheckboxes' Example!", vbCritical
  End If
  GoTo ExitSub

'PJG TODO: Run through all journey combinations!

ExitSub:
  Set WebBrowser = Nothing
  Set SearchBox = Nothing
  Set SearchButton = Nothing
  
  Set Header1Text = Nothing
  Set YesText = Nothing
  Set YesRadioButton = Nothing
  Set NoText = Nothing
  Set NoRadioButton = Nothing
  
  Set Header2Text = Nothing
  Set RocksRadioText = Nothing
  Set RocksRadioButton = Nothing
  Set ReallyRocksRadioText = Nothing
  Set ReallyRocksRadioButton = Nothing
  Set MaybeRocksRadioText = Nothing
  Set MaybeRocksRadioButton = Nothing
  
  Set Header3Text = Nothing
  Set PhosphorusRocksText = Nothing
  Set PhosphorusRocksCheckbox = Nothing
  
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  RadioButtonsAndCheckboxes = Succeeded

End Function
