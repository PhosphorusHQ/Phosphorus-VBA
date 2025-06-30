Attribute VB_Name = "FrameworkTest"
'@Folder _FrameworkTest
Option Explicit

Private Sub test()
  Dim dog As A
  Dim cat As B
  Set dog = New A
  Set cat = New B
'  dog.Speak
'  cat.Speak
  
  Dim C As Object
  Set C = cat
  C.Speak
End Sub
