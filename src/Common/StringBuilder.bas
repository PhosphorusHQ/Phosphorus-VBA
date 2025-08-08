VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Common
Option Explicit

Private buffer As String
Private bufferLength As Long

Private Sub Class_Initialize()
  buffer = Space(1000) ' Initial buffer size
  bufferLength = 0
End Sub

Public Sub Append(Text As String)
  If bufferLength + Len(Text) > Len(buffer) Then
    buffer = buffer & Space(Len(buffer) * 2)
  End If
  Mid$(buffer, bufferLength + 1, Len(Text)) = Text
  bufferLength = bufferLength + Len(Text)
End Sub

Public Function ToString(Optional start As Long = 1, Optional Length As Long = -1) As String
  If Length = -1 Then Length = bufferLength
  If start <= bufferLength And Length > 0 Then
    ToString = Mid$(buffer, start, Length)
  Else
    ToString = ""
  End If
End Function

Public Property Get Length() As Long
  Length = bufferLength
End Property
