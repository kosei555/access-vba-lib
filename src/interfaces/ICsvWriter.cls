VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsvWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Public Sub SetSeparator(separator As String)
End Sub

Public Sub CreateFile(path As String)
End Sub

Public Sub AddLine(data() As Variant)
End Sub

Public Sub Complete()
End Sub

Private Function ExpandEnvironment(str As String) As String
End Function

Private Sub Clean()
End Sub
