VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 終了前の後始末（Bookが壊れることがあるので・・・)
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If oProcess Is Nothing = False Then
        oProcess.Terminate
    End If
    With New FileSystemObject
        If .FileExists(sTmpPath) Then .DeleteFile sTmpPath
    End With
End Sub

