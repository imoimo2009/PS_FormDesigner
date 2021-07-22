VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    Me.TextBox1.Text = GenPSCode(ThisWorkbook.Worksheets("Design"))
    Me.TextBox1.Activate
    SendKeys "^{HOME}"
End Sub

Private Sub CommandButton2_Click()
    With New DataObject
        .SetText Me.TextBox1.Text
        .PutInClipboard
    End With
    MsgBox "クリップボードにコピーしました。", vbInformation, ThisWorkbook.Name
End Sub

Private Sub CommandButton3_Click()
    Dim lsPath              As String
    
    With New FileSystemObject
        lsPath = .GetSpecialFolder(2).Path & "\" & .GetBaseName(.GetTempName) & ".ps1"
        With .CreateTextFile(lsPath, True)
            .Write Me.TextBox1.Text
            .Close
        End With
        With New WshShell
            .Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden " & lsPath, 1, True
        End With
        .DeleteFile lsPath
    End With
End Sub
