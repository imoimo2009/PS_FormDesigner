Attribute VB_Name = "Module1"
Option Explicit

Public Const cVersion           As String = "1.0.0"

Private Const cFormName         As String = "Form1"

Private Const cForm             As String = "System.Windows.Forms"
Private Const cDraw             As String = "System.Drawing"

'�w��̃V�[�g�ɔz�u�����R���g���[�����X�L�������APowerShell�X�N���v�g�ɕϊ�����
Public Function GenPSCode(loWs As Worksheet) As String
    Dim lsStr                   As String
    Dim lsCtl                   As String
    Dim lvCtl                   As Variant
    Dim o                       As OLEObject
    Dim i                       As Integer
    
    ReDim lvCtl(0)
    AddLine lsStr, "Add-Type -AssemblyName System.Windows.Forms"
    AddLine lsStr, "Add-Type -AssemblyName System.Drawing"
    AddLine lsStr
    With loWs
        For Each o In .OLEObjects
            If o.OLEType = 2 Then
                Select Case TypeName(o.Object)
                    Case "Label"
                        lsCtl = CreateLabel(o, lvCtl)
                    Case "CommandButton"
                        lsCtl = CreateButton(o, lvCtl)
                    Case "TextBox"
                        lsCtl = CreateTextBox(o, lvCtl)
                    Case "ComboBox"
                        lsCtl = CreateComboBox(o, lvCtl)
                    Case "ListBox"
                        lsCtl = CreateListBox(o, lvCtl)
                    Case "CheckBox"
                        lsCtl = CreateCheckBox(o, lvCtl)
                    Case "OptionButton"
                        lsCtl = CreateRadioButton(o, lvCtl)
                    Case "ProgressBar"
                        lsCtl = CreateProgressBar(o, lvCtl)
                    Case "Image"
                        lsCtl = CreatePictureBox(o, lvCtl)
                    Case Else
                        lsCtl = ""
                End Select
                lsStr = lsStr & lsCtl
            End If
        Next
        AddLine lsStr, "[" & cForm & ".Form]$form = New-Object " & cForm & ".Form"
        AddLine lsStr, "$form.Text = """ & cFormName & """"
        If .Shapes.Count > 0 Then
            AddLine lsStr, "$form.ClientSize = New-Object " & cDraw & ".Size(" & Pts2Pxl(.Shapes(1).Width) & "," & Pts2Pxl(.Shapes(1).Height) & ")"
        End If
        AddLine lsStr
    End With
    For i = 1 To UBound(lvCtl)
        AddLine lsStr, "$form.Controls.Add(" & lvCtl(i) & ")"
    Next
    AddLine lsStr
    AddLine lsStr, "$form.ShowDialog()"
    GenPSCode = lsStr
End Function

' ���x���R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateLabel(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".Label]" & lsVar & " = New-Object " & cForm & ".Label"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & CtrlFont(loCtl)
    AddLine lsStr, lsVar & ".Text = """ & loCtl.Object.Caption & """"
    AddLine lsStr
    CreateLabel = lsStr
End Function

' �{�^���R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateButton(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".Button]" & lsVar & " = New-Object " & cForm & ".Button"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & CtrlFont(loCtl)
    AddLine lsStr, lsVar & ".Text = """ & loCtl.Object.Caption & """"
    AddLine lsStr
    CreateButton = lsStr
End Function

' �e�L�X�g�{�b�N�X�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateTextBox(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".TextBox]" & lsVar & " = New-Object " & cForm & ".TextBox"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & CtrlFont(loCtl)
    lsStr = lsStr & lsVar & ".MultiLine = "
    If loCtl.Object.MultiLine Then
        AddLine lsStr, "$true"
    Else
        AddLine lsStr, "$false"
    End If
    lsStr = lsStr & lsVar & ".ScrollBars = [" & cForm & ".ScrollBars]::"
    Select Case loCtl.Object.ScrollBars
        Case fmScrollBarsNone
            AddLine lsStr, "None"
        Case fmScrollBarsHorizontal
            AddLine lsStr, "Horizontal"
        Case fmScrollBarsVertical
            AddLine lsStr, "Vertical"
        Case fmScrollBarsBoth
            AddLine lsStr, "Both"
    End Select
    AddLine lsStr
    CreateTextBox = lsStr
End Function

' �R���{�{�b�N�X�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateComboBox(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".ComboBox]" & lsVar & " = New-Object " & cForm & ".ComboBox"
    lsStr = lsStr & CtrlGeometry(loCtl)
    AddLine lsStr, CtrlFont(loCtl)
    CreateComboBox = lsStr
End Function

' ���X�g�{�b�N�X�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateListBox(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".ListBox]" & lsVar & " = New-Object " & cForm & ".ListBox"
    lsStr = lsStr & CtrlGeometry(loCtl)
    AddLine lsStr, CtrlFont(loCtl)
    CreateListBox = lsStr
End Function

' �`�F�b�N�{�b�N�X�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateCheckBox(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".CheckBox]" & lsVar & " = New-Object " & cForm & ".CheckBox"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & CtrlFont(loCtl)
    AddLine lsStr, lsVar & ".Text = """ & loCtl.Object.Caption & """"
    AddLine lsStr
    CreateCheckBox = lsStr
End Function

' ���W�I�{�^���R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateRadioButton(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".RadioButton]" & lsVar & " = New-Object " & cForm & ".RadioButton"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & CtrlFont(loCtl)
    AddLine lsStr, lsVar & ".Text = """ & loCtl.Object.Caption & """"
    AddLine lsStr
    CreateRadioButton = lsStr
End Function

' �v���O���X�o�[�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreateProgressBar(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".ProgressBar]" & lsVar & " = New-Object " & cForm & ".ProgressBar"
    AddLine lsStr, CtrlGeometry(loCtl)
    CreateProgressBar = lsStr
End Function

' �s�N�`���{�b�N�X�R���g���[����PowerShell�X�N���v�g�ɕϊ�����
Private Function CreatePictureBox(loCtl As OLEObject, lvCtl As Variant) As String
    Dim lsStr                   As String
    Dim lsVar                   As String

    lsVar = "$" & loCtl.Name
    PushArray lvCtl, lsVar
    AddLine lsStr, "[" & cForm & ".PictureBox]" & lsVar & " = New-Object " & cForm & ".PictureBox"
    lsStr = lsStr & CtrlGeometry(loCtl)
    lsStr = lsStr & lsVar & ".BorderStyle = [" & cForm & ".BorderStyle]::"
    Select Case loCtl.Object.BorderStyle
        Case fmBorderStyleNone
            AddLine lsStr, "None"
        Case fmBorderStyleSingle
            AddLine lsStr, "FixedSingle"
    End Select
    AddLine lsStr
    CreatePictureBox = lsStr
End Function

' �R���g���[���̃W�I���g������PowerShell�X�N���v�g�ɕϊ�����
Private Function CtrlGeometry(loCtl As OLEObject) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    AddLine lsStr, lsVar & ".Location = New-Object " & cDraw & ".Point(" & Pts2Pxl(loCtl.Left) & "," & Pts2Pxl(loCtl.Top) & ")"
    AddLine lsStr, lsVar & ".Size = New-Object " & cDraw & ".Size(" & Pts2Pxl(loCtl.Width) & "," & Pts2Pxl(loCtl.Height) & ")"
    CtrlGeometry = lsStr
End Function

' �R���g���[���̃t�H���g����PowerShell�X�N���v�g�ɕϊ�����
Private Function CtrlFont(loCtl As OLEObject) As String
    Dim lsStr                   As String
    Dim lsVar                   As String
    
    lsVar = "$" & loCtl.Name
    AddLine lsStr, lsVar & ".Font = New-Object " & cDraw & ".Font(""" & loCtl.Object.Font.Name & """," & loCtl.Object.Font.Size & ")"
    CtrlFont = lsStr
End Function

' �z����g�����A�����ɒl������
Private Sub PushArray(lvArr, lvVal)
    ReDim Preserve lvArr(UBound(lvArr) + 1)
    lvArr(UBound(lvArr)) = lvVal
End Sub

'�������A�����A�����ɉ��s�R�[�h��}������
Private Sub AddLine(lsSrc As String, Optional ByVal lsStr As String = "")
    lsSrc = lsSrc & lsStr & vbCrLf
End Sub

'�|�C���g�����s�N�Z���ɕϊ�
Private Function Pts2Pxl(lfPoint As Double) As Long
    Pts2Pxl = Int(lfPoint / 0.75)
End Function

