VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Formのベース(Shape)の原点を元の位置に戻す
Private Sub Worksheet_Activate()
    Me.Shapes(1).Top = 0.75
    Me.Shapes(1).Left = 0.75
End Sub


