VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index As Integer
Public WithEvents btn As CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents txt As TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents vsb As VScrollBar
Attribute vsb.VB_VarHelpID = -1

Private Sub btn_Click()
    MsgBox "Button Name: " & btn.Name & ", Text: " & clsBtn(Index + 1).txt.Text
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Debug.Print "TextBox: " & txt.Name & ", Key: " & KeyAscii
End Sub

Private Sub vsb_Change()
    Form1.Controls("picobj").Top = 0 - vsb.Value
End Sub

Private Sub vsb_Scroll()
    Form1.Controls("picobj").Top = 0 - vsb.Value
End Sub
