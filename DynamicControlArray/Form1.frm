VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer, iLeft As Integer, iTop As Integer
    iLeft = 240
    
    LoadControl Me, "VB.PictureBox", "picObj", "", Me.Height - 400, Me.Width - 400, 0, 0
    LoadControl Me, "VB.VScrollBar", "scrollObj", "", Me.Height - 400, 285, 0, Me.Controls("picObj").Width
    
    For i = 0 To 100
        iTop = iTop + 320
        LoadControl Me, "VB.CommandButton", "btnObj" & i, "Command" & i, 315, 2000, iTop, iLeft
        LoadControl Me, "VB.TextBox", "txtObj" & i, "", 315, 2000, iTop, iLeft + 2240
    Next i

    Me.Controls("picObj").Height = Me.Controls("btnObj" & "100").Top + 400
    Me.Controls("scrollObj").Max = Abs(Me.Controls("picObj").Height - Me.Height) + 500
    Me.Controls("scrollObj").LargeChange = Me.Controls("scrollObj").Max / 100
    Me.Controls("scrollObj").SmallChange = Me.Controls("scrollObj").Max / 100
    
End Sub

Sub LoadControl(frm As Form, sControlType As String, sControlName As String, sCaption As String, iHeight As Integer, iWidth As Integer, iTop As Integer, iLeft As Integer)
    On Error Resume Next
    Dim newBtn As Control
    
    Set newBtn = frm.Controls.Add(sControlType, sControlName)
    With newBtn
        .Height = iHeight
        .Width = iWidth
        .Top = iTop
        .Left = iLeft
        .Visible = True
    End With
    
    ReDim Preserve clsBtn(UBound(clsBtn) + 1)
    If Err Then ReDim Preserve clsBtn(0): Err.Clear
    
    Set clsBtn(UBound(clsBtn)) = New clsControl
    clsBtn(UBound(clsBtn)).Index = UBound(clsBtn)
    If (TypeOf newBtn Is CommandButton) Then
        newBtn.Caption = sCaption
        Set newBtn.Container = frm.Controls("picObj")
        Set clsBtn(UBound(clsBtn)).btn = newBtn
    ElseIf (TypeOf newBtn Is TextBox) Then
        newBtn.Text = sCaption
        Set newBtn.Container = frm.Controls("picObj")
        Set clsBtn(UBound(clsBtn)).txt = newBtn
    ElseIf (TypeOf newBtn Is VScrollBar) Then
        Set clsBtn(UBound(clsBtn)).vsb = newBtn
    End If
    
    If Err Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub

Private Sub Form_Resize()
    Me.Controls("picObj").Width = Me.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase clsBtn
End Sub
