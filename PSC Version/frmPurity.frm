VERSION 5.00
Begin VB.Form frmPurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4995
   ClientLeft      =   6585
   ClientTop       =   16200
   ClientWidth     =   4995
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   1940
   Icon            =   "frmPurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   Tag             =   "120"
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmPurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub Form_Activate()
    'clears mousepointer

    Call ShowCursor(0)

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    'reinstate default mousepointer

    Call ShowCursor(1)
    
    Unload Me
    
End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, Nothing, False
    
    If KeyAscii = vbKeyEscape Then frmPurity.Hide
    If KeyAscii = vbKeySpace Then frmPurity.Hide
    
End Sub

Private Sub Form_Load()

    frmPurity.Width = Screen.Width
    frmPurity.Height = Screen.Height
    frmPurity.BackColor = vbRed

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmPurity.Hide
    
End Sub

