VERSION 5.00
Begin VB.Form frmTestcard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   HelpContextID   =   1930
   Icon            =   "frmTestcard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmTestcard.frx":000C
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   ShowInTaskbar   =   0   'False
   Tag             =   "119"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BorderColor     =   16777215
      FillColor       =   8388608
      Caption         =   "Help"
      HoverForeColor  =   65535
      HoverFillColor  =   16744576
      HoverBorderColor=   65535
      PictureAlignment=   2
      ActiveBorderColor=   65535
      ActiveForeColor =   65535
      ActiveFillColor =   16761024
      BackColor       =   8421504
      CornerRadius    =   10
   End
End
Attribute VB_Name = "frmTestcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form uses a graphic imported from PSPro 7

Private Sub cmdHelp_Click()

Call ShowHelpTopic(Hlp_Testcard)

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, Nothing, False

    If KeyAscii = vbKeyEscape Then frmTestcard.Hide
    If KeyAscii = vbKeySpace Then frmTestcard.Hide

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

