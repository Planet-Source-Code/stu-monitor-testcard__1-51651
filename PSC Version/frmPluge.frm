VERSION 5.00
Begin VB.Form frmPluge 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2.45745e5
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2.45745e5
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   HelpContextID   =   1900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   16383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16383
   ShowInTaskbar   =   0   'False
   Tag             =   "116"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton frText 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderColor     =   65535
      FillColor       =   0
      Caption         =   ""
      RedrawOnHover   =   0   'False
      ActiveBorderColor=   65535
      ActiveFillColor =   0
      CheckedFillColor=   0
      BackColor       =   0
      CornerRadius    =   10
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   8400
         TabIndex        =   2
         Top             =   180
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
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin VB.Label lblWhite 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   $"frmPluge.frx":0000
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8130
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPluge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_PLUGE)

End Sub

Private Sub Form_Activate()

frText.Left = (ScaleWidth - frText.Width) / 2
    frText.Top = (ScaleHeight - frText.Height) - 50
    
End Sub

Private Sub Form_Paint()
    'draws display

    ScaleWidth = 100
    ScaleHeight = 100

    BackColor = RGB(6, 6, 6)

    Line (64, 6)-(71, 9), vbRed, BF  'primary colour bars
    Line (71, 6)-(78, 9), vbGreen, BF
    Line (78, 6)-(85, 9), vbBlue, BF

    Line (64, 11)-(71, 14), vbMagenta, BF 'secondary colour bars
    Line (71, 11)-(78, 14), vbCyan, BF
    Line (78, 11)-(85, 14), vbYellow, BF

    Line (17, 10)-(25, 80), vbBlack, BF 'left bar
    Line (37, 10)-(45, 80), RGB(12, 12, 12), BF 'right bar
    Line (64, 17)-(85, 46), RGB(255, 255, 255), BF 'white box
    Line (64, 46)-(85, 79), RGB(128, 128, 128), BF 'grey box

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, Nothing, False
    
    If KeyAscii = vbKeyEscape Then frmPluge.Hide
    If KeyAscii = vbKeySpace Then frmPluge.Hide
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

