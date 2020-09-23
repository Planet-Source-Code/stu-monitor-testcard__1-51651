VERSION 5.00
Begin VB.Form frmReg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1.99995e5
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1.99995e5
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   1860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   13333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13333
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton frHelp 
      Height          =   1575
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
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
      FillColor       =   8388608
      Caption         =   ""
      HoverForeColor  =   65535
      RedrawOnHover   =   0   'False
      HoverFillColor  =   65535
      HoverBorderColor=   65535
      ActiveForeColor =   -2147483643
      ActiveFillColor =   -2147483634
      CheckedFillColor=   -2147483634
      BackColor       =   12582912
      Begin Testcard.MultiButton fraDraw 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   65535
         BorderColor     =   65535
         FillColor       =   8388608
         Caption         =   "Speed"
         RedrawOnHover   =   0   'False
         Alignment       =   0
         VerticalAlignment=   0
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   10
         Begin Testcard.MultiButton optSpeed 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            Picture         =   "frmReg.frx":0000
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "Slow"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   1
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            CheckedPicture  =   "frmReg.frx":059A
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optSpeed 
            Default         =   -1  'True
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   6
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            Picture         =   "frmReg.frx":06F4
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "Normal"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   1
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            Value           =   -1  'True
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            CheckedPicture  =   "frmReg.frx":0C8E
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optSpeed 
            Height          =   315
            Index           =   2
            Left            =   3660
            TabIndex        =   5
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            Picture         =   "frmReg.frx":0DE8
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "Fast"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   1
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            CheckedPicture  =   "frmReg.frx":1382
            BackColor       =   8388608
            CornerRadius    =   10
         End
      End
      Begin Testcard.MultiButton cmdHelp 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
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
         BorderColor     =   65535
         FillColor       =   8388608
         Caption         =   "Help"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         PictureAlignment=   2
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         CheckedFillColor=   8388608
         CheckedForeColor=   -2147483643
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdRestart 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
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
         BorderColor     =   65535
         FillColor       =   8388608
         Caption         =   "Restart"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         CheckedBorderColor=   16777215
         CheckedFillColor=   8388608
         CheckedForeColor=   16777215
         BackColor       =   8388608
         CornerRadius    =   10
      End
   End
   Begin VB.Timer tmrFlash 
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VB.Label lblWhite 
      BackColor       =   &H00FFFFFF&
      Height          =   8100
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_Regulation)

    tmrFlash.Enabled = False
    cmdHelp.Visible = False
 
End Sub

Private Sub cmdRestart_Click()

    tmrFlash.Enabled = True
    cmdHelp.Visible = True

    Call QuitHelp

End Sub

Private Sub Form_Activate()

    frHelp.Left = (ScaleWidth - frHelp.Width) / 2
    frHelp.Top = (ScaleHeight - frHelp.Height) / 2
    
End Sub

Private Sub Form_Click()

    frmReg.Hide

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, Nothing, False
    
    If KeyAscii = vbKeyEscape Then frmReg.Hide
    If KeyAscii = vbKeySpace Then frmReg.Hide
    
End Sub

Private Sub Form_Paint()
    'draws outer line

    ScaleWidth = 100
    ScaleHeight = 100

    DrawWidth = 3
    
    Line (1, 1)-(99, 1)
    Line (99, 1)-(99, 99)
    Line (99, 99)-(1, 99)
    Line (1, 99)-(1, 1)
       
    'size of label
    lblWhite.Height = 90
    lblWhite.Width = 90
    lblWhite.Left = 5
    lblWhite.Top = 5

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub lblWhite_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub

Private Sub optSpeed_Click(Index As Integer)

Select Case Index

        Case 0
            tmrFlash.Interval = 2000
        Case 1
            tmrFlash.Interval = 1000
        Case 2
            tmrFlash.Interval = 250
    End Select
    
End Sub

Private Sub tmrFlash_Timer()
    'changes background from black to white

    If lblWhite.BackColor = vbWhite Then
        lblWhite.BackColor = vbBlack

    Else

        lblWhite.BackColor = vbWhite

    End If
    
End Sub

