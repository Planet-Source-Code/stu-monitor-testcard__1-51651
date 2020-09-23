VERSION 5.00
Begin VB.Form frmGreyscale 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1.99995e5
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1.99995e5
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   1910
   Icon            =   "frmGreyscale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   13333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "117"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraCol 
      Height          =   1590
      Left            =   3720
      TabIndex        =   0
      Top             =   4605
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   2805
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
      Caption         =   "    Display Colour"
      HoverForeColor  =   65535
      RedrawOnHover   =   0   'False
      HoverFillColor  =   16744576
      HoverBorderColor=   65535
      Alignment       =   0
      PictureAlignment=   2
      ActiveBorderColor=   65535
      ActiveForeColor =   65535
      ActiveFillColor =   16761024
      VerticalAlignment=   0
      BackColor       =   0
      CornerRadius    =   10
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   3
         Left            =   5040
         TabIndex        =   6
         Top             =   420
         Width           =   1485
         _ExtentX        =   2619
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
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Blue"
         HoverForeColor  =   65535
         HoverFillColor  =   16711680
         HoverBorderColor=   65535
         PictureAlignment=   2
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         CheckedBorderColor=   65535
         CheckedFillColor=   16711680
         CheckedForeColor=   16777215
         ButtonMode      =   1
         OptionName      =   "Colour"
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   2
         Left            =   3420
         TabIndex        =   5
         Top             =   420
         Width           =   1485
         _ExtentX        =   2619
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
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Green"
         HoverForeColor  =   0
         HoverFillColor  =   65280
         HoverBorderColor=   65535
         PictureAlignment=   2
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         CheckedBorderColor=   65535
         CheckedFillColor=   65280
         ButtonMode      =   1
         OptionName      =   "Colour"
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   420
         Width           =   1485
         _ExtentX        =   2619
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
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Red"
         HoverForeColor  =   16777215
         HoverFillColor  =   255
         HoverBorderColor=   65535
         PictureAlignment=   2
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         CheckedBorderColor=   65535
         CheckedFillColor=   255
         CheckedForeColor=   65535
         ButtonMode      =   1
         OptionName      =   "Colour"
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   1485
         _ExtentX        =   2619
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
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "GreyScale"
         HoverForeColor  =   0
         HoverFillColor  =   16777215
         HoverBorderColor=   65535
         PictureAlignment=   2
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         Value           =   -1  'True
         CheckedBorderColor=   65535
         CheckedFillColor=   16777215
         ButtonMode      =   1
         OptionName      =   "Colour"
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   960
         Width           =   1485
         _ExtentX        =   2619
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
      Begin Testcard.MultiButton cmdClear 
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   4755
         _ExtentX        =   8387
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
         Caption         =   "Clear Toolbar (Shift Key to Toggle)"
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
   End
End
Attribute VB_Name = "frmGreyscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim R As Integer
Dim g As Integer
Dim x As Integer
Dim y As Integer

Private Sub cmdClear_Click()
    'clear toolbar
    
    fraCol.Visible = False

End Sub

Private Sub cmdHelp_Click()
    'help
    
    Call ShowHelpTopic(Hlp_Greyscale)

End Sub

Private Sub Form_Activate()

    fraCol.Left = (ScaleWidth - fraCol.Width) / 2
    fraCol.Top = (ScaleHeight - fraCol.Height) / 2
    fraCol.Visible = True
        
    GreyScale

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, fraCol, False
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub
Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub
Private Sub GreyScale()
    'draws bars using RGB values of shades of grey

    ScaleWidth = 80
    ScaleHeight = 80

    Line (0, 0)-(10, 80), vbWhite, BF
    Line (10, 0)-(20, 80), RGB(219.43, 219.43, 219.43), BF
    Line (20, 0)-(30, 80), RGB(182.86, 182.86, 182.86), BF
    Line (30, 0)-(40, 80), RGB(146.29, 146.29, 146.29), BF
    Line (40, 0)-(50, 80), RGB(109.71, 109.71, 109.71), BF
    Line (50, 0)-(60, 80), RGB(73.14, 73.14, 73.14), BF
    Line (60, 0)-(70, 80), RGB(36.57, 36.57, 36.57), BF
    Line (70, 0)-(80, 80), vbBlack, BF

End Sub
Private Sub RedScale()
    'red only

    ScaleWidth = 80
    ScaleHeight = 80

    Line (0, 0)-(10, 80), vbRed, BF
    Line (10, 0)-(20, 80), RGB(219.43, 0, 0), BF
    Line (20, 0)-(30, 80), RGB(182.86, 0, 0), BF
    Line (30, 0)-(40, 80), RGB(146.29, 0, 0), BF
    Line (40, 0)-(50, 80), RGB(109.71, 0, 0), BF
    Line (50, 0)-(60, 80), RGB(73.14, 0, 0), BF
    Line (60, 0)-(70, 80), RGB(36.57, 0, 0), BF
    Line (70, 0)-(80, 80), vbBlack, BF

End Sub

Private Sub GreenScale()
    'green only

    ScaleWidth = 80
    ScaleHeight = 80

    Line (0, 0)-(10, 80), vbGreen, BF
    Line (10, 0)-(20, 80), RGB(0, 219.43, 0), BF
    Line (20, 0)-(30, 80), RGB(0, 182.86, 0), BF
    Line (30, 0)-(40, 80), RGB(0, 146.29, 0), BF
    Line (40, 0)-(50, 80), RGB(0, 109.71, 0), BF
    Line (50, 0)-(60, 80), RGB(0, 73.14, 0), BF
    Line (60, 0)-(70, 80), RGB(0, 36.57, 0), BF
    Line (70, 0)-(80, 80), vbBlack, BF

End Sub
Private Sub BlueScale()
    'blue only

    ScaleWidth = 80
    ScaleHeight = 80

    Line (0, 0)-(10, 80), vbBlue, BF
    Line (10, 0)-(20, 80), RGB(0, 0, 219.43), BF
    Line (20, 0)-(30, 80), RGB(0, 0, 182.86), BF
    Line (30, 0)-(40, 80), RGB(0, 0, 146.29), BF
    Line (40, 0)-(50, 80), RGB(0, 0, 109.71), BF
    Line (50, 0)-(60, 80), RGB(0, 0, 73.14), BF
    Line (60, 0)-(70, 80), RGB(0, 0, 36.57), BF
    Line (70, 0)-(80, 80), vbBlack, BF

End Sub

Private Sub Form_Paint()

    GreyScale

End Sub

Private Sub optColour_Click(Index As Integer)

    Cls
    Select Case Index
        Case 0
            GreyScale
        Case 1
            RedScale
        Case 2
            GreenScale
        Case 3
            BlueScale
    End Select
End Sub

