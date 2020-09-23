VERSION 5.00
Begin VB.Form frmGeometry 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   628
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   ShowInTaskbar   =   0   'False
   Tag             =   "113"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraCol 
      Height          =   2550
      Left            =   3780
      TabIndex        =   0
      Top             =   3750
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4498
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
      Caption         =   "     Display Colour"
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
      Begin Testcard.MultiButton cmdClear 
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   1920
         Width           =   3195
         _ExtentX        =   5636
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
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   3780
         TabIndex        =   9
         Top             =   1920
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
      Begin Testcard.MultiButton fraDraw 
         Height          =   855
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Line Width"
         RedrawOnHover   =   0   'False
         Alignment       =   0
         VerticalAlignment=   0
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   5
         Begin Testcard.MultiButton chkCircle 
            Height          =   315
            Left            =   3840
            TabIndex        =   11
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            Picture         =   "frmGeometry.frx":0000
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "Circles"
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
            CheckedPicture  =   "frmGeometry.frx":059A
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optLineWidth 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "Single"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   2
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            Value           =   -1  'True
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optLineWidth 
            Height          =   315
            Index           =   1
            Left            =   1380
            TabIndex        =   7
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "Width x 2"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   2
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optLineWidth 
            Height          =   315
            Index           =   2
            Left            =   2580
            TabIndex        =   6
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "Width x 3"
            HoverForeColor  =   65535
            HoverFillColor  =   16744576
            HoverBorderColor=   65535
            PictureAlignment=   2
            ActiveBorderColor=   65535
            ActiveForeColor =   65535
            ActiveFillColor =   16761024
            CheckedBorderColor=   65535
            CheckedFillColor=   16761024
            ButtonMode      =   1
            OptionName      =   "Line"
            BackColor       =   8388608
            CornerRadius    =   10
         End
      End
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "White"
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
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   3
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
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
         Index           =   2
         Left            =   2820
         TabIndex        =   2
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
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
         Index           =   3
         Left            =   4140
         TabIndex        =   1
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
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
   End
End
Attribute VB_Name = "frmGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As Integer
Dim y As Integer

Private Sub chkCircle_Click()

    Draw

End Sub

Private Sub Form_Activate()

    fraCol.Left = (ScaleWidth - fraCol.Width) / 2
    fraCol.Top = (ScaleHeight - fraCol.Height) / 2
    fraCol.Visible = True

    ForeColor = vbWhite

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, fraCol, True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

Private Sub Form_Paint()

    Draw

End Sub

Private Sub optColour_Click(Index As Integer)

    Select Case Index
        Case 0
            ForeColor = vbWhite
        Case 1
            ForeColor = vbRed
        Case 2
            ForeColor = vbGreen
        Case 3
            ForeColor = vbBlue
    End Select
    
    Draw
    
End Sub

Private Sub optLineWidth_Click(Index As Integer)

    DrawWidth = Index + 1
    Draw
    
End Sub

Private Sub Draw()
    'draw grid

    ScaleWidth = 801
    ScaleHeight = 601

    BackColor = vbBlack

    Cls

    For x = 0 To 800 Step 40

        Line (x, 600)-(x, 0)

    Next x

    For y = 0 To 600 Step 40

        Line (0, y)-(800, y)

    Next y

    'draw circles if requested
    If chkCircle.Value = True Then

        Circle (400, 300), 400
        Circle (400, 300), 299
        Circle (400, 300), 199

    End If

End Sub

Private Sub cmdClear_Click()

    fraCol.Visible = False

End Sub

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_Geometry1)

End Sub
