VERSION 5.00
Begin VB.Form frmCircles 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15090
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   HelpContextID   =   1970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   655
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1006
   ShowInTaskbar   =   0   'False
   Tag             =   "114"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraCol 
      Height          =   3495
      Left            =   4320
      TabIndex        =   0
      Top             =   3780
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   6165
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
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   3
         Left            =   4140
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
         Left            =   2820
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
         Left            =   1500
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
      Begin Testcard.MultiButton fraRes 
         Height          =   855
         Left            =   180
         TabIndex        =   9
         Top             =   1860
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
         Caption         =   "Line Mode"
         RedrawOnHover   =   0   'False
         Alignment       =   0
         VerticalAlignment=   0
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   5
         Begin Testcard.MultiButton optLineMode 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   10
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
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "LOW"
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
            OptionName      =   "Mode"
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optLineMode 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   11
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
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "NORMAL"
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
            OptionName      =   "Mode"
            BackColor       =   8388608
            CornerRadius    =   10
         End
         Begin Testcard.MultiButton optLineMode 
            Height          =   315
            Index           =   2
            Left            =   3660
            TabIndex        =   12
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
            BorderColor     =   16777215
            FillColor       =   8388608
            Caption         =   "HIGH"
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
            OptionName      =   "Mode"
            BackColor       =   8388608
            CornerRadius    =   10
         End
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
         Begin Testcard.MultiButton optLineWidth 
            Height          =   315
            Index           =   2
            Left            =   3660
            TabIndex        =   8
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
         Begin Testcard.MultiButton optLineWidth 
            Height          =   315
            Index           =   1
            Left            =   1920
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
            Index           =   0
            Left            =   180
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
      End
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   3780
         TabIndex        =   14
         Top             =   2940
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
      Begin Testcard.MultiButton cmdClear 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   2940
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
   End
End
Attribute VB_Name = "frmCircles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x
Dim y

Private Sub Draw()

    'reset screen
    Cls
    ScaleWidth = 100
    ScaleHeight = 100
    
    'center cross
    Line (46, 50)-(54, 50)
    Line (50, 45)-(50, 55)
          
    'corners
    Line (0, 0)-(5, 5)
    Line (100, 0)-(95, 5)
    Line (0, 100)-(5, 95)
    Line (100, 100)-(95, 95)
 
    '------------------------------------
    'ensure single line when High Mode (other sizes are meaningless)
    If y = -0.2 Then

        DrawWidth = 1
        optLineWidth(0).Value = True
        optLineWidth(1).Caption = "Unavailable"
        optLineWidth(2).Caption = "Unavailable"
    Else
        optLineWidth(1).Caption = "Width x 2"
        optLineWidth(2).Caption = "Width x 3"
    End If
       
    'circles
        
    For x = 12 To 0 Step y
        Circle (20, 20), x
        Circle (80, 20), x
        Circle (20, 80), x
        Circle (80, 80), x
    Next x

    'center circle

    Circle (50, 50), 10

End Sub

Private Sub cmdClear_Click()
    'clear toolbar
    
    fraCol.Visible = False

End Sub

Private Sub cmdHelp_Click()
    'help
    
    Call ShowHelpTopic(Hlp_Distortion)

End Sub

Private Sub Form_Activate()
    'load startup settings

    fraCol.Left = (ScaleWidth - fraCol.Width) / 2
    fraCol.Top = (ScaleHeight - fraCol.Height) / 2
    fraCol.Visible = True
        
    y = -2 'default medium circles
    
    Draw
        
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

Private Sub optLineMode_Click(Index As Integer)

    Select Case Index
        Case 0
            y = -12
        Case 1
            y = -2
        Case 2
            y = -0.2
    End Select
    
    Draw
End Sub

Private Sub optLineWidth_Click(Index As Integer)

    DrawWidth = Index + 1
    Draw
    
End Sub

