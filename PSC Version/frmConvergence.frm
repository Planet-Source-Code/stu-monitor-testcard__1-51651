VERSION 5.00
Begin VB.Form frmConvergence 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   FillColor       =   &H80000000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000B&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Tag             =   "115"
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraCol 
      Height          =   4095
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   7223
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
      Caption         =   "      Display Colour"
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
      CheckedBorderColor=   16777215
      BackColor       =   0
      CornerRadius    =   10
      Begin Testcard.MultiButton cmdClear 
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   3420
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
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   3420
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
         Height          =   2355
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4154
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
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Grid Select"
         RedrawOnHover   =   0   'False
         Alignment       =   0
         VerticalAlignment=   0
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   5
         Begin Testcard.MultiButton MultiButton2 
            Height          =   915
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   1614
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
            Caption         =   "Font Size"
            RedrawOnHover   =   0   'False
            Alignment       =   0
            VerticalAlignment=   0
            ButtonMode      =   2
            BackColor       =   8388608
            CornerRadius    =   5
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   17
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
               Caption         =   "8 Point"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   5
               Left            =   1560
               TabIndex        =   16
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
               Caption         =   "12 Point"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   6
               Left            =   2940
               TabIndex        =   15
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
               Caption         =   "16 Point"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   7
               Left            =   4320
               TabIndex        =   14
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
               Caption         =   "24 Point"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
         End
         Begin Testcard.MultiButton MultiButton1 
            Height          =   915
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   1614
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
            Caption         =   "Dot Size"
            RedrawOnHover   =   0   'False
            Alignment       =   0
            VerticalAlignment=   0
            ButtonMode      =   2
            BackColor       =   8388608
            CornerRadius    =   5
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   3
               Left            =   4320
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
               Caption         =   "Point x 8"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   2
               Left            =   2940
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
               Caption         =   "Point x 4"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   1
               Left            =   1560
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
               Caption         =   "Point x 2"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
            Begin Testcard.MultiButton optDot 
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   9
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
               Caption         =   "Single Point"
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
               OptionName      =   "Size"
               BackColor       =   8388608
               CornerRadius    =   10
            End
         End
      End
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   0
         Left            =   180
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
         Left            =   1800
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
         Left            =   3360
         TabIndex        =   2
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
         Index           =   3
         Left            =   4980
         TabIndex        =   1
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
   End
End
Attribute VB_Name = "frmConvergence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'checks for dots or text
Dim DisplayState As Boolean

Dim b As String
Dim c As String
Dim d As Integer

Dim x As Integer
Dim y As Integer
Dim i As Integer

Dim size

Private Sub Form_Activate()
    'startup settings

    fraCol.Left = (Width - fraCol.Width) / 2
    fraCol.Top = (Height - fraCol.Height) / 2
    fraCol.Visible = True
    fraCol.Visible = True
            
    FillStyle = 0
    FillColor = vbWhite
     
    BackColor = vbBlack
    ForeColor = vbWhite
    
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

Private Sub Text()
    
    'displays text as test

    DisplayState = False
    Cls

    For i = optDot.LBound To optDot.UBound
        If optDot(i).Value = True Then
            b = "..TEXT " & Format$(Val(Left$(optDot(i).Caption, 2)), "@@") & "  Point....Convergence Test.."
            Exit For
        End If
    Next
    
    c = b + b + b + b + b + b

    For d = 1 To 100
        Print c
    Next d

End Sub
Private Sub Draw()

    'displays dots as test

    DisplayState = True
    Cls

    ScaleWidth = 100
    ScaleHeight = 100
 
    For x = 2 To 100 Step 5
        For y = 2 To 100 Step 5

            Circle (x, y), size
 
        Next y
    Next x

End Sub

Private Sub Form_Paint()
   
    For i = optDot.LBound To optDot.UBound
        If optDot(i).Value = True Then
            optDot_Click i
            Exit For
        End If
    Next

End Sub

Private Sub optColour_Click(Index As Integer)

    Select Case Index
        Case 0
            ForeColor = vbWhite
            FillColor = vbWhite
        Case 1
            ForeColor = vbRed
            FillColor = vbRed
        Case 2
            ForeColor = vbGreen
            FillColor = vbGreen
        Case 3
            ForeColor = vbBlue
            FillColor = vbBlue
    End Select

    If DisplayState = True Then Draw Else Text

End Sub

Private Sub optDot_Click(Index As Integer)

    If Index < 4 Then
        Select Case Index
            Case 0
                size = 0.1
            Case 1
                size = 0.2
            Case 2
                size = 0.4
            Case 3
                size = 0.8
        End Select
        
        Draw
    Else
        Debug.Print "FontSize " & Val(Left$(optDot(Index).Caption, 2))
        FontSize = Val(Left$(optDot(Index).Caption, 2))
        Text
    End If
End Sub

Private Sub cmdClear_Click()

    fraCol.Visible = False

End Sub

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_Convergence)

End Sub
