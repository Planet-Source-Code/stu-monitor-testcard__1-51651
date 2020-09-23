VERSION 5.00
Begin VB.Form frmMoire 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13230
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   2030
   Icon            =   "frmMoire.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   882
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraRes 
      Height          =   2610
      Left            =   4500
      TabIndex        =   0
      Top             =   3660
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4604
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
      Caption         =   "   "
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
      BackColor       =   8388608
      CornerRadius    =   10
      Begin Testcard.MultiButton fraCol 
         Height          =   810
         Left            =   180
         TabIndex        =   7
         Top             =   60
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1429
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
         BorderColor     =   8388608
         FillColor       =   8388608
         Caption         =   "Display Colour"
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
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   10
         Begin Testcard.MultiButton optColour 
            Height          =   315
            Index           =   3
            Left            =   3780
            TabIndex        =   11
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
            Left            =   2520
            TabIndex        =   10
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
            Left            =   1260
            TabIndex        =   9
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
            Left            =   0
            TabIndex        =   8
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
      End
      Begin Testcard.MultiButton fraDraw 
         Height          =   855
         Left            =   180
         TabIndex        =   3
         Top             =   1020
         Width           =   4875
         _ExtentX        =   8599
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
         Caption         =   "  Spacing"
         RedrawOnHover   =   0   'False
         Alignment       =   0
         VerticalAlignment=   0
         ButtonMode      =   2
         BackColor       =   8388608
         CornerRadius    =   5
         Begin Testcard.MultiButton optSpace 
            Height          =   315
            Index           =   2
            Left            =   3600
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
            Caption         =   "3 Pixels"
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
         Begin Testcard.MultiButton optSpace 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   5
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
            Caption         =   "2 Pixels"
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
         Begin Testcard.MultiButton optSpace 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   4
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
            Caption         =   "1 Pixel"
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
      Begin Testcard.MultiButton cmdHelp 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   2100
         Width           =   1455
         _ExtentX        =   2566
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
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   2100
         Width           =   3195
         _ExtentX        =   5636
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
Attribute VB_Name = "frmMoire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim v As Integer
Dim w As Integer
Dim x As Integer
Dim y As Integer
Dim Z As Integer

Dim Index As Integer
Dim i As Integer

Private Sub cmdClear_Click()

    fraRes.Visible = False
 
End Sub

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_Moire)

End Sub

Private Sub Form_Activate()
    'default settings
    
    pDraw

    fraRes.Left = (ScaleWidth - fraRes.Width) / 2
        fraRes.Top = (ScaleHeight - fraRes.Height) / 2
            fraRes.Visible = True
        
End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, fraRes, True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub pDraw()
    
'    For i = optType.LBound To optType.UBound
'        If optType(i).Value = True Then
'            Index = i
'            Exit For
'        End If
'    Next
    
    For i = optSpace.LBound To optSpace.UBound
        If optSpace(i).Value = True Then
            v = Val(Left$(optSpace(i).Caption, 1)) + 1
            Exit For
        End If
    Next
    
    Cls
'    Select Case Index
'        Case 0
'            Vert
'        Case 1
            Mix
       
   ' End Select
End Sub

Private Sub Horiz()
    'draw horizontal lines

    ScaleMode = 3

    w = Screen.Height
    x = Screen.Height / v
    y = Screen.Height / x

    For Z = 0 To x Step y
    
        Line (0, Z)-(Screen.Width, Z)

    Next Z

End Sub

Private Sub Vert()
    'draw vertical lines

    ScaleMode = 3

    w = Screen.Width
    x = Screen.Width / v
    y = Screen.Width / x

    For Z = 0 To w Step y

        Line (Z, 0)-(Z, Screen.Height)

    Next Z

End Sub

Private Sub Mix()

    Horiz
    Vert
    
End Sub

Private Sub optSpace_Click(Index As Integer)

    pDraw
    
End Sub
'
'Private Sub optType_Click(Index As Integer)
'
'    pDraw
'
'End Sub

Private Sub optColour_Click(Index As Integer)

    Cls
    
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
    
    pDraw
    
End Sub

