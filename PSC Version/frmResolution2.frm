VERSION 5.00
Begin VB.Form frmChangeResolution 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Change Screen Resolution and Colour"
   ClientHeight    =   1.99995e5
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1.99995e5
   HasDC           =   0   'False
   HelpContextID   =   1950
   Icon            =   "frmResolution2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1.99995e5
   ScaleWidth      =   1.99995e5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOuter 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Screen Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6105
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   9540
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Video Card Specification"
         Top             =   240
         Width           =   9435
      End
      Begin Testcard.MultiButton cmdWinProp 
         Height          =   315
         Left            =   5340
         TabIndex        =   9
         Top             =   5100
         Width           =   3375
         _ExtentX        =   5953
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
         ForeColor       =   0
         BorderColor     =   0
         FillColor       =   16761024
         Caption         =   "Exit Testcard + Display Windows Properties"
         HoverFillColor  =   16744576
         HoverBorderColor=   16744576
         ActiveFillColor =   16761024
         BackColor       =   12648447
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdHelp 
         Height          =   315
         Left            =   7500
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1200
         _ExtentX        =   2117
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
         ForeColor       =   0
         BorderColor     =   0
         FillColor       =   16761024
         Caption         =   "Help"
         HoverFillColor  =   16744576
         ActiveFillColor =   16761024
         BackColor       =   12648447
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdCancel 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   315
         Left            =   5340
         TabIndex        =   7
         Top             =   4560
         Width           =   1200
         _ExtentX        =   2117
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
         ForeColor       =   0
         BorderColor     =   0
         FillColor       =   16761024
         Caption         =   "Cancel"
         HoverFillColor  =   16744576
         ActiveFillColor =   16761024
         BackColor       =   12648447
         CornerRadius    =   10
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "WARNING !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   4755
         Left            =   240
         TabIndex        =   3
         Top             =   1020
         Width           =   4215
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1035
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "frmResolution2.frx":030A
            Top             =   3660
            Width           =   4095
         End
         Begin VB.Label lblText3 
            BackColor       =   &H00C0FFFF&
            Caption         =   $"frmResolution2.frx":03D7
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1515
            Left            =   120
            TabIndex        =   6
            Top             =   2100
            Width           =   3975
         End
         Begin VB.Label lblText 
            BackColor       =   &H00C0FFFF&
            Caption         =   $"frmResolution2.frx":04C7
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1035
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   3975
         End
         Begin VB.Label lblWarn 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Only change the Display Properties prior to any adjustments. Altering the display may require you to readjust certain parameters."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   675
            Left            =   120
            TabIndex        =   4
            Top             =   420
            Width           =   3705
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   2460
         Left            =   5385
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3330
      End
      Begin VB.Label lblModes 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Unknown No. of Modes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4905
         TabIndex        =   2
         Top             =   4020
         Width           =   4215
      End
      Begin VB.Shape shpBlue 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Height          =   4575
         Left            =   4620
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmChangeResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i          As Integer
Dim dblreturn  As Variant
Dim Index      As Long

Private Sub Form_Load()
    'centre frame

    fraOuter.Left = (ScaleWidth - fraOuter.Width) / 2
    fraOuter.Top = (ScaleHeight - fraOuter.Height) / 2
    
    'get info & fill in details
    FillList List1

    lblModes.Caption = "Total Number of Video Card Modes =  " & List1.ListCount

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'clears down any Help file
    Call QuitHelp
    
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    'resets to main form
 
    Unload Me
    frmMain.Show
    frmMain.Enabled = True
    frmMain.WindowState = 2

End Sub

Private Sub cmdWinProp_Click()
    'opens Windows Properties Settings
 
    dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", vbNormalFocus)
 
    Ex ' call exit

End Sub

Private Sub cmdHelp_Click()

    ShowHelpTopic Hlp_Screen_Properties

End Sub

Private Sub Form_Activate()

    frmMain.Enabled = False

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

List1.ListIndex = -1

End Sub
