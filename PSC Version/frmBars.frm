VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBars 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   1.99995e5
   ClientLeft      =   0
   ClientTop       =   495
   ClientWidth     =   1.99995e5
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000004&
   HelpContextID   =   1980
   Icon            =   "frmBars.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   13333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "118"
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraRGB 
      Height          =   2835
      Left            =   3900
      TabIndex        =   0
      Top             =   4200
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5001
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
      Caption         =   "    Colour Select"
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
      BackColor       =   12632256
      CornerRadius    =   10
      Begin VB.TextBox txtSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   9
         Text            =   " %"
         Top             =   1020
         Width           =   435
      End
      Begin VB.TextBox txtSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   8
         Top             =   1020
         Width           =   435
      End
      Begin MSComctlLib.Slider sldSat 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   1440
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   10
         Min             =   80
         Max             =   100
         SelectRange     =   -1  'True
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin Testcard.MultiButton chkColour 
         Height          =   315
         Index           =   2
         Left            =   3600
         TabIndex        =   5
         Top             =   420
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
         Picture         =   "frmBars.frx":000C
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "ENABLED"
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         Value           =   -1  'True
         CheckedBorderColor=   65535
         CheckedFillColor=   12582912
         CheckedForeColor=   16777215
         ButtonMode      =   1
         CheckedPicture  =   "frmBars.frx":05A6
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton chkColour 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   420
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
         Picture         =   "frmBars.frx":0700
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "ENABLED"
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         Value           =   -1  'True
         CheckedBorderColor=   65535
         CheckedFillColor=   49152
         CheckedForeColor=   16777215
         ButtonMode      =   1
         CheckedPicture  =   "frmBars.frx":0C9A
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton chkColour 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   420
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
         Picture         =   "frmBars.frx":0DF4
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "ENABLED"
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         Value           =   -1  'True
         CheckedBorderColor=   65535
         CheckedFillColor=   255
         CheckedForeColor=   16777215
         ButtonMode      =   1
         CheckedPicture  =   "frmBars.frx":138E
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   2160
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
         TabIndex        =   1
         Top             =   2160
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
      Begin VB.TextBox txtSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Text            =   "Percentage of Saturation :"
         Top             =   1020
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'used for calculations
Dim i As Integer
Dim x As Integer
Dim y As Integer
Dim lVal As Long

'colour values
Dim Ra As Integer
Dim Ga As Integer
Dim Ba As Integer

Private Sub chkColour_Click(Index As Integer)
    
    If chkColour(Index).Value = True Then
        lVal = sldSat.Value * 2.55
    Else
        lVal = 0
    End If
    
    If chkColour(Index).Value = True Then
        chkColour(Index).Caption = "ENABLED"
    Else
        chkColour(Index).Caption = "DISABLED"
    End If
    
    Select Case Index
        Case 0
            Ra = lVal
        Case 1
            Ga = lVal
        Case 2
            Ba = lVal
    End Select
    
    CB
End Sub

Private Sub cmdClear_Click()
    'clear toolbar with mouse
    
    fraRGB.Visible = False

End Sub

Private Sub cmdHelp_Click()
    'helpfile
    
    Call ShowHelpTopic(Hlp_Colour_Bars)

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, fraRGB, False
    
End Sub

Private Sub Form_Load()

txtSlider(1).Text = 100

    Ra = sldSat.Value * 2.55
    Ga = sldSat.Value * 2.55
    Ba = sldSat.Value * 2.55

    'frame position
    fraRGB.Left = (ScaleWidth - fraRGB.Width) / 2
    fraRGB.Top = (ScaleHeight - fraRGB.Height) / 2

    'call colour bars
    CB

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub
Private Sub CB()
    'draws the colours

    ScaleWidth = 80
    ScaleHeight = 80

If chkColour(0).Value = True And chkColour(1).Value = True And chkColour(2).Value = True Then

    Line (0, 0)-(10, 80), vbWhite, BF
    
Else

    Line (0, 0)-(10, 80), RGB(Ra, Ga, Ba), BF
    
End If


    Line (10, 0)-(20, 80), RGB(Ra, Ga, 0), BF
    Line (20, 0)-(30, 80), RGB(0, Ga, Ba), BF
    Line (30, 0)-(40, 80), RGB(0, Ga, 0), BF
    Line (40, 0)-(50, 80), RGB(Ra, 0, Ba), BF
    Line (50, 0)-(60, 80), RGB(Ra, 0, 0), BF
    Line (60, 0)-(70, 80), RGB(0, 0, Ba), BF
    Line (70, 0)-(80, 80), vbBlack, BF

End Sub

Private Sub Form_Paint()

    CB
   
End Sub

Private Sub sldSat_Scroll()

    If chkColour(0).Value = True Then Ra = sldSat.Value * 2.55
    If chkColour(1).Value = True Then Ga = sldSat.Value * 2.55
    If chkColour(2).Value = True Then Ba = sldSat.Value * 2.55

    txtSlider(1).Text = sldSat.Value

    CB

End Sub

