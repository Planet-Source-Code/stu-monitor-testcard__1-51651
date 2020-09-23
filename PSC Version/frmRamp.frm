VERSION 5.00
Begin VB.Form frmRamp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2.45745e5
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2.45745e5
   ForeColor       =   &H00000000&
   HelpContextID   =   1890
   Icon            =   "frmRamp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   16383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16383
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Testcard.MultiButton fraCol 
      Height          =   1575
      Left            =   2280
      TabIndex        =   0
      Top             =   2880
      Width           =   5475
      _ExtentX        =   9657
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
      Begin Testcard.MultiButton optColour 
         Height          =   315
         Index           =   3
         Left            =   4140
         TabIndex        =   6
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
         TabIndex        =   5
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
      Begin Testcard.MultiButton cmdHelp 
         Height          =   375
         Left            =   3780
         TabIndex        =   2
         Top             =   960
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
         Top             =   960
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
Attribute VB_Name = "frmRamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R As Integer
Dim g As Integer
Dim b As Integer

Dim intY As Integer
Dim EndColour As Long

Private Sub cmdClear_Click()
 'clear toolbar
 
    fraCol.Visible = False
    
End Sub

Private Sub cmdHelp_Click()

    Call ShowHelpTopic(Hlp_Ramp)

End Sub

Private Sub Form_Activate()

    Ramp Me, , , 2000

End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    'SwitchTestScreen KeyAscii, Nothing, False
    SwitchTestScreen KeyAscii, fraCol, False
    
    If KeyAscii = vbKeyEscape Then frmRamp.Hide
    If KeyAscii = vbKeySpace Then frmRamp.Hide
    
End Sub

Private Sub Form_Load()

    'start with white
    R = 255
    g = 255
    b = 255
   
    fraCol.Left = (ScaleWidth - fraCol.Width) / 2
    fraCol.Top = (ScaleHeight - fraCol.Height) / 2
    fraCol.Visible = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then frmMain.Show

End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
    Unload Me

End Sub

Private Sub Ramp(pObject As Object, Optional Colour As Integer, Optional Orientation As Integer = 0, Optional Range As Integer = 2000)

    Cls

    MousePointer = 11
    
    pObject.Scale (0, 0)-(Range, Range)

    For intY = 0 To Range

        'this line dictates the colour scheme
        
            EndColour = RGB(CInt((intY / Range) * R), CInt((intY / Range) * g), CInt((intY / Range) * b))
        'left to right
        pObject.Line (intY, 0)-(intY, Range), EndColour

    Next intY

    MousePointer = 0

End Sub

Private Sub optColour_Click(Index As Integer)

  Select Case Index
  
        Case 0
            R = 255
            g = 255
            b = 255
            
        Case 1
            R = 255
            g = 0
            b = 0
            
        Case 2
            R = 0
            g = 255
            b = 0
            
        Case 3
            R = 0
            g = 0
            b = 255
            
    End Select
    
 Form_Activate
 
End Sub
