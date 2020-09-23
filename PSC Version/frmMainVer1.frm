VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " Computer Monitor Testcard "
   ClientHeight    =   2.45085e5
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2.45655e5
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   160
   Icon            =   "frmMainVer1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   16339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16377
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtLoad 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   46
      Text            =   "LOADING SOUND FILE............................"
      Top             =   1200
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.Frame frEnd 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   555
      Left            =   4620
      TabIndex        =   43
      Top             =   7680
      Width           =   2535
      Begin Testcard.MultiButton cmdMin 
         Height          =   555
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   65535
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Minimize"
         HoverForeColor  =   65535
         HoverFillColor  =   255
         HoverBorderColor=   16777215
         BackColor       =   0
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton cmdExit 
         Height          =   555
         Left            =   840
         TabIndex        =   44
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   65535
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "End Program"
         HoverForeColor  =   65535
         HoverFillColor  =   255
         HoverBorderColor=   16777215
         BackColor       =   0
         CornerRadius    =   10
      End
   End
   Begin VB.Frame frSound 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   3180
      MouseIcon       =   "frmMainVer1.frx":014A
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
      Begin Testcard.MultiButton Sound 
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         Picture         =   "frmMainVer1.frx":029C
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "     Tone"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         PictureAlignment=   1
         ActiveFillColor =   8454016
         CheckedFillColor=   8454016
         ButtonMode      =   1
         OptionName      =   "snd"
         CheckedPicture  =   "frmMainVer1.frx":0836
         BackColor       =   8388608
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton Sound 
         Height          =   435
         Index           =   2
         Left            =   0
         TabIndex        =   40
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         MousePointer    =   99
         Picture         =   "frmMainVer1.frx":0990
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "     Music"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         PictureAlignment=   1
         ActiveFillColor =   8454016
         MouseIcon       =   "frmMainVer1.frx":0F2A
         CheckedFillColor=   8454016
         ButtonMode      =   1
         OptionName      =   "snd"
         CheckedPicture  =   "frmMainVer1.frx":108C
         BackColor       =   8388608
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton Sound 
         Default         =   -1  'True
         Height          =   435
         Index           =   3
         Left            =   0
         TabIndex        =   41
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   255
         Caption         =   "Sound Off"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         PictureAlignment=   1
         ActiveFillColor =   255
         CheckedFillColor=   255
         ButtonMode      =   1
         OptionName      =   "snd"
         BackColor       =   8388608
         CornerRadius    =   5
      End
   End
   Begin VB.Frame frHelp 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   4395
      Left            =   1980
      MouseIcon       =   "frmMainVer1.frx":11E6
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      Begin Testcard.MultiButton tool 
         Height          =   440
         Index           =   8
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "Video Card Properties"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16777215
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   440
         Index           =   9
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "System Properties"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16777215
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   440
         Index           =   10
         Left            =   0
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "System Info"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16777215
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   11
         Left            =   0
         TabIndex        =   33
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "Sound Properties"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16777215
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   12
         Left            =   0
         TabIndex        =   34
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         FillColor       =   0
         Caption         =   "Help"
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   13
         Left            =   0
         TabIndex        =   35
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         FillColor       =   0
         Caption         =   "E-Mail Us"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   14
         Left            =   0
         TabIndex        =   36
         Top             =   3960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         FillColor       =   0
         Caption         =   "About"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   15
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         Picture         =   "frmMainVer1.frx":1338
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "Sound"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         PictureAlignment=   1
         ActiveFillColor =   16777215
         CheckedFillColor=   16777215
         ButtonMode      =   1
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton tool 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   3000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         FillColor       =   0
         Caption         =   " Monitor Tutorial"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16777215
         BackColor       =   0
         CornerRadius    =   5
      End
   End
   Begin Testcard.MultiButton frExit 
      Height          =   1635
      Left            =   5160
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   2884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65535
      BorderColor     =   65535
      FillColor       =   8388608
      Caption         =   "                                                   Finished Testing ?"
      HoverForeColor  =   65535
      HoverFillColor  =   8388608
      HoverBorderColor=   65535
      ActiveBorderColor=   65535
      ActiveForeColor =   65535
      ActiveFillColor =   8388608
      VerticalAlignment=   0
      Value           =   -1  'True
      BackColor       =   8388608
      CornerRadius    =   20
      Begin Testcard.MultiButton lblNO 
         Height          =   495
         Left            =   2280
         TabIndex        =   28
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         Picture         =   "frmMainVer1.frx":178A
         BorderColor     =   255
         FillColor       =   8388608
         Caption         =   "      NO"
         HoverFillColor  =   255
         HoverBorderColor=   255
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   8388608
         Value           =   -1  'True
         BackColor       =   8388608
         CornerRadius    =   10
      End
      Begin Testcard.MultiButton lblYES 
         Height          =   495
         Left            =   300
         TabIndex        =   27
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         Picture         =   "frmMainVer1.frx":1D24
         BorderColor     =   65280
         FillColor       =   8388608
         Caption         =   "      YES"
         HoverFillColor  =   65280
         HoverBorderColor=   65280
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   8388608
         Value           =   -1  'True
         BackColor       =   8388608
         CornerRadius    =   10
      End
   End
   Begin VB.Frame frToolbar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   1725
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   10
         Left            =   0
         TabIndex        =   12
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":1E7E
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "  Purity Screens  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   0
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":22D0
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Geometry  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   12632256
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":2562
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Distortion  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":287C
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Regulation  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":2B96
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Convergence  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   4
         Left            =   0
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":37E8
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Resolution  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   5
         Left            =   0
         TabIndex        =   7
         Top             =   3600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":3B02
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Ramp  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   6
         Left            =   0
         TabIndex        =   8
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":3D94
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Pluge  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   7
         Left            =   0
         TabIndex        =   9
         Top             =   4800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":40AE
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Greyscale  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   8
         Left            =   0
         TabIndex        =   10
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":43C8
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Colour Bars  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   9
         Left            =   0
         TabIndex        =   11
         Top             =   6000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":46E2
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Testcard  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   11
         Left            =   0
         TabIndex        =   22
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":49FC
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "Moiré  "
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton cmd 
         Height          =   555
         Index           =   12
         Left            =   0
         TabIndex        =   42
         Top             =   7200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
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
         Picture         =   "frmMainVer1.frx":4D16
         BorderColor     =   16777215
         FillColor       =   8388608
         Caption         =   "  Tools / Help"
         HoverForeColor  =   65535
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         Alignment       =   0
         PictureAlignment=   1
         ActiveBorderColor=   65535
         ActiveForeColor =   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2.44710e5
      Width           =   2.45655e5
      _ExtentX        =   433308
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   425556
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRuntime 
      Interval        =   1000
      Left            =   6660
      Top             =   2940
   End
   Begin VB.Timer tmrIcon 
      Interval        =   400
      Left            =   6060
      Top             =   2940
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5220
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":5168
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":527A
            Key             =   "sound"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":5594
            Key             =   "hlp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":59E6
            Key             =   "comp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":5D00
            Key             =   "dist"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":5E5A
            Key             =   "about"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":64A4
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":67BE
            Key             =   "screen"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":6D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":71AA
            Key             =   "GEOico"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":7304
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":789E
            Key             =   "grey"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":7BBA
            Key             =   "bars"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":7ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":822A
            Key             =   "test"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":8546
            Key             =   "r"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":8862
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":8CB6
            Key             =   "g"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":8FD2
            Key             =   "b"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":92EE
            Key             =   "m"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":95AA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":9866
            Key             =   "y"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":9B22
            Key             =   "w"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":9E3E
            Key             =   "bl"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":A15A
            Key             =   "geo"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":A4AE
            Key             =   "prog"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":A902
            Key             =   "reg"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":AC2A
            Key             =   "abc"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":B07E
            Key             =   "conv"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":B3A6
            Key             =   "resol"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":B6CE
            Key             =   "pluge"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainVer1.frx":B9F6
            Key             =   "ramp"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frColour 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   3795
      Left            =   1980
      MouseIcon       =   "frmMainVer1.frx":BC88
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   65280
         FillColor       =   65280
         Caption         =   "Green"
         HoverFillColor  =   16744576
         HoverBorderColor=   8454016
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16711680
         FillColor       =   16711680
         Caption         =   "Blue"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   16711680
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16711935
         FillColor       =   16711935
         Caption         =   "Magenta"
         HoverFillColor  =   16744576
         HoverBorderColor=   16711935
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16776960
         FillColor       =   16776960
         Caption         =   "Cyan"
         HoverFillColor  =   16744576
         HoverBorderColor=   16776960
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   5
         Left            =   0
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   65535
         FillColor       =   65535
         Caption         =   "Yellow"
         HoverFillColor  =   16744576
         HoverBorderColor=   65535
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   6
         Left            =   0
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   16777215
         FillColor       =   16777215
         Caption         =   "White"
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   7
         Left            =   0
         TabIndex        =   20
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
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
         FillColor       =   0
         Caption         =   "Black"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   16777215
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
      Begin Testcard.MultiButton col 
         Height          =   440
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   767
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
         BorderColor     =   255
         FillColor       =   255
         Caption         =   "Red"
         HoverForeColor  =   0
         HoverFillColor  =   16744576
         HoverBorderColor=   255
         ActiveFillColor =   16761024
         BackColor       =   0
         CornerRadius    =   5
      End
   End
   Begin VB.Frame frTitle 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   240
      TabIndex        =   23
      Top             =   60
      Width           =   9675
      Begin VB.Label lblAuthor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "© SJS TV Services Ltd.  2004"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   6420
         TabIndex        =   25
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label lblCMT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPUTER MONITOR TESTCARD   "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   0
         MouseIcon       =   "frmMainVer1.frx":BDDA
         TabIndex        =   24
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.Label LblShare 
      BackColor       =   &H00000000&
      Caption         =   "SHAREWARE VERSION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   47
      Top             =   780
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   5
      Left            =   7440
      Picture         =   "frmMainVer1.frx":C0E4
      Top             =   5700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   4
      Left            =   6840
      Picture         =   "frmMainVer1.frx":C3EE
      Top             =   5700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   6240
      Picture         =   "frmMainVer1.frx":C6F8
      Top             =   5700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   5940
      Picture         =   "frmMainVer1.frx":CA02
      Top             =   5820
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   5580
      Picture         =   "frmMainVer1.frx":CB4C
      Top             =   5820
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   5220
      Picture         =   "frmMainVer1.frx":CC96
      Top             =   5820
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API-Functions to move cursor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'API to locate the taskbar
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'API to hide/show taskbar
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'Constants used to get window handles
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const GW_NEXT = 2
Private Const GW_CHILD = 5

'Used for screen area & colour bit rate
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Const BITSPIXEL = 12
Dim lBits As Long, lWidth As Long, lHeight As Long

'Declare API call sndPlaySound in winmm.dll
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

'API for send e-mail
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

'Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim Menu As Long
Dim i As Integer
Dim dblreturn
Dim sTip As String

'timer
Dim howlong As String
Dim newsec As Integer
Dim cursec As Integer
Dim newmin As String
Dim curmin As Integer
Dim newhour As Integer
Dim curhour As Integer
Dim Minutes As Integer

'msg boxes
Dim Msg As String
Dim Msg1 As String
Dim Msg2 As String
Dim Style As String
Dim Title As String
Dim Ctxt As String
Dim response As String

'used for retrieving state of taskbar
Dim Desktop As Long
Dim mhandle As Long
Dim Temp As String * 16
Dim SrchString As String

Dim WaveCheck As Boolean 'Checks state of sound file
Dim y As Integer 'Used in image counter

'centres exit buttons
Dim Rec As RECT
Dim R As Integer
Dim t As Integer
 
'sound
Dim Start As Long
Dim wFlags%
Dim x As Integer
Dim SoundName As String
Dim file As String
    
Dim help As String

'send e-mail
Dim lngResult As Long

Private Function GetTrayHandle(mType As Integer) As Long
    'Get the window handle
    
    Desktop = GetDesktopWindow()
    mhandle = GetWindow(Desktop, GW_CHILD)
    Do While mhandle <> 0
        GetClassName mhandle, Temp, 14
        If Left$(Temp, 13) = "Shell_TrayWnd" Then
            If mType = 4 Then 'entire taskbar
                GetTrayHandle = mhandle
                Exit Do
            End If
            mhandle = GetWindow(mhandle, GW_CHILD)
           
        End If
        mhandle = GetWindow(mhandle, GW_NEXT)
    Loop

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            frmGeometry.Show
        Case 1
            frmCircles.Show
        Case 2
            frmReg.Show
        Case 3
            frmConvergence.Show
        Case 4
            frmResolutionLines.Show
        Case 5
            Ramp
        Case 6
            frmPluge.Show
        Case 7
            frmGreyscale.Show
        Case 8
            frmBars.Show
        Case 9
            frmTestcard.Show
        Case 11
            frmMoire.Show
      End Select
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Select Case Index
        Case 0
            sTip = " Set Display Position"
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 1
            sTip = " Check for Eliptical Errors "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 2
            sTip = " Frame Error Check (The outer box should not move!)"
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 3
            sTip = " Dots or Text to check for sharpness "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 4
            sTip = " Screen Resolution. Choose Horizontal or Vertical"
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 5
            sTip = " Uniform Black to White (Not available under 24 Bit Colours)"
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 6
            sTip = " Set Brightness + Contrast "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 7
            sTip = " White to Black No-Colour Test "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 8
            sTip = " 100% Bars Colour Test "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 9
            sTip = " General Purpose Testcard "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 10
            sTip = " Select Purity Screen "
            frColour.Visible = True
             frHelp.Visible = False
              frSound.Visible = False
        Case 11
            sTip = " Check Screen for Patterning "
            frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
        Case 12
            sTip = " System Tools and Help "
            frColour.Visible = False
             frHelp.Visible = True
              frSound.Visible = False
     
              
    End Select
    
    StatusBar1.Panels(1).Text = sTip
    
End Sub

Private Sub cmdMin_Click()

WindowState = vbMinimized

End Sub

Private Sub Form_Click()
'clears down any Help file

    Call QuitHelp
    
End Sub

Private Sub Form_Deactivate()
'clears down any Help file

    Call QuitHelp
    
End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, Shift As Integer)

    SwitchTestScreen KeyAscii, Nothing, False
    
End Sub

Private Sub frToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Call QuitHelp

End Sub

Private Sub tool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call QuitHelp

Select Case Index
        
        Case 0
        sTip = " Learn more about your monitor"
            frSound.Visible = False
        Case 8
        sTip = " Displays Available Video Card Resolutions"
            frSound.Visible = False
        Case 15
        sTip = " Choose a Sound Test"
            frSound.Visible = True
        Case 9
        sTip = " Displays System Properties"
            frSound.Visible = False
        Case 10
        sTip = " Advanced System Infomation"
            frSound.Visible = False
        Case 11
        sTip = " Displays Sound Properties"
            frSound.Visible = False
        Case 12
        sTip = " Help Files with a Brief Tutorial"
            frSound.Visible = False
        Case 13
        sTip = " E-mail us with any suggestions"
            frSound.Visible = False
        Case 14
        sTip = " Pretty Credits & who done what :)"
            frSound.Visible = False

End Select

StatusBar1.Panels(1).Text = sTip

End Sub

Private Sub cmdExit_Click()
    'Get Left, Right, Top and Bottom of frmMain
    
    GetWindowRect frmMain.hwnd, Rec
  
   
    frToolbar.Enabled = False
    frEnd.Enabled = False
    
    frExit.Left = (ScaleWidth - frExit.Width) / 2
        frExit.Top = (ScaleHeight - frExit.Height) / 2
        
    frExit.Visible = True
    
    R = (lWidth / 2) + 50
    t = (lHeight / 2) + 20
 
    SetCursorPos Rec.Left + R, Rec.Top + t

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
              
End Sub

Private Sub cmdMin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    frColour.Visible = False
             frHelp.Visible = False
              frSound.Visible = False
              
End Sub

Private Sub col_Click(Index As Integer)

        Select Case Index
            
        Case 0
            frmPurity.BackColor = vbRed
        Case 1
            frmPurity.BackColor = vbGreen
        Case 2
            frmPurity.BackColor = vbBlue
        Case 3
            frmPurity.BackColor = vbMagenta
        Case 4
            frmPurity.BackColor = vbCyan
        Case 5
            frmPurity.BackColor = vbYellow
        Case 6
            frmPurity.BackColor = vbWhite
        Case 7
            frmPurity.BackColor = vbBlack
        
        End Select
    
    frmPurity.Show
    
End Sub

Private Sub Sound_Click(Index As Integer)

        Select Case Index
            
        Case 1
        Call WAVStop
        Call WAVLoop1K
    
        Case 2
        Call WAVStop
        Call WAVLoopMusic
         
        Case 3
        Call WAVStop
         
        End Select
        
End Sub


Private Sub tool_click(Index As Integer)
            Select Case Index
            
        Case 0
    
        frHelp.Visible = False
        frSound.Visible = False
   
        Call ShowHelpTopic2(Hlp_Monitor_Basics)
       
        WAVStop
     
        Sound(1).Value = False
        Sound(2).Value = False
        
'-----------------------------------
        Case 8

        WAVStop
        
        frmChangeResolution.Show
        frmRamp.Hide
        frmMain.Enabled = False
        frmMain.WindowState = 1
    
        Sound(1).Value = False
        Sound(2).Value = False
        Sound(3).Value = True
        
'-----------------------------------
        Case 9
        
        WAVStop
        
        On Error GoTo 0
        On Error Resume Next

        dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)

        If Err.Number Then

        Style = vbCritical
        Title = "System Error"
        Msg = " Sorry, Can not find Computer Specifications. "
        response = MsgBox(Msg, Style, Title)
        
        Err.Clear

        End If
    
        WindowState = vbMinimized

        Resume
    
        Sound(1).Value = False
        Sound(2).Value = False
'----------------------------------
        Case 10
        
        WAVStop
        StartSysInfo
        
        Sound(1).Value = False
        Sound(2).Value = False
'----------------------------------
        Case 11
        
        WAVStop
      
        On Error GoTo 0
        On Error Resume Next

        dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)

        If Err.Number Then

        Style = vbCritical
        Title = "System Error"
        Msg = " Sorry, Can not find Sound Specifications. "
        response = MsgBox(Msg, Style, Title)
        
        Err.Clear

        End If

        Resume
        
        Sound(1).Value = False
        Sound(2).Value = False
'---------------------------------
        Case 12
        
        WAVStop
      
        frHelp.Visible = False
        frSound.Visible = False
            
        ShowHelpTopic Hlp_Computer_Monitor
        
        Sound(1).Value = False
        Sound(2).Value = False
'---------------------------------
        Case 13
        
        WAVStop
        
        'calls host e-mail if available

        On Error GoTo 0
        On Error Resume Next
    
        WindowState = vbMinimized
    
        frHelp.Visible = False
        frSound.Visible = False
    
        ShellExecute Me.hwnd, "Open", "mailto:enquires@sjstv.co.uk", "", "", vbNormalFocus

        If Err.Number Then

        Style = vbCritical
        Title = "Mail Location Error"
        Msg = " Sorry, Can not find your E-Mail program. "
        response = MsgBox(Msg, Style, Title)
        
        Err.Clear

        End If
    
        Sound(1).Value = False
        Sound(2).Value = False
        
        Resume
'---------------------------------
        Case 14
        
        WAVStop
      
        frmAbout.Show
        frHelp.Visible = False
        frSound.Visible = False
        
        End Select
    
        Sound(1).Value = False
        Sound(2).Value = False
        
End Sub

Private Sub Form_Activate()

        GradeForm Me, , , 2000 'sets background from basBlueGrad

        frEnd.Left = (ScaleWidth - frEnd.Width) - 50
        frEnd.Top = (ScaleHeight - ScaleHeight) + 40

    'screen property stuff -------------------------------------------------------
    'get screen area + colour
    
        lBits = GetDeviceCaps(hDC, BITSPIXEL)
        lWidth = Screen.Width \ Screen.TwipsPerPixelX
        lHeight = Screen.Height \ Screen.TwipsPerPixelY

    'display screen area + colour
        StatusBar1.Panels(3).Text = "  Screen Area: " & lWidth & " x " & _
        lHeight & "...Colours:" & "  " & lBits & " Bit  "
    
    'Refresh Rate
        If GetDeviceCaps(hDC, VREFRESH) > 2 Then
            StatusBar1.Panels(4) = "  Refresh Rate: " & GetDeviceCaps(hDC, VREFRESH) & " Hz  "
        Else
            StatusBar1.Panels(4) = "  Refresh Rate:  (System Default)  "
        End If

End Sub

Private Sub Form_Load()

    If frmMain.WindowState = vbMaximized Then
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_HIDEWINDOW

    Else
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
 
    End If

    GradeForm Me, , , 2000 'sets background from basBlueGrad
 

    'Code added for Help
    SetAppHelp Me.hwnd
    Call SetAppHelp(Me.hwnd)
    
    '----------------------------------------------------

    'sets runtime on statusbar
    tmrRuntime.Enabled = True
    howlong = 0
    newsec = 0
    cursec = 0
    newmin = 0
    curmin = 0
    newhour = 0
    curhour = 0

End Sub

Private Sub tim()

    'sets prog time counter
    cursec = cursec + 1

    If cursec >= 60 Then
        curmin = curmin + 1
        cursec = 0

    End If
    
    If cursec < 10 Then
        newsec = "0" & cursec

    Else
        newsec = cursec

    End If

    If curmin >= 60 Then
        curhour = curhour + 1
        curmin = 0

    End If

    If curmin < 10 Then
        newmin = "0" & curmin

    Else

        newmin = curmin

    End If

    newhour = curhour
    
    howlong = newhour & ":" & newmin
    StatusBar1.Panels(2).Text = "Runtime:  " & howlong
   
    'tried various layouts to avoid flicker
    'seconds have been left out until I can find cure
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    StatusBar1.Panels(1).Text = ""
            
End Sub

Private Sub lblNO_Click()

    frToolbar.Enabled = True
    frEnd.Enabled = True

    frExit.Visible = False

End Sub

Private Sub lblYES_Click()

    'reinstate taskbar
    SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
   
    Ex

End Sub

Private Sub runicon()

    ' Advance animation one frame.
    y = y + 1: If y = 6 Then y = 0

    ' Icon animation:  This will only be evident when the
    ' form is minimized.

    frmMain.Icon = imgIcon(y)

End Sub

Private Sub tmrIcon_Timer()

    'sets interval for minimised icons

    If frmMain.WindowState = vbMinimized Then
        runicon

    Else

        frmMain.Icon = imgIcon(0)

    End If

End Sub

Sub WAVStop()

Dim wFlags%


    'ends tone.wav if running
    file = " "
    SoundName$ = file
    x = sndPlaySound(SoundName$, wFlags%)
   
    WaveCheck = False

End Sub

Sub WAVLoop1K()

    DoEvents

    TxtLoad.Visible = True
    Screen.MousePointer = 11
    Start = Timer

    Do While Timer < Start + 1
        DoEvents
    Loop

    On Error GoTo 0
    On Error Resume Next
    'check tone.wav exists
    If Dir(App.Path & "\tone.wav") <> "" Then SoundName$ = Dir(App.Path & "\tone.wav")

    wFlags% = SND_ASYNC Or SND_LOOP
    x = sndPlaySound(SoundName$, wFlags%)
    
    WaveCheck = True

    TxtLoad.Visible = False
 
    Screen.MousePointer = 0
    
    file = Dir(SoundName$)

    'error msg if not found

    If Err.Number Or x = 0 Then


        Style = vbOKOnly + vbCritical
        Title = "Sound Location Error"
        Msg = " Sorry. Unable to carry out Sound Check. "
        Msg1 = " The file is either missing, corrupt, "
        Msg2 = " or the system is unable play wav sound files. "
        response = MsgBox(Msg + vbNewLine + vbNewLine + Msg1 + vbNewLine + Msg2, Style, Title)
        
        Err.Clear
        
        WAVStop
                    
        Sound(1).Value = False
        Sound(2).Value = False
    End If
    
    Resume

End Sub

Sub WAVLoopMusic()

    DoEvents

    TxtLoad.Visible = True
    Screen.MousePointer = 11
    Start = Timer

    Do While Timer < Start + 1
        DoEvents
    Loop

    On Error GoTo 0
    On Error Resume Next
    'check tone.wav exists
    If Dir(App.Path & "\music.wav") <> "" Then SoundName$ = Dir(App.Path & "\music.wav")

    wFlags% = SND_ASYNC Or SND_LOOP
    x = sndPlaySound(SoundName$, wFlags%)
    
    WaveCheck = True

    TxtLoad.Visible = False
 
    Screen.MousePointer = 0
    
    file = Dir(SoundName$)

    'error msg if not found

    If Err.Number Or x = 0 Then


        Style = vbOKOnly + vbCritical
        Title = "Sound Location Error"
        Msg = " Sorry. Unable to carry out Music Test. "
        Msg1 = " The file is either missing, corrupt, "
        Msg2 = " or the system is unable play wav sound files. "
        response = MsgBox(Msg + vbNewLine + vbNewLine + Msg1 + vbNewLine + Msg2, Style, Title)
        
        Err.Clear
        
        WAVStop
        
        Sound(1).Value = False
        Sound(2).Value = False
    End If
    
    Resume

End Sub

Private Sub tmrRuntime_Timer()
    'sets interval for prog run time
    
    tim

End Sub

Private Sub Ramp()

    'resolution pixel check
    If Val(GetDeviceCaps(hDC, BITSPIXEL)) < 24 Then
    
        Msg = "THIS TEST IS NOT VALID !"
        Msg1 = "The Colour Bit Rate is not set to at least 24 Bits Per Pixel."
        Msg2 = "Please refer to HELP for more information."

        Style = vbOKOnly + vbCritical + vbMsgBoxHelpButton
        Title = "Screen Property Error"
        help = "CMT Help.HLP"   ' Define Help file.
        Ctxt = 1950   ' Define topic
      
        response = MsgBox(Msg + Chr(13) + Msg1 + Chr(13) + Chr(13) + Msg2, Style, Title, help, Ctxt)

        Else: frmRamp.Show

    End If

End Sub

Public Sub StartSysInfo()
    'standard VB Sysinfo stuff

    Dim rc As Long
    Dim SysInfoPath As String

    On Error GoTo SysInfoErr
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, 2)
        WindowState = vbMinimized
    Exit Sub

SysInfoErr:
       
    MsgBox "Advanced System Information Is Unavailable On This Computer", vbOKOnly

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable

    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
        KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win98 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    
        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value
    
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
            Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:          ' Cleanup After An Error Has Occured...
    
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    
End Function


