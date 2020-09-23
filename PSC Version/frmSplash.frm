VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00AF0511&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picLoad 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   -75
      ScaleHeight     =   330
      ScaleMode       =   0  'User
      ScaleWidth      =   6510.857
      TabIndex        =   0
      Top             =   2580
      Width           =   6480
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ".........................................................................................................................LOADING"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   0
         Width           =   6240
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

    Dim Msg As String
    Dim Style As String
    Dim Title As String
    Dim response As String

    'if Monitor Testcard is already running then quit
    If App.PrevInstance = True Then

        Style = vbCritical
        Title = "Critical Error"
        Msg = "Testcard is already running !"
        response = MsgBox(Msg, Style, Title)
        
          Beep
        End
    End If

    GradeSplash picLoad, , , 500   'sets timer from basBlueGrad

    Unload Me
      frmMain.Show
      
 Call ShowHelpTopic(Hlp_General)
     
        
End Sub

