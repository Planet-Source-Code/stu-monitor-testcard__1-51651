VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3270
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2257.012
   ScaleMode       =   0  'User
   ScaleWidth      =   6409.028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   1
      Left            =   240
      Top             =   1320
   End
   Begin VB.PictureBox picText 
      Align           =   1  'Align Top
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   6825
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6825
   End
   Begin Testcard.MultiButton cmdOK 
      Default         =   -1  'True
      Height          =   319
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "OK"
      HoverFillColor  =   16744576
      ActiveFillColor =   16761024
      BackColor       =   12648447
      CornerRadius    =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5324.423
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Serial No:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please read the Legal section of the Help file."
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2820
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1656.523
      Y2              =   1656.523
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Warning: This program is copyright of SJS Television Services (Kent) Ltd."
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrAllText As String
Private mblnStart As Boolean

Private Sub cmdOK_Click()

    frmMain.Enabled = True

    Unload Me
  
End Sub

Private Sub Form_Load()

    frmMain.Enabled = False

    Call SetText
    
    lblTitle.Caption = "Serial No:   091203 " & "---" & App.Major & "." & _
        App.Minor & App.Revision

    Me.Caption = "About " & App.Title
    
End Sub

Private Sub timText_Timer()
    'animated text
    
    Const Wait = 25 'wait before drawing the next frame
    
    Dim udtFont As FontStruc
    Dim udtBmp As BitmapStruc
    Dim udtMask As BitmapStruc
    Dim udtBmpSize As RECT
    Dim intResult As Integer
    Dim intTextHeight As Integer
    Dim lngStartingTick As Long
    
    Static udtSurphase As BitmapStruc
    Static intScroll As Integer

    'find out how much time it takes to draw a frame
    lngStartingTick = GetTickCount
    
    'set the bitmap dimensions and create them
    udtBmpSize.Right = picText.ScaleWidth
    udtBmpSize.Bottom = picText.ScaleHeight
    
    Call RectToPixels(udtBmpSize)
    
    udtMask.Area = udtBmpSize
    udtSurphase.Area = udtBmpSize
    udtBmp.Area = udtBmpSize
    
    'set font variables
    udtFont.Alignment = vbCentreAlign
    udtFont.Name = picText.FontName
    udtFont.Bold = picText.FontBold
    udtFont.Colour = vbWhite 'picText.ForeColor
    udtFont.Italic = picText.FontItalic
    udtFont.StrikeThru = picText.FontStrikethru
    udtFont.PointSize = picText.FontSize
    udtFont.Underline = picText.FontUnderline
    
    intTextHeight = GetTextHeight(picText.hDC) * LineCount(mstrAllText)
    
    intScroll = intScroll - Screen.TwipsPerPixelY
    If (intScroll < -(intTextHeight * Screen.TwipsPerPixelY)) _
        Or (Not mblnStart) Then
    intScroll = picText.ScaleHeight
    mblnStart = True
End If
    
'only create the surphase if necessary
If udtSurphase.hDcMemory = 0 Then
    Call CreateNewBitmap(udtSurphase.hDcMemory, _
        udtSurphase.hDcBitmap, _
        udtSurphase.hDcPointer, _
        udtSurphase.Area, _
        frmAbout.hDC, _
        picText.ForeColor, _
        InPixels)
        
    'create the surphase
    'text fade in
    Call Gradient(udtSurphase.hDcMemory, _
        picText.ForeColor, _
        picText.FillColor, _
        0, _
        (udtSurphase.Area.Bottom - ((intTextHeight / LineCount(mstrAllText)) * 2)), _
        udtSurphase.Area.Right, _
        (intTextHeight / LineCount(mstrAllText) * 2), _
        GradHorizontal, InPixels)
    'text fade out
    Call Gradient(udtSurphase.hDcMemory, _
        picText.FillColor, _
        picText.ForeColor, _
        0, _
        0, _
        udtSurphase.Area.Right, _
        (intTextHeight / LineCount(mstrAllText)) * 2, _
        GradHorizontal, _
        InPixels)
End If
    
    Call CreateNewBitmap(udtMask.hDcMemory, _
        udtMask.hDcBitmap, _
        udtMask.hDcPointer, _
        udtMask.Area, _
        frmAbout.hDC, _
        vbBlue, _
        InPixels)
    Call CreateNewBitmap(udtBmp.hDcMemory, _
        udtBmp.hDcBitmap, _
        udtBmp.hDcPointer, _
        udtBmp.Area, _
        frmAbout.hDC, _
        vbWhite, _
        InPixels)
    
    'draw the text onto the mask in black
    Call MakeText(udtMask.hDcMemory, _
        mstrAllText, _
        (intScroll / Screen.TwipsPerPixelY), _
        0, _
        intTextHeight, _
        udtBmp.Area.Right, _
        udtFont, _
        InPixels)
    
    'copy the surphase onto the background
    intResult = BitBlt(udtBmp.hDcMemory, _
        0, _
        0, _
        udtBmp.Area.Right, _
        udtBmp.Area.Bottom, _
        udtSurphase.hDcMemory, _
        0, _
        0, _
        SRCCOPY)
    
    'place the mask onto the background
    intResult = BitBlt(udtBmp.hDcMemory, _
        0, _
        0, _
        udtBmp.Area.Right, _
        udtBmp.Area.Bottom, _
        udtMask.hDcMemory, _
        0, _
        0, _
        SRCAND)
    
    'copy the result to the screen
    intResult = BitBlt(frmAbout.hDC, _
        0, _
        0, _
        udtBmp.Area.Right, _
        udtBmp.Area.Bottom, _
        udtBmp.hDcMemory, _
        0, _
        0, _
        SRCCOPY)
    
    'remove the bitmaps created
    Call DeleteBitmap(udtBmp.hDcMemory, _
        udtBmp.hDcBitmap, _
        udtBmp.hDcPointer)
    Call DeleteBitmap(udtMask.hDcMemory, _
        udtMask.hDcBitmap, _
        udtMask.hDcPointer)
    
    'wait X ticks minus the time it took to draw the frame
    Call Pause(Wait - (GetTickCount - lngStartingTick))
End Sub

Private Sub SetText()
   
    mstrAllText = App.ProductName & vbCrLf & _
        "Version " & App.Major & "." & _
        App.Minor & "." & _
        App.Revision & vbCrLf & _
        "" & vbCrLf & _
        "This program was constructed by " & vbCrLf & _
        "Stu Tyler" & vbCrLf & vbCrLf & _
        "Additional material provided by" & vbCrLf & _
        "Greg Holdys" & vbCrLf & _
        "Paul Sanders" & vbCrLf & _
        "Vic Richardson" & vbCrLf & _
         vbCrLf & vbCrLf & _
        "Copyright 2004" & vbCrLf & _
        "All rights reserved" & vbCrLf & _
        "" & vbCrLf & _
        "enquires@sjstv.co.uk" & vbCrLf & _
        "" & vbCrLf & _
        "WWW.SJSTV.CO.UK"
End Sub

Public Function LineCount(ByVal strText As String) _
        As Integer
    'This function will return the number of lines
    'in the strText
    
    Dim intTemp As Integer
    Dim intCounter As Integer
    Dim intLastPos As Integer

    intLastPos = 1
    
    Do
        intTemp = intLastPos
        intLastPos = InStr(intLastPos + Len(vbCrLf), strText, vbCrLf)
        
        If intTemp <> intLastPos Then
            'a line was found
            intCounter = intCounter + 1
        End If
    Loop Until intLastPos = 0

    LineCount = intCounter
End Function

