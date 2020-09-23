VERSION 5.00
Begin VB.UserControl MultiButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "MultiButton.ctx":0000
   Begin VB.Timer timCheck 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2040
      Top             =   1800
   End
End
Attribute VB_Name = "MultiButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------'
' By Paul Sanders, pa_sanders@hotmail.com '
'--------------------------------------------------------------------------------------------
'            :
' Project    : MultiButtonControl
' Module     : MultiButton
'            :
' Created    : 01-Apr-02 23:10
'            :
' Notes      : This is a complete replacement control for commandbutton/option button/checkbox/label/menu.frame controls.
'            : It exposes various properties to allow it to look and feel much like .Net/XP/Explorer buttons
'            : In this implementation I have chosen to use a simple timer control to
'            : determine when the user has moved off the control.  This gives two advantages
'            : over the usual SetCapture implementation.
'            : 1    It doesn't fail - SetCapture needs re-applying after the user performs
'            :      certain actions ie mousedown in the control, mouseup outside
'            : 2    Tooltips don't display when using SetCapture, with a timer
'            :      there's no problem.
'            : The control is pretty simple.  It took me an afternoon to write the core.  The main
'            : routine is pDraw which handles all the drawing and is called after setting
'            : most properties, and when various events occur.
'            :
' References : None
'            :
'--------------------------------------------------------------------------------------------

Private Const MODULENAME = "MultiButton"

'A few API's
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor& Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long)
Private Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long)
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function FillRect& Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long)
Private Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_EXPANDTABS = &H40
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

'Default Property Values:
Const m_def_OwnerDrawn = False
Const m_def_BorderStyle = 0
Const m_def_CornerRadius = 0
Const m_def_OptionName = ""
Const m_def_ButtonMode = 0
Const m_def_CheckedBorderColor = vbWindowBackground
Const m_def_CheckedFillColor = vbButtonFace
Const m_def_CheckedForeColor = vbWindowText
Const m_def_Value = False
Const m_def_VerticalAlignment = 1
Const m_def_ActiveBorderColor = vbWhite
Const m_def_ActiveForeColor = vbWindowText
Const m_def_ActiveFillColor = vbButtonFace
Const m_def_PictureAlignment = 0
Const m_def_Alignment = vbCenter
Const m_def_HoverFillColor = vbButtonFace
Const m_def_HoverBorderColor = vbWindowText
Const m_def_RedrawOnHover = True
Const m_def_HoverForeColor = vbWindowText
Const m_def_BorderColor = vbButtonShadow
Const m_def_FillColor = vbButtonFace
Const m_def_Caption = "MultiButton"

'Property Variables:
Dim m_OwnerDrawn As Boolean
Dim m_BorderStyle As ButtonBorderStyle
Dim m_CornerRadius As Integer
Dim m_CheckedPicture As StdPicture
Dim m_OptionName As String
Dim m_ButtonMode As ButtonModeConstants
Dim m_CheckedBorderColor As OLE_COLOR
Dim m_CheckedFillColor As OLE_COLOR
Dim m_CheckedForeColor As OLE_COLOR
Dim m_Value As Boolean
Dim m_VerticalAlignment As VerticalAlignmentConstants
Dim m_ActiveBorderColor As OLE_COLOR
Dim m_ActiveForeColor As OLE_COLOR
Dim m_ActiveFillColor As OLE_COLOR
Dim m_PictureAlignment As AlignmentConstants
Dim m_Alignment As AlignmentConstants
Dim m_HoverFillColor As OLE_COLOR
Dim m_HoverBorderColor As OLE_COLOR
Dim m_RedrawOnHover As Boolean
Dim m_HoverForeColor As OLE_COLOR
Dim m_Picture As StdPicture
Dim m_BorderColor As OLE_COLOR
Dim m_FillColor As OLE_COLOR
Dim m_Caption As String
Dim m_ForeColor As OLE_COLOR

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseOut()
Event DrawButton(ByVal hDC As Long, ByVal hwnd As Long, ByVal MouseOver As Boolean, ByVal MouseDown As Boolean, ByVal Value As Boolean)


' DrawState constants - Needed when the control is disabled
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_NORMAL = &H0
Private Const DSS_RIGHT = &H8000
Private Const DSS_UNION = &H10
Private Const DST_BITMAP = &H4
Private Const DST_COMPLEX = &H0
Private Const DST_ICON = &H3
Private Const DST_PREFIXTEXT = &H2
Private Const DST_TEXT = &H1

Private mX As Single
Private mY As Single
Private mbOver As Boolean
Private mbDown As Boolean
Private mbInClick As Boolean

Public Enum VerticalAlignmentConstants
    mbnTop = 0
    mbnCenter = 1
    mbnBottom = 2
End Enum

Public Enum ButtonModeConstants
    mbnButton = 0
    mbnOption = 1
    mbnGroupBox = 2
    mbnMenu = 3
End Enum

Public Enum ButtonBorderStyle
    mbnBox = 0
    mbnTopTab = 1
    mbnBottonTab = 2
End Enum

'--------------------------------------------------------------------------------------------
'Procedure : timCheck_Timer
'Author    : Paul Sanders, pa_sanders@hotmail.com, 01-Apr-02 23:18
'Notes     : The timer is only enabled when the user has moved over the control.  On each
'          : tick it checks to see if the mouse is still over it.  If not the relevent flags
'          : are set and the timer is disabled.
'          : I am using the timer instead of SetCapture() to allow tooltips to be displayed
'          : correctly.
'--------------------------------------------------------------------------------------------
Private Sub timCheck_Timer()
    
    If Not pCursorInWindow Then
        timCheck.Enabled = False
        mbOver = False
        mbDown = False
        pDraw
        RaiseEvent MouseOut
    End If

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    mbInClick = True
    pClick
    mbInClick = False
    timCheck.Enabled = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    pDraw
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    pDraw
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    pDraw
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Private Sub UserControl_DblClick()
    pClick
End Sub

Private Sub UserControl_EnterFocus()
    'If were are emulating a groupbox then pass the focus on
    If m_ButtonMode = mbnGroupBox Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lRet As Long
    
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    'Only redraw if designed to
    If m_RedrawOnHover Then
        mbDown = (Button = vbLeftButton)
        pDraw
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lRet As Long
    
    'This control does nothing if the middle or right button is pressed
    mbDown = (Button = vbLeftButton)
    
    If Not mbOver And m_RedrawOnHover Then
        If pCursorInWindow Then
            mX = x
            mY = y
            mbOver = True
            pDraw
            
            'Start the timer to check the cursor position
            timCheck.Enabled = True
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lRet As Long
    
    If m_RedrawOnHover Then
        mbDown = False
        pDraw
    End If
    
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    
    pDraw
    
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    pDraw
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FillColor.VB_UserMemId = -510
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_FillColor = New_FillColor
    pDraw
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,MultiButton
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    UserControl.AccessKeys = modUserControlHelper.DefineAccessKeys(New_Caption)
    pDraw
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    Set m_Picture = LoadPicture("")
    m_BorderColor = m_def_BorderColor
    m_FillColor = m_def_FillColor
    m_Caption = m_def_Caption
    m_HoverForeColor = m_def_HoverForeColor
    m_RedrawOnHover = m_def_RedrawOnHover
    m_HoverFillColor = m_def_HoverFillColor
    m_HoverBorderColor = m_def_HoverBorderColor
    m_Alignment = m_def_Alignment
    m_PictureAlignment = m_def_PictureAlignment
    m_ActiveBorderColor = m_def_ActiveBorderColor
    m_ActiveForeColor = m_def_ActiveForeColor
    m_ActiveFillColor = m_def_ActiveFillColor
    m_VerticalAlignment = m_def_VerticalAlignment
    m_Value = m_def_Value
    m_CheckedBorderColor = m_def_CheckedBorderColor
    m_CheckedFillColor = m_def_CheckedFillColor
    m_CheckedForeColor = m_def_CheckedForeColor
    m_ButtonMode = m_def_ButtonMode
    m_OptionName = m_def_OptionName
    m_CornerRadius = m_def_CornerRadius
    m_BorderStyle = m_def_BorderStyle
    m_OwnerDrawn = m_def_OwnerDrawn
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_HoverForeColor = PropBag.ReadProperty("HoverForeColor", m_def_HoverForeColor)
    m_RedrawOnHover = PropBag.ReadProperty("RedrawOnHover", m_def_RedrawOnHover)
    m_HoverFillColor = PropBag.ReadProperty("HoverFillColor", m_def_HoverFillColor)
    m_HoverBorderColor = PropBag.ReadProperty("HoverBorderColor", m_def_HoverBorderColor)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_PictureAlignment = PropBag.ReadProperty("PictureAlignment", m_def_PictureAlignment)
    
    UserControl.AccessKeys = modUserControlHelper.DefineAccessKeys(m_Caption)
    m_ActiveBorderColor = PropBag.ReadProperty("ActiveBorderColor", m_def_ActiveBorderColor)
    m_ActiveForeColor = PropBag.ReadProperty("ActiveForeColor", m_def_ActiveForeColor)
    m_ActiveFillColor = PropBag.ReadProperty("ActiveFillColor", m_def_ActiveFillColor)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_VerticalAlignment = PropBag.ReadProperty("VerticalAlignment", m_def_VerticalAlignment)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_CheckedBorderColor = PropBag.ReadProperty("CheckedBorderColor", m_def_CheckedBorderColor)
    m_CheckedFillColor = PropBag.ReadProperty("CheckedFillColor", m_def_CheckedFillColor)
    m_CheckedForeColor = PropBag.ReadProperty("CheckedForeColor", m_def_CheckedForeColor)
    m_ButtonMode = PropBag.ReadProperty("ButtonMode", m_def_ButtonMode)
    m_OptionName = PropBag.ReadProperty("OptionName", m_def_OptionName)
    Set m_CheckedPicture = PropBag.ReadProperty("CheckedPicture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_CornerRadius = PropBag.ReadProperty("CornerRadius", m_def_CornerRadius)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_OwnerDrawn = PropBag.ReadProperty("OwnerDrawn", m_def_OwnerDrawn)
End Sub

Private Sub UserControl_Resize()
    pDraw
End Sub

Private Sub UserControl_Show()
    'Fires when the control is shown in design or run mode
    pDraw
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("HoverForeColor", m_HoverForeColor, m_def_HoverForeColor)
    Call PropBag.WriteProperty("RedrawOnHover", m_RedrawOnHover, m_def_RedrawOnHover)
    Call PropBag.WriteProperty("HoverFillColor", m_HoverFillColor, m_def_HoverFillColor)
    Call PropBag.WriteProperty("HoverBorderColor", m_HoverBorderColor, m_def_HoverBorderColor)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("PictureAlignment", m_PictureAlignment, m_def_PictureAlignment)
    Call PropBag.WriteProperty("ActiveBorderColor", m_ActiveBorderColor, m_def_ActiveBorderColor)
    Call PropBag.WriteProperty("ActiveForeColor", m_ActiveForeColor, m_def_ActiveForeColor)
    Call PropBag.WriteProperty("ActiveFillColor", m_ActiveFillColor, m_def_ActiveFillColor)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("VerticalAlignment", m_VerticalAlignment, m_def_VerticalAlignment)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("CheckedBorderColor", m_CheckedBorderColor, m_def_CheckedBorderColor)
    Call PropBag.WriteProperty("CheckedFillColor", m_CheckedFillColor, m_def_CheckedFillColor)
    Call PropBag.WriteProperty("CheckedForeColor", m_CheckedForeColor, m_def_CheckedForeColor)
    Call PropBag.WriteProperty("ButtonMode", m_ButtonMode, m_def_ButtonMode)
    Call PropBag.WriteProperty("OptionName", m_OptionName, m_def_OptionName)
    Call PropBag.WriteProperty("CheckedPicture", m_CheckedPicture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("CornerRadius", m_CornerRadius, m_def_CornerRadius)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("OwnerDrawn", m_OwnerDrawn, m_def_OwnerDrawn)
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : pDraw
'Author    : Paul Sanders, pa_sanders@hotmail.com, 31-Mar-02 22:03
'Notes     : Draws the control based on its current state
'--------------------------------------------------------------------------------------------
Private Sub pDraw()
    Dim lX As Long
    Dim lY As Long
    Dim lW As Long
    Dim lH As Long
    
    Dim bGotPic As Boolean
    Dim lFlags As Long
    Dim R As RECT
    Dim lRet As Long
    Dim pic As StdPicture
    Dim lBorderOffset As Long
    Dim PT As POINTAPI
    Dim lTextX As Long
    
    Const OFFSET = 0
    
    'Don't bother doing anything if the control is not visible
'    If IsWindowVisible(UserControl.hwnd) = 0 Then Exit Sub '---------------------------->-->-->
    
    If m_OwnerDrawn Then
        Cls
        'If owner drawn then tell the user that the button needs redrawing
        RaiseEvent DrawButton(UserControl.hDC, UserControl.hwnd, mbOver, mbDown, m_Value)
    Else
        'First calculate as much as possible without doing anything with the control
        lW = ScaleWidth
        lH = ScaleHeight - 1
        
        lX = 1
        lY = 0
    
        'Adjust if we are emulating a frame
        If m_ButtonMode = mbnGroupBox Then lBorderOffset = (TextHeight("Xy") / 2)
        
        'Get the right picture to use
        If m_ButtonMode = mbnButton Or (m_ButtonMode = mbnOption And m_Value = False) Or m_ButtonMode = mbnMenu Then
            Set pic = m_Picture
        Else
            Set pic = m_CheckedPicture
            If pic Is Nothing Then
                Set pic = m_Picture
            ElseIf pic.Handle = 0 Then
                Set pic = m_Picture
            End If
        End If
        
        'Adjust if we have a picture to display
        If Not pic Is Nothing Then
            If pic.Handle <> 0 Then
                Select Case m_PictureAlignment
                    Case vbLeftJustify
                        lX = lX + ScaleX(pic.Width, vbHimetric, vbPixels) + 4
                        lW = lW - ScaleX(pic.Width, vbHimetric, vbPixels) - 4
                    
                    Case vbRightJustify
                        lW = lW - ScaleX(pic.Width, vbHimetric, vbPixels) - 4
                End Select
                bGotPic = True
            End If
        End If
        
        'If we have a caption, workout where it needs placing
        If Len(m_Caption) > 0 Then
            Select Case m_Alignment
                Case vbLeftJustify
                    lFlags = DT_LEFT
                    lW = lW - 2
                    If m_ButtonMode = mbnGroupBox Then lX = 6
                Case vbRightJustify
                    lFlags = DT_RIGHT
                    lW = lW - 1
                    If m_ButtonMode = mbnGroupBox Then lW = lW - 5
                Case Else
                    lFlags = DT_CENTER
            End Select
            
            Select Case m_VerticalAlignment
                Case mbnTop
                    lFlags = lFlags Or DT_TOP
                Case mbnCenter
                    lFlags = lFlags Or DT_VCENTER
                Case mbnBottom
                    lFlags = lFlags Or DT_BOTTOM
            End Select
            
            'Calculate position of caption
            If InStr(1, m_Caption, vbCr) > 0 Or TextWidth(m_Caption) > (ScaleWidth - lX - 2) Then
                lFlags = lFlags Or DT_WORDBREAK
                
                SetTextColor UserControl.hDC, 0
                If mbDown Then
                    UserControl.ForeColor = m_ActiveForeColor
                ElseIf mbOver Then
                    UserControl.ForeColor = m_HoverForeColor
                ElseIf m_Value And m_ButtonMode = mbnOption Then
                    UserControl.ForeColor = m_CheckedForeColor
                Else
                    UserControl.ForeColor = m_ForeColor
                End If
                
                'Evaluate height of text
                If m_VerticalAlignment = mbnCenter Then
                    lRet = GetClientRect(UserControl.hwnd, R)
                    R.Left = lX
                    DrawText UserControl.hDC, m_Caption, Len(m_Caption), R, lFlags Or DT_CALCRECT
                    lY = (lH - R.Bottom) / 2
                    lH = R.Bottom
                ElseIf m_VerticalAlignment = mbnBottom Then
                    lRet = GetClientRect(UserControl.hwnd, R)
                    R.Left = lX
                    DrawText UserControl.hDC, m_Caption, Len(m_Caption), R, lFlags Or DT_CALCRECT
                    lY = lH - R.Bottom
                    lH = R.Bottom
                End If
            Else
                lFlags = lFlags Or DT_SINGLELINE
            End If
        End If
    
        Select Case m_Alignment
            Case vbLeftJustify
                lTextX = lX
            Case vbCenter
                lTextX = ((ScaleWidth - TextWidth(m_Caption)) / 2)
            Case vbRightJustify
                lTextX = ScaleWidth - TextWidth(m_Caption) - 5
        End Select
        
        'And were off...
        Cls
        
        If mbOver Or (m_ButtonMode = mbnMenu And mbInClick) Then
            If mbDown Or (m_ButtonMode = mbnMenu And mbInClick) Then
                'Mouse is down
                If Len(m_Caption) > 0 Then
                    lY = lY + 1
                    lX = lX + 1
                End If
                
                If m_BorderStyle = mbnBox Then
                    'If were are a group box/frame then offset to center border through text
                    If m_ButtonMode = mbnGroupBox Then
                        If m_VerticalAlignment = mbnTop Then
                            BoxDC UserControl.hDC, 0, lBorderOffset, ScaleWidth, ScaleHeight - lBorderOffset, m_ActiveBorderColor, m_ActiveFillColor
                        Else
                            BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight - lBorderOffset, m_ActiveBorderColor, m_ActiveFillColor
                        End If
                    Else
                        BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_ActiveBorderColor, m_ActiveFillColor
                    End If
                Else
                    DrawTab UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_BorderStyle = mbnTopTab, m_Value, m_ActiveBorderColor, m_ActiveFillColor
                End If
                
                If Len(m_Caption) > 0 Then
                    'Draw a line 2 pixels longer than the textwidth along the border using the backcolor
                    'so we don't have a line through the text
                    If m_ButtonMode = mbnGroupBox Then
                        If m_VerticalAlignment = mbnTop Then
                            LineDC UserControl.hDC, lTextX, lBorderOffset, lTextX + TextWidth(m_Caption) + 2, lBorderOffset, UserControl.BackColor
                        Else
                            LineDC UserControl.hDC, lTextX - 1, ScaleHeight - lBorderOffset - 1, lTextX + TextWidth(m_Caption) + 2, ScaleHeight - lBorderOffset - 1, UserControl.BackColor
                        End If
                    End If
                    
                    'Now draw the text
                    UserControl.ForeColor = m_ActiveForeColor
                    PaintText UserControl.hDC, m_Caption, lX, lY, lW, lH, lFlags
                End If
            Else
                'Mouse is over control
                If m_BorderStyle = mbnBox Then
                    If m_ButtonMode = mbnGroupBox Then
                        If m_VerticalAlignment = mbnTop Then
                            BoxDC UserControl.hDC, 0, lBorderOffset, ScaleWidth, ScaleHeight - lBorderOffset, m_HoverBorderColor, m_HoverFillColor
                        Else
                            BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight - lBorderOffset, m_HoverBorderColor, m_HoverFillColor
                        End If
                    Else
                        BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_HoverBorderColor, m_HoverFillColor
                    End If
                Else
                    DrawTab UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_BorderStyle = mbnTopTab, m_Value, m_HoverBorderColor, m_HoverFillColor, IIf(m_Value, -1, m_BorderColor)
                End If
        
                If Len(m_Caption) > 0 Then
                    'Draw a line 2 pixels longer than the textwidth along the border using the backcolor
                    'so we don't have a line through the text
                    If m_ButtonMode = mbnGroupBox Then
                        If m_VerticalAlignment = mbnTop Then
                            LineDC UserControl.hDC, lTextX - 1, lBorderOffset, lTextX + TextWidth(m_Caption) + 2, lBorderOffset, UserControl.BackColor
                        Else
                            LineDC UserControl.hDC, lTextX - 1, ScaleHeight - lBorderOffset - 1, lTextX + TextWidth(m_Caption) + 2, ScaleHeight - lBorderOffset - 1, UserControl.BackColor
                        End If
                    End If
                    UserControl.ForeColor = m_HoverForeColor
                    PaintText UserControl.hDC, m_Caption, lX, lY, lW, lH, lFlags
                End If
            End If
        ElseIf m_Value And m_ButtonMode = mbnOption Then
            'If were are an optionbutton/checkbox and are set draw the checked version
            If m_BorderStyle = mbnBox Then
                BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_CheckedBorderColor, m_CheckedFillColor
            Else
                DrawTab UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_BorderStyle = mbnTopTab, m_Value, m_CheckedBorderColor, m_CheckedFillColor
            End If
                    
            If Len(m_Caption) > 0 Then
                'Draw a line 2 pixels longer than the textwidth along the border using the backcolor
                'so we don't have a line through the text
                If m_ButtonMode = mbnGroupBox Then
                    If m_VerticalAlignment = mbnTop Then
                        LineDC UserControl.hDC, lTextX - 1, lBorderOffset, lTextX + TextWidth(m_Caption) + 2, lBorderOffset, UserControl.BackColor
                    Else
                        LineDC UserControl.hDC, lTextX - 1, ScaleHeight - lBorderOffset - 1, lTextX + TextWidth(m_Caption) + 2, ScaleHeight - lBorderOffset - 1, UserControl.BackColor
                    End If
                End If
                UserControl.ForeColor = m_CheckedForeColor
                PaintText UserControl.hDC, m_Caption, lX, lY, lW, lH, lFlags
            End If
        Else
            'Mouse is not over and the control is not selected
            If m_BorderStyle = mbnBox Then
                If m_ButtonMode = mbnGroupBox Then
                    If m_VerticalAlignment = mbnTop Then
                        BoxDC UserControl.hDC, 0, lBorderOffset, ScaleWidth, ScaleHeight - lBorderOffset, m_BorderColor, m_FillColor
                    Else
                        BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight - lBorderOffset, m_BorderColor, m_FillColor
                    End If
                Else
                    BoxDC UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_BorderColor, m_FillColor
                End If
            Else
                DrawTab UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, m_BorderStyle = mbnTopTab, m_Value, m_BorderColor, m_FillColor
            End If
                    
            If Len(m_Caption) > 0 Then
                'Draw a line 2 pixels longer than the textwidth along the border using the backcolor
                'so we don't have a line through the text
                If m_ButtonMode = mbnGroupBox Then
                    If m_VerticalAlignment = mbnTop Then
                        LineDC UserControl.hDC, lTextX - 1, lBorderOffset, lTextX + TextWidth(m_Caption) + 2, lBorderOffset, UserControl.BackColor
                    Else
                        LineDC UserControl.hDC, lTextX - 1, ScaleHeight - lBorderOffset - 1, lTextX + TextWidth(m_Caption) + 2, ScaleHeight - lBorderOffset - 1, UserControl.BackColor
                    End If
                End If
                UserControl.ForeColor = m_ForeColor
                PaintText UserControl.hDC, m_Caption, lX, lY, lW, lH, lFlags
            End If
        End If
        
        'Paint the picture - this will also update the window
        If bGotPic Then
            PaintPic pic
        Else
            UpdateWindow UserControl.hwnd
        End If
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwindowtext
Public Property Get HoverForeColor() As OLE_COLOR
    HoverForeColor = m_HoverForeColor
End Property

Public Property Let HoverForeColor(ByVal New_HoverForeColor As OLE_COLOR)
    m_HoverForeColor = New_HoverForeColor
    PropertyChanged "HoverForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get RedrawOnHover() As Boolean
Attribute RedrawOnHover.VB_Description = "When True the control will be redrawn when the mouse is moved over it."
    RedrawOnHover = m_RedrawOnHover
End Property

Public Property Let RedrawOnHover(ByVal New_RedrawOnHover As Boolean)
    m_RedrawOnHover = New_RedrawOnHover
    PropertyChanged "RedrawOnHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbbuttonface
Public Property Get HoverFillColor() As OLE_COLOR
Attribute HoverFillColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoverFillColor = m_HoverFillColor
End Property

Public Property Let HoverFillColor(ByVal New_HoverFillColor As OLE_COLOR)
    m_HoverFillColor = New_HoverFillColor
    PropertyChanged "HoverFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwindowtext
Public Property Get HoverBorderColor() As OLE_COLOR
    HoverBorderColor = m_HoverBorderColor
End Property

Public Property Let HoverBorderColor(ByVal New_HoverBorderColor As OLE_COLOR)
    m_HoverBorderColor = New_HoverBorderColor
    PropertyChanged "HoverBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    pDraw
    PropertyChanged "Alignment"
End Property

'--------------------------------------------------------------------------------------------
'Procedure : PaintPic
'Author    : Paul Sanders, pa_sanders@hotmail.com, 03-Apr-02 00:11
'Notes     : Paints a regular or disable picture depending upon the UserControl.Enabled
'--------------------------------------------------------------------------------------------
Private Sub PaintPic(pic As StdPicture)
    Dim lW As Long
    Dim lH As Long
    Dim lX As Long
    Dim lY As Long
    
    Const DI_NORMAL = &H3
    
    'Calculate the position
    lW = ScaleX(pic.Width, vbHimetric, vbPixels)
    lH = ScaleY(pic.Height, vbHimetric, vbPixels)
    
    Select Case m_PictureAlignment
        Case vbLeftJustify
            lX = 2
        Case vbRightJustify
            lX = ScaleWidth - 1 - lW
        Case vbCenter
            lX = ((ScaleWidth - lW) / 2) + 1
    End Select
    
    lY = 1 + Int((ScaleHeight - lH) / 2)
    
    If mbDown And Len(m_Caption) > 0 Then
        lX = lX + 1
        lY = lY + 1
    End If
    
    'Now draw the pic
    If UserControl.Enabled Then
        UserControl.PaintPicture pic, lX, lY
    Else
        DrawState UserControl.hDC, 0, 0, pic.Handle, 0, lX, lY, lW, lH, (DST_ICON Or DSS_DISABLED)
    End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=22,0,0,0
Public Property Get PictureAlignment() As AlignmentConstants
    PictureAlignment = m_PictureAlignment
End Property

Public Property Let PictureAlignment(ByVal New_PictureAlignment As AlignmentConstants)
    m_PictureAlignment = New_PictureAlignment
    pDraw
    PropertyChanged "PictureAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWhite
Public Property Get ActiveBorderColor() As OLE_COLOR
    ActiveBorderColor = m_ActiveBorderColor
End Property

Public Property Let ActiveBorderColor(ByVal New_ActiveBorderColor As OLE_COLOR)
    m_ActiveBorderColor = New_ActiveBorderColor
    PropertyChanged "ActiveBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwindowtext
Public Property Get ActiveForeColor() As OLE_COLOR
    ActiveForeColor = m_ActiveForeColor
End Property

Public Property Let ActiveForeColor(ByVal New_ActiveForeColor As OLE_COLOR)
    m_ActiveForeColor = New_ActiveForeColor
    PropertyChanged "ActiveForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbButtonface
Public Property Get ActiveFillColor() As OLE_COLOR
    ActiveFillColor = m_ActiveFillColor
End Property

Public Property Let ActiveFillColor(ByVal New_ActiveFillColor As OLE_COLOR)
    m_ActiveFillColor = New_ActiveFillColor
    PropertyChanged "ActiveFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get VerticalAlignment() As VerticalAlignmentConstants
Attribute VerticalAlignment.VB_Description = "Sets/returns the vertical alignment of the caption."
    VerticalAlignment = m_VerticalAlignment
End Property

Public Property Let VerticalAlignment(ByVal New_VerticalAlignment As VerticalAlignmentConstants)
    If m_ButtonMode = mbnGroupBox And New_VerticalAlignment = mbnCenter Then
        New_VerticalAlignment = mbnTop
    End If
    m_VerticalAlignment = New_VerticalAlignment
    pDraw
    PropertyChanged "VerticalAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Sets/returns the checked value of the control."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    If m_Value Then
        pUncheckControls
    End If
    pDraw
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWindowBackground
Public Property Get CheckedBorderColor() As OLE_COLOR
    CheckedBorderColor = m_CheckedBorderColor
End Property

Public Property Let CheckedBorderColor(ByVal New_CheckedBorderColor As OLE_COLOR)
    m_CheckedBorderColor = New_CheckedBorderColor
    pDraw
    PropertyChanged "CheckedBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbbuttonface
Public Property Get CheckedFillColor() As OLE_COLOR
Attribute CheckedFillColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CheckedFillColor = m_CheckedFillColor
End Property

Public Property Let CheckedFillColor(ByVal New_CheckedFillColor As OLE_COLOR)
    m_CheckedFillColor = New_CheckedFillColor
    pDraw
    PropertyChanged "CheckedFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwindowtext
Public Property Get CheckedForeColor() As OLE_COLOR
    CheckedForeColor = m_CheckedForeColor
End Property

Public Property Let CheckedForeColor(ByVal New_CheckedForeColor As OLE_COLOR)
    m_CheckedForeColor = New_CheckedForeColor
    pDraw
    PropertyChanged "CheckedForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ButtonMode() As ButtonModeConstants
Attribute ButtonMode.VB_Description = "Sets/returns the operating mode of the MultiButton."
    ButtonMode = m_ButtonMode
End Property

Public Property Let ButtonMode(ByVal New_ButtonMode As ButtonModeConstants)
    If New_ButtonMode = mbnGroupBox Then
        If m_VerticalAlignment = mbnCenter Then
            m_VerticalAlignment = mbnTop
            PropertyChanged "VerticalAlignment"
        End If
        m_RedrawOnHover = False
        PropertyChanged "RedrawOnHover"
    End If
    
    m_ButtonMode = New_ButtonMode
    m_Value = False
    
    pDraw
    PropertyChanged "ButtonMode"
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get OptionName() As String
Attribute OptionName.VB_Description = "Sets/returns the option name.  Used when button is in option mode.  Setting all related MultiButtons to the same option name will allow default switching."
    OptionName = m_OptionName
End Property

Public Property Let OptionName(ByVal New_OptionName As String)
    m_OptionName = New_OptionName
    PropertyChanged "OptionName"
End Property

'--------------------------------------------------------------------------------------------
'Procedure : pUncheckControls
'Author    : Paul Sanders, pa_sanders@hotmail.com, 03-Apr-02 00:03
'Notes     : This trawls through other controls with this controls parent and resets their
'          : value if it is a MultiButton control with the same OptionName
'--------------------------------------------------------------------------------------------
Private Sub pUncheckControls()
    Dim oParent As Object
    Dim oPanel As Control
    Dim i As Integer
    Dim nCount As Integer
    
    On Error Resume Next

    If Len(m_OptionName) > 0 Then
        nCount = ParentControls.Count - 1
        
        For i = 0 To nCount
            If TypeName(ParentControls(i)) = "MultiButton" Then
                Set oPanel = Nothing
                Set oPanel = ParentControls(i)
                If Not oPanel Is Nothing Then
                    If oPanel.hwnd <> UserControl.hwnd Then
                        If oPanel.OptionName = m_OptionName Then
                            If oPanel.Value = True Then
                                oPanel.Value = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=22,0,0,0
Public Property Get CheckedPicture() As StdPicture
    Set CheckedPicture = m_CheckedPicture
End Property

Public Property Set CheckedPicture(ByVal New_CheckedPicture As StdPicture)
    Set m_CheckedPicture = New_CheckedPicture
    pDraw
    PropertyChanged "CheckedPicture"
End Property

'--------------------------------------------------------------------------------------------
'Procedure : PaintText
'Author    : Ariad Software
'Notes     : Extracted from basGDI of Ariad Software's excellent toolbar v1.0
'--------------------------------------------------------------------------------------------
Private Sub PaintText(ByVal hDC As Long, ByVal Text As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, Optional ByVal Flags As Long = DT_LEFT)
    Dim R As RECT

    If UserControl.Enabled = False Then UserControl.ForeColor = vbGrayText
    With R
        .Left = x
        .Top = y
        .Right = x + w
        .Bottom = y + h
    End With
    
    DrawText hDC, Text, -1, R, Flags
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : BoxDC
'Author    : Ariad Software
'Notes     : Extracted from basGDI of Ariad Software's excellent toolbar v1.0
'          : and modified to allow for rounded corners
'--------------------------------------------------------------------------------------------
Private Sub BoxDC(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional Color As OLE_COLOR = vbButtonFace, Optional Fill As OLE_COLOR = -1)
    Dim hPen As Long, hPenOld As Long
    Dim PT As POINTAPI
    Dim hBrush As Long
    Dim lOldBr As Long
    
    'Fill
    If Fill <> -1 Then BoxSolidDC hDC, x, y, w, h, Fill
    
    'Box
    hPen = CreatePen(0, 1, TranslateColor(Color))
    hPenOld = SelectObject(hDC, hPen)
    
    If m_CornerRadius = 0 Then
        'Draw a standard box
        MoveToEx hDC, x + w - 1, y, PT
        LineTo hDC, x, y
        LineTo hDC, x, y + h - 1
        LineTo hDC, x + w - 1, y + h - 1
        LineTo hDC, x + w - 1, y
    Else
        'Draw a box with rounded corners
        RoundRect hDC, x, y, x + w, y + h, m_CornerRadius, m_CornerRadius
        If Fill <> -1 Then
            hBrush = CreateSolidBrush(TranslateColor(Fill))
            lOldBr = SelectObject(hDC, hBrush)
            FloodFill hDC, w / 2, h / 2, TranslateColor(Color)
            SelectObject hDC, lOldBr
            DeleteObject hBrush
        End If
    End If
    
    'Clean up
    SelectObject hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : LineDC
'Author    : Paul Sanders, pa_sanders@hotmail.com, 21-Jun-02 18:24
'Notes     :
'--------------------------------------------------------------------------------------------
Private Sub LineDC(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Color As OLE_COLOR)
    Dim hPen As Long, hPenOld As Long
    Dim PT As POINTAPI
        
    'Box
    hPen = CreatePen(0, 1, TranslateColor(Color))
    hPenOld = SelectObject(hDC, hPen)
    
    MoveToEx hDC, x1, y1, PT
    LineTo hDC, x2, y2
    
    SelectObject hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : TranslateColor
'Author    : Paul Sanders, pa_sanders@hotmail.com, 03-Apr-02 00:08
'Notes     :
'--------------------------------------------------------------------------------------------
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function

'--------------------------------------------------------------------------------------------
'Procedure : BoxSolidDC
'Author    : Ariad Software
'Notes     : Extracted from basGDI of Ariad Software's excellent toolbar v1.0
'--------------------------------------------------------------------------------------------
Private Function BoxSolidDC(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional ByVal Fill As OLE_COLOR = vbButtonFace)
    Dim hBrush As Long
    Dim R As RECT
    
    If m_CornerRadius = 0 Then
        hBrush = CreateSolidBrush(TranslateColor(Fill))
        
        With R
            .Left = x
            .Top = y
            .Right = x + w - 1
            .Bottom = y + h - 1
        End With
        FillRect hDC, R, hBrush
        DeleteObject hBrush
    End If
End Function

'--------------------------------------------------------------------------------------------
'Procedure : pClick
'Author    : Paul Sanders, pa_sanders@hotmail.com, 03-Apr-02 00:09
'Notes     : Sets relevent options and redraws the control whenever the control is
'          : clicked or double clicked
'--------------------------------------------------------------------------------------------
Private Sub pClick()
    mbOver = False
    
    If m_ButtonMode = mbnOption And Len(m_OptionName) > 0 Then
        m_Value = True
    Else
        m_Value = Not m_Value
    End If
    
    pDraw
    
    'If were are emulating an option button/check box, make sure any
    'related MultiButtons are reset
    If m_ButtonMode = mbnOption Then
        pUncheckControls
    End If
    
    RaiseEvent Click
End Sub

'--------------------------------------------------------------------------------------------
'Procedure : pCursorInWindow
'Author    : Paul Sanders, pa_sanders@hotmail.com, 23-Jun-02 17:49
'Notes     : Returns True if the cursor is over the control
'--------------------------------------------------------------------------------------------
Private Function pCursorInWindow() As Boolean
    Dim R As RECT
    Dim PT As POINTAPI
    Dim lRet As Long
    
    lRet = GetClientRect(UserControl.hwnd, R)
    lRet = GetCursorPos(PT)
    lRet = ScreenToClient(UserControl.hwnd, PT)
    
    'PtInRect is a bit flaky, so go for the manual method
    pCursorInWindow = Not (PT.x < 0 Or PT.x > ScaleWidth Or PT.y < 0 Or PT.y > ScaleHeight)
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    pDraw
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CornerRadius() As Integer
Attribute CornerRadius.VB_Description = "Sets/returns the radius to be used to draw each corner.  Set to 0 to have square corners."
    CornerRadius = m_CornerRadius
End Property

Public Property Let CornerRadius(ByVal New_CornerRadius As Integer)
    If New_CornerRadius <> m_CornerRadius Then
        m_CornerRadius = New_CornerRadius
        pDraw
        PropertyChanged "CornerRadius"
    End If
End Property

'--------------------------------------------------------------------------------------------
'Procedure : DrawTab
'Author    : Paul Sanders, pa_sanders@hotmail.com, 31-May-02 20:41
'Notes     : Makes the button display like a tab
'--------------------------------------------------------------------------------------------
Private Sub DrawTab(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
                    ByVal w As Long, ByVal h As Long, _
                    ByVal TopTab As Boolean, ByVal Selected As Boolean, _
                    Optional Color As OLE_COLOR = vbButtonFace, Optional Fill As OLE_COLOR = -1, _
                    Optional AltBorderColor As OLE_COLOR = -1)
    Dim hPen As Long, hPenOld As Long
    Dim PT As POINTAPI
    Dim hBrush As Long
    Dim lOldBr As Long
    Dim lOffset As Long
    
    'Box
    hPen = CreatePen(0, 1, TranslateColor(Color))
    hPenOld = SelectObject(hDC, hPen)
    
    If TopTab Then
        'Draw tab so it looks top aligned
        MoveToEx hDC, x, y + h - 1, PT
        LineTo hDC, x + m_CornerRadius, y
        LineTo hDC, x + w - 1 - m_CornerRadius, y
        LineTo hDC, x + w - 1, y + h - 1
        LineTo hDC, x, y + h - 1
    Else
        'Draw tab so it looks bottom aligned
        MoveToEx hDC, x, y, PT
        LineTo hDC, x + m_CornerRadius, y + h - 1
        LineTo hDC, x + w - 1 - m_CornerRadius, y + h - 1
        LineTo hDC, x + w - 1, y
        LineTo hDC, x, y
    End If
    
    'Clean up
    SelectObject hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld

    If Fill <> -1 Then
        hBrush = CreateSolidBrush(TranslateColor(Fill))
        lOldBr = SelectObject(hDC, hBrush)
        FloodFill hDC, w / 2, h / 2, TranslateColor(Color)
        SelectObject hDC, lOldBr
        DeleteObject hBrush
    End If
    
    If Selected Or AltBorderColor <> -1 Then
        If AltBorderColor = -1 Then
            hPen = CreatePen(0, 2, TranslateColor(Fill))
        Else
            hPen = CreatePen(0, 2, TranslateColor(AltBorderColor))
        End If
        hPenOld = SelectObject(hDC, hPen)
    
        If Selected Then
            lOffset = 1
        Else
            lOffset = 2
        End If
        
        If TopTab Then
            MoveToEx hDC, x - lOffset, y + h, PT
            LineTo hDC, x + w, y + h
        Else
            MoveToEx hDC, x - lOffset, y, PT
            LineTo hDC, x + w, y
        End If
    
        SelectObject hDC, hPenOld
        DeleteObject hPen
        DeleteObject hPenOld
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,0
Public Property Get BorderStyle() As ButtonBorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As ButtonBorderStyle)
    m_BorderStyle = New_BorderStyle
    pDraw
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get OwnerDrawn() As Boolean
Attribute OwnerDrawn.VB_Description = "Sets/returns if the control should be drawn by owner"
    OwnerDrawn = m_OwnerDrawn
End Property

Public Property Let OwnerDrawn(ByVal New_OwnerDrawn As Boolean)
    m_OwnerDrawn = New_OwnerDrawn
    pDraw
    PropertyChanged "OwnerDrawn"
End Property

'--------------------------------------------------------------------------------------------
'Procedure : Refresh
'Author    : Paul Sanders, pa_sanders@hotmail.com, 23-Jun-02 19:05
'Notes     : Redraws the control
'--------------------------------------------------------------------------------------------
Public Sub Refresh()
    pDraw
End Sub
