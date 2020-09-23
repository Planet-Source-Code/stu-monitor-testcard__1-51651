Attribute VB_Name = "basAbout"

Option Explicit

Option Private Module

'--------------------------------------------------------------------------
'API calls
'--------------------------------------------------------------------------


'This is used to copy bitmaps
Public Declare Function BitBlt _
    Lib "gdi32" _
    (ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) _
    As Long


'creates a brush object which can be applied to a bitmap
Public Declare Function CreateBrushIndirect _
    Lib "gdi32" _
    (lpLogBrush As LogBrush) _
    As Long


'the will create a bitmap compatable with the passed hDc
Public Declare Function CreateCompatibleBitmap _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long) _
    As Long

'this create a compatable device context with the specified
'windows handle
Public Declare Function CreateCompatibleDC _
    Lib "gdi32" _
    (ByVal hDC As Long) _
    As Long



'creates a font compatable with the specified device context
Public Declare Function CreateFontIndirect _
    Lib "gdi32" _
    Alias "CreateFontIndirectA" _
    (lpLogFont As LogFont) _
    As Long



'creates a pen that can be applied to a hDc
Public Declare Function CreatePenIndirect _
    Lib "gdi32" _
    (lpLogPen As LogPen) _
    As Long



'removes a device context from memory
Public Declare Function DeleteDC _
    Lib "gdi32" _
    (ByVal hDC As Long) _
    As Long

'removes an object such as a brush or bitmap from memory
Public Declare Function DeleteObject _
    Lib "gdi32" _
    (ByVal hObject As Long) _
    As Long



'this draws text onto the bitmap
Public Declare Function DrawText _
    Lib "user32" _
    Alias "DrawTextA" _
    (ByVal hDC As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) _
    As Long



'-----------------------------

'this will get the current devices' capabilities
Public Declare Function GetDeviceCaps _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal nIndex As Long) _
    As Long

'get the last error to occur from within the api
Public Declare Function GetLastError _
    Lib "kernel32" _
    () _
    As Long



'get the dimensions of the applied text metrics for
'the device context
Public Declare Function GetTextMetrics _
    Lib "gdi32" _
    Alias "GetTextMetricsA" _
    (ByVal hDC As Long, _
    lpMetrics As TEXTMETRIC) _
    As Long

'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
'Public Declare Function GetTickCount Lib "kernel32" () As Long

'gets any intersection of two rectangles
Public Declare Function IntersectRect _
    Lib "user32" _
    (lpDestRect As RECT, _
    lpSrc1Rect As RECT, _
    lpSrc2Rect As RECT) _
    As Long




'Pattern Blitter. Used to draw a pattern onto
'a device context
Public Declare Function PatBlt _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal dwRop As Long) _
    As Long


'This will draw a set of lines to the specifed
'points
Public Declare Function Polyline _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    lpPoint As POINTAPI, _
    ByVal nCount As Long) _
    As Long

'This will draw a set of lines starting from
'the current position on the device context.
Public Declare Function PolylineTo _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    lppt As POINTAPI, _
    ByVal cCount As Long) _
    As Long

'This draws a rectangle onto the device
'context
Public Declare Function Rectangle _
    Lib "gdi32" _
    (ByVal hwnd As Long, _
    x1 As Integer, _
    y1 As Integer, _
    x2 As Integer, _
    y2 As Integer) _
    As Long


'-------------
'this will select the specified object to
'a window or device context
Public Declare Function SelectObject _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal hObject As Long) _
    As Long

'This sets the background colour on a bitmap
Public Declare Function SetBkColor _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal crColor As Long) _
    As Long

'This sets the background mode on a bitmap
'(eg, transparent, solid etc)
Public Declare Function SetBkMode _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal nBkMode As Long) _
    As Long



'sets the current text colour
Public Declare Function SetTextColor _
    Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal crColor As Long) _
    As Long

'pauses the execution of the programs thread
'for a specified amount of milliseconds
Public Declare Sub Sleep _
    Lib "kernel32" _
    (ByVal dwMilliseconds As Long)




'--------------------------------------------------------------------------
'enumerator section
'--------------------------------------------------------------------------

'the direction of the gradient
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum

'in twips or pixels
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum

'The key values of the mouse buttons
Public Enum MouseKeys
    MouseLeft = 1
    MouseRight = 2
    MouseMiddle = 4
End Enum

'text alignment constants
Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

'bitmap flip constants
Public Enum FlipType
    FlipHorizontally = 0
    FlipVertically = 1
End Enum

'image load constants
Public Enum LoadType
    IMAGE_BITMAP& = 0
End Enum

'rotate bitmap constants
Public Enum RotateType
    RotateRight = 0
    RotateLeft = 1
    Rotate180 = 2
End Enum

'--------------------------------------------------------------------------
'programr defined data types
'--------------------------------------------------------------------------

'AlphaBlend information for bitmaps
Private Type BLENDFUNCTION
    bytBlendOp As Byte              'currently the only blend op supported by windows 98+ is AC_SRC_OVER
    bytBlendFlags As Byte           'must be left blank
    bytSourceConstantAlpha As Byte  'the amount to blend by. Must be between 0 and 255
    bytAlphaFormat As Byte          'don't set this. If you wish more infor, go to "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/bitmaps_3b3m.asp"
End Type

'Bitmap structue for menu information
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'size structure
Public Type SizeType
    cx As Long
    cy As Long
End Type

'Text metrics
Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type


Public Type COLORADJUSTMENT
    caSize As Integer
    caFlags As Integer
    caIlluminantIndex As Integer
    caRedGamma As Integer
    caGreenGamma As Integer
    caBlueGamma As Integer
    caReferenceBlack As Integer
    caReferenceWhite As Integer
    caContrast As Integer
    caBrightness As Integer
    caColorfulness As Integer
    caRedGreenTint As Integer
End Type

Public Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Public Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type

Public Type LogColorSpace
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSType As Long
    lcsIntent As Long
    lcsEndPoints As CIEXYZTRIPLE
    lcsGammaRed As Long
    lcsGammaGreen As Long
    lcsGammaBlue As Long
    lcsFileName As String * 26 'MAX_PATH
End Type

'display settings (800x600 etc)
Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BitmapStruc
    hDcMemory As Long
    hDcBitmap As Long
    hDcPointer As Long
    Area As RECT
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LogPen
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Public Type LogBrush
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public Type FontStruc
    Name As String
    Alignment As AlignText
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    StrikeThru As Boolean
    PointSize As Byte
    Colour As Long
End Type

Public Type LogFont
    'for the DrawText api call
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type

Public Type Point
    'you'll need this to reference a point on the
    'screen'
    x As Integer
    y As Integer
End Type

'To hold the RGB value
Public Type RGBVal
    Red As Single
    Green As Single
    Blue As Single
End Type

'bitmap structure for the GetObject api call
Public Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'--------------------------------------------------------------------------
'Constants section
'--------------------------------------------------------------------------

'general constants
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const WM_USER = &H400

'alphablend constants
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H0

'Image load constants
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const LR_DEFAULTSIZE As Long = &H40

'PatBlt constants
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

'Display constants
Public Const CDS_FULLSCREEN = 4
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000

'DrawText constants
Public Const DT_CENTER = &H1
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'CreateBrushIndirect constants
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_DIBPATTERNPT = 6
Public Const BS_HATCHED = 2
Public Const BS_HOLLOW = 1
Public Const BS_NULL = 1
Public Const BS_PATTERN = 3
Public Const BS_PATTERN8X8 = 7
Public Const BS_SOLID = 0
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_NOSHADE = 17
Public Const HS_SOLID = 8
Public Const HS_SOLIDBKCLR = 23
Public Const HS_SOLIDCLR = 19
Public Const HS_VERTICAL = 1                '  |||||

'BitBlt constants
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)

'LogFont constants
Public Const LF_FACESIZE = 32
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100
Public Const DEFAULT_CHARSET = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_MASK = &HF
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_TT_ALWAYS = 32
Public Const WM_SETFONT = &H30
Public Const LF_FULLFACESIZE = 64
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_QUALITY = 0
Public Const PROOF_QUALITY = 2

'GetDeviceCaps constants
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

'colourspace constants
Public Const MAX_PATH = 260

'pen constants
Public Const PS_COSMETIC = &H0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DASHDOT = 3                 '  _._._._
    Public Const PS_DASHDOTDOT = 4              '  _.._.._
    Public Const PS_DOT = 2                     '  .......
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_GEOMETRIC = &H10000
Public Const PS_INSIDEFRAME = 6
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0
Public Const PS_SOLID = 0


Private blnResChanged As Boolean

'--------------------------------------------------------------------------
'Procedures/functions section
'--------------------------------------------------------------------------

Public Sub DrawRect(ByVal lngHDC As Long, _
        ByVal lngColour As Long, _
        ByVal intLeft As Integer, _
        ByVal intTop As Integer, _
        ByVal intRight As Integer, _
        ByVal intBottom As Integer, _
        Optional ByVal udtMeasurement As Scaling = InPixels, _
        Optional ByVal lngStyle As Long = BS_SOLID, _
        Optional ByVal lngPattern As Long = HS_SOLID)
    
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Dim StartRect As RECT
    Dim lngResult As Long
    Dim lngJunk  As Long
    Dim lnghBrush As Long
    Dim BrushStuff As LogBrush

    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intRight = intRight / Screen.TwipsPerPixelX
        intBottom = intBottom / Screen.TwipsPerPixelY
    End If
    
    'initalise values
    StartRect.Top = intTop
    StartRect.Left = intLeft
    StartRect.Bottom = intBottom
    StartRect.Right = intRight
    
    'create a brush
    BrushStuff.lbColor = lngColour
    BrushStuff.lbHatch = lngPattern
    BrushStuff.lbStyle = lngStyle
    
    'apply the brush to the device context
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lngHDC, lnghBrush)
    
    'draw a rectangle
    lngResult = PatBlt(lngHDC, _
        intLeft, _
        intTop, _
        (intRight - intLeft), _
        (intBottom - intTop), _
        PATCOPY)
    
    'A "Brush" object was created. It must be removed from memory.
    lngJunk = SelectObject(lngHDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub DrawLine(lngHDC As Long, _
        ByVal intX1 As Integer, _
        ByVal intY1 As Integer, _
        ByVal intX2 As Integer, _
        ByVal intY2 As Integer, _
        Optional ByVal lngColour As Long = 0, _
        Optional ByVal intWidth As Integer = 1, _
        Optional ByVal udtMeasurement As Scaling = InTwips)
                    
    'This will draw a line from point1 to point2
    
    Const NumOfPoints = 2
    
    Dim lngResult As Long
    Dim lnghPen As Long
    Dim PenStuff As LogPen
    Dim Junk  As Long
    Dim Points(NumOfPoints) As POINTAPI

    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intX1 = intX1 / Screen.TwipsPerPixelX
        intX2 = intX2 / Screen.TwipsPerPixelX
        intY1 = intY1 / Screen.TwipsPerPixelY
        intY2 = intY2 / Screen.TwipsPerPixelY
    End If
    
    'Find out if a specific lngColour is to be set. If so set it.
    PenStuff.lopnColor = lngColour
    PenStuff.lopnStyle = PS_GEOMETRIC
    PenStuff.lopnWidth.x = intWidth
    
    'apply the pen settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(lngHDC, lnghPen)
    
    'set the points
    Points(1).x = intX1
    Points(1).y = intY1
    Points(2).x = intX2
    Points(2).y = intY2
    
    'draw the line
    lngResult = Polyline(lngHDC, Points(1), NumOfPoints)
    lngResult = GetLastError
    
    'A "Pen" object was created. It must be removed from memory.
    Junk = SelectObject(lngHDC, lnghPen)
    Junk = DeleteObject(Junk)
End Sub

Public Sub CreateNewBitmap(ByRef hDcMemory As Long, _
        ByRef hDcBitmap As Long, _
        ByRef hDcPointer As Long, _
        ByRef BmpArea As RECT, _
        ByVal CompatableWithhDc As Long, _
        Optional ByVal lngBackColour As Long = 0, _
        Optional ByVal udtMeasurement As Scaling = InPixels)
                           
    'This procedure will create a new bitmap compatable with a given
    'form (you will also be able to then use this bitmap in a picturebox).
    'The space specified in "Area" should be in "Twips" and will be
    'converted into pixels in the following code.
    
    Dim lngResult As Long
    Dim Area As RECT

    'scale the bitmap points if necessary
    Area = BmpArea
    If udtMeasurement = InTwips Then
        Call RectToPixels(Area)
    End If
    
    'create the bitmap and its references
    hDcMemory = CreateCompatibleDC(CompatableWithhDc)
    hDcBitmap = CreateCompatibleBitmap(CompatableWithhDc, _
        (Area.Right - Area.Left), _
        (Area.Bottom - Area.Top))
    hDcPointer = SelectObject(hDcMemory, hDcBitmap)
    
    'set default colours and clear bitmap to selected colour
    lngResult = SetBkMode(hDcMemory, OPAQUE)
    lngResult = SetBkColor(hDcMemory, lngBackColour)
    Call DrawRect(hDcMemory, _
        lngBackColour, _
        0, _
        0, _
        (Area.Right - Area.Left), _
        (Area.Bottom - Area.Top))
End Sub

Public Sub DeleteBitmap(ByRef hDcMemory As Long, _
        ByRef hDcBitmap As Long, _
        ByRef hDcPointer As Long)
                        
    'This will remove the bitmap that stored what was displayed before
    'the text was written to the screen, from memory.
    
    Dim lngJunk As Long

    If hDcMemory = 0 Then
        'there is nothing to delete. Exit the sub-routine
        Exit Sub
    End If
    
    'delete the device context
    lngJunk = SelectObject(hDcMemory, hDcPointer)
    lngJunk = DeleteObject(hDcBitmap)
    lngJunk = DeleteDC(hDcMemory)
    
    'show that the device context has been deleted by setting
    'all parameters passed to the procedure to zero
    hDcMemory = 0
    hDcBitmap = 0
    hDcPointer = 0
End Sub

Public Sub RectToPixels(ByRef TheRect As RECT)
    'converts twips to pixels in a rect structure
    
    TheRect.Left = TheRect.Left \ Screen.TwipsPerPixelX
    TheRect.Right = TheRect.Right \ Screen.TwipsPerPixelX
    TheRect.Top = TheRect.Top \ Screen.TwipsPerPixelY
    TheRect.Bottom = TheRect.Bottom \ Screen.TwipsPerPixelY
End Sub

Public Sub Gradient(ByVal lngDesthDc As Long, _
        ByVal lngStartCol As Long, _
        ByVal FinishCol As Long, _
        ByVal intLeft As Integer, _
        ByVal intTop As Integer, _
        ByVal intWidth As Integer, _
        ByVal intHeight As Integer, _
        ByVal Direction As GradientTo, _
        Optional ByVal udtMeasurement As Scaling = 1, _
        Optional ByVal bytLineWidth As Byte = 1)
                    
    'draws a gradient from colour mblnStart to colour Finish, and assums
    'that all measurments passed to it are in pixels unless otherwise
    'specified.
    
    Dim intCounter As Integer
    Dim intBiggestDiff As Integer
    Dim Colour As RGBVal
    Dim mblnStart As RGBVal
    Dim Finish As RGBVal
    Dim sngAddRed As Single
    Dim sngAddGreen As Single
    Dim sngAddBlue As Single

    'perform all necessary calculations before drawing gradient
    'such as converting long to rgb values, and getting the correct
    'scaling for the bitmap.
    mblnStart = GetRGB(lngStartCol)
    Finish = GetRGB(FinishCol)
    
    If udtMeasurement = InTwips Then
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelY
        intWidth = intWidth / Screen.TwipsPerPixelX
        intHeight = intHeight / Screen.TwipsPerPixelY
    End If
    
    'draw the colour gradient
    Select Case Direction
        Case GradVertical
            intBiggestDiff = intWidth
        Case GradHorizontal
            intBiggestDiff = intHeight
    End Select
    
    'calculate how much to increment/decrement each colour per step
    sngAddRed = (bytLineWidth * ((Finish.Red) - mblnStart.Red) / intBiggestDiff)
    sngAddGreen = (bytLineWidth * ((Finish.Green) - mblnStart.Green) / intBiggestDiff)
    sngAddBlue = (bytLineWidth * ((Finish.Blue) - mblnStart.Blue) / intBiggestDiff)
    Colour = mblnStart
    
    'calculate the colour of each line before drawing it on the bitmap
    For intCounter = 0 To intBiggestDiff Step bytLineWidth
        'find the point between colour mblnStart and Colour Finish that
        'corresponds to the point between 0 and intBiggestDiff
        
        'check for overflow
        If Colour.Red > 255 Then
            Colour.Red = 255
        Else
            If Colour.Red < 0 Then
                Colour.Red = 0
            End If
        End If
        If Colour.Green > 255 Then
            Colour.Green = 255
        Else
            If Colour.Green < 0 Then
                Colour.Green = 0
            End If
        End If
        If Colour.Blue > 255 Then
            Colour.Blue = 255
        Else
            If Colour.Blue < 0 Then
                Colour.Blue = 0
            End If
        End If
        
        'draw the gradient in the proper orientation in the calculated colour
        Select Case Direction
            Case GradVertical
                Call DrawLine(lngDesthDc, _
                    intCounter + intLeft, _
                    intTop, _
                    intCounter + intLeft, _
                    intHeight + intTop, _
                    RGB(Colour.Red, Colour.Green, Colour.Blue), _
                    bytLineWidth, _
                    InPixels)
            Case GradHorizontal
                Call DrawLine(lngDesthDc, _
                    intLeft, _
                    intCounter + intTop, _
                    intLeft + intWidth, _
                    intTop + intCounter, _
                    RGB(Colour.Red, Colour.Green, Colour.Blue), _
                    bytLineWidth, _
                    InPixels)
        End Select
        
        'set next colour
        Colour.Red = Colour.Red + sngAddRed
        Colour.Green = Colour.Green + sngAddGreen
        Colour.Blue = Colour.Blue + sngAddBlue
    Next intCounter

End Sub

Public Function GetRGB(ByVal lngColour As Long) _
        As RGBVal
    'Convert Long to RGB:
    
    'if the lngcolour value is greater than
    'acceptable then half the value
    If (lngColour > RGB(255, 255, 255)) _
        Or (lngColour < (RGB(255, 255, 255) * -1)) Then
    Exit Function
End If
    
    GetRGB.Blue = (lngColour \ 65536)
    GetRGB.Green = ((lngColour - (GetRGB.Blue * 65536)) \ 256)
    GetRGB.Red = (lngColour - (GetRGB.Blue * (65536)) - ((GetRGB.Green) * 256))
End Function

Public Sub Pause(lngTicks As Long)
    'pause execution of the program for a specified number of lngTicks
    
    If lngTicks < 0 Then
        lngTicks = 0
    End If
    Call Sleep(lngTicks)
End Sub

Public Sub MakeText(ByVal hDcSurphase As Long, _
        ByVal strText As String, _
        ByVal intTop As Integer, _
        ByVal intLeft As Integer, _
        ByVal intHeight As Integer, _
        ByVal intWidth As Integer, _
        ByRef udtFont As FontStruc, _
        Optional ByVal udtMeasurement As Scaling = 0)
                    
    'This procedure will draw strText onto the bitmap in the specified udtFont,
    'colour and position.
    
    Dim udtAPIFont As LogFont
    Dim lngAlignment As Long
    Dim udtTextRect As RECT
    Dim lngResult As Long
    Dim lngJunk As Long
    Dim hDcFont As Long
    Dim hDcOldFont As Long
    Dim intCounter As Integer

    'set Measurement values
    udtTextRect.Top = intTop
    udtTextRect.Left = intLeft
    udtTextRect.Right = intLeft + intWidth
    udtTextRect.Bottom = intTop + intHeight
    
    If udtMeasurement = InTwips Then
        'convert to pixels
        Call RectToPixels(udtTextRect)
    End If
    
    'Create details about the udtFont using the udtFont structure
    '====================
    
    'convert point size to pixels
    udtAPIFont.lfHeight = -((udtFont.PointSize * GetDeviceCaps(hDcSurphase, LOGPIXELSY)) / 72)
    udtAPIFont.lfCharSet = DEFAULT_CHARSET
    udtAPIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    udtAPIFont.lfEscapement = 0
    
    'move the name of the udtFont into the array
    For intCounter = 1 To Len(udtFont.Name)
        udtAPIFont.lfFaceName(intCounter) = Asc(Mid(udtFont.Name, intCounter, 1))
    Next intCounter

    'this has to be a Null terminated string
    udtAPIFont.lfFaceName(intCounter) = 0
    
    udtAPIFont.lfItalic = udtFont.Italic
    udtAPIFont.lfUnderline = udtFont.Underline
    udtAPIFont.lfStrikeOut = udtFont.StrikeThru
    udtAPIFont.lfOrientation = 0
    udtAPIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    udtAPIFont.lfPitchAndFamily = DEFAULT_PITCH
    udtAPIFont.lfQuality = PROOF_QUALITY
    
    If udtFont.Bold Then
        udtAPIFont.lfWeight = FW_BOLD
    Else
        udtAPIFont.lfWeight = FW_NORMAL
    End If
    
    udtAPIFont.lfWidth = 0
    hDcFont = CreateFontIndirect(udtAPIFont)
    hDcOldFont = SelectObject(hDcSurphase, hDcFont)
    '====================
    
    Select Case udtFont.Alignment
        Case vbLeftAlign
            lngAlignment = DT_LEFT
        Case vbCentreAlign
            lngAlignment = DT_CENTER
        Case vbRightAlign
            lngAlignment = DT_RIGHT
    End Select
    
    'Draw the strText into the off-screen bitmap before copying the
    'new bitmap (with the strText) onto the screen.
    lngResult = SetBkMode(hDcSurphase, TRANSPARENT)
    lngResult = SetTextColor(hDcSurphase, udtFont.Colour)
    lngResult = DrawText(hDcSurphase, _
        strText, _
        Len(strText), _
        udtTextRect, _
        lngAlignment)
    
    'clean up by deleting the off-screen bitmap and udtFont
    lngJunk = SelectObject(hDcSurphase, hDcOldFont)
    lngJunk = DeleteObject(hDcFont)
End Sub

Public Function GetTextHeight(ByVal hDC As Long) _
        As Integer
    'This function will return the height of the text using the point size
    
    Dim udtMetrics As TEXTMETRIC
    Dim lngResult As Long

    lngResult = GetTextMetrics(hDC, udtMetrics)
    
    GetTextHeight = udtMetrics.tmHeight
End Function

