Attribute VB_Name = "basBlueGrad"
Option Explicit

' Declaration necessary to use Sleep API
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const Pause As Integer = 1

Dim intY As Integer
Dim sColour As Long

Public Sub GradeForm(pObject As Object, Optional Colour As Integer, Optional Orientation As Integer = 0, Optional Range As Integer = 500)
    'grades the screen from black to blue

    pObject.Scale (0, 0)-(Range, Range)

    For intY = 0 To Range

        'this line dictates the colour scheme
        sColour = RGB(0, 0, CInt((intY / Range) * 255))

        'dictates direction of shading
        pObject.Line (0, intY)-(Range, intY), sColour

    Next intY

End Sub

Public Sub GradeSplash(pObject As Object, Optional Colour As Integer, Optional Orientation As Integer = 0, Optional Range As Integer = 1000)
    'grades the progress bar in frmSplash

    pObject.Scale (0, 0)-(Range, Range)

    For intY = 0 To Range
        
        Sleep Pause

        'this line dictates the colour scheme
        sColour = RGB(0, CInt((intY / Range) * 255), 0)

        'left to right
        pObject.Line (intY, 0)-(intY, Range), sColour
    Next intY

End Sub
