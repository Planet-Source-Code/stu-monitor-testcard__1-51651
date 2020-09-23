Attribute VB_Name = "basScreen"


Option Explicit

Public Const DM_BITSPERPEL As Long = &H40000
Public Const DM_PELSWIDTH As Long = &H80000
Public Const DM_PELSHEIGHT As Long = &H100000
Public Const CDS_FORCE As Long = &H80000000
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10
Public Const BITSPIXEL As Long = 12
Public Const VREFRESH As Long = 116

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

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal modeIndex As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Public lpDevMode()  As DEVMODE
Public CurrentIndex As Long

Public Sub GetDisplaySettings(displayDescr() As String)
    
    Dim Index        As Long
    Dim displayCount As Long
    Dim Colours      As String
    Dim scnHeight    As String
    Dim scnWidth     As String

    ' set the DEVMODE flags and structure size
    ReDim lpDevMode(0 To 1) As DEVMODE
    lpDevMode(0).dmSize = Len(lpDevMode(0))
    lpDevMode(0).dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    
    ' collect display settings
    Do While EnumDisplaySettings(0, displayCount, lpDevMode(0)) > 0
        displayCount = displayCount + 1
    Loop

    ' now displayCount holds the number of display settings
    ' and we can DIMension the result arrays
    ReDim displayDescr(0 To displayCount) As String
    ReDim lpDevMode(0 To displayCount) As DEVMODE
    
    For Index = 0 To displayCount
        
        ' retrieve info on the index-th display mode
        EnumDisplaySettings 0, Index, lpDevMode(Index)
        
        Select Case lpDevMode(Index).dmBitsPerPel
            Case 4
                Colours = "    16 "
            Case 8
                Colours = "  256 "
            Case 16
                Colours = "16bit "
            Case 24
                Colours = "24bit "
            Case 32
                Colours = "32bit "
        End Select
        
        'aligns results in listbox
        scnWidth = lpDevMode(Index).dmPelsWidth
        scnHeight = lpDevMode(Index).dmPelsHeight

        With lpDevMode(Index)
            If .dmPelsWidth < 1000 Then scnWidth = "  " & .dmPelsWidth
            If .dmPelsHeight < 1000 Then scnHeight = "  " & .dmPelsHeight

            displayDescr(Index) = " " & scnWidth & " x " & scnHeight & " x  " & Colours

            'tests for result of Freq

            If .dmDisplayFrequency > 1 Then
                displayDescr(Index) = displayDescr(Index) & " - " & .dmDisplayFrequency & " Hz"
          
            Else
                displayDescr(Index) = displayDescr(Index) & " -   (Def)"
            
            End If
     
        End With
    Next

End Sub

Public Function ChangeScreenResolution(ByRef Index As Long) As Boolean
    
    If ChangeDisplaySettings(lpDevMode(Index), CDS_FORCE) = 0 Then _
        ChangeScreenResolution = True
   
End Function

Public Function lookupCurrent() As Long
    'get the system settings

    Dim currHRes   As Long
    Dim currVRes   As Long
    Dim currBPP    As Long
    Dim currVFreq  As Long
    Dim sBPPtype   As String
    Dim sFreqtype  As String
    Dim hDC        As Long
    Dim i          As Long

    lookupCurrent = -1
   
    hDC = GetDC(0)
   
    currHRes = GetDeviceCaps(hDC, HORZRES)
    currVRes = GetDeviceCaps(hDC, VERTRES)
    currBPP = GetDeviceCaps(hDC, BITSPIXEL)
    currVFreq = GetDeviceCaps(hDC, VREFRESH)
   
    Call DeleteDC(hDC)
   
    For i = 0 To UBound(lpDevMode) - 1
   
        If lpDevMode(i).dmPelsWidth = currHRes Then
        If (lpDevMode(i).dmPelsHeight = currVRes) Then
        If (lpDevMode(i).dmBitsPerPel = currBPP) Then
        If (lpDevMode(i).dmDisplayFrequency = currVFreq) Then
        lookupCurrent = i
                    
        Exit Function

        End If
        End If
        End If
        End If
   
    Next

End Function

Public Function FillList(List As Object)
    Dim stringList() As String
    Dim i As Long

    List.Clear

    Call GetDisplaySettings(stringList)

    For i = 0 To UBound(stringList) - 1
        List.AddItem stringList(i)
    Next

    CurrentIndex = lookupCurrent()

    If CurrentIndex <> -1 Then
        List.ListIndex = CurrentIndex
    Else
    
        MsgBox "Error: Could not read current settings!", vbCritical
        
    End If

End Function

Public Property Get CurrentResolution() As Long

    CurrentResolution = CurrentIndex

End Property

Public Property Let CurrentResolution(ByVal vNewValue As Long)

    Dim Msg1 As String
    Dim Msg2 As String
    Dim Msg3 As String

    Msg1 = "This monitor or some programs may not be able to cope with these Screen Settings without Re-Booting."
    Msg2 = "If this change is temporary, then revert back to the original after you have finished testing."
    Msg3 = "Screen Settings will now change, press Esc if the system is unable to display the new settings."

    MsgBox Msg1 + vbNewLine + Msg2 + vbNewLine + vbNewLine + Msg3, vbInformation
 
    If ChangeScreenResolution(vNewValue) Then
        If Not (MsgBox("Keep current setting?", vbOKCancel + vbDefaultButton2 + vbQuestion) = vbOK) Then
        
        Call ChangeScreenResolution(CurrentIndex)
    
        Else
        
        CurrentIndex = vNewValue

    End If
    End If

End Property
