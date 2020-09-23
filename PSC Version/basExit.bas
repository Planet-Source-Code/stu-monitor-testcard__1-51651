Attribute VB_Name = "basExit"
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub Ex()
    Dim file As String
    Dim wFlags%
    Dim x As Integer
    Dim SoundName As String
    Dim i As Integer
    

    'ends tone.wav if running
    file = " "
    SoundName$ = file
    x = sndPlaySound(SoundName$, wFlags%)
        
    'clears down any Help file
    Call QuitHelp

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next

    End
    
End Sub
