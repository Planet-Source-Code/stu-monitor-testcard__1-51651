Attribute VB_Name = "modPS"
Option Explicit

'-----------------------------------------'
' By Paul Sanders, pa_sanders@hotmail.com '
'--------------------------------------------------------------------------------------------
'            :
' Project    : Testcard
' Module     : modPS
'            :
' Created    : 30-Jul-02 11:20
'            :
' Notes      :
'            :
' References : None
'            :
'--------------------------------------------------------------------------------------------

Private Const MODULENAME = "modPS::"

'--------------------------------------------------------------------------------------------
'Procedure : SwitchTestScreen
'Author    : Paul Sanders, pa_sanders@hotmail.com, 30-Jul-02 11:22
'Notes     : Common method to switch between test screens
'--------------------------------------------------------------------------------------------
Public Sub SwitchTestScreen(KeyAscii As Integer, fraControl As Control, bUnload As Boolean)
    Dim frm As Form
    Dim nKey As Integer
    
    'toggle toolbar
    
     If Not fraControl Is Nothing Then
        If KeyAscii = vbKeyShift Then
            fraControl.Visible = Not (fraControl.Visible)
            KeyAscii = 0
            Exit Sub '---------------------------->-->-->
        End If
        Set frm = fraControl.Parent
    End If


    
    'change form with function key
    
    If Not frm Is Nothing Then
        If frm.Tag <> "" Then
            nKey = CInt(frm.Tag)
        End If
    End If
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeySpace Then
        If Not frm Is Nothing Then
            If bUnload Then
                Unload frm
            Else
                frm.Hide
            End If
            Set frm = Nothing
        End If
    End If
 
    If KeyAscii = vbKeyF2 And nKey <> vbKeyF2 Then frmGeometry.Show
    If KeyAscii = vbKeyF3 And nKey <> vbKeyF3 Then frmCircles.Show
    If KeyAscii = vbKeyF4 And nKey <> vbKeyF4 Then frmConvergence.Show
    If KeyAscii = vbKeyF5 And nKey <> vbKeyF5 Then frmPluge.Show
    If KeyAscii = vbKeyF6 And nKey <> vbKeyF6 Then frmGreyscale.Show
    If KeyAscii = vbKeyF7 And nKey <> vbKeyF7 Then frmBars.Show
    If KeyAscii = vbKeyF8 And nKey <> vbKeyF8 Then frmTestcard.Show
    If KeyAscii = vbKeyF9 And nKey <> vbKeyF9 Then frmPurity.BackColor = vbRed: frmPurity.Show
End Sub
