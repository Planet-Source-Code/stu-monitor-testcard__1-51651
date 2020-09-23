Attribute VB_Name = "modUserControlHelper"
Option Explicit
'-----------------------------------------'
' By Paul Sanders, pa_sanders@hotmail.com '
'--------------------------------------------------------------------------------------------
'            :
' Project    : n/a
' Module     : modUserControlHelper
'            :
' Created    : 16-Jan-99 15:16
'            :
' Notes      :
'            :
' References : None
'            :
'--------------------------------------------------------------------------------------------

Private Const MODULENAME = "modUserControlHelper::"


'--------------------------------------------------------------------------------------------
'Procedure : DefineAccessKeys
'Author    : Paul Sanders, pa_sanders@hotmail.com, 16-May-01 15:15
'Notes     : Determines the access key from a string
'--------------------------------------------------------------------------------------------
Public Function DefineAccessKeys(ByVal sCaption As String) As String
    Dim i As Integer
    Dim sKeys As String
    Dim nPos As Integer
    
    If Len(sCaption) = 0 Then
        DefineAccessKeys = ""
    Else
        nPos = InStr(sCaption, "&&")
        If nPos = 0 Then
            nPos = InStr(sCaption, "&")
            If nPos > 0 Then
                sKeys = Mid$(sCaption, nPos + 1, 1)
            End If
        End If
        DefineAccessKeys = sKeys
    End If
End Function

