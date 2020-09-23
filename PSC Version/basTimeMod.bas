Attribute VB_Name = "basTimeMod"
Option Explicit

Public Enum TimeFormatType
    DaysHoursMinutesSecondsMilliseconds = 0
    DaysHoursMinutesSeconds = 1
    DHMSMColonSeparated = 2
    DaysHoursMinutes = 3
End Enum

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long



