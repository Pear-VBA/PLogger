Attribute VB_Name = "PLoggerCstr"
'@Folder "PLoggerProject.src.PLogger"
Option Explicit

Public Function NewPLogger(ByVal Name As String) As PLogger
    Set NewPLogger = New PLogger
    NewPLogger.Name = Name
End Function
