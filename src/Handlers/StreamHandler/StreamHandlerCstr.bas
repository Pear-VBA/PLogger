Attribute VB_Name = "StreamHandlerCstr"
'@Folder "PLoggerProject.src.Handlers.StreamHandler"
Option Explicit

Public Function NewStreamHandler() As StreamHandler
    Set NewStreamHandler = New StreamHandler
End Function
