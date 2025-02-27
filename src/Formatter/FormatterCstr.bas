Attribute VB_Name = "FormatterCstr"
'@Folder "PLoggerProject.src.Formatter"
Option Explicit

Public Function NewFormatter(Optional ByVal Fmt As String = "{levelname}:{name}:{message}") As Formatter
    Set NewFormatter = New Formatter
    NewFormatter.Fmt = Fmt
End Function
