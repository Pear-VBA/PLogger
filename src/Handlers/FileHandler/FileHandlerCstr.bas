Attribute VB_Name = "FileHandlerCstr"
'@Folder "PLoggerProject.src.Handlers.FileHandler"
Option Explicit

Public Function NewFileHandler( _
    ByVal FilePath As String, _
    Optional ByVal Mode As IOMode = IOMode.Appending, _
    Optional ByVal Encoding As String = "utf-8" _
) As FileHandler
    Dim Buff As FileHandler
    Set Buff = New FileHandler
    Buff.FilePath = FilePath
    Buff.Mode = Mode
    Buff.Encoding = Encoding
    Set NewFileHandler = Buff
End Function
