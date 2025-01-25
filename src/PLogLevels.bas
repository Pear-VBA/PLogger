Attribute VB_Name = "PLogLevels"
'@Folder "PLoggerProject.src"
Option Explicit

Public Enum LogLevels
    llNotSet = 0
    llTrace = 10
    llDebug = 20
    llInfo = 30
    llWarn = 40
    llError = 50
    llFatal = 60
End Enum

Public Function LevelToString(ByVal Level As LogLevels) As String
    Dim Levels As Object
    Set Levels = NewDictionary()
    Levels(LogLevels.llNotSet) = ""
    Levels(LogLevels.llTrace) = "TRACE"
    Levels(LogLevels.llDebug) = "DEBUG"
    Levels(LogLevels.llInfo) = "INFO"
    Levels(LogLevels.llWarn) = "WARN"
    Levels(LogLevels.llError) = "ERROR"
    Levels(LogLevels.llFatal) = "FATAL"

    LevelToString = Levels(Level)
End Function
