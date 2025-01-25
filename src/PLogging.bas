Attribute VB_Name = "PLogging"
'@Folder "PLoggerProject.src"
Option Explicit

Private Type TPLogging
    Loggers As Object ' Dictionary
    Root As PLogger
End Type

Private this As TPLogging

Public Property Get Root() As PLogger
    Set Root = this.Root
End Property

Public Property Get StreamHandler() As Handler
    Set StreamHandler = NewStreamHandler()
End Property

Public Property Get LoggersCount() As Long
    If this.Loggers Is Nothing Then
        CreateRoot
    End If

    LoggersCount = this.Loggers.Count
End Property

Public Function GetLogger(ByVal Name As String) As PLogger
    If this.Loggers Is Nothing Then
        CreateRoot
    End If

    If Not this.Loggers.Exists(Name) Then
        Set this.Loggers(Name) = NewPLogger(Name)
    End If

    Set GetLogger = this.Loggers(Name)
End Function

Private Sub CreateRoot()
    Set this.Loggers = NewDictionary()
    Set this.Root = NewPLogger("Root")
    this.Root.SetLevel LogLevels.llError
    this.Root.AddHandler NewStreamHandler()

    Set this.Loggers(this.Root.Name) = this.Root
End Sub
