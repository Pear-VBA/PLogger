Attribute VB_Name = "Example"
'@Folder "PLoggerProject"
Option Explicit

Public Sub Test()
    Dim Log As PLogger
    Set Log = PLogging.GetLogger("example")

    Dim SH As Handler
    Set SH = NewStreamHandler()
    SH.SetFormatter NewFormatter()

    FileSystem.ChDir ThisWorkbook.Path
    Dim FH As Handler
    Set FH = NewFileHandler("example.log", Mode:=Writing)
    FH.SetFormatter NewFormatter("[{name}] - {levelname} - {message}")

    Log.AddHandler SH
    Log.AddHandler FH
    Log.SetLevel llWarn

    Log.Debug_ "Not showed"
    Log.Info "Not showed too"
    Log.Warn "Hello, World!"
    Log.Error "Lets make some ERROR"
    Log.Fatal "omg..."

    PLogging.Clean
End Sub
