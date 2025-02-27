VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PLoggerProject.src.Handlers.FileHandler"
Option Explicit

Implements Handler

Private Type TFileHandler
    Level As Long
    Formatter As Formatter
    FilePath As String
    Mode As IOMode
    Encoding As String
    Stream As Object
End Type

Private this As TFileHandler

Private Sub Class_Initialize()
    If Not this.Stream Is Nothing Then
        CloseStream
    End If
End Sub

Private Sub Class_Terminate()
    If Not this.Stream Is Nothing Then
        CloseStream
    End If
End Sub

Private Property Get Handler_Name() As String
    Handler_Name = Me.Name
End Property

Private Sub Handler_Emit(ByRef Record As Object)
    Me.Emit Record
End Sub

Private Sub Handler_SetFormatter(ByRef Formatter As Formatter)
    Me.SetFormatter Formatter
End Sub

Public Property Get FilePath() As String
    FilePath = this.FilePath
End Property

Public Property Let FilePath(ByVal RHS As String)
    this.FilePath = RHS
End Property

Public Property Get Mode() As IOMode
    Mode = this.Mode
End Property

Public Property Let Mode(ByVal RHS As IOMode)
    this.Mode = RHS
End Property

Public Property Get Encoding() As String
    Encoding = this.Encoding
End Property

Public Property Let Encoding(ByVal RHS As String)
    this.Encoding = RHS
End Property

Public Property Get Name() As String
    Name = Information.TypeName(Me)
End Property

Public Sub SetFormatter(ByRef Formatter As Formatter)
    Set this.Formatter = Formatter
End Sub

Public Sub Emit(ByRef Record As Object)
    If this.Formatter Is Nothing Then
        Set this.Formatter = NewFormatter()
    End If

    If this.Stream Is Nothing Then
        OpenStream
    End If

    Dim Msg As String
    Msg = this.Formatter.Format(Record)
    this.Stream.WriteText Msg, 1 ' adWriteLine
End Sub

Private Sub OpenStream()
    ' Type 2 - adTypeText
    Set this.Stream = NewStream( _
        StreamType:=2, _
        Charset:=this.Encoding _
    )
    this.Stream.Open

    If this.Mode = IOMode.Appending Then
        On Error Resume Next
        this.Stream.LoadFromFile this.FilePath
        this.Stream.Position = this.Stream.Size
        On Error GoTo 0
    End If
End Sub

Private Sub CloseStream()
    If Not this.Stream Is Nothing Then
        this.Stream.SaveToFile this.FilePath, 2 ' adSaveCreateOverWrite
        this.Stream.Close
        Set this.Stream = Nothing
    End If
End Sub
