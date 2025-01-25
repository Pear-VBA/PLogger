VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PLoggerProject.src.PLogger"
Option Explicit

Private Type TPLogger
    Name As String
    Level As LogLevels
    Handlers As Object
    Trace As Collection
End Type

Private this As TPLogger

Public Sub PushTrace(ByVal TraceInfo As String)
    If this.Trace Is Nothing Then Set this.Trace = New Collection
    this.Trace.Add TraceInfo
End Sub

Public Sub PopTrace()
    If this.Trace Is Nothing Then Exit Sub
    this.Trace.Remove this.Trace.Count
End Sub

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property

Public Sub SetLevel(ByVal Level As LogLevels)
    this.Level = Level
End Sub

Public Sub AddHandler(ByRef Handler As Handler)
    If this.Handlers Is Nothing Then
        Set this.Handlers = NewDictionary()
    End If
    Set this.Handlers(Handler.Name) = Handler
End Sub

Public Sub RemoveHandler(ByRef Handler As Handler)
    If this.Handlers Is Nothing Then
        Information.Err.Raise _
            Number:=9, _
            Source:=Information.TypeName(Me), _
            Description:="Handlers not initialized"
    End If

    If Not this.Handlers.Exists(Handler.Name) Then
        Information.Err.Raise _
            Number:=9, _
            Source:=Information.TypeName(Me), _
            Description:="Handler type not exists: " & Handler.Name
    End If

    this.Handlers.Remove Handler.Name
End Sub

Public Function IsEnabledFor(ByVal Level As LogLevels) As Boolean
    IsEnabledFor = Level >= this.Level
End Function

Public Sub Trace(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llTrace) Then
        Log_ LogLevels.llTrace, Strings.Join(Message, " ")
    End If
End Sub

Public Sub Debug_(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llDebug) Then
        Log_ LogLevels.llDebug, Strings.Join(Message, " ")
    End If
End Sub

Public Sub Info(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llInfo) Then
        Log_ LogLevels.llInfo, Strings.Join(Message, " ")
    End If
End Sub

Public Sub Warn(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llWarn) Then
        Log_ LogLevels.llWarn, Strings.Join(Message, " ")
    End If
End Sub

Public Sub Error(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llError) Then
        Log_ LogLevels.llError, Strings.Join(Message, " ")
    End If
End Sub

Public Sub Fatal(ParamArray Message() As Variant)
    If Me.IsEnabledFor(LogLevels.llFatal) Then
        Log_ LogLevels.llFatal, Strings.Join(Message, " ")
    End If
End Sub

Private Sub Log_(ByVal Level As LogLevels, ByVal Msg As String)
    Dim TraceInfo As String
    If Not this.Trace Is Nothing Then
        If this.Trace.Count > 0 Then
            TraceInfo = this.Trace(this.Trace.Count)
        End If
    End If

    Dim Record As Record
    Set Record = NewRecord(this.Name, Level, Msg, TraceInfo)

    Handle Record
End Sub

Private Sub Handle(ByRef Record As Record)
    If this.Handlers Is Nothing Then
        Me.AddHandler NewStreamHandler()
    End If

    Dim h As Variant
    For Each h In this.Handlers.Items()
        h.Emit Record
    Next
End Sub
