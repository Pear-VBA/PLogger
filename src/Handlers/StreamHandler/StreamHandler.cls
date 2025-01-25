VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StreamHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PLoggerProject.src.Handlers.StreamHandler"
Option Explicit

Implements Handler

Private Type TStreamHandler
    Level As Long
    Formatter As Formatter
End Type

Private this As TStreamHandler

Private Property Get Handler_Name() As String
    Handler_Name = Me.Name
End Property

Private Sub Handler_Emit(ByRef Record As Object)
    Me.Emit Record
End Sub

Private Sub Handler_SetFormatter(ByRef Formatter As Formatter)
    Me.SetFormatter Formatter
End Sub

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

    Dim Msg As String
    Msg = this.Formatter.Format(Record)
    Debug.Print Msg
End Sub
