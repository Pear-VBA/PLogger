VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "PLoggerProject.src.Handlers"
Option Explicit

Public Property Get Name() As String
End Property

Public Sub Emit(ByRef Record As Object)
End Sub

Public Sub SetFormatter(ByRef Formatter As Formatter)
End Sub
