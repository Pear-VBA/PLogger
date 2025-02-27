VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PLoggerProject.src.Formatter"
Option Explicit

Private Type TFormatter
    Fmt As String
End Type

Private this As TFormatter

Public Property Get Fmt() As String
    Fmt = this.Fmt
End Property

Public Property Let Fmt(ByVal RHS As String)
    this.Fmt = RHS
End Property

Public Function Format(ByRef Record As Record) As String
    Dim Compare As VbCompareMethod
    Compare = VbCompareMethod.vbTextCompare

    Dim s As String
    s = this.Fmt
    s = Strings.Replace(s, "{name}", Record.Name, Compare:=Compare)
    s = Strings.Replace(s, "{levelno}", Record.Level, Compare:=Compare)
    s = Strings.Replace(s, "{levelname}", LevelToString(Record.Level), Compare:=Compare)
    s = Strings.Replace(s, "{message}", Record.Message, Compare:=Compare)
    s = Strings.Replace(s, "{traceinfo}", Record.TraceInfo, Compare:=Compare)

    Format = s
End Function
