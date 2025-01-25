Attribute VB_Name = "RecordCstr"
'@Folder "PLoggerProject.src.Record"
Option Explicit

Public Function NewRecord( _
    ByVal Name As String, _
    ByVal Level As String, _
    ByVal Message As String, _
    ByVal TraceInfo As String _
) As Record
    Set NewRecord = New Record
    NewRecord.Time = DateTime.Now
    NewRecord.Name = Name
    NewRecord.Level = Level
    NewRecord.Message = Message
    NewRecord.TraceInfo = TraceInfo
End Function
