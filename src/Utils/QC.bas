Attribute VB_Name = "QC"
Attribute VB_Description = "Quick constructors. Short hands for CreateObject(""..."")"
'@ModuleDescription "Quick constructors. Short hands for CreateObject(""..."")"
'@Folder "PLoggerProject.src.Utils"
Option Explicit

Public Function NewDictionary() As Object
    Set NewDictionary = Interaction.CreateObject("Scripting.Dictionary")
End Function

Public Function NewStream( _
    ByVal StreamType As Long, _
    ByVal Charset As String _
) As Object
    Set NewStream = Interaction.CreateObject("ADODB.Stream")
    NewStream.Type = StreamType
    NewStream.Charset = Charset
End Function
