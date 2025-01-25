Attribute VB_Name = "QC"
'@Folder "PLoggerProject.src.Utils"
Option Explicit

' Quick constructors.
' Short hands for CreateObject("...")

Public Function NewDictionary() As Object
    Set NewDictionary = Interaction.CreateObject("Scripting.Dictionary")
End Function
