Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

