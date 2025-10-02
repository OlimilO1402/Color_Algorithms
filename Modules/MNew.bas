Attribute VB_Name = "MNew"
Option Explicit

Public Function ColorSelector(aTimer As Timer, aButton As CommandButton, aColorView As PictureBox, aLabel As Label) As ColorSelector
    Set ColorSelector = New ColorSelector: ColorSelector.New_ aTimer, aButton, aColorView, aLabel
End Function

Public Function AlphaPB(ForePB As PictureBox, BackPB As PictureBox) As AlphaPB
    Set AlphaPB = New AlphaPB: AlphaPB.New_ ForePB, BackPB
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function




