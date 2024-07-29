Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    MMath.Init
    MColor.Init
    MString.Init
    MMunsell.Init
    MMunsell.FilterChromaValues
    
    FMain.Show
End Sub

Public Function AlphaPB(ForePB As PictureBox, BackPB As PictureBox) As AlphaPB
    Set AlphaPB = New AlphaPB: AlphaPB.New_ ForePB, BackPB
End Function

