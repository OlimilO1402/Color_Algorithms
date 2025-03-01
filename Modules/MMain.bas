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
