VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Private WithEvents mTimer  As VB.Timer
Attribute mTimer.VB_VarHelpID = -1
Private WithEvents mButton As VB.CommandButton
Attribute mButton.VB_VarHelpID = -1
Private m_PBColorView As VB.PictureBox
Private m_LblCurColor As VB.Label
Private m_CurMousePos As POINTAPI
Private m_CurColor    As Long

Public Event ColorSelected(Color As Long)

Friend Sub New_(aTimer As VB.Timer, aButton As CommandButton, aColorView As PictureBox, aLabel As Label)
    Set mTimer = aTimer: Set mButton = aButton: Set m_PBColorView = aColorView: Set m_LblCurColor = aLabel
    mTimer.Enabled = False: mTimer.Interval = 50
    mButton.Caption = "Pick a Color"
End Sub
'copy to MNew:
'Public Function ColorSelector(aTimer As Timer, aButton As CommandButton, aColorView As PictureBox, aLabel As Label) As ColorSelector
'    Set ColorSelector = New ColorSelector: ColorSelector.New_ aTimer, aButton, aColorView, aLabel
'End Function

Public Property Get ColorSelected() As Long
    ColorSelected = m_CurColor
End Property

Private Sub mButton_Click()
    mTimer.Enabled = Not mTimer.Enabled
    mButton.Caption = IIf(mTimer.Enabled, "Press Enter", "Pick a Color")
    If Not mTimer.Enabled Then
        RaiseEvent ColorSelected(m_CurColor)
    End If
End Sub

Private Sub mTimer_Timer()
    GetCursorPos m_CurMousePos
    Dim Color As Long: Color = ColorUnderMouse(m_CurMousePos.X, m_CurMousePos.Y)
    If Not m_PBColorView Is Nothing Then m_PBColorView.BackColor = Color
    If Not m_LblCurColor Is Nothing Then m_LblCurColor.Caption = MouseCoordsNColor_ToStr(m_CurMousePos.X, m_CurMousePos.Y, Color)
End Sub

Private Function ColorUnderMouse(ByVal X As Long, ByVal Y As Long) As Long
    ColorUnderMouse = GetPixel(GetDC(0), X, Y)
End Function

Private Function MouseCoordsNColor_ToStr(ByVal X As Single, ByVal Y As Single, ByVal Color As Long) As String
    MouseCoordsNColor_ToStr = "X;Y: " & X & ";" & Y & vbCrLf & Color_ToStr(Color)
End Function

Private Function Color_ToStr(ByVal this As Long) As String
    Dim r As Long: r = (this And &HFF&)
    Dim G As Long: G = (this And &HFF00&) \ &H100&
    Dim b As Long: b = (this And &HFF0000) \ &H10000
    Dim hexprefix As String: hexprefix = "&&H"
    Dim sr As String: sr = CStr(r): sr = Space$(3 - Len(sr)) & sr
    Dim sG As String: sG = CStr(G): sG = Space$(3 - Len(sG)) & sG
    Dim sB As String: sB = CStr(b): sB = Space$(3 - Len(sB)) & sB
    Color_ToStr = "R=" & sr & " (" & hexprefix & MString.Hex2(CByte(r)) & ")" & vbCrLf & _
                  "G=" & sG & " (" & hexprefix & MString.Hex2(CByte(G)) & ")" & vbCrLf & _
                  "B=" & sB & " (" & hexprefix & MString.Hex2(CByte(b)) & ")"
End Function

