Attribute VB_Name = "MGGUColors"
Option Explicit

Private m_GGUColors(0 To 24) As Long

Public Sub Init()
    Dim i As Integer
    m_GGUColors(i) = RGB(&HFF, &HFF, &H84): i = i + 1
    m_GGUColors(i) = RGB(&HFF, &HC0, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HFF, &HC5, &H59): i = i + 1
    m_GGUColors(i) = RGB(&HAC, &HAC, &H51): i = i + 1
    m_GGUColors(i) = RGB(&HC2, &H8C, &HFF): i = i + 1
    m_GGUColors(i) = RGB(&HC8, &HC8, &H8C): i = i + 1
    m_GGUColors(i) = RGB(&HF8, &HC8, &H65): i = i + 1
    m_GGUColors(i) = RGB(&HE1, &HFF, &HC8): i = i + 1
    m_GGUColors(i) = RGB(&HC8, &HC8, &HB4): i = i + 1
    m_GGUColors(i) = RGB(&HF4, &HEA, &HFF): i = i + 1
    m_GGUColors(i) = RGB(&HFF, &HD0, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HFF, &HE5, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HFF, &HFA, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HEF, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HDB, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HC6, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HB1, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H9C, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H87, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H72, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H5E, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H49, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H34, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&H1F, &HFF, &H0):  i = i + 1
    m_GGUColors(i) = RGB(&HB, &HFF, &H0)
End Sub

Public Property Get Color(ByVal Index As Byte) As Long
    If m_GGUColors(0) = 0 Then Init
    Color = m_GGUColors(Index)
End Property

Public Property Get IndexFromColor(ByVal aColor As Long) As Integer
    Dim i As Integer, c As Long
    For i = 0 To 24
        'Debug.Print i
        c = m_GGUColors(i)
        If c = aColor Then
            IndexFromColor = i: Exit Property
        End If
    Next
    IndexFromColor = IndexOfClosestColorTo(aColor)
End Property

Public Property Get IndexOfClosestColorTo(ByVal aColor As Long) As Integer
    Dim i As Integer, i_minEd As Long, edi As Double
    Dim c As Long, lc As LngColor: lc = LngColor(aColor)
    Dim ed0 As Double: ed0 = LngColor_EuclidRMean(LngColor((&HFFFFFF And m_GGUColors(0))), lc)
    For i = 0 To 24
        c = m_GGUColors(i)
        edi = MColor.LngColor_EuclidRMean(LngColor((&HFFFFFF And c)), lc)
        If edi < ed0 Then
            i_minEd = i
            ed0 = edi
        End If
    Next
    IndexOfClosestColorTo = i_minEd
End Property

Public Property Get ClosestColorTo(ByVal aColor As Long) As Long
    Dim i As Byte: i = IndexOfClosestColorTo(aColor)
    ClosestColorTo = m_GGUColors(i)
End Property

Public Property Get ColorArray() As Long()
    If m_GGUColors(0) = 0 Then Init
    ColorArray = m_GGUColors
End Property

