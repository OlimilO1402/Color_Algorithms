Attribute VB_Name = "MInfoGColors"
Option Explicit

Private m_InfoGColors(0 To 255) As Long

Public Sub Init()
    Dim i As Byte
    'ReDim m_InfoGColors(0 To 255)
    m_InfoGColors(i) = RGB(&HFF, &HFF, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H0):    i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HBD, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HFF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &HFF):   i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HFF, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H7F, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HFF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HFF, &H7F):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H7F, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &H7F):  i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &H0, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &H7F, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &HE5, &HE5): i = i + 1
    m_InfoGColors(i) = RGB(&H26, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H26, &HA, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H26, &H19, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H26, &H26, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H19, &H26, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HA, &H26, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H26, &HA):   i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H26, &H19):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H26, &H26):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H19, &H26):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HA, &H26):   i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H26):   i = i + 1
    m_InfoGColors(i) = RGB(&HA, &H0, &H26):   i = i + 1
    m_InfoGColors(i) = RGB(&H19, &H0, &H26):  i = i + 1
    m_InfoGColors(i) = RGB(&H26, &H0, &H26):  i = i + 1
    m_InfoGColors(i) = RGB(&H26, &H0, &H19):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H17, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H33, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H4C, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H4C, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H17, &H4C, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H4C, &H17):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H4C, &H33):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H4C, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H33, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H17, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H4C):   i = i + 1
    m_InfoGColors(i) = RGB(&H17, &H0, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H0, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H0, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H0, &H33):  i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H24, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H4C, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H73, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H73, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H24, &H73, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H73, &H24):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H73, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H73, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H4C, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H24, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H73):   i = i + 1
    m_InfoGColors(i) = RGB(&H24, &H0, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H4C, &H0, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H0, &H73):  i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H0, &H4C):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H30, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H66, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H99, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H99, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H30, &H99, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H99, &H30):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H99, &H66):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H99, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H66, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H30, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H99):   i = i + 1
    m_InfoGColors(i) = RGB(&H30, &H0, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H0, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H0, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H0, &H66):  i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &H3D, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &H7F, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &HBF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HBF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H3D, &HBF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HBF, &H3D):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HBF, &H7F):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HBF, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H7F, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H3D, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &HBF):   i = i + 1
    m_InfoGColors(i) = RGB(&H3D, &H0, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &H0, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &H0, &HBF):  i = i + 1
    m_InfoGColors(i) = RGB(&HBF, &H0, &H7F):  i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &H4A, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &H99, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &HE5, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &HE5, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H4A, &HE5, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HE5, &H4A):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HE5, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HE5, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H99, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H4A, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &HE5):   i = i + 1
    m_InfoGColors(i) = RGB(&H4A, &H0, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H0, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &H0, &HE5):  i = i + 1
    m_InfoGColors(i) = RGB(&HE5, &H0, &H99):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &H0):   i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H54, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HAB, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HFF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&HAB, &HFF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H54, &HFF, &H0):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HFF, &H54):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HFF, &HAB):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HFF, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &HAB, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H54, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &HFF):   i = i + 1
    m_InfoGColors(i) = RGB(&H54, &H0, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HAB, &H0, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &HFF):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H0, &HAB):  i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H7F, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HA8, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HD4, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HFF, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HD4, &HFF, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&HA8, &HFF, &H7F): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HFF, &HA8): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HFF, &HD4): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HFF, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HD4, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &HA8, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&H7F, &H7F, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&HA8, &H7F, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&HD4, &H7F, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H7F, &HFF): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &H7F, &HD4): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &H6B, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &H8F, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &HB2, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &HD9, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &HD9, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&H8F, &HD9, &H6B): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &HD9, &H8F): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &HD9, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &HD9, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &HB2, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &H8F, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&H6B, &H6B, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&H8F, &H6B, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H6B, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &H6B, &HD9): i = i + 1
    m_InfoGColors(i) = RGB(&HD9, &H6B, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H59, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H75, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H94, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &HB2, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&H94, &HB2, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&H75, &HB2, &H59): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &HB2, &H75): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &HB2, &H94): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &HB2, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &H94, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &H75, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H59, &H59, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H75, &H59, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&H94, &H59, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H59, &HB2): i = i + 1
    m_InfoGColors(i) = RGB(&HB2, &H59, &H94): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H45, &H45): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H5C, &H44): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H73, &H45): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H8C, &H45): i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H8C, &H45): i = i + 1
    m_InfoGColors(i) = RGB(&H5C, &H8C, &H45): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H8C, &H5C): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H8C, &H73): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H8C, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H73, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H5C, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H45, &H45, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H5C, &H45, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H73, &H45, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H45, &H8C): i = i + 1
    m_InfoGColors(i) = RGB(&H8C, &H45, &H73): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H33, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H42, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H54, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H66, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H54, &H66, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H42, &H66, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H66, &H42): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H66, &H54): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H66, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H54, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H42, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H33, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H42, &H33, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H54, &H33, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H33, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H33, &H54): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H1F, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H29, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H33, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H40, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H40, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H29, &H40, &H1F): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H40, &H29): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H40, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H40, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H33, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H29, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H1F, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H29, &H1F, &H36): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H1F, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H1F, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&H40, &H1F, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H19, &HD, &HD):   i = i + 1
    m_InfoGColors(i) = RGB(&H19, &HF, &HD):   i = i + 1
    m_InfoGColors(i) = RGB(&H19, &HA, &HD):   i = i + 1
    m_InfoGColors(i) = RGB(&H19, &H19, &HD):  i = i + 1
    m_InfoGColors(i) = RGB(&H14, &H19, &HD):  i = i + 1
    m_InfoGColors(i) = RGB(&HF, &H19, &HD):   i = i + 1
    m_InfoGColors(i) = RGB(&HD, &H19, &HF):   i = i + 1
    m_InfoGColors(i) = RGB(&HD, &H19, &H14):  i = i + 1
    m_InfoGColors(i) = RGB(&HD, &H19, &H19):  i = i + 1
    m_InfoGColors(i) = RGB(&H1F, &H40, &H40): i = i + 1
    m_InfoGColors(i) = RGB(&HD, &HF, &H19):   i = i + 1
    m_InfoGColors(i) = RGB(&HD, &HD, &H19):   i = i + 1
    m_InfoGColors(i) = RGB(&HF, &HD, &H19):   i = i + 1
    m_InfoGColors(i) = RGB(&H14, &HD, &H19):  i = i + 1
    m_InfoGColors(i) = RGB(&H19, &HD, &H19):  i = i + 1
    m_InfoGColors(i) = RGB(&H19, &HD, &H14):  i = i + 1
    m_InfoGColors(i) = RGB(&H0, &H0, &H0):    i = i + 1
    m_InfoGColors(i) = RGB(&HF, &HF, &HF):    i = i + 1
    m_InfoGColors(i) = RGB(&H21, &H21, &H21): i = i + 1
    m_InfoGColors(i) = RGB(&H33, &H33, &H33): i = i + 1
    m_InfoGColors(i) = RGB(&H42, &H42, &H42): i = i + 1
    m_InfoGColors(i) = RGB(&H54, &H54, &H54): i = i + 1
    m_InfoGColors(i) = RGB(&H66, &H66, &H66): i = i + 1
    m_InfoGColors(i) = RGB(&H75, &H75, &H75): i = i + 1
    m_InfoGColors(i) = RGB(&H87, &H87, &H87): i = i + 1
    m_InfoGColors(i) = RGB(&H99, &H99, &H99): i = i + 1
    m_InfoGColors(i) = RGB(&HA8, &HA8, &HA8): i = i + 1
    m_InfoGColors(i) = RGB(&HBA, &HBA, &HBA): i = i + 1
    m_InfoGColors(i) = RGB(&HCC, &HCC, &HCC): i = i + 1
    m_InfoGColors(i) = RGB(&HDB, &HDB, &HDB): i = i + 1
    m_InfoGColors(i) = RGB(&HED, &HED, &HED): i = i + 1
    m_InfoGColors(i) = RGB(&HFF, &HFF, &HFF): i = i + 1
    
End Sub

Public Property Get Color(ByVal Index As Byte) As Long
    If m_InfoGColors(0) = 0 Then Init
    Color = m_InfoGColors(Index)
End Property

Public Property Get IndexFromColor(ByVal aColor As Long) As Integer
    Dim i As Integer, c As Long
    For i = 0 To 255
        Debug.Print i
        c = m_InfoGColors(i)
        If c = aColor Then
            IndexFromColor = i: Exit Property
        End If
    Next
    IndexFromColor = IndexOfClosestColorTo(aColor)
End Property

Public Property Get IndexOfClosestColorTo(ByVal aColor As Long) As Integer
    Dim i As Integer, i_minEd As Long, edi As Double
    Dim c As Long, lc As LngColor: lc = LngColor(aColor)
    Dim ed0 As Double: ed0 = LngColor_EuclidRMean(LngColor((&HFFFFFF And m_InfoGColors(0))), lc)
    For i = 0 To 255
        c = m_InfoGColors(i)
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
    ClosestColorTo = m_InfoGColors(i)
End Property

Public Property Get ColorArray() As Long()
    If m_InfoGColors(0) = 0 Then Init
    ColorArray = m_InfoGColors
End Property
