Attribute VB_Name = "MKnownColors"
Option Explicit

'https://en.wikipedia.org/wiki/X11_color_names

Public Enum X11KnownColor
    AliceBlue = &HFFFFF8F0          '
    AntiqueWhite = &HFFD7EBFA       '
    Aqua = &HFFFFFF00               ' aka: Cyan
    Aquamarine = &HFFD4FF7F         '
    Azure = &HFFFFFFF0              '
    Beige = &HFFDCF5F5              '
    Bisque = &HFFC4E4FF             '
    Black = &HFF000000              '
    BlanchedAlmond = &HFFCDEBFF     '
    Blue = &HFFFF0000               '
    BlueViolet = &HFFE22B8A         '
    Brown = &HFF2A2AA5              '
    Burlywood = &HFF87B8DE          '
    CadetBlue = &HFFA09E5F          '
    Chartreuse = &HFF00FF7F         '
    Chocolate = &HFF1E69D2          '
    Coral = &HFF507FFF              '
    CornflowerBlue = &HFFED9564     '
    Cornsilk = &HFFDCF8FF           '
    Crimson = &HFF3C14DC            '
    Cyan = &HFFFFFF00               ' aka: Aqua
    DarkBlue = &HFF8B0000           '
    DarkCyan = &HFF8B8B00           '
    DarkGoldenrod = &HFF0B86B8      '
    DarkGray = &HFFA9A9A9           ' aka: Dark Grey
    DarkGreen = &HFF006400          '
    DarkKhaki = &HFF6BB7BD          '
    DarkMagenta = &HFF8B008B        '
    DarkOliveGreen = &HFF2F6B55     '
    DarkOrange = &HFF008CFF         '
    DarkOrchid = &HFFCC3299         '
    DarkRed = &HFF00008B            '
    DarkSalmon = &HFF7A96E9         '
    DarkSeaGreen = &HFF8FBC8F       '
    DarkSlateBlue = &HFF8B3D48      '
    DarkSlateGray = &HFF4F4F2F      ' aka: Dark Slate Grey
    DarkTurquoise = &HFFD1CE00      '
    DarkViolet = &HFFD30094         '
    DeepPink = &HFF9314FF           '
    DeepSkyBlue = &HFFFFBF00        '
    DimGray = &HFF696969            ' aka: Dim Grey
    DodgerBlue = &HFFFF901E         '
    Firebrick = &HFF2222B2          '
    FloralWhite = &HFFF0FAFF        '
    ForestGreen = &HFF228B22        '
    Fuchsia = &HFFFF00FF            ' aka: Magenta
    Gainsboro = &HFFDCDCDC          '
    GhostWhite = &HFFFFF8F8         '
    Gold = &HFF00D7FF               '
    Goldenrod = &HFF20A5DA          '
    Gray = &HFFBEBEBE               ' aka: Grey, X11 Gray, X11 Grey
    WebGray = &HFF808080            ' aka: Web Grey
    Green = &HFF00FF00              ' aka: X11 Green, Lime
    WebGreen = &HFF008000           '
    GreenYellow = &HFF2FFFAD        '
    Honeydew = &HFFF0FFF0           '
    HotPink = &HFFB469FF            '
    IndianRed = &HFF5C5CCD          '
    Indigo = &HFF82004B             '
    Ivory = &HFFF0FFFF              '
    Khaki = &HFF8CE6F0              '
    Lavender = &HFFFAE6E6           '
    LavenderBlush = &HFFF5F0FF      '
    LawnGreen = &HFF00FC7C          '
    LemonChiffon = &HFFCDFAFF       '
    LightBlue = &HFFE6D8AD          '
    LightCoral = &HFF8080F0         '
    LightCyan = &HFFFFFFE0          '
    LightGoldenrod = &HFFD2FAFA     '
    LightGray = &HFFD3D3D3          ' aka: Light Grey
    LightGreen = &HFF90EE90         '
    LightPink = &HFFC1B6FF          '
    LightSalmon = &HFF7AA0FF        '
    LightSeaGreen = &HFFAAB220      '
    LightSkyBlue = &HFFFACE87       '
    LightSlateGray = &HFF998877     ' aka: Light Slate Grey
    LightSteelBlue = &HFFDEC4B0     '
    LightYellow = &HFFE0FFFF        '
    Lime = &HFF00FF00               '
    LimeGreen = &HFF32CD32          '
    Linen = &HFFE6F0FA              '
    Magenta = &HFFFF00FF            ' aka: Fuchsia
    Maroon = &HFF6030B0             ' aka: X11 Maroon
    WebMaroon = &HFF000080          '
    MediumAquamarine = &HFFAACD66   '
    MediumBlue = &HFFCD0000         '
    MediumOrchid = &HFFD355BA       '
    MediumPurple = &HFFDB7093       '
    MediumSeaGreen = &HFF71B33C     '
    MediumSlateBlue = &HFFEE687B    '
    MediumSpringGreen = &HFF9AFA00  '
    MediumTurquoise = &HFFCCD148    '
    MediumVioletRed = &HFF8515C7    '
    MidnightBlue = &HFF701919       '
    MintCream = &HFFFAFFF5          '
    MistyRose = &HFFE1E4FF          '
    Moccasin = &HFFB5E4FF           '
    NavajoWhite = &HFFADDEFF        '
    NavyBlue = &HFF800000           ' aka: Navy
    OldLace = &HFFE6F5FD            '
    Olive = &HFF008080              '
    OliveDrab = &HFF238E6B          '
    Orange = &HFF00A5FF             '
    OrangeRed = &HFF0045FF          '
    Orchid = &HFFD670DA             '
    PaleGoldenrod = &HFFAAE8EE      '
    PaleGreen = &HFF98FB98          '
    PaleTurquoise = &HFFEEEEAF      '
    PaleVioletRed = &HFF9370DB      '
    PapayaWhip = &HFFD5EFFF         '
    PeachPuff = &HFFB9DAFF          '
    Peru = &HFF3F85CD               '
    Pink = &HFFCBC0FF               '
    Plum = &HFFDDA0DD               '
    PowderBlue = &HFFE6E0B0         '
    Purple = &HFFF020A0             ' aka: X11 Purple
    WebPurple = &HFF800080          '
    RebeccaPurple = &HFF993366      '
    Red = &HFF0000FF                '
    RosyBrown = &HFF8F8FBC          '
    RoyalBlue = &HFFE16941          '
    SaddleBrown = &HFF13458B        '
    Salmon = &HFF7280FA             '
    SandyBrown = &HFF60A4F4         '
    SeaGreen = &HFF578B2E           '
    Seashell = &HFFEEF5FF           '
    Sienna = &HFF2D52A0             '
    Silver = &HFFC0C0C0             '
    SkyBlue = &HFFEBCE87            '
    SlateBlue = &HFFCD5A6A          '
    SlateGray = &HFF908070          ' aka: Slate Grey
    Snow = &HFFFAFAFF               '
    SpringGreen = &HFF7FFF00        '
    SteelBlue = &HFFB48246          '
    Tan = &HFF8CB4D2                '
    Teal = &HFF808000               '
    Thistle = &HFFD8BFD8            '
    Tomato = &HFF4763FF             '
    Turquoise = &HFFD0E040          '
    Violet = &HFFEE82EE             '
    Wheat = &HFFB3DEF5              '
    White = &HFFFFFFFF              '
    WhiteSmoke = &HFFF5F5F5         '
    Yellow = &HFF00FFFF             '
    YellowGreen = &HFF32CD9A        '
End Enum

Public Type TNamedColor
    Name   As String
    X11Col As X11KnownColor
    'Color  As ColorRGBA
End Type

Private m_Arr() As TNamedColor
Private m_Count As Long

Private Function TNamedColor(aName As String, ByVal X11Color As X11KnownColor) As TNamedColor
    With TNamedColor
        .Name = aName
        .X11Col = X11Color
        'Set .Color = MNew.ColorKC(X11Color)
    End With
End Function

Public Function X11KnownColor_Contains(ByVal aColor As Long) As Boolean
    If m_Count = 0 Then X11KnownColor_Init
    Dim i As Long
    For i = LBound(m_Arr) To m_Count - 1
        X11KnownColor_Contains = m_Arr(i).X11Col = aColor
        If X11KnownColor_Contains Then Exit Function
    Next
End Function

Public Function X11KnownColor_ClosestColorTo(ByVal aColor As Long) As TNamedColor
    Dim i As Long, i_minEd As Long, edi As Double
    Dim lc As LngColor: lc = LngColor(aColor)
    Dim ed0 As Double: ed0 = LngColor_EuclidRMean(LngColor((&HFFFFFF And m_Arr(0).X11Col)), lc)
    For i = 1 To m_Count - 1
        edi = LngColor_EuclidRMean(LngColor((&HFFFFFF And m_Arr(i).X11Col)), lc)
        If edi < ed0 Then
            i_minEd = i
            ed0 = edi
        End If
    Next
    X11KnownColor_ClosestColorTo = m_Arr(i_minEd)
End Function

'Public Property Get ColorByName(aName As String) As Color
'    If m_Count = 0 Then X11KnownColor_Init
'    Dim i As Long
'    For i = LBound(m_Arr) To UBound(m_Arr)
'        If m_Arr(i).Name = aName Then
'            Set ColorByName = m_Arr(i).Color
'            Exit Property
'        End If
'    Next
'End Property
Public Property Get ColorByName(aName As String) As Long
    If m_Count = 0 Then X11KnownColor_Init
    Dim i As Long
    For i = LBound(m_Arr) To m_Count - 1
        If m_Arr(i).Name = aName Then
            ColorByName = &HFFFFFF And m_Arr(i).X11Col
            Exit Property
        End If
    Next
End Property

Public Property Get NameFromColor(ByVal aColor As Long) As String
    Dim i As Long
    Dim c1 As Long: c1 = &HFFFFFF And aColor
    Dim c2 As Long
    For i = LBound(m_Arr) To m_Count - 1
        c2 = &HFFFFFF And m_Arr(i).X11Col
        If c1 = c2 Then
            NameFromColor = m_Arr(i).Name
            Exit Property
        End If
    Next
End Property
Public Sub X11KnownColor_ToCB(CB As ComboBox)
    If m_Count = 0 Then X11KnownColor_Init
    Dim i As Long
    CB.Clear
    For i = LBound(m_Arr) To m_Count - 1
        CB.AddItem m_Arr(i).Name
    Next
End Sub
'Public Function X11KnownColor_IsClose(toColor As Color) As Boolean
'    Dim i As Long
'    For i = LBound(m_Arr) To UBound(m_Arr)
'        X11KnownColor_IsClose = m_Arr(i).Color.IsClose(toColor)
'        If X11KnownColor_IsClose Then Exit Function
'    Next
'End Function
Public Sub X11KnownColor_Init()
    ReDim m_Arr(0 To 144): Dim i As Long
    m_Arr(i) = TNamedColor("AliceBlue", X11KnownColor.AliceBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("AntiqueWhite", X11KnownColor.AntiqueWhite):           i = i + 1
    m_Arr(i) = TNamedColor("Aqua", X11KnownColor.Aqua):                           i = i + 1
    m_Arr(i) = TNamedColor("Aquamarine", X11KnownColor.Aquamarine):               i = i + 1
    m_Arr(i) = TNamedColor("Azure", X11KnownColor.Azure):                         i = i + 1
    m_Arr(i) = TNamedColor("Beige", X11KnownColor.Beige):                         i = i + 1
    m_Arr(i) = TNamedColor("Bisque", X11KnownColor.Bisque):                       i = i + 1
    m_Arr(i) = TNamedColor("Black", X11KnownColor.Black):                         i = i + 1
    m_Arr(i) = TNamedColor("BlanchedAlmond", X11KnownColor.BlanchedAlmond):       i = i + 1
    m_Arr(i) = TNamedColor("Blue", X11KnownColor.Blue):                           i = i + 1
    m_Arr(i) = TNamedColor("BlueViolet", X11KnownColor.BlueViolet):               i = i + 1
    m_Arr(i) = TNamedColor("Brown", X11KnownColor.Brown):                         i = i + 1
    m_Arr(i) = TNamedColor("Burlywood", X11KnownColor.Burlywood):                 i = i + 1
    m_Arr(i) = TNamedColor("CadetBlue", X11KnownColor.CadetBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("Chartreuse", X11KnownColor.Chartreuse):               i = i + 1
    m_Arr(i) = TNamedColor("Chocolate", X11KnownColor.Chocolate):                 i = i + 1
    m_Arr(i) = TNamedColor("Coral", X11KnownColor.Coral):                         i = i + 1
    m_Arr(i) = TNamedColor("CornflowerBlue", X11KnownColor.CornflowerBlue):       i = i + 1
    m_Arr(i) = TNamedColor("Cornsilk", X11KnownColor.Cornsilk):                   i = i + 1
    m_Arr(i) = TNamedColor("Crimson", X11KnownColor.Crimson):                     i = i + 1
    m_Arr(i) = TNamedColor("Cyan", X11KnownColor.Cyan):                           i = i + 1
    m_Arr(i) = TNamedColor("DarkBlue", X11KnownColor.DarkBlue):                   i = i + 1
    m_Arr(i) = TNamedColor("DarkCyan", X11KnownColor.DarkCyan):                   i = i + 1
    m_Arr(i) = TNamedColor("DarkGoldenrod", X11KnownColor.DarkGoldenrod):         i = i + 1
    m_Arr(i) = TNamedColor("DarkGray", X11KnownColor.DarkGray):                   i = i + 1
    m_Arr(i) = TNamedColor("DarkGreen", X11KnownColor.DarkGreen):                 i = i + 1
    m_Arr(i) = TNamedColor("DarkKhaki", X11KnownColor.DarkKhaki):                 i = i + 1
    m_Arr(i) = TNamedColor("DarkMagenta", X11KnownColor.DarkMagenta):             i = i + 1
    m_Arr(i) = TNamedColor("DarkOliveGreen", X11KnownColor.DarkOliveGreen):       i = i + 1
    m_Arr(i) = TNamedColor("DarkOrange", X11KnownColor.DarkOrange):               i = i + 1
    m_Arr(i) = TNamedColor("DarkOrchid", X11KnownColor.DarkOrchid):               i = i + 1
    m_Arr(i) = TNamedColor("DarkRed", X11KnownColor.DarkRed):                     i = i + 1
    m_Arr(i) = TNamedColor("DarkSalmon", X11KnownColor.DarkSalmon):               i = i + 1
    m_Arr(i) = TNamedColor("DarkSeaGreen", X11KnownColor.DarkSeaGreen):           i = i + 1
    m_Arr(i) = TNamedColor("DarkSlateBlue", X11KnownColor.DarkSlateBlue):         i = i + 1
    m_Arr(i) = TNamedColor("DarkSlateGray", X11KnownColor.DarkSlateGray):         i = i + 1
    m_Arr(i) = TNamedColor("DarkTurquoise", X11KnownColor.DarkTurquoise):         i = i + 1
    m_Arr(i) = TNamedColor("DarkViolet", X11KnownColor.DarkViolet):               i = i + 1
    m_Arr(i) = TNamedColor("DeepPink", X11KnownColor.DeepPink):                   i = i + 1
    m_Arr(i) = TNamedColor("DeepSkyBlue", X11KnownColor.DeepSkyBlue):             i = i + 1
    m_Arr(i) = TNamedColor("DimGray", X11KnownColor.DimGray):                     i = i + 1
    m_Arr(i) = TNamedColor("DodgerBlue", X11KnownColor.DodgerBlue):               i = i + 1
    m_Arr(i) = TNamedColor("Firebrick", X11KnownColor.Firebrick):                 i = i + 1
    m_Arr(i) = TNamedColor("FloralWhite", X11KnownColor.FloralWhite):             i = i + 1
    m_Arr(i) = TNamedColor("ForestGreen", X11KnownColor.ForestGreen):             i = i + 1
    m_Arr(i) = TNamedColor("Fuchsia", X11KnownColor.Fuchsia):                     i = i + 1
    m_Arr(i) = TNamedColor("Gainsboro", X11KnownColor.Gainsboro):                 i = i + 1
    m_Arr(i) = TNamedColor("GhostWhite", X11KnownColor.GhostWhite):               i = i + 1
    m_Arr(i) = TNamedColor("Gold", X11KnownColor.Gold):                           i = i + 1
    m_Arr(i) = TNamedColor("Goldenrod", X11KnownColor.Goldenrod):                 i = i + 1
    m_Arr(i) = TNamedColor("Gray", X11KnownColor.Gray):                           i = i + 1
    m_Arr(i) = TNamedColor("WebGray", X11KnownColor.WebGray):                     i = i + 1
    m_Arr(i) = TNamedColor("Green", X11KnownColor.Green):                         i = i + 1
    m_Arr(i) = TNamedColor("WebGreen", X11KnownColor.WebGreen):                   i = i + 1
    m_Arr(i) = TNamedColor("GreenYellow", X11KnownColor.GreenYellow):             i = i + 1
    m_Arr(i) = TNamedColor("Honeydew", X11KnownColor.Honeydew):                   i = i + 1
    m_Arr(i) = TNamedColor("HotPink", X11KnownColor.HotPink):                     i = i + 1
    m_Arr(i) = TNamedColor("IndianRed", X11KnownColor.IndianRed):                 i = i + 1
    m_Arr(i) = TNamedColor("Indigo", X11KnownColor.Indigo):                       i = i + 1
    m_Arr(i) = TNamedColor("Ivory", X11KnownColor.Ivory):                         i = i + 1
    m_Arr(i) = TNamedColor("Khaki", X11KnownColor.Khaki):                         i = i + 1
    m_Arr(i) = TNamedColor("Lavender", X11KnownColor.Lavender):                   i = i + 1
    m_Arr(i) = TNamedColor("LavenderBlush", X11KnownColor.LavenderBlush):         i = i + 1
    m_Arr(i) = TNamedColor("LawnGreen", X11KnownColor.LawnGreen):                 i = i + 1
    m_Arr(i) = TNamedColor("LemonChiffon", X11KnownColor.LemonChiffon):           i = i + 1
    m_Arr(i) = TNamedColor("LightBlue", X11KnownColor.LightBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("LightCoral", X11KnownColor.LightCoral):               i = i + 1
    m_Arr(i) = TNamedColor("LightCyan", X11KnownColor.LightCyan):                 i = i + 1
    m_Arr(i) = TNamedColor("LightGoldenrod", X11KnownColor.LightGoldenrod):       i = i + 1
    m_Arr(i) = TNamedColor("LightGray", X11KnownColor.LightGray):                 i = i + 1
    m_Arr(i) = TNamedColor("LightGreen", X11KnownColor.LightGreen):               i = i + 1
    m_Arr(i) = TNamedColor("LightPink", X11KnownColor.LightPink):                 i = i + 1
    m_Arr(i) = TNamedColor("LightSalmon", X11KnownColor.LightSalmon):             i = i + 1
    m_Arr(i) = TNamedColor("LightSeaGreen", X11KnownColor.LightSeaGreen):         i = i + 1
    m_Arr(i) = TNamedColor("LightSkyBlue", X11KnownColor.LightSkyBlue):           i = i + 1
    m_Arr(i) = TNamedColor("LightSlateGray", X11KnownColor.LightSlateGray):       i = i + 1
    m_Arr(i) = TNamedColor("LightSteelBlue", X11KnownColor.LightSteelBlue):       i = i + 1
    m_Arr(i) = TNamedColor("LightYellow", X11KnownColor.LightYellow):             i = i + 1
    m_Arr(i) = TNamedColor("Lime", X11KnownColor.Lime):                           i = i + 1
    m_Arr(i) = TNamedColor("LimeGreen", X11KnownColor.LimeGreen):                 i = i + 1
    m_Arr(i) = TNamedColor("Linen", X11KnownColor.Linen):                         i = i + 1
    m_Arr(i) = TNamedColor("Magenta", X11KnownColor.Magenta):                     i = i + 1
    m_Arr(i) = TNamedColor("Maroon", X11KnownColor.Maroon):                       i = i + 1
    m_Arr(i) = TNamedColor("WebMaroon", X11KnownColor.WebMaroon):                 i = i + 1
    m_Arr(i) = TNamedColor("MediumAquamarine", X11KnownColor.MediumAquamarine):   i = i + 1
    m_Arr(i) = TNamedColor("MediumBlue", X11KnownColor.MediumBlue):               i = i + 1
    m_Arr(i) = TNamedColor("MediumOrchid", X11KnownColor.MediumOrchid):           i = i + 1
    m_Arr(i) = TNamedColor("MediumPurple", X11KnownColor.MediumPurple):           i = i + 1
    m_Arr(i) = TNamedColor("MediumSeaGreen", X11KnownColor.MediumSeaGreen):       i = i + 1
    m_Arr(i) = TNamedColor("MediumSlateBlue", X11KnownColor.MediumSlateBlue):     i = i + 1
    m_Arr(i) = TNamedColor("MediumSpringGreen", X11KnownColor.MediumSpringGreen): i = i + 1
    m_Arr(i) = TNamedColor("MediumTurquoise", X11KnownColor.MediumTurquoise):     i = i + 1
    m_Arr(i) = TNamedColor("MediumVioletRed", X11KnownColor.MediumVioletRed):     i = i + 1
    m_Arr(i) = TNamedColor("MidnightBlue", X11KnownColor.MidnightBlue):           i = i + 1
    m_Arr(i) = TNamedColor("MintCream", X11KnownColor.MintCream):                 i = i + 1
    m_Arr(i) = TNamedColor("MistyRose", X11KnownColor.MistyRose):                 i = i + 1
    m_Arr(i) = TNamedColor("Moccasin", X11KnownColor.Moccasin):                   i = i + 1
    m_Arr(i) = TNamedColor("NavajoWhite", X11KnownColor.NavajoWhite):             i = i + 1
    m_Arr(i) = TNamedColor("NavyBlue", X11KnownColor.NavyBlue):                   i = i + 1
    m_Arr(i) = TNamedColor("OldLace", X11KnownColor.OldLace):                     i = i + 1
    m_Arr(i) = TNamedColor("Olive", X11KnownColor.Olive):                         i = i + 1
    m_Arr(i) = TNamedColor("OliveDrab", X11KnownColor.OliveDrab):                 i = i + 1
    m_Arr(i) = TNamedColor("Orange", X11KnownColor.Orange):                       i = i + 1
    m_Arr(i) = TNamedColor("OrangeRed", X11KnownColor.OrangeRed):                 i = i + 1
    m_Arr(i) = TNamedColor("Orchid", X11KnownColor.Orchid):                       i = i + 1
    m_Arr(i) = TNamedColor("PaleGoldenrod", X11KnownColor.PaleGoldenrod):         i = i + 1
    m_Arr(i) = TNamedColor("PaleGreen", X11KnownColor.PaleGreen):                 i = i + 1
    m_Arr(i) = TNamedColor("PaleTurquoise", X11KnownColor.PaleTurquoise):         i = i + 1
    m_Arr(i) = TNamedColor("PaleVioletRed", X11KnownColor.PaleVioletRed):         i = i + 1
    m_Arr(i) = TNamedColor("PapayaWhip", X11KnownColor.PapayaWhip):               i = i + 1
    m_Arr(i) = TNamedColor("PeachPuff", X11KnownColor.PeachPuff):                 i = i + 1
    m_Arr(i) = TNamedColor("Peru", X11KnownColor.Peru):                           i = i + 1
    m_Arr(i) = TNamedColor("Pink", X11KnownColor.Pink):                           i = i + 1
    m_Arr(i) = TNamedColor("Plum", X11KnownColor.Plum):                           i = i + 1
    m_Arr(i) = TNamedColor("PowderBlue", X11KnownColor.PowderBlue):               i = i + 1
    m_Arr(i) = TNamedColor("Purple", X11KnownColor.Purple):                       i = i + 1
    m_Arr(i) = TNamedColor("WebPurple", X11KnownColor.WebPurple):                 i = i + 1
    m_Arr(i) = TNamedColor("RebeccaPurple", X11KnownColor.RebeccaPurple):         i = i + 1
    m_Arr(i) = TNamedColor("Red", X11KnownColor.Red):                             i = i + 1
    m_Arr(i) = TNamedColor("RosyBrown", X11KnownColor.RosyBrown):                 i = i + 1
    m_Arr(i) = TNamedColor("RoyalBlue", X11KnownColor.RoyalBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("SaddleBrown", X11KnownColor.SaddleBrown):             i = i + 1
    m_Arr(i) = TNamedColor("Salmon", X11KnownColor.Salmon):                       i = i + 1
    m_Arr(i) = TNamedColor("SandyBrown", X11KnownColor.SandyBrown):               i = i + 1
    m_Arr(i) = TNamedColor("SeaGreen", X11KnownColor.SeaGreen):                   i = i + 1
    m_Arr(i) = TNamedColor("Seashell", X11KnownColor.Seashell):                   i = i + 1
    m_Arr(i) = TNamedColor("Sienna", X11KnownColor.Sienna):                       i = i + 1
    m_Arr(i) = TNamedColor("Silver", X11KnownColor.Silver):                       i = i + 1
    m_Arr(i) = TNamedColor("SkyBlue", X11KnownColor.SkyBlue):                     i = i + 1
    m_Arr(i) = TNamedColor("SlateBlue", X11KnownColor.SlateBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("SlateGray", X11KnownColor.SlateGray):                 i = i + 1
    m_Arr(i) = TNamedColor("Snow", X11KnownColor.Snow):                           i = i + 1
    m_Arr(i) = TNamedColor("SpringGreen", X11KnownColor.SpringGreen):             i = i + 1
    m_Arr(i) = TNamedColor("SteelBlue", X11KnownColor.SteelBlue):                 i = i + 1
    m_Arr(i) = TNamedColor("Tan", X11KnownColor.Tan):                             i = i + 1
    m_Arr(i) = TNamedColor("Teal", X11KnownColor.Teal):                           i = i + 1
    m_Arr(i) = TNamedColor("Thistle", X11KnownColor.Thistle):                     i = i + 1
    m_Arr(i) = TNamedColor("Tomato", X11KnownColor.Tomato):                       i = i + 1
    m_Arr(i) = TNamedColor("Turquoise", X11KnownColor.Turquoise):                 i = i + 1
    m_Arr(i) = TNamedColor("Violet", X11KnownColor.Violet):                       i = i + 1
    m_Arr(i) = TNamedColor("Wheat", X11KnownColor.Wheat):                         i = i + 1
    m_Arr(i) = TNamedColor("White", X11KnownColor.White):                         i = i + 1
    m_Arr(i) = TNamedColor("WhiteSmoke", X11KnownColor.WhiteSmoke):               i = i + 1
    m_Arr(i) = TNamedColor("Yellow", X11KnownColor.Yellow):                       i = i + 1
    m_Arr(i) = TNamedColor("YellowGreen", X11KnownColor.YellowGreen):             i = i + 1
    m_Count = i
End Sub
' #################### '   ^^^  mit einem Array   ^^^   ' #################### '



' #################### '  vvv mit einer Collection vvv  ' #################### '
'Public Function X11KnownColor_Contains(ByVal c As Long) As Boolean
'    On Error Resume Next
'    If IsEmpty(m_Col(CStr(c))) Then: 'donothing
'    X11KnownColor_Contains = Err.Number = 0
'    On Error GoTo 0
'End Function

'Public Sub X11KnownColor_Init()
'    m_Col.Add X11KnownColor_Name(X11KnownColor.AliceBlue), CStr(X11KnownColor.AliceBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.AntiqueWhite), CStr(X11KnownColor.AntiqueWhite)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Aqua), CStr(X11KnownColor.Aqua)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Aquamarine), CStr(X11KnownColor.Aquamarine)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Azure), CStr(X11KnownColor.Azure)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Beige), CStr(X11KnownColor.Beige)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Bisque), CStr(X11KnownColor.Bisque)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Black), CStr(X11KnownColor.Black)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.BlanchedAlmond), CStr(X11KnownColor.BlanchedAlmond)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Blue), CStr(X11KnownColor.Blue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.BlueViolet), CStr(X11KnownColor.BlueViolet)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Brown), CStr(X11KnownColor.Brown)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Burlywood), CStr(X11KnownColor.Burlywood)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.CadetBlue), CStr(X11KnownColor.CadetBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Chartreuse), CStr(X11KnownColor.Chartreuse)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Chocolate), CStr(X11KnownColor.Chocolate)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Coral), CStr(X11KnownColor.Coral)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.CornflowerBlue), CStr(X11KnownColor.CornflowerBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Cornsilk), CStr(X11KnownColor.Cornsilk)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Crimson), CStr(X11KnownColor.Crimson)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Cyan), CStr(X11KnownColor.Cyan)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkBlue), CStr(X11KnownColor.DarkBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkCyan), CStr(X11KnownColor.DarkCyan)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkGoldenrod), CStr(X11KnownColor.DarkGoldenrod)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkGray), CStr(X11KnownColor.DarkGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkGreen), CStr(X11KnownColor.DarkGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkKhaki), CStr(X11KnownColor.DarkKhaki)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkMagenta), CStr(X11KnownColor.DarkMagenta)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkOliveGreen), CStr(X11KnownColor.DarkOliveGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkOrange), CStr(X11KnownColor.DarkOrange)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkOrchid), CStr(X11KnownColor.DarkOrchid)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkRed), CStr(X11KnownColor.DarkRed)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkSalmon), CStr(X11KnownColor.DarkSalmon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkSeaGreen), CStr(X11KnownColor.DarkSeaGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkSlateBlue), CStr(X11KnownColor.DarkSlateBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkSlateGray), CStr(X11KnownColor.DarkSlateGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkTurquoise), CStr(X11KnownColor.DarkTurquoise)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DarkViolet), CStr(X11KnownColor.DarkViolet)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DeepPink), CStr(X11KnownColor.DeepPink)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DeepSkyBlue), CStr(X11KnownColor.DeepSkyBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DimGray), CStr(X11KnownColor.DimGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.DodgerBlue), CStr(X11KnownColor.DodgerBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Firebrick), CStr(X11KnownColor.Firebrick)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.FloralWhite), CStr(X11KnownColor.FloralWhite)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.ForestGreen), CStr(X11KnownColor.ForestGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Fuchsia), CStr(X11KnownColor.Fuchsia)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Gainsboro), CStr(X11KnownColor.Gainsboro)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.GhostWhite), CStr(X11KnownColor.GhostWhite)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Gold), CStr(X11KnownColor.Gold)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Goldenrod), CStr(X11KnownColor.Goldenrod)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Gray), CStr(X11KnownColor.Gray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.WebGray), CStr(X11KnownColor.WebGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Green), CStr(X11KnownColor.Green)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.WebGreen), CStr(X11KnownColor.WebGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.GreenYellow), CStr(X11KnownColor.GreenYellow)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Honeydew), CStr(X11KnownColor.Honeydew)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.HotPink), CStr(X11KnownColor.HotPink)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.IndianRed), CStr(X11KnownColor.IndianRed)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Indigo), CStr(X11KnownColor.Indigo)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Ivory), CStr(X11KnownColor.Ivory)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Khaki), CStr(X11KnownColor.Khaki)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Lavender), CStr(X11KnownColor.Lavender)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LavenderBlush), CStr(X11KnownColor.LavenderBlush)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LawnGreen), CStr(X11KnownColor.LawnGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LemonChiffon), CStr(X11KnownColor.LemonChiffon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightBlue), CStr(X11KnownColor.LightBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightCoral), CStr(X11KnownColor.LightCoral)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightCyan), CStr(X11KnownColor.LightCyan)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightGoldenrod), CStr(X11KnownColor.LightGoldenrod)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightGray), CStr(X11KnownColor.LightGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightGreen), CStr(X11KnownColor.LightGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightPink), CStr(X11KnownColor.LightPink)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightSalmon), CStr(X11KnownColor.LightSalmon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightSeaGreen), CStr(X11KnownColor.LightSeaGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightSkyBlue), CStr(X11KnownColor.LightSkyBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightSlateGray), CStr(X11KnownColor.LightSlateGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightSteelBlue), CStr(X11KnownColor.LightSteelBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LightYellow), CStr(X11KnownColor.LightYellow)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Lime), CStr(X11KnownColor.Lime)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.LimeGreen), CStr(X11KnownColor.LimeGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Linen), CStr(X11KnownColor.Linen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Magenta), CStr(X11KnownColor.Magenta)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Maroon), CStr(X11KnownColor.Maroon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.WebMaroon), CStr(X11KnownColor.WebMaroon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumAquamarine), CStr(X11KnownColor.MediumAquamarine)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumBlue), CStr(X11KnownColor.MediumBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumOrchid), CStr(X11KnownColor.MediumOrchid)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumPurple), CStr(X11KnownColor.MediumPurple)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumSeaGreen), CStr(X11KnownColor.MediumSeaGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumSlateBlue), CStr(X11KnownColor.MediumSlateBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumSpringGreen), CStr(X11KnownColor.MediumSpringGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumTurquoise), CStr(X11KnownColor.MediumTurquoise)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MediumVioletRed), CStr(X11KnownColor.MediumVioletRed)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MidnightBlue), CStr(X11KnownColor.MidnightBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MintCream), CStr(X11KnownColor.MintCream)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.MistyRose), CStr(X11KnownColor.MistyRose)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Moccasin), CStr(X11KnownColor.Moccasin)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.NavajoWhite), CStr(X11KnownColor.NavajoWhite)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.NavyBlue), CStr(X11KnownColor.NavyBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.OldLace), CStr(X11KnownColor.OldLace)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Olive), CStr(X11KnownColor.Olive)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.OliveDrab), CStr(X11KnownColor.OliveDrab)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Orange), CStr(X11KnownColor.Orange)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.OrangeRed), CStr(X11KnownColor.OrangeRed)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Orchid), CStr(X11KnownColor.Orchid)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PaleGoldenrod), CStr(X11KnownColor.PaleGoldenrod)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PaleGreen), CStr(X11KnownColor.PaleGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PaleTurquoise), CStr(X11KnownColor.PaleTurquoise)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PaleVioletRed), CStr(X11KnownColor.PaleVioletRed)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PapayaWhip), CStr(X11KnownColor.PapayaWhip)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PeachPuff), CStr(X11KnownColor.PeachPuff)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Peru), CStr(X11KnownColor.Peru)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Pink), CStr(X11KnownColor.Pink)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Plum), CStr(X11KnownColor.Plum)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.PowderBlue), CStr(X11KnownColor.PowderBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Purple), CStr(X11KnownColor.Purple)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.WebPurple), CStr(X11KnownColor.WebPurple)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.RebeccaPurple), CStr(X11KnownColor.RebeccaPurple)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Red), CStr(X11KnownColor.Red)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.RosyBrown), CStr(X11KnownColor.RosyBrown)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.RoyalBlue), CStr(X11KnownColor.RoyalBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SaddleBrown), CStr(X11KnownColor.SaddleBrown)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Salmon), CStr(X11KnownColor.Salmon)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SandyBrown), CStr(X11KnownColor.SandyBrown)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SeaGreen), CStr(X11KnownColor.SeaGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Seashell), CStr(X11KnownColor.Seashell)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Sienna), CStr(X11KnownColor.Sienna)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Silver), CStr(X11KnownColor.Silver)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SkyBlue), CStr(X11KnownColor.SkyBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SlateBlue), CStr(X11KnownColor.SlateBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SlateGray), CStr(X11KnownColor.SlateGray)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Snow), CStr(X11KnownColor.Snow)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SpringGreen), CStr(X11KnownColor.SpringGreen)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.SteelBlue), CStr(X11KnownColor.SteelBlue)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Tan), CStr(X11KnownColor.Tan)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Teal), CStr(X11KnownColor.Teal)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Thistle), CStr(X11KnownColor.Thistle)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Tomato), CStr(X11KnownColor.Tomato)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Turquoise), CStr(X11KnownColor.Turquoise)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Violet), CStr(X11KnownColor.Violet)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Wheat), CStr(X11KnownColor.Wheat)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.White), CStr(X11KnownColor.White)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.WhiteSmoke), CStr(X11KnownColor.WhiteSmoke)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.Yellow), CStr(X11KnownColor.Yellow)
'    m_Col.Add X11KnownColor_Name(X11KnownColor.YellowGreen), CStr(X11KnownColor.YellowGreen)
'End Sub
'
'Public Property Get X11KnownColor_Name(c As Long) As String
'    Dim s As String
'    Select Case c
'    Case X11KnownColor.AliceBlue:         s = "AliceBlue"
'    Case X11KnownColor.AntiqueWhite:      s = "AntiqueWhite"
'    Case X11KnownColor.Aqua:              s = "Aqua"
'    Case X11KnownColor.Aquamarine:        s = "Aquamarine"
'    Case X11KnownColor.Azure:             s = "Azure"
'    Case X11KnownColor.Beige:             s = "Beige"
'    Case X11KnownColor.Bisque:            s = "Bisque"
'    Case X11KnownColor.Black:             s = "Black"
'    Case X11KnownColor.BlanchedAlmond:    s = "BlanchedAlmond"
'    Case X11KnownColor.Blue:              s = "Blue"
'    Case X11KnownColor.BlueViolet:        s = "BlueViolet"
'    Case X11KnownColor.Brown:             s = "Brown"
'    Case X11KnownColor.Burlywood:         s = "Burlywood"
'    Case X11KnownColor.CadetBlue:         s = "CadetBlue"
'    Case X11KnownColor.Chartreuse:        s = "Chartreuse"
'    Case X11KnownColor.Chocolate:         s = "Chocolate"
'    Case X11KnownColor.Coral:             s = "Coral"
'    Case X11KnownColor.CornflowerBlue:    s = "CornflowerBlue"
'    Case X11KnownColor.Cornsilk:          s = "Cornsilk"
'    Case X11KnownColor.Crimson:           s = "Crimson"
'    Case X11KnownColor.Cyan:              s = "Cyan"
'    Case X11KnownColor.DarkBlue:          s = "DarkBlue"
'    Case X11KnownColor.DarkCyan:          s = "DarkCyan"
'    Case X11KnownColor.DarkGoldenrod:     s = "DarkGoldenrod"
'    Case X11KnownColor.DarkGray:          s = "DarkGray"
'    Case X11KnownColor.DarkGreen:         s = "DarkGreen"
'    Case X11KnownColor.DarkKhaki:         s = "DarkKhaki"
'    Case X11KnownColor.DarkMagenta:       s = "DarkMagenta"
'    Case X11KnownColor.DarkOliveGreen:    s = "DarkOliveGreen"
'    Case X11KnownColor.DarkOrange:        s = "DarkOrange"
'    Case X11KnownColor.DarkOrchid:        s = "DarkOrchid"
'    Case X11KnownColor.DarkRed:           s = "DarkRed"
'    Case X11KnownColor.DarkSalmon:        s = "DarkSalmon"
'    Case X11KnownColor.DarkSeaGreen:      s = "DarkSeaGreen"
'    Case X11KnownColor.DarkSlateBlue:     s = "DarkSlateBlue"
'    Case X11KnownColor.DarkSlateGray:     s = "DarkSlateGray"
'    Case X11KnownColor.DarkTurquoise:     s = "DarkTurquoise"
'    Case X11KnownColor.DarkViolet:        s = "DarkViolet"
'    Case X11KnownColor.DeepPink:          s = "DeepPink"
'    Case X11KnownColor.DeepSkyBlue:       s = "DeepSkyBlue"
'    Case X11KnownColor.DimGray:           s = "DimGray"
'    Case X11KnownColor.DodgerBlue:        s = "DodgerBlue"
'    Case X11KnownColor.Firebrick:         s = "Firebrick"
'    Case X11KnownColor.FloralWhite:       s = "FloralWhite"
'    Case X11KnownColor.ForestGreen:       s = "ForestGreen"
'    Case X11KnownColor.Fuchsia:           s = "Fuchsia"
'    Case X11KnownColor.Gainsboro:         s = "Gainsboro"
'    Case X11KnownColor.GhostWhite:        s = "GhostWhite"
'    Case X11KnownColor.Gold:              s = "Gold"
'    Case X11KnownColor.Goldenrod:         s = "Goldenrod"
'    Case X11KnownColor.Gray:              s = "Gray"
'    Case X11KnownColor.WebGray:           s = "WebGray"
'    Case X11KnownColor.Green:             s = "Green"
'    Case X11KnownColor.WebGreen:          s = "WebGreen"
'    Case X11KnownColor.GreenYellow:       s = "GreenYellow"
'    Case X11KnownColor.Honeydew:          s = "Honeydew"
'    Case X11KnownColor.HotPink:           s = "HotPink"
'    Case X11KnownColor.IndianRed:         s = "IndianRed"
'    Case X11KnownColor.Indigo:            s = "Indigo"
'    Case X11KnownColor.Ivory:             s = "Ivory"
'    Case X11KnownColor.Khaki:             s = "Khaki"
'    Case X11KnownColor.Lavender:          s = "Lavender"
'    Case X11KnownColor.LavenderBlush:     s = "LavenderBlush"
'    Case X11KnownColor.LawnGreen:         s = "LawnGreen"
'    Case X11KnownColor.LemonChiffon:      s = "LemonChiffon"
'    Case X11KnownColor.LightBlue:         s = "LightBlue"
'    Case X11KnownColor.LightCoral:        s = "LightCoral"
'    Case X11KnownColor.LightCyan:         s = "LightCyan"
'    Case X11KnownColor.LightGoldenrod:    s = "LightGoldenrod"
'    Case X11KnownColor.LightGray:         s = "LightGray"
'    Case X11KnownColor.LightGreen:        s = "LightGreen"
'    Case X11KnownColor.LightPink:         s = "LightPink"
'    Case X11KnownColor.LightSalmon:       s = "LightSalmon"
'    Case X11KnownColor.LightSeaGreen:     s = "LightSeaGreen"
'    Case X11KnownColor.LightSkyBlue:      s = "LightSkyBlue"
'    Case X11KnownColor.LightSlateGray:    s = "LightSlateGray"
'    Case X11KnownColor.LightSteelBlue:    s = "LightSteelBlue"
'    Case X11KnownColor.LightYellow:       s = "LightYellow"
'    Case X11KnownColor.Lime:              s = "Lime"
'    Case X11KnownColor.LimeGreen:         s = "LimeGreen"
'    Case X11KnownColor.Linen:             s = "Linen"
'    Case X11KnownColor.Magenta:           s = "Magenta"
'    Case X11KnownColor.Maroon:            s = "Maroon"
'    Case X11KnownColor.WebMaroon:         s = "WebMaroon"
'    Case X11KnownColor.MediumAquamarine:  s = "MediumAquamarine"
'    Case X11KnownColor.MediumBlue:        s = "MediumBlue"
'    Case X11KnownColor.MediumOrchid:      s = "MediumOrchid"
'    Case X11KnownColor.MediumPurple:      s = "MediumPurple"
'    Case X11KnownColor.MediumSeaGreen:    s = "MediumSeaGreen"
'    Case X11KnownColor.MediumSlateBlue:   s = "MediumSlateBlue"
'    Case X11KnownColor.MediumSpringGreen: s = "MediumSpringGreen"
'    Case X11KnownColor.MediumTurquoise:   s = "MediumTurquoise"
'    Case X11KnownColor.MediumVioletRed:   s = "MediumVioletRed"
'    Case X11KnownColor.MidnightBlue:      s = "MidnightBlue"
'    Case X11KnownColor.MintCream:         s = "MintCream"
'    Case X11KnownColor.MistyRose:         s = "MistyRose"
'    Case X11KnownColor.Moccasin:          s = "Moccasin"
'    Case X11KnownColor.NavajoWhite:       s = "NavajoWhite"
'    Case X11KnownColor.NavyBlue:          s = "NavyBlue"
'    Case X11KnownColor.OldLace:           s = "OldLace"
'    Case X11KnownColor.Olive:             s = "Olive"
'    Case X11KnownColor.OliveDrab:         s = "OliveDrab"
'    Case X11KnownColor.Orange:            s = "Orange"
'    Case X11KnownColor.OrangeRed:         s = "OrangeRed"
'    Case X11KnownColor.Orchid:            s = "Orchid"
'    Case X11KnownColor.PaleGoldenrod:     s = "PaleGoldenrod"
'    Case X11KnownColor.PaleGreen:         s = "PaleGreen"
'    Case X11KnownColor.PaleTurquoise:     s = "PaleTurquoise"
'    Case X11KnownColor.PaleVioletRed:     s = "PaleVioletRed"
'    Case X11KnownColor.PapayaWhip:        s = "PapayaWhip"
'    Case X11KnownColor.PeachPuff:         s = "PeachPuff"
'    Case X11KnownColor.Peru:              s = "Peru"
'    Case X11KnownColor.Pink:              s = "Pink"
'    Case X11KnownColor.Plum:              s = "Plum"
'    Case X11KnownColor.PowderBlue:        s = "PowderBlue"
'    Case X11KnownColor.Purple:            s = "Purple"
'    Case X11KnownColor.WebPurple:         s = "WebPurple"
'    Case X11KnownColor.RebeccaPurple:     s = "RebeccaPurple"
'    Case X11KnownColor.Red:               s = "Red"
'    Case X11KnownColor.RosyBrown:         s = "RosyBrown"
'    Case X11KnownColor.RoyalBlue:         s = "RoyalBlue"
'    Case X11KnownColor.SaddleBrown:       s = "SaddleBrown"
'    Case X11KnownColor.Salmon:            s = "Salmon"
'    Case X11KnownColor.SandyBrown:        s = "SandyBrown"
'    Case X11KnownColor.SeaGreen:          s = "SeaGreen"
'    Case X11KnownColor.Seashell:          s = "Seashell"
'    Case X11KnownColor.Sienna:            s = "Sienna"
'    Case X11KnownColor.Silver:            s = "Silver"
'    Case X11KnownColor.SkyBlue:           s = "SkyBlue"
'    Case X11KnownColor.SlateBlue:         s = "SlateBlue"
'    Case X11KnownColor.SlateGray:         s = "SlateGray"
'    Case X11KnownColor.Snow:              s = "Snow"
'    Case X11KnownColor.SpringGreen:       s = "SpringGreen"
'    Case X11KnownColor.SteelBlue:         s = "SteelBlue"
'    Case X11KnownColor.Tan:               s = "Tan"
'    Case X11KnownColor.Teal:              s = "Teal"
'    Case X11KnownColor.Thistle:           s = "Thistle"
'    Case X11KnownColor.Tomato:            s = "Tomato"
'    Case X11KnownColor.Turquoise:         s = "Turquoise"
'    Case X11KnownColor.Violet:            s = "Violet"
'    Case X11KnownColor.Wheat:             s = "Wheat"
'    Case X11KnownColor.White:             s = "White"
'    Case X11KnownColor.WhiteSmoke:        s = "WhiteSmoke"
'    Case X11KnownColor.Yellow:            s = "Yellow"
'    Case X11KnownColor.YellowGreen:       s = "YellowGreen"
'    End Select
'    X11KnownColor_Name = s
'End Property
    
