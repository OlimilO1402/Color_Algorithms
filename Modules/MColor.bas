Attribute VB_Name = "MColor"
Option Explicit 'lines of code: 2024-08-14: 1286; 2025-03-01: 1404;
'VB Type-Kürzel
' $ = String
' % = Integer
' & = Long
' ! = Single
' # = Double
' @ = Decimal in a Variant
'
'Dim VBStr$
'Dim VBInt%
'Dim VBLng&
'Dim VBSng!
'Dim VBDbl#
'Dim VBDec@


'Public Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
'Public Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long

Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue_out As Integer, ByRef pwLuminance_out As Integer, ByRef pwSaturation_out As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long

Public Type ColorValueRange
    RangeMin   As Single
    RangeMax   As Single
    RangeSteps As Long
    FormatStr  As String
End Type

Public Type LngColor
    Value As Long
End Type

Public Type RGBA
    r As Byte '0..255
    G As Byte '0..255
    b As Byte '0..255
    a As Byte '0..255
End Type

Private m_CVR_RGBA_R As ColorValueRange
Private m_CVR_RGBA_G As ColorValueRange
Private m_CVR_RGBA_B As ColorValueRange
Private m_CVR_RGBA_A As ColorValueRange

Public Type ARGB
    a As Byte '0..255
    r As Byte '0..255
    G As Byte '0..255
    b As Byte '0..255
End Type

Public Type RGBAf
    r As Single '0..1
    G As Single '0..1
    b As Single '0..1
    a As Single '0..1
End Type

Private m_CVR_RGBAf_R As ColorValueRange
Private m_CVR_RGBAf_G As ColorValueRange
Private m_CVR_RGBAf_B As ColorValueRange
Private m_CVR_RGBAf_A As ColorValueRange

Public Type CMYK
    c As Single '0..1
    M As Single '0..1
    Y As Single '0..1
    k As Single '0..1
    a As Single '0..1
End Type

Private m_CVR_CMYK_C As ColorValueRange
Private m_CVR_CMYK_M As ColorValueRange
Private m_CVR_CMYK_Y As ColorValueRange
Private m_CVR_CMYK_K As ColorValueRange
Private m_CVR_CMYK_A As ColorValueRange

Public Type HSLA
    H As Byte '0..239
    s As Byte '0..240
    L As Byte '0..240
    a As Byte '0..255
End Type

Private m_CVR_HSLA_H As ColorValueRange
Private m_CVR_HSLA_S As ColorValueRange
Private m_CVR_HSLA_L As ColorValueRange
Private m_CVR_HSLA_A As ColorValueRange

'Public Type HSLA
'    H As Byte '0..255
'    S As Byte '0..255
'    L As Byte '0..255
'    A As Byte '0..255
'End Type

Public Type HSLAf
    H As Single '0..1
    s As Single '0..1
    L As Single '0..1
    a As Single '0..1
End Type

Private m_CVR_HSLAf_H As ColorValueRange
Private m_CVR_HSLAf_S As ColorValueRange
Private m_CVR_HSLAf_L As ColorValueRange
Private m_CVR_HSLAf_A As ColorValueRange

Public Type HSV
    H As Single '0..1
    s As Single '0..1
    v As Single '0..1
    a As Single '0..1
End Type

Private m_CVR_HSV_H As ColorValueRange
Private m_CVR_HSV_S As ColorValueRange
Private m_CVR_HSV_V As ColorValueRange
Private m_CVR_HSV_A As ColorValueRange

'Public Type HSB
'    H As Single '0..1
'    S As Single '0..1
'    B As Single '0..1
'    A As Single '0..1
'End Type

Public Type XYZ
    x As Single '0..1
    Y As Single '0..1
    z As Single '0..1
    a As Single '0..1
End Type

Private m_CVR_XYZ_X As ColorValueRange
Private m_CVR_XYZ_Y As ColorValueRange
Private m_CVR_XYZ_Z As ColorValueRange
Private m_CVR_XYZ_A As ColorValueRange

'https://de.wikipedia.org/wiki/YCbCr-Farbmodell

Public Type YCbCr '(CCIR 601-256 levels)
    Y  As Single  ' E'Y  =  0..1     ' Grundhelligkeit Y
    cb As Single  ' E'Cb = -0.5..0.5 ' Blue-Yellow Chrominance
    Cr As Single  ' E'Cr = -0.5..0.5 ' Red-Green Chrominance
    a  As Single  ' 0..1
End Type

Private m_CVR_YCbCr_Y  As ColorValueRange
Private m_CVR_YCbCr_Cb As ColorValueRange
Private m_CVR_YCbCr_Cr As ColorValueRange
Private m_CVR_YCbCr_A  As ColorValueRange

Public Type CIELab
    L  As Single ' 0 .. 100
    aa As Single '-128 .. 128
    bb As Single '-128 .. 128
    a  As Single  '0..1
End Type

Private m_CVR_CIELab_L  As ColorValueRange
Private m_CVR_CIELab_aa As ColorValueRange
Private m_CVR_CIELab_bb As ColorValueRange
Private m_CVR_CIELab_A  As ColorValueRange

Public Enum CIELabLight
    D50_2 = 0
    D65_2 = 1
    D50_10 = 2
    D65_10 = 3
End Enum

Private Type XYZMatrix
    x(0 To 2) As Single
    Y(0 To 2) As Single
    z(0 To 2) As Single
    r(0 To 2) As Single
    G(0 To 2) As Single
    b(0 To 2) As Single
End Type

Private M As XYZMatrix
Private CIELabLights(0 To 3) As XYZ

Public Sub Init()
    'https://github.com/PitPik/colorPicker/blob/master/colors.js
    '// Observer = 2° (CIE 1931), Illuminant = D65
    With M
        .x(0) = 0.4124564: .x(1) = 0.3575761:  .x(2) = 0.1804375
        .Y(0) = 0.2126729: .Y(1) = 0.7151522:  .Y(2) = 0.072175
        .z(0) = 0.0193339: .z(1) = 0.119192:   .z(2) = 0.9503041
        
'
'    rgb[0] = (xyz[0] * 3.240479f) + (xyz[1] * -1.537150f) + (xyz[2] * -.498535f);
'    rgb[1] = (xyz[0] * -.969256f) + (xyz[1] *  1.875992f) + (xyz[2] * .041556f);
'    rgb[2] = (xyz[0] * .055648f)  + (xyz[1] * -.204043f) + (xyz[2] * 1.057311f);
        
        .r(0) = 3.240479:  .r(1) = -1.53715:  .r(2) = -0.498535
        .G(0) = -0.969256: .G(1) = 1.8760108: .G(2) = 0.041556
        .b(0) = 0.055648:  .b(1) = -0.204043: .b(2) = 1.057311
'
'        .R(0) = 3.2404542: .R(1) = -1.5371385: .R(2) = -0.4985314
'        .G(0) = -0.969266: .G(1) = 1.8760108:  .G(2) = 0.041556
'        .B(0) = 0.0556434: .B(1) = -0.2040259: .B(2) = 1.0572252
    End With
'    CIELabLights(CIELabLight.D50_2) = XYZ(96.422, 100, 82.521)
'    CIELabLights(CIELabLight.D65_2) = XYZ(95.047, 100, 108.883)
'    CIELabLights(CIELabLight.D50_10) = XYZ(96.72, 100, 81.427)
'    CIELabLights(CIELabLight.D65_10) = XYZ(94.811, 100, 107.304)
    
    CIELabLights(CIELabLight.D50_2) = XYZ(0.96422, 1, 0.82521)
    CIELabLights(CIELabLight.D65_2) = XYZ(0.95047, 1, 1.08883)
    CIELabLights(CIELabLight.D50_10) = XYZ(0.9672, 1, 0.81427)
    CIELabLights(CIELabLight.D65_10) = XYZ(0.94811, 1, 1.07304)
    
    'init the ValueRanges
    InitValueRanges
End Sub

'https://web.archive.org/web/20111111073514/http://www.easyrgb.com/index.php?X=MATH&H=08#text8

Private Sub InitValueRanges()
    m_CVR_RGBA_R = CVR_RGBA_R:      m_CVR_RGBA_G = CVR_RGBA_G:        m_CVR_RGBA_B = CVR_RGBA_B:       m_CVR_RGBA_A = CVR_RGBA_A
    m_CVR_RGBAf_R = CVR_RGBAf_R:    m_CVR_RGBAf_G = CVR_RGBAf_G:      m_CVR_RGBAf_B = CVR_RGBAf_B:     m_CVR_RGBAf_A = CVR_RGBAf_A
    m_CVR_CMYK_C = CVR_CMYK_C:      m_CVR_CMYK_M = CVR_CMYK_M:        m_CVR_CMYK_Y = CVR_CMYK_Y:       m_CVR_CMYK_K = CVR_CMYK_K:    m_CVR_CMYK_A = CVR_CMYK_A
    m_CVR_HSLA_H = CVR_HSL_H:       m_CVR_HSLA_S = CVR_HSL_S:         m_CVR_HSLA_L = CVR_HSL_L:        m_CVR_HSLA_A = CVR_HSL_A
    m_CVR_HSLAf_H = CVR_HSLAf_H:    m_CVR_HSLAf_S = CVR_HSLAf_S:      m_CVR_HSLAf_L = CVR_HSLAf_L:     m_CVR_HSLAf_A = CVR_HSLAf_A
    m_CVR_HSV_H = CVR_HSV_H:        m_CVR_HSV_S = CVR_HSV_S:          m_CVR_HSV_V = CVR_HSV_V:         m_CVR_HSV_A = CVR_HSV_A
    m_CVR_XYZ_X = CVR_XYZ_X:        m_CVR_XYZ_Y = CVR_XYZ_Y:          m_CVR_XYZ_Z = CVR_XYZ_Z:         m_CVR_XYZ_A = CVR_XYZ_A
    m_CVR_CIELab_L = CVR_CIELab_L:  m_CVR_CIELab_aa = CVR_CIELab_aa:  m_CVR_CIELab_bb = CVR_CIELab_bb: m_CVR_CIELab_A = CVR_CIELab_A
    m_CVR_YCbCr_Y = CVR_YCbCr_Y:    m_CVR_YCbCr_Cb = CVR_YCbCr_Cb:    m_CVR_YCbCr_Cr = CVR_YCbCr_Cr:   m_CVR_YCbCr_A = CVR_YCbCr_A
End Sub


' #################### ' LngColor ' #################### '
Public Function LngColor(ByVal aColor As Long) As LngColor
    LngColor.Value = aColor
End Function

Public Function LngColor_EuclidRMean(this As LngColor, other As LngColor) As Double
    LngColor_EuclidRMean = RGBA_EuclidRMean(LngColor_ToRGBA(this), LngColor_ToRGBA(other))
End Function

Public Function LngColor_ToRGBA(this As LngColor) As RGBA
    LSet LngColor_ToRGBA = this
End Function
Public Function LngColor_ToRGBAf(this As LngColor) As RGBAf
    LngColor_ToRGBAf = RGBA_ToRGBAf(LngColor_ToRGBA(this))
End Function
Public Function LngColor_ToCMYK(this As LngColor) As CMYK
    LngColor_ToCMYK = RGBAf_ToCMYK(LngColor_ToRGBAf(this))
End Function

Public Function LngColor_ToWebHex(this As LngColor) As String
    LngColor_ToWebHex = RGBA_ToWebHex(LngColor_ToRGBA(this))
End Function
Public Function LngColor_ParseWebHex(HashtagColor As String) As LngColor
    LngColor_ParseWebHex = RGBA_ToLngColor(RGBA_ParseWebHex(HashtagColor))
End Function
Public Function LngColor_Read(TB As TextBox) As LngColor
    LngColor_Read = LngColor_ParseWebHex(TB.Text)
End Function
Public Function LngColor_Write(this As LngColor, aTB_out As TextBox)
    aTB_out.Text = LngColor_ToWebHex(this)
End Function

' #################### ' RGBA ' #################### '
'Public Type RGBA
'    R As Byte '0..255
'    G As Byte '0..255
'    B As Byte '0..255
'    A As Byte '0..255
'End Type
Public Function RGBA(ByVal r As Byte, ByVal G As Byte, ByVal b As Byte, ByVal a As Byte) As RGBA
    With RGBA: .r = r: .G = G: .b = b: .a = a: End With
End Function
Public Function RGBA_Read(this_out As RGBA, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Byte, s As String
    With this_out
        s = TB_R.Text: If Byte_TryParse(s, v) Then .r = v Else err_out = s: Exit Function
        s = TB_G.Text: If Byte_TryParse(s, v) Then .G = v Else err_out = s: Exit Function
        s = TB_B.Text: If Byte_TryParse(s, v) Then .b = v Else err_out = s: Exit Function
        s = TB_A.Text: If Byte_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    RGBA_Read = True
End Function
Public Function RGBA_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBA)
    With this: TB_R.Text = .r: TB_G.Text = .G: TB_B.Text = .b: TB_A.Text = .a: End With
End Function

Public Function RGBA_TryParse(ByVal s As String, RGBA_out As RGBA) As Boolean
Try: On Error GoTo Catch
    Dim sv As String, v As Byte, sa() As String: sa = Split(s, ",")
    Dim i As Long, u As Long: u = UBound(sa)
    If u >= i Then sv = sa(i): If Byte_TryParse(sv, v) Then RGBA_out.r = v
    i = i + 1
    If u >= i Then sv = sa(i): If Byte_TryParse(sv, v) Then RGBA_out.G = v
    i = i + 1
    If u >= i Then sv = sa(i): If Byte_TryParse(sv, v) Then RGBA_out.b = v
    i = i + 1
    If u >= i Then sv = sa(i): If Byte_TryParse(sv, v) Then RGBA_out.a = v
    RGBA_TryParse = True
Catch:
End Function

Public Function RGBA_ToStr(RGBA As RGBA) As String
    RGBA_ToStr = RGBA.r & "," & RGBA.G & "," & RGBA.b
End Function

'https://en.wikipedia.org/wiki/Color_difference
Public Function RGBA_EuclidRMean(this As RGBA, other As RGBA) As Double
    Dim dR As Double: dR = CDbl(this.r) - CDbl(other.r)
    Dim dG As Double: dG = CDbl(this.G) - CDbl(other.G)
    Dim dB As Double: dB = CDbl(this.b) - CDbl(other.b)
    Dim sr As Double: sr = 0.5 * (CDbl(this.r) + CDbl(other.r))
    RGBA_EuclidRMean = Math.Sqr((2 + sr / 256#) * dR * dR + 4 * dG * dG + (2 + (255# - sr) / 256) * dB * dB)
End Function

Public Function RGBA_ToARGB(this As RGBA) As ARGB
    With RGBA_ToARGB: .r = this.r: .G = this.G: .b = this.b: .a = this.a: End With
End Function
Public Function RGBA_ToLngColor(this As RGBA) As LngColor
    LSet RGBA_ToLngColor = this
End Function
Public Function RGBA_ToWebHex(this As RGBA) As String
    With this: RGBA_ToWebHex = "#" & Hex2(.a) & Hex2(.r) & Hex2(.G) & Hex2(.b): End With
End Function

Public Function RGBA_ParseWebHex(ByVal HashtagColor As String) As RGBA
    If Left(HashtagColor, 1) <> "#" Then Exit Function
    HashtagColor = Mid$(HashtagColor, 2)
    Dim s As String: s = Mid$(HashtagColor, 1, 2)
    With RGBA_ParseWebHex
        If 7 < Len(HashtagColor) Then     'ARGB
            .a = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .r = CByte("&H" & s): s = Mid$(HashtagColor, 5, 2)
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 7, 2)
            .b = CByte("&H" & s)
        ElseIf 5 < Len(HashtagColor) Then 'RGB
            .r = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 5, 2)
            .b = CByte("&H" & s)
        ElseIf 3 < Len(HashtagColor) Then 'GB
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .b = CByte("&H" & s)
        ElseIf 1 < Len(HashtagColor) Then 'B
            .b = CByte("&H" & s)
        End If
    End With
End Function

Public Function RGBA_ToRGBAf(this As RGBA) As RGBAf
    With this: RGBA_ToRGBAf = RGBAf(.r / 255, .G / 255, .b / 255, .a / 255): End With
End Function

Public Function RGBA_ToCMYK(this As RGBA) As CMYK
    RGBA_ToCMYK = RGBAf_ToCMYK(RGBA_ToRGBAf(this))
End Function

'Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
'Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Public Function RGBA_ToHSLA(this As RGBA) As HSLA
    Dim L As LngColor: L = RGBA_ToLngColor(this)
    Dim iiH As Integer, iiL As Integer, iiS As Integer
    With RGBA_ToHSLA
        .a = this.a
        ColorRGBToHLS L.Value, iiH, iiL, iiS
        .H = CByte(iiH)
        .s = CByte(iiS)
        .L = CByte(iiL)
    End With
End Function

'https://de.wikipedia.org/wiki/HSV-Farbraum
'Gelb = RGBA(255, 255, 0, 0) = HSL(40, 240, 120)
Public Function RGBA_ToHSLAf(this As RGBA) As HSLAf
    RGBA_ToHSLAf = RGBAf_ToHSLAf(RGBA_ToRGBAf(this))
End Function
Public Function RGBA_ToHSV(this As RGBA) As HSV
    RGBA_ToHSV = RGBAf_ToHSV(RGBA_ToRGBAf(this))
End Function
'        XYZ2rgb: function(XYZ, skip) {
'            var _Math = _math,
'                M = _instance.options.XYZMatrix,
'                X = XYZ.X,
'                Y = XYZ.Y,
'                Z = XYZ.Z,
'                r = X * M.R[0] + Y * M.R[1] + Z * M.R[2],
'                g = X * M.G[0] + Y * M.G[1] + Z * M.G[2],
'                b = X * M.B[0] + Y * M.B[1] + Z * M.B[2],
'                N = 1 / 2.4;
'
'            M = 0.0031308;
'
'            r = (r > M ? 1.055 * _Math.pow(r, N) - 0.055 : 12.92 * r);
'            g = (g > M ? 1.055 * _Math.pow(g, N) - 0.055 : 12.92 * g);
'            b = (b > M ? 1.055 * _Math.pow(b, N) - 0.055 : 12.92 * b);
'
'            if (!skip) { // out of gammut
'                _colors._rgb = {r: r, g: g, b: b};
'            }
'
'            return {
'                r: limitValue(r, 0, 1),
'                g: limitValue(g, 0, 1),
'                b: limitValue(b, 0, 1)
'            };
'        },
'
'        rgb2XYZ: function(rgb) {
'            var _Math = _math,
'                M = _instance.options.XYZMatrix,
'                r = rgb.r,
'                g = rgb.g,
'                b = rgb.b,
'                N = 0.04045;
'
'            r = (r > N ? _Math.pow((r + 0.055) / 1.055, 2.4) : r / 12.92);
'            g = (g > N ? _Math.pow((g + 0.055) / 1.055, 2.4) : g / 12.92);
'            b = (b > N ? _Math.pow((b + 0.055) / 1.055, 2.4) : b / 12.92);
'
'            return {
'                X: r * M.X[0] + g * M.X[1] + b * M.X[2],
'                Y: r * M.Y[0] + g * M.Y[1] + b * M.Y[2],
'                Z: r * M.Z[0] + g * M.Z[1] + b * M.Z[2]
'            };
'        },
'Public Function RGBA_ToCMYK(this As RGBA) As CMYK
'    With RGBA_ToCMYK
'        .C = 255 - this.R
'        .M = 255 - this.G
'        .Y = 255 - this.B
'        .K = MinB(.C, MinB(.M, .Y))
'        If .K = 255 Then Exit Function
'        .C = ((.C - .K) / (255 - .K)) * 255
'        .M = ((.M - .K) / (255 - .K)) * 255
'        .Y = ((.Y - .K) / (255 - .K)) * 255
'    End With
'End Function

'https://de.wikipedia.org/wiki/YCbCr-Farbmodell
Public Function RGBA_ToYCbCr(this As RGBA) As YCbCr
    'ITU-R BT 601 (=CCIR 601)
    Const Kb As Single = 0.114!
    Const Kr As Single = 0.299!
    With this
        RGBA_ToYCbCr.Y = Kr * .r + (1 - Kb - Kr) * .G + Kb * .b
        RGBA_ToYCbCr.cb = -0.168736 * .r - 0.331264 * .G + 0.5 * .b + 128
        RGBA_ToYCbCr.Cr = 0.5 * .r - 0.418688 * .G - 0.081312 * .b + 128
        RGBA_ToYCbCr.a = .a / 255
    End With
End Function

Public Function CVR_RGBA_R() As ColorValueRange
    CVR_RGBA_R = ColorValueRange(0, 255, 256, "0")
End Function

Public Function CVR_RGBA_G() As ColorValueRange
    CVR_RGBA_G = ColorValueRange(0, 255, 256, "0")
End Function

Public Function CVR_RGBA_B() As ColorValueRange
    CVR_RGBA_B = ColorValueRange(0, 255, 256, "0")
End Function

Public Function CVR_RGBA_A() As ColorValueRange
    CVR_RGBA_A = ColorValueRange(0, 255, 256, "0")
End Function

' #################### ' ARGB ' #################### '
'Public Type ARGB
'    A As Byte '0..255
'    R As Byte '0..255
'    G As Byte '0..255
'    B As Byte '0..255
'End Type
Public Function ARGB_ToRGBA(this As ARGB) As RGBA
    With this: ARGB_ToRGBA.r = .r: ARGB_ToRGBA.G = .G: ARGB_ToRGBA.b = .b: ARGB_ToRGBA.a = .a: End With
End Function

' #################### ' RGBAf ' #################### '
'Public Type RGBAf
'    R As Single '0..1
'    G As Single '0..1
'    B As Single '0..1
'    A As Single '0..1
'End Type
Public Function RGBAf(r As Single, G As Single, b As Single, a As Single) As RGBAf
    With RGBAf: .r = r: .G = G: .b = b: .a = a: End With
End Function
Public Function RGBAf_Read(this_out As RGBAf, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_R.Text: If Single_TryParse(s, v) Then .r = v Else err_out = s: Exit Function
        s = TB_G.Text: If Single_TryParse(s, v) Then .G = v Else err_out = s: Exit Function
        s = TB_B.Text: If Single_TryParse(s, v) Then .b = v Else err_out = s: Exit Function
        s = TB_A.Text: If Single_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    RGBAf_Read = True
End Function
Public Function RGBAf_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBAf)
    Dim fmt As String: fmt = "0.000"
    With this
        TB_R.Text = Format(.r, fmt)
        TB_G.Text = Format(.G, fmt)
        TB_B.Text = Format(.b, fmt)
        TB_A.Text = Format(.a, fmt)
    End With
End Function

Public Function RGBAf_EuclidRMean(this As RGBAf, other As RGBAf) As Double
    Dim dR As Double: dR = this.r - other.r
    Dim dG As Double: dG = this.G - other.G
    Dim dB As Double: dB = this.b - other.b
    Dim sr As Double: sr = 0.5 * (this.r + other.r)
    RGBAf_EuclidRMean = Math.Sqr((2 + sr / 256) * dR * dR + 4 * dG * dG + (2 + (255 - sr) / 256) * dB * dB)
End Function

Public Function RGBAf_ToRGBA(this As RGBAf) As RGBA
    With this
        RGBAf_ToRGBA.r = CByte(Min(.r * 255, 255))
        RGBAf_ToRGBA.G = CByte(Min(.G * 255, 255))
        RGBAf_ToRGBA.b = CByte(Min(.b * 255, 255))
        RGBAf_ToRGBA.a = CByte(Min(.a * 255, 255))
    End With
End Function

Public Function RGBAf_ToCMYK(this As RGBAf) As CMYK
    With RGBAf_ToCMYK
        .a = this.a
        .c = 1 - this.r
        .M = 1 - this.G
        .Y = 1 - this.b
        .k = MinSng3(.c, .M, .Y)
        If .k = 1 Then Exit Function
        Dim kf As Single: kf = 1 - .k
        .c = ((.c - .k) / kf)
        .M = ((.M - .k) / kf)
        .Y = ((.Y - .k) / kf)
    End With
End Function

'https://de.wikipedia.org/wiki/HSV-Farbraum
'Gelb = RGBA(255, 255, 0, 0) = HSL(40, 240, 120)
Public Function RGBAf_ToHSLAf(this As RGBAf) As HSLAf
    With this
        Dim MaxRGB As Single: MaxRGB = MaxSng3(.r, .G, .b)
        Dim MinRGB As Single: MinRGB = MinSng3(.r, .G, .b)
    End With
    With RGBAf_ToHSLAf
        .a = this.a
        .L = (MaxRGB + MinRGB) / 2
        If MaxRGB = MinRGB Then
            .H = 0: .s = 0 'achromatic
        Else
            Dim Delta As Single: Delta = MaxRGB - MinRGB
            If .L > 0.5 Then
                .s = Delta / (2 - MaxRGB - MinRGB)
            Else
                .s = Delta / (MaxRGB + MinRGB)
            End If
            Select Case MaxRGB
            Case this.r: If this.G < this.b Then .H = 6 Else .H = 0
                         .H = (this.G - this.b) / Delta + .H
            Case this.G: .H = (this.b - this.r) / Delta + 2
            Case this.b: .H = (this.r - this.G) / Delta + 4
            End Select
            .H = .H / 6
        End If
    End With
End Function

Function RGBAf_ToHSV(this As RGBAf) As HSV
    With this
        Dim MaxRGB As Single: MaxRGB = MaxSng3(.r, .G, .b)
        Dim MinRGB As Single: MinRGB = MinSng3(.r, .G, .b)
    End With
    With RGBAf_ToHSV
        .a = this.a
        .v = MaxRGB
        Dim Delta As Single: Delta = MaxRGB - MinRGB
        If MaxRGB <> 0 Then .s = Delta / MaxRGB
        If MaxRGB = MinRGB Then
            .H = 0 'achromatic
        Else
            Select Case MaxRGB
            Case this.r: .H = (this.G - this.b) / Delta
                         If this.G < this.b Then .H = .H + 6
            Case this.G: .H = (this.b - this.r) / Delta + 2
            Case this.b: .H = (this.r - this.G) / Delta + 4
            End Select
            .H = .H / 6
        End If
    End With
End Function

'Function RGBAf_ToHSB(this As RGBAf) As HSB
'    '
'End Function

'rgb2XYZ: function(rgb) {
'    var _Math = _math,
'        M = _instance.options.XYZMatrix,
'        r = rgb.r,
'        g = rgb.g,
'        b = rgb.b,
'        N = 0.04045;
'
'    r = (r > N ? _Math.pow((r + 0.055) / 1.055, 2.4) : r / 12.92);
'    g = (g > N ? _Math.pow((g + 0.055) / 1.055, 2.4) : g / 12.92);
'    b = (b > N ? _Math.pow((b + 0.055) / 1.055, 2.4) : b / 12.92);
'
'    return {
'        X: r * M.X[0] + g * M.X[1] + b * M.X[2],
'        Y: r * M.Y[0] + g * M.Y[1] + b * M.Y[2],
'        Z: r * M.Z[0] + g * M.Z[1] + b * M.Z[2]
'    };
'},

Function RGBAf_ToXYZ(this As RGBAf) As XYZ
    Dim r As Single: r = this.r
    Dim G As Single: G = this.G
    Dim b As Single: b = this.b
    Dim n As Single: n = 0.04045
    
    If r > n Then r = ((r + 0.055) / 1.055) ^ (2.4) Else r = r / 12.92
    If G > n Then G = ((G + 0.055) / 1.055) ^ (2.4) Else G = G / 12.92
    If b > n Then b = ((b + 0.055) / 1.055) ^ (2.4) Else b = b / 12.92
    
    With RGBAf_ToXYZ
        .x = r * M.x(0) + G * M.x(1) + b * M.x(2)
        .Y = r * M.Y(0) + G * M.Y(1) + b * M.Y(2)
        .z = r * M.z(0) + G * M.z(1) + b * M.z(2)
        .a = this.a
    End With
End Function

'https://de.wikipedia.org/wiki/YCbCr-Farbmodell
'https://de.wikipedia.org/wiki/ITU-R_BT_601
Function RGBAf_ToYCbCr(this As RGBAf) As YCbCr
    'ITU-R BT 601 (=CCIR 601)
    Const Kb As Single = 0.114
    Const Kr As Single = 0.299
    With this
        Dim Y As Single: Y = Kr * .r + (1 - Kb - Kr) * .G + Kb * .b
        RGBAf_ToYCbCr.Y = Y
        RGBAf_ToYCbCr.cb = 0.5 * (.b - Y) / (1 - Kb)
        RGBAf_ToYCbCr.Cr = 0.5 * (.r - Y) / (1 - Kr)
        RGBAf_ToYCbCr.a = .a
    End With
End Function

Public Function CVR_RGBAf_R() As ColorValueRange
    CVR_RGBAf_R = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_RGBAf_G() As ColorValueRange
    CVR_RGBAf_G = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_RGBAf_B() As ColorValueRange
    CVR_RGBAf_B = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_RGBAf_A() As ColorValueRange
    CVR_RGBAf_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' CMYK ' #################### '
'Public Type CMYK
'    C As Single '0..1
'    M As Single '0..1
'    Y As Single '0..1
'    K As Single '0..1
'    A As Single '0..1
'End Type
Public Function CMYK(ByVal c As Single, ByVal M As Single, ByVal Y As Single, ByVal k As Single, ByVal a As Single) As CMYK
    With CMYK: .c = c: .M = M: .Y = Y: .k = k: .a = a: End With
End Function

Public Function CMYK_Read(this_out As CMYK, TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_C.Text: If Single_TryParse(s, v) Then .c = v Else err_out = s: Exit Function
        s = TB_M.Text: If Single_TryParse(s, v) Then .M = v Else err_out = s: Exit Function
        s = TB_Y.Text: If Single_TryParse(s, v) Then .Y = v Else err_out = s: Exit Function
        s = TB_K.Text: If Single_TryParse(s, v) Then .k = v Else err_out = s: Exit Function
        s = TB_A.Text: If Single_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    CMYK_Read = True
End Function
Public Function CMYK_ToView(TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, this As CMYK)
    With this
        TB_C.Text = Format(.c, m_CVR_CMYK_C.FormatStr)
        TB_M.Text = Format(.M, m_CVR_CMYK_M.FormatStr)
        TB_Y.Text = Format(.Y, m_CVR_CMYK_Y.FormatStr)
        TB_K.Text = Format(.k, m_CVR_CMYK_K.FormatStr)
        TB_A.Text = Format(.a, m_CVR_CMYK_A.FormatStr)
    End With
End Function

Public Function CMYK_Euclidean(this As CMYK, other As CMYK) As Double
    Dim dC As Double: dC = this.c - other.c
    Dim dM As Double: dM = this.M - other.M
    Dim dY As Double: dY = this.Y - other.Y
    Dim dK As Double: dK = this.k - other.k
    CMYK_Euclidean = VBA.Math.Sqr(dC * dC + dM * dM + dY * dY + dK * dK)
End Function

Public Function CMYK_ToRGBAf(this As CMYK) As RGBAf
    With this
        Dim kf As Single: kf = 1 - .k
        CMYK_ToRGBAf.r = 1 - MinSng(1, .c * kf + .k)
        CMYK_ToRGBAf.G = 1 - MinSng(1, .M * kf + .k)
        CMYK_ToRGBAf.b = 1 - MinSng(1, .Y * kf + .k)
        CMYK_ToRGBAf.a = .a
    End With
End Function
Public Function CMYK_ToRGBA(this As CMYK) As RGBA
'    With CMYK_ToRGBA
'        .R = 255 - MinB(255, (this.C / 255) * (255 - this.K) + this.K)
'        .G = 255 - MinB(255, (this.M / 255) * (255 - this.K) + this.K)
'        .B = 255 - MinB(255, (this.Y / 255) * (255 - this.K) + this.K)
'    End With
    CMYK_ToRGBA = RGBAf_ToRGBA(CMYK_ToRGBAf(this))
End Function

Public Function CVR_CMYK_C() As ColorValueRange
    CVR_CMYK_C = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_CMYK_M() As ColorValueRange
    CVR_CMYK_M = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_CMYK_Y() As ColorValueRange
    CVR_CMYK_Y = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_CMYK_K() As ColorValueRange
    CVR_CMYK_K = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_CMYK_A() As ColorValueRange
    CVR_CMYK_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' HSLA  ' #################### '
'Public Type HSLA
'    H As Byte '0..239
'    S As Byte '0..240
'    L As Byte '0..240
'    A As Byte '0..255
'End Type
Public Function HSLA(ByVal H As Byte, ByVal s As Byte, ByVal L As Byte, ByVal a As Byte) As HSLA
    With HSLA: .H = H: .s = s: .L = L: .a = a: End With
End Function

Public Function HSLA_ToRGBA(this As HSLA) As RGBA
    Dim L As LngColor
    With this
        L.Value = ColorHLSToRGB(.H, .L, .s)
    End With
    HSLA_ToRGBA = LngColor_ToRGBA(L)
    HSLA_ToRGBA.a = this.a
End Function
Public Function HSLA_Read(this_out As HSLA, TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Byte, s As String
    With this_out
        s = TB_H.Text: If Byte_TryParse(s, v) Then .H = MinByt(v, 240) Else err_out = s: Exit Function
        s = TB_S.Text: If Byte_TryParse(s, v) Then .s = MinByt(v, 240) Else err_out = s: Exit Function
        s = TB_L.Text: If Byte_TryParse(s, v) Then .L = MinByt(v, 240) Else err_out = s: Exit Function
        s = TB_A.Text: If Byte_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    HSLA_Read = True
End Function
Public Function HSLA_ToView(TBHSLA_H As TextBox, TBHSLA_S As TextBox, TBHSLA_L As TextBox, TBHSLA_A As TextBox, this As HSLA)
    With this
        TBHSLA_H.Text = .H
        TBHSLA_S.Text = .s
        TBHSLA_L.Text = .L
        TBHSLA_A.Text = .a
    End With
End Function

Public Function HSLA_Euclidean(this As HSLA, other As HSLA) As Double
    '
End Function

Public Function HSLA_ToHSLAf(this As HSLA) As HSLAf
    With HSLA_ToHSLAf
        .H = this.H / 240
        .s = this.s / 240
        .L = this.L / 240
        .a = this.a / 255
    End With
End Function

Public Function CVR_HSL_H() As ColorValueRange
    CVR_HSL_H = ColorValueRange(0, 239, 239, "0")
End Function

Public Function CVR_HSL_S() As ColorValueRange
    CVR_HSL_S = ColorValueRange(0, 240, 240, "0")
End Function

Public Function CVR_HSL_L() As ColorValueRange
    CVR_HSL_L = ColorValueRange(0, 240, 240, "0")
End Function

Public Function CVR_HSL_A() As ColorValueRange
    CVR_HSL_A = ColorValueRange(0, 255, 255, "0")
End Function

' #################### ' HSLAf ' #################### '
'Public Type HSLAf
'    H As Single '0..1
'    S As Single '0..1
'    L As Single '0..1
'    A As Single '0..1
'End Type
Public Function HSLAf(ByVal H As Single, ByVal s As Single, ByVal L As Single, ByVal a As Single) As HSLAf
    With HSLAf: .H = H: .s = s: .L = L: .a = a: End With
End Function

Public Function HSLAf_ToRGBAf(this As HSLAf) As RGBAf
    With this
        HSLAf_ToRGBAf.a = .a
        If .s = 0 Then 'achromatic
            HSLAf_ToRGBAf.r = .L
            HSLAf_ToRGBAf.G = .L
            HSLAf_ToRGBAf.b = .L
        Else
            Dim q As Single: If .L < 0.5 Then q = .L * (1 + .s) Else q = .L + .s - .L * .s
            Dim p As Single: p = 2 * .L - q
            HSLAf_ToRGBAf.r = Hue_ToRGB(p, q, .H + 1 / 3)
            HSLAf_ToRGBAf.G = Hue_ToRGB(p, q, .H)
            HSLAf_ToRGBAf.b = Hue_ToRGB(p, q, .H - 1 / 3)
        End If
    End With
End Function
Public Function Hue_ToRGB(ByVal p As Single, ByVal q As Single, ByVal t As Single) As Single
    If t < 0 Then t = t + 1
    If t > 1 Then t = t - 1
    If t < 1 / 6 Then
        Hue_ToRGB = p + (q - p) * 6 * t
    ElseIf t < 1 / 2 Then
        Hue_ToRGB = q
    ElseIf t < 2 / 3 Then
        Hue_ToRGB = p + (q - p) * (2 / 3 - t) * 6
    Else
        Hue_ToRGB = p
    End If
End Function

Public Function HSLAf_Read(this_out As HSLAf, TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_H.Text: If Single_TryParse(s, v) Then .H = v Else err_out = s: Exit Function
        s = TB_S.Text: If Single_TryParse(s, v) Then .s = v Else err_out = s: Exit Function
        s = TB_L.Text: If Single_TryParse(s, v) Then .L = v Else err_out = s: Exit Function
        s = TB_A.Text: If Single_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    HSLAf_Read = True
End Function
Public Function HSLAf_ToView(TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, this As HSLAf)
    With this
        TB_H.Text = Format(.H, m_CVR_HSLAf_H.FormatStr)
        TB_S.Text = Format(.s, m_CVR_HSLAf_S.FormatStr)
        TB_L.Text = Format(.L, m_CVR_HSLAf_L.FormatStr)
        TB_A.Text = Format(.a, m_CVR_HSLAf_A.FormatStr)
    End With
End Function

Public Function CVR_HSLAf_H() As ColorValueRange
    CVR_HSLAf_H = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSLAf_S() As ColorValueRange
    CVR_HSLAf_S = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSLAf_L() As ColorValueRange
    CVR_HSLAf_L = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSLAf_A() As ColorValueRange
    CVR_HSLAf_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' HSV ' #################### '
'https://www.rapidtables.com/convert/color/hsv-to-rgb.html
'Public Type HSV
'    H As Single '0..1
'    S As Single '0..1
'    v As Single '0..1
'    A As Single '0..1
'End Type
Public Function HSV(ByVal Hue As Single, ByVal Saturation As Single, ByVal Value As Single, ByVal Alpha As Single) As HSV
    With HSV
        .H = Hue
        .s = Saturation
        .v = Value
        .a = Alpha
    End With
End Function

Public Function HSV_ToRGBAf(this As HSV) As RGBAf
    With this
        Dim i As Single: i = CSng(Int(.H * 6)) 'Floor
        Dim f As Single: f = .H * 6 - i
        Dim p As Single: p = .v * (1 - .s)
        Dim q As Single: q = .v * (1 - f * .s)
        Dim t As Single: t = .v * (1 - (1 - f) * .s)
    End With
    With HSV_ToRGBAf
        .a = this.a
        Select Case i Mod 6
        Case 0: .r = this.v: .G = t:      .b = p
        Case 1: .r = q:      .G = this.v: .b = p
        Case 2: .r = p:      .G = this.v: .b = t
        Case 3: .r = p:      .G = q:      .b = this.v
        Case 4: .r = t:      .G = p:      .b = this.v
        Case 5: .r = this.v: .G = p:      .b = q
        End Select
    End With
End Function

Public Function HSV_Read(this_out As HSV, TB_H As TextBox, TB_S As TextBox, TB_V As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_H.Text: If Single_TryParse(s, v) Then .H = v Else err_out = s: Exit Function
        s = TB_S.Text: If Single_TryParse(s, v) Then .s = v Else err_out = s: Exit Function
        s = TB_V.Text: If Single_TryParse(s, v) Then .v = v Else err_out = s: Exit Function
        s = TB_A.Text: If Single_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    HSV_Read = True
End Function
Public Function HSV_ToView(TB_H As TextBox, TB_S As TextBox, TB_V As TextBox, TB_A As TextBox, this As HSV)
    With this
        TB_H.Text = Format(.H, m_CVR_HSV_H.FormatStr)
        TB_S.Text = Format(.s, m_CVR_HSV_S.FormatStr)
        TB_V.Text = Format(.v, m_CVR_HSV_V.FormatStr)
        TB_A.Text = Format(.a, m_CVR_HSV_A.FormatStr)
    End With
End Function

Public Function CVR_HSV_H() As ColorValueRange
    CVR_HSV_H = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSV_S() As ColorValueRange
    CVR_HSV_S = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSV_V() As ColorValueRange
    CVR_HSV_V = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_HSV_A() As ColorValueRange
    CVR_HSV_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' HSB ' #################### '
'Public Function HSB_ToRGBAf(this As HSB) As RGBAf
'    '
'End Function
'
'Public Function HSB_Read(this_out As HSB, TB_H As TextBox, TB_S As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
'    '
'End Function
'Public Function HSB_ToView(TB_H As TextBox, TB_S As TextBox, TB_B As TextBox, TB_A As TextBox, this As HSB)
'    With this
'        TB_H.Text = Format(.H, "0.#####")
'        TB_S.Text = Format(.S, "0.#####")
'        TB_B.Text = Format(.B, "0.#####")
'        TB_A.Text = Format(.A, "0.#####")
'    End With
'End Function
'

' #################### ' XYZ ' #################### '
'XYZ2rgb: function(XYZ, skip) {
'    var _Math = _math,
'        M = _instance.options.XYZMatrix,
'        X = XYZ.X,
'        Y = XYZ.Y,
'        Z = XYZ.Z,
'        r = X * M.R[0] + Y * M.R[1] + Z * M.R[2],
'        g = X * M.G[0] + Y * M.G[1] + Z * M.G[2],
'        b = X * M.B[0] + Y * M.B[1] + Z * M.B[2],
'        N = 1 / 2.4;
'
'    M = 0.0031308;
'
'    r = (r > M ? 1.055 * _Math.pow(r, N) - 0.055 : 12.92 * r);
'    g = (g > M ? 1.055 * _Math.pow(g, N) - 0.055 : 12.92 * g);
'    b = (b > M ? 1.055 * _Math.pow(b, N) - 0.055 : 12.92 * b);
'
'    if (!skip) { // out of gammut
'        _colors._rgb = {r: r, g: g, b: b};
'    }
'
'    return {
'        r: limitValue(r, 0, 1),
'        g: limitValue(g, 0, 1),
'        b: limitValue(b, 0, 1)
'    };
'},
'Public Type XYZ
'    x As Single '0..1
'    Y As Single '0..1
'    z As Single '0..1
'    a As Single '0..1
'End Type

Public Function XYZ(ByVal aX As Single, ByVal aY As Single, ByVal aZ As Single, Optional ByVal aa As Single = 1!) As XYZ
    With XYZ: .x = aX: .Y = aY: .z = aZ: .a = aa: End With
End Function
Public Function XYZ_Read(this_out As XYZ, TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_X.Text: If Single_TryParse(s, v) Then .x = v Else err_out = s: Exit Function
        s = TB_Y.Text: If Single_TryParse(s, v) Then .Y = v Else err_out = s: Exit Function
        s = TB_Z.Text: If Single_TryParse(s, v) Then .z = v Else err_out = s: Exit Function
        s = TB_A.Text: If Single_TryParse(s, v) Then .a = v Else err_out = s: Exit Function
    End With
    XYZ_Read = True
End Function
Public Function XYZ_ToView(TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, this As XYZ)
    With this
        TB_X.Text = Format(.x, m_CVR_XYZ_X.FormatStr)
        TB_Y.Text = Format(.Y, m_CVR_XYZ_Y.FormatStr)
        TB_Z.Text = Format(.z, m_CVR_XYZ_Z.FormatStr)
        TB_A.Text = Format(.a, m_CVR_XYZ_A.FormatStr)
    End With
End Function

Public Function XYZ_Euclidean(this As XYZ, other As XYZ) As Double
    Dim dx As Double: dx = this.x - other.x
    Dim dY As Double: dY = this.Y - other.Y
    Dim dZ As Double: dZ = this.z - other.z
    XYZ_Euclidean = Math.Sqr(dx * dx + dY * dY + dZ * dZ)
End Function

Public Function XYZ_ToRGBAf(this As XYZ) As RGBAf
    With this
        Dim var_X As Single: var_X = .x '/ 100        '//X from 0 to  95.047      (Observer = 2°, Illuminant = D65)
        Dim var_Y As Single: var_Y = .Y '/ 100        '//Y from 0 to 100.000
        Dim var_Z As Single: var_Z = .z '/ 100        '//Z from 0 to 108.883
    End With
    Dim var_R As Single: var_R = var_X * 3.2406 + var_Y * -1.5372 + var_Z * -0.4986
    Dim var_G As Single: var_G = var_X * -0.9689 + var_Y * 1.8758 + var_Z * 0.0415
    Dim var_B As Single: var_B = var_X * 0.0557 + var_Y * -0.204 + var_Z * 1.057

    With XYZ_ToRGBAf
        If (var_R > 0.0031308) Then .r = 1.055 * (var_R ^ (1 / 2.4)) - 0.055 Else .r = 12.92 * var_R
        If (var_G > 0.0031308) Then .G = 1.055 * (var_G ^ (1 / 2.4)) - 0.055 Else .G = 12.92 * var_G
        If (var_B > 0.0031308) Then .b = 1.055 * (var_B ^ (1 / 2.4)) - 0.055 Else .b = 12.92 * var_B
    End With
    'r = var_R * 255
    'G = var_G * 255
    'b = var_B * 255

End Function


'Public Function XYZ_ToRGBAf(this As XYZ) As RGBAf
'    Dim x As Single: x = this.x
'    Dim Y As Single: Y = this.Y
'    Dim z As Single: z = this.z
'    Dim r As Single: r = x * M.r(0) + Y * M.r(1) + z * M.r(2)
'    Dim G As Single: G = x * M.G(0) + Y * M.G(1) + z * M.G(2)
'    Dim b As Single: b = x * M.b(0) + Y * M.b(1) + z * M.b(2)
'    Dim n As Single: n = 1 / 2.4
'    Dim MM As Single: MM = 0.0031308
'
'    If r > MM Then r = 1.055 * r ^ n - 0.055 Else r = 12.92 * r
'    If G > MM Then G = 1.055 * G ^ n - 0.055 Else G = 12.92 * G
'    If b > MM Then b = 1.055 * b ^ n - 0.055 Else b = 12.92 * b
'
'    With XYZ_ToRGBAf
'        .r = r * 0.13844015
'
'        .G = G * 0.13844015
'        .b = b * 0.0913844
'
'        .a = this.a
''        .R = MinSng(MaxSng(R, 0), 1) 'limitValue 0..1
''        .G = MinSng(MaxSng(G, 0), 1)
''        .B = MinSng(MaxSng(B, 0), 1)
''        .A = this.A
'        Debug.Print "RGB:(" & .r & ", " & .G & ", " & .b & ")"
'    End With
'End Function

Public Function XYZ_ToCIELab(this As XYZ, Optional lighttype As CIELabLight = CIELabLight.D65_2) As CIELab
    'https://de.wikipedia.org/wiki/Lab-Farbraum
    Dim n As XYZ: n = CIELabLights(lighttype)
    Dim XXN As Double: If n.x <> 0 Then XXN = this.x / n.x
    Dim YYN As Double: If n.Y <> 0 Then YYN = this.Y / n.Y
    Dim ZZN As Double: If n.z <> 0 Then ZZN = this.z / n.z
    Dim root3_XXN As Double: If XXN < 216 / 24389 Then root3_XXN = 1 / 116 * (24389 / 27 * XXN + 16) Else root3_XXN = (this.x / n.x) ^ (1 / 3)
    Dim root3_YYN As Double: If YYN < 216 / 24389 Then root3_YYN = 1 / 116 * (24389 / 27 * YYN + 16) Else root3_YYN = (this.Y / n.Y) ^ (1 / 3)
    Dim root3_ZZN As Double: If ZZN < 216 / 24389 Then root3_ZZN = 1 / 116 * (24389 / 27 * ZZN + 16) Else root3_ZZN = (this.z / n.z) ^ (1 / 3)
    With XYZ_ToCIELab
        .L = 116 * root3_YYN - 16
        .aa = 500 * (root3_XXN - root3_YYN)
        .bb = 200 * (root3_YYN - root3_ZZN)
        .a = this.a
    End With
End Function

Public Function CVR_XYZ_X() As ColorValueRange
    CVR_XYZ_X = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_XYZ_Y() As ColorValueRange
    CVR_XYZ_Y = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_XYZ_Z() As ColorValueRange
    CVR_XYZ_Z = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_XYZ_A() As ColorValueRange
    CVR_XYZ_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' CIELab ' #################### '
'https://de.wikipedia.org/wiki/Lab-Farbraum
'CIE = International Commission on Illumination
'Public Type CIELab
'    L  As Single ' 0 .. 100
'    aA As Single '-128 .. 128
'    bb As Single '-128 .. 128
'    a  As Single  '0..1
'End Type

Public Function CIELab(ByVal L As Single, ByVal aa As Single, ByVal bb As Single) As CIELab
    With CIELab: .L = L: .aa = aa: .bb = bb: .a = 1: End With
End Function
Public Function CIELab_Read(this_out As CIELab, TB_L As TextBox, TB_aa As TextBox, TB_bb As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_L.Text:  If Single_TryParse(s, v) Then .L = v Else err_out = s:  Exit Function
        s = TB_aa.Text: If Single_TryParse(s, v) Then .aa = v Else err_out = s: Exit Function
        s = TB_bb.Text: If Single_TryParse(s, v) Then .bb = v Else err_out = s: Exit Function
        s = TB_A.Text:  If Single_TryParse(s, v) Then .a = v Else err_out = s:  Exit Function
    End With
    CIELab_Read = True
End Function
Public Function CIELab_ToView(TB_L As TextBox, TB_aa As TextBox, TB_bb As TextBox, TB_A As TextBox, this As CIELab)
    With this
        TB_L.Text = Format(.L, m_CVR_CIELab_L.FormatStr)
        TB_aa.Text = Format(.aa, m_CVR_CIELab_aa.FormatStr)
        TB_bb.Text = Format(.bb, m_CVR_CIELab_bb.FormatStr)
        TB_A.Text = Format(.a, m_CVR_CIELab_A.FormatStr)
    End With
End Function

Function CIELabLight_ToStr(ByVal L As CIELabLight) As String
    Dim s As String
    Select Case L
    Case CIELabLight.D50_2:  s = "D-50 2°"
    Case CIELabLight.D65_2:  s = "D-65 2°"
    Case CIELabLight.D50_10: s = "D-50 10°"
    Case CIELabLight.D65_10: s = "D-65 10°"
    End Select
    CIELabLight_ToStr = s
End Function

Public Sub CIELabLight_ToCmb(aCBLB As ComboBox)
    Dim i As Long, L As CIELabLight
    With aCBLB
        .Clear
        For i = 0 To 3
            L = i
            .AddItem CIELabLight_ToStr(L)
        Next
        .ListIndex = 3
    End With
End Sub

'    CIELabLights(CIELabLight.D50_2) = XYZ(0.96422, 1, 0.82521)
'    CIELabLights(CIELabLight.D65_2) = XYZ(0.95047, 1, 1.08883)
'    CIELabLights(CIELabLight.D50_10) = XYZ(0.9672, 1, 0.81427)
'    CIELabLights(CIELabLight.D65_10) = XYZ(0.94811, 1, 1.07304)

'https://stackoverflow.com/questions/58973890/lab-to-xyz-and-xyz-to-rgb-color-space-conversion-algorithm
'https://stackoverflow.com/questions/58952430/rgb-xyz-and-xyz-lab-color-space-conversion-algorithm
Public Function CIELab_ToXYZ(this As CIELab, Optional ByVal lighttype As CIELabLight = CIELabLight.D65_2) As XYZ
    With this
        Dim var_Y As Single: var_Y = (.L + 16) / 116
        Dim var_X As Single: var_X = .aa / 500 + var_Y
        Dim var_Z As Single: var_Z = var_Y - .bb / 200
    End With

    If (var_Y ^ 3 > 0.008856) Then var_Y = var_Y ^ 3 Else var_Y = (var_Y - 16 / 116) / 7.787
    If (var_X ^ 3 > 0.008856) Then var_X = var_X ^ 3 Else var_X = (var_X - 16 / 116) / 7.787
    If (var_Z ^ 3 > 0.008856) Then var_Z = var_Z ^ 3 Else var_Z = (var_Z - 16 / 116) / 7.787

    With CIELab_ToXYZ
        .x = CIELabLights(lighttype).x * var_X  '//ref_X =  95.047     Observer= 2°, Illuminant= D65
        .Y = CIELabLights(lighttype).Y * var_Y  '//ref_Y = 100.000
        .z = CIELabLights(lighttype).z * var_Z  '//ref_Z = 108.883
    End With
End Function

'Public Function CIELab_ToXYZ(this As CIELab, Optional ByVal lighttype As CIELabLight = CIELabLight.D65_2) As XYZ
'    'TODO!!!
'    Dim pow As Double, ratio As Double: ratio = 0.206892707 '6! / 29!
'    With CIELab_ToXYZ
'        .Y = (this.L + 16!) / 116!
'        .x = (this.aa / 500!) + .Y
'        .z = .Y - (this.bb / 200!)
'        .a = this.a * 255
'
'        pow = .x * .x * .x
'        If .x > ratio Then
'            .x = pow
'        Else
'            .x = (3! * (6! / 29!) * (6! / 29!) * (.x - (4! / 29!)))
'        End If
'
'        pow = .Y * .Y * .Y
'        If .Y > ratio Then
'            .Y = pow
'        Else
'            .Y = (3! * (6! / 29!) * (6! / 29!) * (.Y - (4! / 29!)))
'        End If
'
'        pow = .z * .z * .z
'        If .z > ratio Then
'            .z = pow
'        Else
'            .z = (3! * (6! / 29!) * (6! / 29!) * (.z - (4! / 29!)))
'        End If
'
'        .x = .x * 95.047!
'        .Y = .Y * 100!
'        .z = .z * 108.883!
'    End With
'End Function


'public static Vector4 LabToXYZ(Vector4 color)
'        {
'            float[] xyz = new float[3];
'            float[] col = new float[] { color[0], color[1], color[2], color[3]};
'
'            xyz[1] = (col[0] + 16.0f) / 116.0f;
'            xyz[0] = (col[1] / 500.0f) + xyz[1];
'            xyz[2] = xyz[1] - (col[2] / 200.0f);
'
'             for (int i = 0; i < 3; i++)
'            {
'                float pow = xyz[i] * xyz[i] * xyz[i];
'                float ratio = (6.0f / 29.0f);
'                if (xyz[i] > ratio)
'                {
'                    xyz[i] = pow;
'                }
'                Else
'                {
'                    xyz[i] = (3.0f * (6.0f / 29.0f) * (6.0f / 29.0f) * (xyz[i] - (4.0f / 29.0f)));
'                }
'            }
'            xyz[0] = xyz[0] * 95.047f;
'            xyz[1] = xyz[1] * 100.0f;
'            xyz[2] = xyz[2] * 108.883f;
'
'            return new Vector4(xyz[0], xyz[1], xyz[2], color[3]);
'        }
'
'Public Type CIELab
'    L  As Single ' 0 .. 100
'    aA As Single '-128 .. 128
'    bb As Single '-128 .. 128
'    a  As Single  '0..1
'End Type


Public Function CVR_CIELab_L() As ColorValueRange
    CVR_CIELab_L = ColorValueRange(0, 100, 1000, "0.000")
End Function

Public Function CVR_CIELab_aa() As ColorValueRange
    CVR_CIELab_aa = ColorValueRange(-128, 128, 1000, "0.000")
End Function

Public Function CVR_CIELab_bb() As ColorValueRange
    CVR_CIELab_bb = ColorValueRange(-128, 128, 1000, "0.000")
End Function

Public Function CVR_CIELab_A() As ColorValueRange
    CVR_CIELab_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' #################### ' YCbCr ' #################### '
' E'Y  =  0..1
' E'Cb = -0.5..0.5
' E'Cr = -0.5..0.5

Public Function YCbCr(ByVal Y As Single, ByVal cb As Single, ByVal Cr As Single, ByVal a As Single) As YCbCr
    With YCbCr: .Y = Y: .cb = cb: .Cr = Cr: .a = a: End With
End Function

Public Function YCbCr_Read(this_out As YCbCr, TB_Y As TextBox, TB_Cb As TextBox, TB_Cr As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim v As Single, s As String
    With this_out
        s = TB_Y.Text:  If Single_TryParse(s, v) Then .Y = v Else err_out = s:  Exit Function
        s = TB_Cb.Text: If Single_TryParse(s, v) Then .cb = v Else err_out = s: Exit Function
        s = TB_Cr.Text: If Single_TryParse(s, v) Then .Cr = v Else err_out = s: Exit Function
        s = TB_A.Text:  If Single_TryParse(s, v) Then .a = v Else err_out = s:  Exit Function
    End With
    YCbCr_Read = True
End Function
Public Function YCbCr_ToView(TB_Y As TextBox, TB_Cb As TextBox, TB_Cr As TextBox, TB_A As TextBox, this As YCbCr)
    With this
        TB_Y.Text = Format(.Y, m_CVR_YCbCr_Y.FormatStr)
        TB_Cb.Text = Format(.cb, m_CVR_YCbCr_Cb.FormatStr)
        TB_Cr.Text = Format(.Cr, m_CVR_YCbCr_Cr.FormatStr)
        TB_A.Text = Format(.a, m_CVR_YCbCr_A.FormatStr)
    End With
End Function

Public Function YCbCr_ToRGBA(this As YCbCr) As RGBA
    With this
        YCbCr_ToRGBA.r = .Y + 1.402 * (.Cr - 128)
        YCbCr_ToRGBA.G = .Y - 0.34414 * (.cb - 128) - 0.71414 * (.Cr - 128)
        YCbCr_ToRGBA.b = .Y + 1.772 * (.cb - 128)
        
        YCbCr_ToRGBA.a = .a * 255
    End With
End Function

Public Function YCbCr_ToRGBAf(this As YCbCr) As RGBAf
    YCbCr_ToRGBAf = RGBA_ToRGBAf(YCbCr_ToRGBA(this))
End Function

Public Function CVR_YCbCr_Y() As ColorValueRange
    CVR_YCbCr_Y = ColorValueRange(0, 1, 1000, "0.000")
End Function

Public Function CVR_YCbCr_Cb() As ColorValueRange
    CVR_YCbCr_Cb = ColorValueRange(-0.5, 0.5, 1000, "0.000")
End Function

Public Function CVR_YCbCr_Cr() As ColorValueRange
    CVR_YCbCr_Cr = ColorValueRange(-0.5, 0.5, 1000, "0.000")
End Function

Public Function CVR_YCbCr_A() As ColorValueRange
    CVR_YCbCr_A = ColorValueRange(0, 1, 1000, "0.000")
End Function

' v #################### v '    ColorValueRange ' v #################### v '
Public Function ColorValueRange(ByVal RMin As Single, ByVal RMax As Single, ByVal CountSteps As Long, ByVal FormatStr As String) As ColorValueRange
    With ColorValueRange: .RangeMin = RMin:     .RangeMax = RMax:   .RangeSteps = CountSteps: .FormatStr = FormatStr: End With
End Function

Public Function ColorValueRange_IsEqual(this As ColorValueRange, other As ColorValueRange) As Boolean
    With this
        ColorValueRange_IsEqual = .RangeMin = other.RangeMin And .RangeMax = other.RangeMax And .RangeSteps = other.RangeSteps
    End With
End Function

Public Property Get ColorValueRange_dx(this As ColorValueRange) As Double
    With this
        ColorValueRange_dx = (.RangeMax - .RangeMin) / .RangeSteps
    End With
End Property

Public Sub ColorValueRange_ToComboBox(this As ColorValueRange, aCmb As ComboBox, ByVal sCurrentVal As String)
    Dim dx As Double: dx = ColorValueRange_dx(this)
    Dim cv As Double
    ' wir müssen den Wert mit der geringsten Differenz suchen
    'und davon das i merken
    If Double_TryParse(sCurrentVal, cv) Then
        'sCurrentVal = Format(d, this.FormatStr)
        'OK weiter
    End If
    Dim s As String, li As Long, vi As Double
    Dim dif_cvvi0 As Double, dif_cvvi1 As Double: dif_cvvi1 = this.RangeMax
    With aCmb
        .Clear
        Dim i As Long
        For i = this.RangeSteps To 0 Step -1
            vi = i * dx
            dif_cvvi0 = Abs(vi - cv)
            If dif_cvvi0 <= dif_cvvi1 Then
                dif_cvvi1 = dif_cvvi0
                li = i
            End If
            s = Format(vi, this.FormatStr)
            .AddItem s
        Next
        .ListIndex = .ListCount - li - 1
    End With
End Sub
