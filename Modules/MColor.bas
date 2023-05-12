Attribute VB_Name = "MColor"
Option Explicit
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

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Public Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
'Public Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long

Public CurMousePos As POINTAPI

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type LngColor
    Value As Long
End Type
Public Type RGBA
    R As Byte '0..255
    G As Byte '0..255
    B As Byte '0..255
    A As Byte '0..255
End Type
Public Type ARGB
    A As Byte '0..255
    R As Byte '0..255
    G As Byte '0..255
    B As Byte '0..255
End Type
Public Type RGBAf
    R As Single '0..1
    G As Single '0..1
    B As Single '0..1
    A As Single '0..1
End Type
Public Type CMYK
    C As Single '0..1
    M As Single '0..1
    Y As Single '0..1
    K As Single '0..1
    A As Single '0..1
End Type
Public Type HSLA
    H As Byte '0..255
    s As Byte '0..255
    l As Byte '0..255
    A As Byte '0..255
End Type
Public Type HSLAf
    H As Single '0..1
    s As Single '0..1
    l As Single '0..1
    A As Single '0..1
End Type
Public Type HSV
    H As Single '0..1
    s As Single '0..1
    V As Single '0..1
    A As Single '0..1
End Type
'Public Type HSB
'    H As Single '0..1
'    S As Single '0..1
'    B As Single '0..1
'    A As Single '0..1
'End Type
Public Type XYZ
    X As Single
    Y As Single
    Z As Single
    A As Single
End Type

Public Type CIELab
    l  As Single
    aa As Single
    bb As Single
    A  As Single
End Type

Public Enum CIELabLight
    D50_2 = 0
    D65_2 = 1
    D50_10 = 2
    D65_10 = 3
End Enum
Private Type XYZMatrix
    X(0 To 2) As Single
    Y(0 To 2) As Single
    Z(0 To 2) As Single
    R(0 To 2) As Single
    G(0 To 2) As Single
    B(0 To 2) As Single
End Type

Private M As XYZMatrix
Private CIELabLights(0 To 3) As XYZ

Public Sub Init()
    'https://github.com/PitPik/colorPicker/blob/master/colors.js
    '// Observer = 2° (CIE 1931), Illuminant = D65
    With M
        .X(0) = 0.4124564: .X(1) = 0.3575761:  .X(2) = 0.1804375
        .Y(0) = 0.2126729: .Y(1) = 0.7151522:  .Y(2) = 0.072175
        .Z(0) = 0.0193339: .Z(1) = 0.119192:   .Z(2) = 0.9503041
        .R(0) = 3.2404542: .R(1) = -1.5371385: .R(2) = -0.4985314
        .G(0) = -0.969266: .G(1) = 1.8760108:  .G(2) = 0.041556
        .B(0) = 0.0556434: .B(1) = -0.2040259: .B(2) = 1.0572252
    End With
'    CIELabLights(CIELabLight.D50_2) = XYZ(96.422, 100, 82.521)
'    CIELabLights(CIELabLight.D65_2) = XYZ(95.047, 100, 108.883)
'    CIELabLights(CIELabLight.D50_10) = XYZ(96.72, 100, 81.427)
'    CIELabLights(CIELabLight.D65_10) = XYZ(94.811, 100, 107.304)
    
    CIELabLights(CIELabLight.D50_2) = XYZ(0.96422, 1, 0.82521)
    CIELabLights(CIELabLight.D65_2) = XYZ(0.95047, 1, 1.08883)
    CIELabLights(CIELabLight.D50_10) = XYZ(0.9672, 1, 0.81427)
    CIELabLights(CIELabLight.D65_10) = XYZ(0.94811, 1, 1.07304)
End Sub

' #################### ' Single ' #################### '
Public Function FloatS_TryParse(ByVal s As String, v_out As Single) As Boolean
Try: On Error GoTo Catch
    v_out = CSng(Val(Replace(s, ",", ".")))
    FloatS_TryParse = True
Catch:
End Function

' #################### ' Byte ' #################### '
Public Function Byte_TryParse(ByVal s As String, v_out As Byte) As Boolean
Try: On Error GoTo Catch
    If Not IsNumeric(s) Then Exit Function
    v_out = CByte(s)
    Byte_TryParse = True
Catch:
End Function

' #################### ' String ' #################### '
Private Function Hex2(ByVal B As Byte) As String
    Hex2 = Hex(B): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

' #################### ' Math ' #################### '
Public Function MinB(ByVal V1 As Byte, ByVal V2 As Byte) As Byte
    If V1 < V2 Then MinB = V1 Else MinB = V2
End Function
Public Function MinB3(ByVal V1 As Byte, ByVal V2 As Byte, ByVal V3 As Byte) As Byte
    If V1 < V2 Then
        If V1 < V3 Then MinB3 = V1 Else MinB3 = V3
    Else
        If V2 < V3 Then MinB3 = V2 Else MinB3 = V3
    End If
End Function
Public Function MaxB(ByVal V1 As Byte, ByVal V2 As Byte) As Byte
    If V1 > V2 Then MaxB = V1 Else MaxB = V2
End Function
Public Function MaxB3(ByVal V1 As Byte, ByVal V2 As Byte, ByVal V3 As Byte) As Byte
    If V1 > V2 Then
        If V1 > V3 Then MaxB3 = V1 Else MaxB3 = V3
    Else
        If V2 > V3 Then MaxB3 = V2 Else MaxB3 = V3
    End If
End Function

Public Function MinS(V1 As Single, V2 As Single) As Single
    If V1 < V2 Then MinS = V1 Else MinS = V2
End Function
Public Function MinS3(V1 As Single, V2 As Single, V3 As Single) As Single
    If V1 < V2 Then
        If V1 < V3 Then MinS3 = V1 Else MinS3 = V3
    Else
        If V2 < V3 Then MinS3 = V2 Else MinS3 = V3
    End If
End Function
Public Function MaxS(V1 As Single, V2 As Single) As Single
    If V1 > V2 Then MaxS = V1 Else MaxS = V2
End Function
Public Function MaxS3(V1 As Single, V2 As Single, V3 As Single) As Single
    If V1 > V2 Then
        If V1 > V3 Then MaxS3 = V1 Else MaxS3 = V3
    Else
        If V2 > V3 Then MaxS3 = V2 Else MaxS3 = V3
    End If
End Function

'Public Function MinD(V1 As Double, V2 As Double) As Double
'    If V1 < V2 Then MinD = V1 Else MinD = V2
'End Function
'Public Function MaxD(V1 As Double, V2 As Double) As Double
'    If V1 > V2 Then MaxD = V1 Else MaxD = V2
'End Function

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
Public Function RGBA(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal A As Byte) As RGBA
    With RGBA: .R = R: .G = G: .B = B: .A = A: End With
End Function
Public Function RGBA_Read(this_out As RGBA, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Byte, s As String
    With this_out
        s = TB_R.Text: If Byte_TryParse(s, V) Then .R = V Else err_out = s: Exit Function
        s = TB_G.Text: If Byte_TryParse(s, V) Then .G = V Else err_out = s: Exit Function
        s = TB_B.Text: If Byte_TryParse(s, V) Then .B = V Else err_out = s: Exit Function
        s = TB_A.Text: If Byte_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    RGBA_Read = True
End Function
Public Function RGBA_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBA)
    With this: TB_R.Text = .R: TB_G.Text = .G: TB_B.Text = .B: TB_A.Text = .A: End With
End Function

'https://en.wikipedia.org/wiki/Color_difference
Public Function RGBA_EuclidRMean(this As RGBA, other As RGBA) As Double
    Dim dR As Double: dR = CDbl(this.R) - CDbl(other.R)
    Dim dG As Double: dG = CDbl(this.G) - CDbl(other.G)
    Dim dB As Double: dB = CDbl(this.B) - CDbl(other.B)
    Dim sR As Double: sR = 0.5 * (CDbl(this.R) + CDbl(other.R))
    RGBA_EuclidRMean = Math.Sqr((2 + sR / 256#) * dR * dR + 4 * dG * dG + (2 + (255# - sR) / 256) * dB * dB)
End Function

Public Function RGBA_ToARGB(this As RGBA) As ARGB
    With RGBA_ToARGB: .R = this.R: .G = this.G: .B = this.B: .A = this.A: End With
End Function
Public Function RGBA_ToLngColor(this As RGBA) As LngColor
    LSet RGBA_ToLngColor = this
End Function
Public Function RGBA_ToWebHex(this As RGBA) As String
    With this: RGBA_ToWebHex = "#" & Hex2(.A) & Hex2(.R) & Hex2(.G) & Hex2(.B): End With
End Function

Public Function RGBA_ParseWebHex(ByVal HashtagColor As String) As RGBA
    If Left(HashtagColor, 1) <> "#" Then Exit Function
    HashtagColor = Mid$(HashtagColor, 2)
    Dim s As String: s = Mid$(HashtagColor, 1, 2)
    With RGBA_ParseWebHex
        If 7 < Len(HashtagColor) Then     'ARGB
            .A = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .R = CByte("&H" & s): s = Mid$(HashtagColor, 5, 2)
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 7, 2)
            .B = CByte("&H" & s)
        ElseIf 5 < Len(HashtagColor) Then 'RGB
            .R = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 5, 2)
            .B = CByte("&H" & s)
        ElseIf 3 < Len(HashtagColor) Then 'GB
            .G = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
            .B = CByte("&H" & s)
        ElseIf 1 < Len(HashtagColor) Then 'B
            .B = CByte("&H" & s)
        End If
    End With
End Function

Public Function RGBA_ToRGBAf(this As RGBA) As RGBAf
    With this: RGBA_ToRGBAf = RGBAf(.R / 255, .G / 255, .B / 255, .A / 255): End With
End Function

Public Function RGBA_ToCMYK(this As RGBA) As CMYK
    RGBA_ToCMYK = RGBAf_ToCMYK(RGBA_ToRGBAf(this))
End Function

'Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
'Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Public Function RGBA_ToHSLA(this As RGBA) As HSLA
    Dim l As LngColor: l = RGBA_ToLngColor(this)
    Dim iiH As Integer, iiL As Integer, iiS As Integer
    With RGBA_ToHSLA
        .A = this.A
        ColorRGBToHLS l.Value, iiH, iiL, iiS
        .H = CByte(iiH)
        .s = CByte(iiS)
        .l = CByte(iiL)
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

' #################### ' ARGB ' #################### '
'Public Type ARGB
'    A As Byte '0..255
'    R As Byte '0..255
'    G As Byte '0..255
'    B As Byte '0..255
'End Type
Public Function ARGB_ToRGBA(this As ARGB) As RGBA
    With this: ARGB_ToRGBA.R = .R: ARGB_ToRGBA.G = .G: ARGB_ToRGBA.B = .B: ARGB_ToRGBA.A = .A: End With
End Function

' #################### ' RGBAf ' #################### '
'Public Type RGBAf
'    R As Single '0..1
'    G As Single '0..1
'    B As Single '0..1
'    A As Single '0..1
'End Type
Public Function RGBAf(R As Single, G As Single, B As Single, A As Single) As RGBAf
    With RGBAf: .R = R: .G = G: .B = B: .A = A: End With
End Function
Public Function RGBAf_Read(this_out As RGBAf, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_R.Text: If FloatS_TryParse(s, V) Then .R = V Else err_out = s: Exit Function
        s = TB_G.Text: If FloatS_TryParse(s, V) Then .G = V Else err_out = s: Exit Function
        s = TB_B.Text: If FloatS_TryParse(s, V) Then .B = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    RGBAf_Read = True
End Function
Public Function RGBAf_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBAf)
    Dim fmt As String: fmt = "0.#####"
    With this
        TB_R.Text = Format(.R, fmt)
        TB_G.Text = Format(.G, fmt)
        TB_B.Text = Format(.B, fmt)
        TB_A.Text = Format(.A, fmt)
    End With
End Function

Public Function RGBAf_EuclidRMean(this As RGBAf, other As RGBAf) As Double
    Dim dR As Double: dR = this.R - other.R
    Dim dG As Double: dG = this.G - other.G
    Dim dB As Double: dB = this.B - other.B
    Dim sR As Double: sR = 0.5 * (this.R + other.R)
    RGBAf_EuclidRMean = Math.Sqr((2 + sR / 256) * dR * dR + 4 * dG * dG + (2 + (255 - sR) / 256) * dB * dB)
End Function

Public Function RGBAf_ToRGBA(this As RGBAf) As RGBA
    With this
        RGBAf_ToRGBA.R = CByte(.R * 255)
        RGBAf_ToRGBA.G = CByte(.G * 255)
        RGBAf_ToRGBA.B = CByte(.B * 255)
        RGBAf_ToRGBA.A = CByte(.A * 255)
    End With
End Function

Public Function RGBAf_ToCMYK(this As RGBAf) As CMYK
    With RGBAf_ToCMYK
        .A = this.A
        .C = 1 - this.R
        .M = 1 - this.G
        .Y = 1 - this.B
        .K = MinS3(.C, .M, .Y)
        If .K = 1 Then Exit Function
        Dim kf As Single: kf = 1 - .K
        .C = ((.C - .K) / kf)
        .M = ((.M - .K) / kf)
        .Y = ((.Y - .K) / kf)
    End With
End Function

'https://de.wikipedia.org/wiki/HSV-Farbraum
'Gelb = RGBA(255, 255, 0, 0) = HSL(40, 240, 120)
Public Function RGBAf_ToHSLAf(this As RGBAf) As HSLAf
    With this
        Dim MaxRGB As Single: MaxRGB = MaxS3(.R, .G, .B)
        Dim MinRGB As Single: MinRGB = MinS3(.R, .G, .B)
    End With
    With RGBAf_ToHSLAf
        .A = this.A
        .l = (MaxRGB + MinRGB) / 2
        If MaxRGB = MinRGB Then
            .H = 0: .s = 0 'achromatic
        Else
            Dim Delta As Single: Delta = MaxRGB - MinRGB
            If .l > 0.5 Then
                .s = Delta / (2 - MaxRGB - MinRGB)
            Else
                .s = Delta / (MaxRGB + MinRGB)
            End If
            Select Case MaxRGB
            Case this.R: If this.G < this.B Then .H = 6 Else .H = 0
                         .H = (this.G - this.B) / Delta + .H
            Case this.G: .H = (this.B - this.R) / Delta + 2
            Case this.B: .H = (this.R - this.G) / Delta + 4
            End Select
            .H = .H / 6
        End If
    End With
End Function

Function RGBAf_ToHSV(this As RGBAf) As HSV
    With this
        Dim MaxRGB As Single: MaxRGB = MaxS3(.R, .G, .B)
        Dim MinRGB As Single: MinRGB = MinS3(.R, .G, .B)
    End With
    With RGBAf_ToHSV
        .A = this.A
        .V = MaxRGB
        Dim Delta As Single: Delta = MaxRGB - MinRGB
        If MaxRGB <> 0 Then .s = Delta / MaxRGB
        If MaxRGB = MinRGB Then
            .H = 0 'achromatic
        Else
            Select Case MaxRGB
            Case this.R: .H = (this.G - this.B) / Delta
                         If this.G < this.B Then .H = .H + 6
            Case this.G: .H = (this.B - this.R) / Delta + 2
            Case this.B: .H = (this.R - this.G) / Delta + 4
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
    Dim R As Single: R = this.R
    Dim G As Single: G = this.G
    Dim B As Single: B = this.B
    Dim N As Single: N = 0.04045
    
    If R > N Then R = ((R + 0.055) / 1.055) ^ (2.4) Else R = R / 12.92
    If G > N Then G = ((G + 0.055) / 1.055) ^ (2.4) Else G = G / 12.92
    If B > N Then B = ((B + 0.055) / 1.055) ^ (2.4) Else B = B / 12.92
    
    With RGBAf_ToXYZ
        .X = R * M.X(0) + G * M.X(1) + B * M.X(2)
        .Y = R * M.Y(0) + G * M.Y(1) + B * M.Y(2)
        .Z = R * M.Z(0) + G * M.Z(1) + B * M.Z(2)
        .A = this.A
    End With
End Function

' #################### ' CMYK ' #################### '
'Public Type CMYK
'    c As Single '0..1
'    M As Single '0..1
'    Y As Single '0..1
'    K As Single '0..1
'    A As Single '0..1
'End Type
Public Function CMYK(ByVal C As Single, ByVal M As Single, ByVal Y As Single, ByVal K As Single, ByVal A As Single) As CMYK
    With CMYK: .C = C: .M = M: .Y = Y: .K = K: .A = A: End With
End Function

Public Function CMYK_Read(this_out As CMYK, TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_C.Text: If FloatS_TryParse(s, V) Then .C = V Else err_out = s: Exit Function
        s = TB_M.Text: If FloatS_TryParse(s, V) Then .M = V Else err_out = s: Exit Function
        s = TB_Y.Text: If FloatS_TryParse(s, V) Then .Y = V Else err_out = s: Exit Function
        s = TB_K.Text: If FloatS_TryParse(s, V) Then .K = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    CMYK_Read = True
End Function
Public Function CMYK_ToView(TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, this As CMYK)
    With this
        TB_C.Text = Format(.C, "0.#####")
        TB_M.Text = Format(.M, "0.#####")
        TB_Y.Text = Format(.Y, "0.#####")
        TB_K.Text = Format(.K, "0.#####")
        TB_A.Text = Format(.A, "0.#####")
    End With
End Function

Public Function CMYK_Euclidean(this As CMYK, other As CMYK) As Double
    Dim dC As Double: dC = this.C - other.C
    Dim dM As Double: dM = this.M - other.M
    Dim dY As Double: dY = this.Y - other.Y
    Dim dK As Double: dK = this.K - other.K
    CMYK_Euclidean = Math.Sqr(dC * dC + dM * dM + dY * dY + dK * dK)
End Function

Public Function CMYK_ToRGBAf(this As CMYK) As RGBAf
    With this
        Dim kf As Single: kf = 1 - .K
        CMYK_ToRGBAf.R = 1 - MinS(1, .C * kf + .K)
        CMYK_ToRGBAf.G = 1 - MinS(1, .M * kf + .K)
        CMYK_ToRGBAf.B = 1 - MinS(1, .Y * kf + .K)
        CMYK_ToRGBAf.A = .A
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

' #################### ' HSLA  ' #################### '
'Public Type HSLA
'    H As Byte '0..239
'    S As Byte '0..240
'    L As Byte '0..240
'    A As Byte '0..255
'End Type
Public Function HSLA_ToRGBA(this As HSLA) As RGBA
    Dim l As LngColor
    With this
        l.Value = ColorHLSToRGB(.H, .l, .s)
    End With
    HSLA_ToRGBA = LngColor_ToRGBA(l)
    HSLA_ToRGBA.A = this.A
End Function
Public Function HSLA_Read(this_out As HSLA, TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Byte, s As String
    With this_out
        s = TB_H.Text: If Byte_TryParse(s, V) Then .H = MinB(V, 240) Else err_out = s: Exit Function
        s = TB_S.Text: If Byte_TryParse(s, V) Then .s = MinB(V, 240) Else err_out = s: Exit Function
        s = TB_L.Text: If Byte_TryParse(s, V) Then .l = MinB(V, 240) Else err_out = s: Exit Function
        s = TB_A.Text: If Byte_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    HSLA_Read = True
End Function
Public Function HSLA_ToView(TBHSLA_H As TextBox, TBHSLA_S As TextBox, TBHSLA_L As TextBox, TBHSLA_A As TextBox, this As HSLA)
    With this
        TBHSLA_H.Text = .H
        TBHSLA_S.Text = .s
        TBHSLA_L.Text = .l
        TBHSLA_A.Text = .A
    End With
End Function

Public Function HSLA_Euclidean(this As HSLA, other As HSLA) As Double

End Function

Public Function HSLA_ToHSLAf(this As HSLA) As HSLAf
    With HSLA_ToHSLAf
        .H = this.H / 240
        .s = this.s / 240
        .l = this.l / 240
        .A = this.A / 255
    End With
End Function

' #################### ' HSLAf ' #################### '
'Public Type HSLAf
'    H As Single '0..1
'    S As Single '0..1
'    L As Single '0..1
'    A As Single '0..1
'End Type
Public Function HSLAf_ToRGBAf(this As HSLAf) As RGBAf
    With this
        HSLAf_ToRGBAf.A = .A
        If .s = 0 Then 'achromatic
            HSLAf_ToRGBAf.R = .l
            HSLAf_ToRGBAf.G = .l
            HSLAf_ToRGBAf.B = .l
        Else
            Dim q As Single: If .l < 0.5 Then q = .l * (1 + .s) Else q = .l + .s - .l * .s
            Dim p As Single: p = 2 * .l - q
            HSLAf_ToRGBAf.R = Hue_ToRGB(p, q, .H + 1 / 3)
            HSLAf_ToRGBAf.G = Hue_ToRGB(p, q, .H)
            HSLAf_ToRGBAf.B = Hue_ToRGB(p, q, .H - 1 / 3)
        End If
    End With
End Function
Public Function Hue_ToRGB(p As Single, q As Single, t As Single) As Single
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
    Dim V As Single, s As String
    With this_out
        s = TB_H.Text: If FloatS_TryParse(s, V) Then .H = V Else err_out = s: Exit Function
        s = TB_S.Text: If FloatS_TryParse(s, V) Then .s = V Else err_out = s: Exit Function
        s = TB_L.Text: If FloatS_TryParse(s, V) Then .l = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    HSLAf_Read = True
End Function
Public Function HSLAf_ToView(TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, this As HSLAf)
    With this
        TB_H.Text = Format(.H, "0.#####")
        TB_S.Text = Format(.s, "0.#####")
        TB_L.Text = Format(.l, "0.#####")
        TB_A.Text = Format(.A, "0.#####")
    End With
End Function

' #################### ' HSV ' #################### '
'https://www.rapidtables.com/convert/color/hsv-to-rgb.html
Public Function HSV_ToRGBAf(this As HSV) As RGBAf
    With this
        Dim i As Single: i = CSng(Int(.H * 6)) 'Floor
        Dim f As Single: f = .H * 6 - i
        Dim p As Single: p = .V * (1 - .s)
        Dim q As Single: q = .V * (1 - f * .s)
        Dim t As Single: t = .V * (1 - (1 - f) * .s)
    End With
    With HSV_ToRGBAf
        .A = this.A
        Select Case i Mod 6
        Case 0: .R = this.V: .G = t:      .B = p
        Case 1: .R = q:      .G = this.V: .B = p
        Case 2: .R = p:      .G = this.V: .B = t
        Case 3: .R = p:      .G = q:      .B = this.V
        Case 4: .R = t:      .G = p:      .B = this.V
        Case 5: .R = this.V: .G = p:      .B = q
        End Select
    End With
End Function

Public Function HSV_Read(this_out As HSV, TB_H As TextBox, TB_S As TextBox, TB_V As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_H.Text: If FloatS_TryParse(s, V) Then .H = V Else err_out = s: Exit Function
        s = TB_S.Text: If FloatS_TryParse(s, V) Then .s = V Else err_out = s: Exit Function
        s = TB_V.Text: If FloatS_TryParse(s, V) Then .V = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    HSV_Read = True
End Function
Public Function HSV_ToView(TB_H As TextBox, TB_S As TextBox, TB_V As TextBox, TB_A As TextBox, this As HSV)
    With this
        TB_H.Text = Format(.H, "0.#####")
        TB_S.Text = Format(.s, "0.#####")
        TB_V.Text = Format(.V, "0.#####")
        TB_A.Text = Format(.A, "0.#####")
    End With
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
Public Function XYZ(ByVal aX As Single, ByVal aY As Single, ByVal aZ As Single) As XYZ
    With XYZ: .X = aX: .Y = aY: .Z = aZ: .A = 1: End With
End Function
Public Function XYZ_Read(this_out As XYZ, TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_X.Text: If FloatS_TryParse(s, V) Then .X = V Else err_out = s: Exit Function
        s = TB_Y.Text: If FloatS_TryParse(s, V) Then .Y = V Else err_out = s: Exit Function
        s = TB_Z.Text: If FloatS_TryParse(s, V) Then .Z = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .A = V Else err_out = s: Exit Function
    End With
    XYZ_Read = True
End Function
Public Function XYZ_ToView(TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, this As XYZ)
    With this
        TB_X.Text = Format(.X, "0.#####")
        TB_Y.Text = Format(.Y, "0.#####")
        TB_Z.Text = Format(.Z, "0.#####")
        TB_A.Text = Format(.A, "0.#####")
    End With
End Function

Public Function XYZ_Euclidean(this As XYZ, other As XYZ) As Double
    Dim dX As Double: dX = this.X - other.X
    Dim dY As Double: dY = this.Y - other.Y
    Dim dZ As Double: dZ = this.Z - other.Z
    XYZ_Euclidean = Math.Sqr(dX * dX + dY * dY + dZ * dZ)
End Function

Public Function XYZ_ToRGBAf(this As XYZ) As RGBAf
    Dim X As Single: X = this.X
    Dim Y As Single: Y = this.Y
    Dim Z As Single: Z = this.Z
    Dim R As Single: R = X * M.R(0) + Y * M.R(1) + Z * M.R(2)
    Dim G As Single: G = X * M.G(0) + Y * M.G(1) + Z * M.G(2)
    Dim B As Single: B = X * M.B(0) + Y * M.B(1) + Z * M.B(2)
    Dim N As Single: N = 1 / 2.4
    Dim MM As Single: MM = 0.0031308
    
    If R > MM Then R = 1.055 * R ^ N - 0.055 Else R = 12.92 * R
    If G > MM Then G = 1.055 * G ^ N - 0.055 Else G = 12.92 * G
    If B > MM Then B = 1.055 * B ^ N - 0.055 Else B = 12.92 * B
    
    With XYZ_ToRGBAf
        .R = MinS(MaxS(R, 0), 1) 'limitValue 0..1
        .G = MinS(MaxS(G, 0), 1)
        .B = MinS(MaxS(B, 0), 1)
        .A = this.A
    End With
End Function

Public Function XYZ_ToCIELab(this As XYZ, Optional lighttype As CIELabLight = CIELabLight.D65_2) As CIELab
    'https://de.wikipedia.org/wiki/Lab-Farbraum
    Dim N As XYZ: N = CIELabLights(lighttype)
    Dim XXN As Double: If N.X <> 0 Then XXN = this.X / N.X
    Dim YYN As Double: If N.Y <> 0 Then YYN = this.Y / N.Y
    Dim ZZN As Double: If N.Z <> 0 Then ZZN = this.Z / N.Z
    Dim root3_XXN As Double: If XXN < 216 / 24389 Then root3_XXN = 1 / 116 * (24389 / 27 * XXN + 16) Else root3_XXN = (this.X / N.X) ^ (1 / 3)
    Dim root3_YYN As Double: If YYN < 216 / 24389 Then root3_YYN = 1 / 116 * (24389 / 27 * YYN + 16) Else root3_YYN = (this.Y / N.Y) ^ (1 / 3)
    Dim root3_ZZN As Double: If ZZN < 216 / 24389 Then root3_ZZN = 1 / 116 * (24389 / 27 * ZZN + 16) Else root3_ZZN = (this.Z / N.Z) ^ (1 / 3)
    With XYZ_ToCIELab
        .l = 116 * root3_YYN - 16
        .aa = 500 * (root3_XXN - root3_YYN)
        .bb = 200 * (root3_YYN - root3_ZZN)
        .A = this.A
    End With
End Function

' #################### ' CIELab ' #################### '
'https://de.wikipedia.org/wiki/Lab-Farbraum
Public Function CIELab(ByVal l As Single, ByVal aa As Single, ByVal bb As Single) As CIELab
    With CIELab: .l = l: .aa = aa: .bb = bb: .A = 1: End With
End Function
Public Function CIELab_Read(this_out As CIELab, TB_L As TextBox, TB_aa As TextBox, TB_bb As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_L.Text:  If FloatS_TryParse(s, V) Then .l = V Else err_out = s:  Exit Function
        s = TB_aa.Text: If FloatS_TryParse(s, V) Then .aa = V Else err_out = s: Exit Function
        s = TB_bb.Text: If FloatS_TryParse(s, V) Then .bb = V Else err_out = s: Exit Function
        s = TB_A.Text:  If FloatS_TryParse(s, V) Then .A = V Else err_out = s:  Exit Function
    End With
    CIELab_Read = True
End Function
Public Function CIELab_ToView(TB_L As TextBox, TB_aa As TextBox, TB_bb As TextBox, TB_A As TextBox, this As CIELab)
    With this
        TB_L.Text = Format(.l, "0.#####")
        TB_aa.Text = Format(.aa, "0.#####")
        TB_bb.Text = Format(.bb, "0.#####")
        TB_A.Text = Format(.A, "0.#####")
    End With
End Function

Function CIELabLight_ToStr(ByVal l As CIELabLight) As String
    Dim s As String
    Select Case l
    Case CIELabLight.D50_2:  s = "D-50 2°"
    Case CIELabLight.D65_2:  s = "D-65 2°"
    Case CIELabLight.D50_10: s = "D-50 10°"
    Case CIELabLight.D65_10: s = "D-65 10°"
    End Select
    CIELabLight_ToStr = s
End Function

Public Sub CIELabLight_ToCmb(aCBLB As ComboBox)
    Dim i As Long, l As CIELabLight
    With aCBLB
        .Clear
        For i = 0 To 3
            l = i
            .AddItem CIELabLight_ToStr(l)
        Next
        .ListIndex = 3
    End With
End Sub

Public Function CIELab_ToXYZ(this As CIELab, Optional ByVal lighttype As CIELabLight = CIELabLight.D65_2) As XYZ
    '
End Function
