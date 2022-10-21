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
    a As Byte '0..255
End Type
Public Type ARGB
    a As Byte '0..255
    R As Byte '0..255
    G As Byte '0..255
    B As Byte '0..255
End Type
Public Type RGBAf
    R As Single '0..1
    G As Single '0..1
    B As Single '0..1
    a As Single '0..1
End Type
Public Type CMYK
    c As Single '0..1
    M As Single '0..1
    Y As Single '0..1
    K As Single '0..1
    a As Single '0..1
End Type
Public Type HSL
    H As Single '0..1
    s As Single '0..1
    L As Single '0..1
    a As Single '0..1
End Type
Public Type HSV
    H As Single '0..1
    s As Single '0..1
    V As Single '0..1
    a As Single '0..1
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
    a As Single
End Type

Private Type XYZMatrix
    X(0 To 2) As Single
    Y(0 To 2) As Single
    Z(0 To 2) As Single
    R(0 To 2) As Single
    G(0 To 2) As Single
    B(0 To 2) As Single
End Type

Private M As XYZMatrix

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
Public Function RGBA(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal a As Byte) As RGBA
    With RGBA: .R = R: .G = G: .B = B: .a = a: End With
End Function
Public Function RGBA_Read(this_out As RGBA, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Byte, s As String
    With this_out
        s = TB_R.Text: If Byte_TryParse(s, V) Then .R = V Else err_out = s: Exit Function
        s = TB_G.Text: If Byte_TryParse(s, V) Then .G = V Else err_out = s: Exit Function
        s = TB_B.Text: If Byte_TryParse(s, V) Then .B = V Else err_out = s: Exit Function
        s = TB_A.Text: If Byte_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    RGBA_Read = True
End Function
Public Function RGBA_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBA)
    With this: TB_R.Text = .R: TB_G.Text = .G: TB_B.Text = .B: TB_A.Text = .a: End With
End Function

Public Function RGBA_ToARGB(this As RGBA) As ARGB
    With RGBA_ToARGB: .R = this.R: .G = this.G: .B = this.B: .a = this.a: End With
End Function
Public Function RGBA_ToLngColor(this As RGBA) As LngColor
    LSet RGBA_ToLngColor = this
End Function
Public Function RGBA_ToWebHex(this As RGBA) As String
    With this: RGBA_ToWebHex = "#" & Hex2(.a) & Hex2(.R) & Hex2(.G) & Hex2(.B): End With
End Function

Public Function RGBA_ParseWebHex(ByVal HashtagColor As String) As RGBA
    If Left(HashtagColor, 1) <> "#" Then Exit Function
    HashtagColor = Mid$(HashtagColor, 2)
    Dim s As String: s = Mid$(HashtagColor, 1, 2)
    With RGBA_ParseWebHex
        If 7 < Len(HashtagColor) Then     'ARGB
            .a = CByte("&H" & s): s = Mid$(HashtagColor, 3, 2)
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
    With this: RGBA_ToRGBAf = RGBAf(.R / 255, .G / 255, .B / 255, .a / 255): End With
End Function

Public Function RGBA_ToCMYK(this As RGBA) As CMYK
    RGBA_ToCMYK = RGBAf_ToCMYK(RGBA_ToRGBAf(this))
End Function

'https://de.wikipedia.org/wiki/HSV-Farbraum
'Gelb = RGBA(255, 255, 0, 0) = HSL(40, 240, 120)
Public Function RGBA_ToHSL(this As RGBA) As HSL
    RGBA_ToHSL = RGBAf_ToHSL(RGBA_ToRGBAf(this))
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
Public Function ARGB_ToRGBA(this As ARGB) As RGBA
    With this: ARGB_ToRGBA.R = .R: ARGB_ToRGBA.G = .G: ARGB_ToRGBA.B = .B: ARGB_ToRGBA.a = .a: End With
End Function

' #################### ' RGBAf ' #################### '
Public Function RGBAf(R As Single, G As Single, B As Single, a As Single) As RGBAf
    With RGBAf: .R = R: .G = G: .B = B: .a = a: End With
End Function
Public Function RGBAf_Read(this_out As RGBAf, TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_R.Text: If FloatS_TryParse(s, V) Then .R = V Else err_out = s: Exit Function
        s = TB_G.Text: If FloatS_TryParse(s, V) Then .G = V Else err_out = s: Exit Function
        s = TB_B.Text: If FloatS_TryParse(s, V) Then .B = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    RGBAf_Read = True
End Function
Public Function RGBAf_ToView(TB_R As TextBox, TB_G As TextBox, TB_B As TextBox, TB_A As TextBox, this As RGBAf)
    Dim fmt As String: fmt = "0.#####"
    With this
        TB_R.Text = Format(.R, fmt)
        TB_G.Text = Format(.G, fmt)
        TB_B.Text = Format(.B, fmt)
        TB_A.Text = Format(.a, fmt)
    End With
End Function

Public Function RGBAf_ToRGBA(this As RGBAf) As RGBA
    With this
        RGBAf_ToRGBA.R = CByte(.R * 255)
        RGBAf_ToRGBA.G = CByte(.G * 255)
        RGBAf_ToRGBA.B = CByte(.B * 255)
        RGBAf_ToRGBA.a = CByte(.a * 255)
    End With
End Function

Public Function RGBAf_ToCMYK(this As RGBAf) As CMYK
    With RGBAf_ToCMYK
        .a = this.a
        .c = 1 - this.R
        .M = 1 - this.G
        .Y = 1 - this.B
        .K = MinS3(.c, .M, .Y)
        If .K = 1 Then Exit Function
        Dim kf As Single: kf = 1 - .K
        .c = ((.c - .K) / kf)
        .M = ((.M - .K) / kf)
        .Y = ((.Y - .K) / kf)
    End With
End Function

'https://de.wikipedia.org/wiki/HSV-Farbraum
'Gelb = RGBA(255, 255, 0, 0) = HSL(40, 240, 120)
Public Function RGBAf_ToHSL(this As RGBAf) As HSL
    With this
        Dim MaxRGB As Single: MaxRGB = MaxS3(.R, .G, .B)
        Dim MinRGB As Single: MinRGB = MinS3(.R, .G, .B)
    End With
    With RGBAf_ToHSL
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
        .a = this.a
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
    Dim n As Single: n = 0.04045
    
    If R > n Then R = ((R + 0.055) / 1.055) ^ (2.4) Else R = R / 12.92
    If G > n Then G = ((G + 0.055) / 1.055) ^ (2.4) Else G = G / 12.92
    If B > n Then B = ((B + 0.055) / 1.055) ^ (2.4) Else B = B / 12.92
    
    With RGBAf_ToXYZ
        .X = R * M.X(0) + G * M.X(1) + B * M.X(2)
        .Y = R * M.Y(0) + G * M.Y(1) + B * M.Y(2)
        .Z = R * M.Z(0) + G * M.Z(1) + B * M.Z(2)
    End With
End Function

' #################### ' CMYK ' #################### '
Public Function CMYK(c As Single, M As Single, Y As Single, K As Single, a As Single) As CMYK
    With CMYK: .c = c: .M = M: .Y = Y: .K = K: .a = a: End With
End Function

Public Function CMYK_Read(this_out As CMYK, TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_C.Text: If FloatS_TryParse(s, V) Then .c = V Else err_out = s: Exit Function
        s = TB_M.Text: If FloatS_TryParse(s, V) Then .M = V Else err_out = s: Exit Function
        s = TB_Y.Text: If FloatS_TryParse(s, V) Then .Y = V Else err_out = s: Exit Function
        s = TB_K.Text: If FloatS_TryParse(s, V) Then .K = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    CMYK_Read = True
End Function
Public Function CMYK_ToView(TB_C As TextBox, TB_M As TextBox, TB_Y As TextBox, TB_K As TextBox, TB_A As TextBox, this As CMYK)
    With this
        TB_C.Text = Format(.c, "0.#####")
        TB_M.Text = Format(.M, "0.#####")
        TB_Y.Text = Format(.Y, "0.#####")
        TB_K.Text = Format(.K, "0.#####")
        TB_A.Text = Format(.a, "0.#####")
    End With
End Function

Public Function CMYK_ToRGBAf(this As CMYK) As RGBAf
    With this
        Dim kf As Single: kf = 1 - .K
        CMYK_ToRGBAf.R = 1 - MinS(1, .c * kf + .K)
        CMYK_ToRGBAf.G = 1 - MinS(1, .M * kf + .K)
        CMYK_ToRGBAf.B = 1 - MinS(1, .Y * kf + .K)
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

' #################### ' HSL ' #################### '
Public Function HSL_ToRGBAf(this As HSL) As RGBAf
    With this
        HSL_ToRGBAf.a = .a
        If .s = 0 Then 'achromatic
            HSL_ToRGBAf.R = .L
            HSL_ToRGBAf.G = .L
            HSL_ToRGBAf.B = .L
        Else
            Dim q As Single: If .L < 0.5 Then q = .L * (1 + .s) Else q = .L + .s - .L * .s
            Dim p As Single: p = 2 * .L - q
            HSL_ToRGBAf.R = Hue_ToRGB(p, q, .H + 1 / 3)
            HSL_ToRGBAf.G = Hue_ToRGB(p, q, .H)
            HSL_ToRGBAf.B = Hue_ToRGB(p, q, .H - 1 / 3)
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

Public Function HSL_Read(this_out As HSL, TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_H.Text: If FloatS_TryParse(s, V) Then .H = V Else err_out = s: Exit Function
        s = TB_S.Text: If FloatS_TryParse(s, V) Then .s = V Else err_out = s: Exit Function
        s = TB_L.Text: If FloatS_TryParse(s, V) Then .L = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    HSL_Read = True
End Function
Public Function HSL_ToView(TB_H As TextBox, TB_S As TextBox, TB_L As TextBox, TB_A As TextBox, this As HSL)
    With this
        TB_H.Text = Format(.H, "0.#####")
        TB_S.Text = Format(.s, "0.#####")
        TB_L.Text = Format(.L, "0.#####")
        TB_A.Text = Format(.a, "0.#####")
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
        .a = this.a
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
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    HSV_Read = True
End Function
Public Function HSV_ToView(TB_H As TextBox, TB_S As TextBox, TB_V As TextBox, TB_A As TextBox, this As HSV)
    With this
        TB_H.Text = Format(.H, "0.#####")
        TB_S.Text = Format(.s, "0.#####")
        TB_V.Text = Format(.V, "0.#####")
        TB_A.Text = Format(.a, "0.#####")
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
    With XYZ: .X = aX: .Y = aY: .Z = aZ: .a = 1: End With
End Function
Public Function XYZ_ToRGBAf(this As XYZ) As RGBAf
    Dim X As Single: X = this.X
    Dim Y As Single: Y = this.Y
    Dim Z As Single: Z = this.Z
    Dim R As Single: R = X * M.R(0) + Y * M.R(1) + Z * M.R(2)
    Dim G As Single: G = X * M.G(0) + Y * M.G(1) + Z * M.G(2)
    Dim B As Single: B = X * M.B(0) + Y * M.B(1) + Z * M.B(2)
    Dim n As Single: n = 1 / 2.4
    Dim MM As Single: MM = 0.0031308
    
    If R > MM Then R = 1.055 * R ^ n - 0.055 Else R = 12.92 * R
    If G > MM Then G = 1.055 * G ^ n - 0.055 Else G = 12.92 * G
    If B > MM Then B = 1.055 * B ^ n - 0.055 Else B = 12.92 * B
    
    With XYZ_ToRGBAf
        .R = MinS(MaxS(R, 0), 1) 'limitValue 0..1
        .G = MinS(MaxS(G, 0), 1)
        .B = MinS(MaxS(B, 0), 1)
        .a = this.a
    End With
End Function

Public Function XYZ_Read(this_out As XYZ, TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, err_out As String) As Boolean
    Dim V As Single, s As String
    With this_out
        s = TB_X.Text: If FloatS_TryParse(s, V) Then .X = V Else err_out = s: Exit Function
        s = TB_Y.Text: If FloatS_TryParse(s, V) Then .Y = V Else err_out = s: Exit Function
        s = TB_Z.Text: If FloatS_TryParse(s, V) Then .Z = V Else err_out = s: Exit Function
        s = TB_A.Text: If FloatS_TryParse(s, V) Then .a = V Else err_out = s: Exit Function
    End With
    XYZ_Read = True
End Function
Public Function XYZ_ToView(TB_X As TextBox, TB_Y As TextBox, TB_Z As TextBox, TB_A As TextBox, this As XYZ)
    With this
        TB_X.Text = Format(.X, "0.#####")
        TB_Y.Text = Format(.Y, "0.#####")
        TB_Z.Text = Format(.Z, "0.#####")
        TB_A.Text = Format(.a, "0.#####")
    End With
End Function

