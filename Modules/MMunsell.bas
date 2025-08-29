Attribute VB_Name = "MMunsell"
Option Explicit

Public Const Count_HuePrefix As Long = 10  '  R - RP
Public Const Count_HueValue  As Long = 4   '2.5 - 10.0
Public Const Count_ValValue  As Long = 9   '  1 -  9
Public Const Count_ChromaMax As Long = 19  '  2 - 38

'Munsell aka Pantone Colors
Public Enum EHuePrefix
    Hue_R = 1
    Hue_YR '=2
    Hue_Y  '=3
    Hue_GY '=4
    Hue_G  '=5
    Hue_BG '=6
    Hue_B  '=7
    Hue_PB '=8
    Hue_P  '=9
    Hue_RP '=10
End Enum
'no need for a
Public Type TMunsellColor
    HuePrefix As Byte ' 1 ' As EHuePrefix
    HueValue  As Byte ' 1 ' Hue Values are 2.5, 5.0, 7.5, 10.0 = 1 (*2,5), 2 (*2,5), 3 (*2,5), 4 (*2,5)
    ValValue  As Byte ' 1 ' Values are: 1-9
    Chroma    As Byte ' 1 ' Values are: 2,4,6,8,10,12,14,16,18,20,22,24,26,28,30,32,34,36,38
    RGBA      As RGBA ' 4
End Type          'Sum: 8
Private MunsellColorZero  As TMunsellColor
Private m_MunsellColors() As TMunsellColor 'in file sorted to Chroma-Value

'Chroma-Value-Types:
Public Type ValValueCV
    Chromas() As TMunsellColor 'Chroma
End Type
Public Type HueValueCV
    ValValues()   As ValValueCV
End Type
Public Type HuePrefixCV
    HueValues() As HueValueCV
End Type
Private m_HuePrefixesCV() As HuePrefixCV

'Chroma-HuePrefrix-Types:
Public Type HuePrefixCHP
    Chromas() As TMunsellColor
End Type
Public Type HueValueCHP
    HuePrefixes() As HuePrefixCHP
End Type
Public Type ValValueCHP
    HueValues() As HueValueCHP
End Type
Private m_ValValuesCHP() As ValValueCHP

'Chroma-Hue-Types:
Public Type HueValueCH
    Chromas() As TMunsellColor
End Type
Public Type ValValueCH
    HueValues() As HueValueCH
End Type
Public Type HuePrefixCH
    ValValues() As ValValueCH
End Type
Private m_HuePrefixesCH() As HuePrefixCH

Private Sub Filter()
    FilterChromaValues
    FilterChromaHuePrefixes
    FilterChromaHues
End Sub

Private Sub FilterChromaValues()
    Dim i As Long, j As Long, k As Long, L As Long
    ReDim m_HuePrefixesCV(1 To Count_HuePrefix)
    For i = 1 To Count_HuePrefix
        ReDim m_HuePrefixesCV(i).HueValues(1 To Count_HueValue)
        For j = 1 To Count_HueValue
            ReDim m_HuePrefixesCV(i).HueValues(j).ValValues(1 To Count_ValValue)
            For k = 1 To Count_ValValue
                ReDim m_HuePrefixesCV(i).HueValues(j).ValValues(k).Chromas(1 To Count_ChromaMax)
            Next
        Next
    Next
    i = 1: j = 1: k = 1: L = 1
    Dim c As Long, u As Long: u = UBound(m_MunsellColors)
    Dim mc0 As TMunsellColor, mc1 As TMunsellColor
    For c = 0 To u - 1
        mc0 = m_MunsellColors(c)
        mc1 = m_MunsellColors(c + 1)
        m_HuePrefixesCV(i).HueValues(j).ValValues(k).Chromas(L) = mc0
        If mc1.Chroma = 2 Then
            ReDim Preserve m_HuePrefixesCV(i).HueValues(j).ValValues(k).Chromas(1 To L)
            L = 1
            If mc1.ValValue = 1 Then
                k = 1
                If mc1.HueValue = 1 Then
                    j = 1
                    If mc1.HuePrefix = 1 Then
                        i = 1
                    Else
                        i = i + 1
                    End If
                Else
                    j = j + 1
                End If
            Else
                k = k + 1
            End If
        Else
            L = L + 1
        End If
    Next
    m_HuePrefixesCV(i).HueValues(j).ValValues(k).Chromas(L) = mc1
    ReDim Preserve m_HuePrefixesCV(i).HueValues(j).ValValues(k).Chromas(1 To L)
End Sub

Private Sub FilterChromaHuePrefixes()
    Dim i As Long, j As Long, k As Long, L As Long
    ReDim m_ValValuesCHP(1 To Count_ValValue)
    For i = 1 To Count_ValValue
        ReDim m_ValValuesCHP(i).HueValues(1 To Count_HueValue)
        For j = 1 To Count_HueValue
            ReDim m_ValValuesCHP(i).HueValues(j).HuePrefixes(1 To Count_HuePrefix)
            For k = 1 To Count_HuePrefix
                ReDim m_ValValuesCHP(i).HueValues(j).HuePrefixes(k).Chromas(1 To Count_ChromaMax)
            Next
        Next
    Next
    Dim ucm As Long
    For i = 1 To Count_ValValue
        For j = 1 To Count_HueValue
            For k = 1 To Count_HuePrefix
                ucm = UBound(m_HuePrefixesCV(k).HueValues(j).ValValues(i).Chromas)
                For L = 1 To ucm
                    m_ValValuesCHP(i).HueValues(j).HuePrefixes(k).Chromas(L) = m_HuePrefixesCV(k).HueValues(j).ValValues(i).Chromas(L)
                Next
            Next
        Next
    Next
End Sub

Private Sub FilterChromaHues()
    Dim i As Long, j As Long, k As Long, L As Long
    ReDim m_HuePrefixesCH(1 To Count_HuePrefix)
    For i = 1 To Count_HuePrefix
        ReDim m_HuePrefixesCH(i).ValValues(1 To Count_ValValue)
        For j = 1 To Count_ValValue
            ReDim m_HuePrefixesCH(i).ValValues(j).HueValues(1 To Count_HueValue)
            For k = 1 To Count_HueValue
                ReDim m_HuePrefixesCH(i).ValValues(j).HueValues(k).Chromas(1 To Count_ChromaMax)
            Next
        Next
    Next
    Dim ucm As Long
    For i = 1 To Count_HuePrefix
        For j = 1 To Count_ValValue
            For k = 1 To Count_HueValue
                ucm = UBound(m_HuePrefixesCV(i).HueValues(k).ValValues(j).Chromas)
                For L = 1 To ucm
                    m_HuePrefixesCH(i).ValValues(j).HueValues(k).Chromas(L) = m_HuePrefixesCV(i).HueValues(k).ValValues(j).Chromas(L)
                Next
            Next
        Next
    Next
End Sub

Public Sub Init()
    
    'OK at first program launch there normally is no file "Munsell.bin"
    'so we want to read from resource and then write the data to a file
    'Once there is a file, we read from the file.
    'Why do we need a file?
    'Because the user should have the ability to change the data
    'If User makes any mistake, he should just delete the file.
    
    Dim FN As String: FN = "Munsell.bin"
    Dim AppPFN As PathFileName: Set AppPFN = MNew.PathFileName(App.Path, FN)
    Dim TmpPFN As PathFileName: Set TmpPFN = MNew.PathFileName(AppPFN.TempPath, FN)
    Dim PFN    As PathFileName
    Set PFN = IIf(AppPFN.Exists, AppPFN, IIf(TmpPFN.Exists, TmpPFN, Nothing))
    'Set PFN = IIf(TmpPFN.Exists, TmpPFN, IIf(AppPFN.Exists, AppPFN, Nothing))
    Dim ba() As Byte
    If PFN Is Nothing Then
        ba = LoadResData(10, "CUSTOM")
        Set PFN = AppPFN
        'Set PFN = TmpPFN
        If Not TryWriteToPFN(ba, PFN) Then
        
            MsgBox "Data read from resource, but could not write to file, maybe try again later." & vbCrLf & _
                    PFN.Value
            'Set PFN = AppPFN
            Set PFN = TmpPFN
            'Debug.Print PFN.Value
            If Not TryWriteToPFN(ba, PFN) Then
                'could not write data to file
                MsgBox "Data read from resource, but could not write to file, maybe try again later." & vbCrLf & _
                       PFN.Value
            End If
        End If
    Else
        'Debug.Print PFN.Value
        PFN.ReadAllBuffer ba
    End If
    If MMunsell.ReadFromMemoryStream(ba) Then
        Filter
        Exit Sub
    Else
        MsgBox "Could not read data, maybe try again later!"
    End If
End Sub

Private Function TryWriteToPFN(ba() As Byte, PFN As PathFileName) As Boolean
Try: On Error GoTo Catch
    PFN.WriteBytes ba
    TryWriteToPFN = True
    Exit Function
Catch:
End Function

' v ' ############################## ' v '    EHuePrefix    ' v ' ############################## ' v '
Public Property Get EHuePrefix_Name(ByVal E As EHuePrefix) As String
    Dim s As String
    Select Case E
    Case EHuePrefix.Hue_R:  s = "Red"
    Case EHuePrefix.Hue_YR: s = "Yellow-Red"
    Case EHuePrefix.Hue_Y:  s = "Yellow"
    Case EHuePrefix.Hue_GY: s = "Green-Yellow"
    Case EHuePrefix.Hue_G:  s = "Green"
    Case EHuePrefix.Hue_BG: s = "Blue-Green"
    Case EHuePrefix.Hue_B:  s = "Blue"
    Case EHuePrefix.Hue_PB: s = "Purple-Blue"
    Case EHuePrefix.Hue_P:  s = "Purple"
    Case EHuePrefix.Hue_RP: s = "Red-Purple"
    End Select
    EHuePrefix_Name = s
End Property

Public Function EHuePrefixName_TryParse(ByVal s As String, e_out As EHuePrefix) As Boolean
    EHuePrefixName_TryParse = True
    Select Case UCase(s)
    Case "R", "RED":           e_out = EHuePrefix.Hue_R
    Case "YR", "YELLOW-RED":   e_out = EHuePrefix.Hue_YR
    Case "Y", "YELLOW":        e_out = EHuePrefix.Hue_Y
    Case "GY", "GREEN-YELLOW": e_out = EHuePrefix.Hue_GY
    Case "G", "GREEN":         e_out = EHuePrefix.Hue_G
    Case "BG", "BLUE-GREEN":   e_out = EHuePrefix.Hue_BG
    Case "B", "BLUE":          e_out = EHuePrefix.Hue_B
    Case "PB", "PURPLE-BLUE":  e_out = EHuePrefix.Hue_PB
    Case "P", "PURPLE":        e_out = EHuePrefix.Hue_P
    Case "RP", "RED-PURPLE":   e_out = EHuePrefix.Hue_RP
    Case Else
        Dim i As Long, ci As String, c1 As String
        For i = 1 To Len(s) '- 1
            ci = Mid(s, i, 1)
            If i < Len(s) Then c1 = Mid(s, i + 1, 1)
            Select Case ci
            Case "R": If c1 = "P" Then e_out = EHuePrefix.Hue_RP Else e_out = EHuePrefix.Hue_R
            Case "Y": If c1 = "R" Then e_out = EHuePrefix.Hue_YR Else e_out = EHuePrefix.Hue_Y
            Case "G": If c1 = "Y" Then e_out = EHuePrefix.Hue_GY Else e_out = EHuePrefix.Hue_G
            Case "B": If c1 = "G" Then e_out = EHuePrefix.Hue_BG Else e_out = EHuePrefix.Hue_B
            Case "P": If c1 = "B" Then e_out = EHuePrefix.Hue_PB Else e_out = EHuePrefix.Hue_P
            Case Else
                'go ahead until eofstr
            End Select
        Next
        If e_out = 0 Then EHuePrefixName_TryParse = False
    End Select
End Function

Public Function EHuePrefix_ToStr(ByVal ehp As EHuePrefix) As String
    Dim s As String
    Select Case ehp
    Case EHuePrefix.Hue_R:  s = "R"
    Case EHuePrefix.Hue_YR: s = "YR"
    Case EHuePrefix.Hue_Y:  s = "Y"
    Case EHuePrefix.Hue_GY: s = "GY"
    Case EHuePrefix.Hue_G:  s = "G"
    Case EHuePrefix.Hue_BG: s = "BG"
    Case EHuePrefix.Hue_B:  s = "B"
    Case EHuePrefix.Hue_PB: s = "PB"
    Case EHuePrefix.Hue_P:  s = "P"
    Case EHuePrefix.Hue_RP: s = "RP"
    End Select
    EHuePrefix_ToStr = s
End Function
' ^ ' ############################## ' ^ '    EHuePrefix    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '     HueValue     ' v ' ############################## ' v '
Public Function HueValue_TryParse(ByVal s As String, hv_out As Byte) As Boolean
    HueValue_TryParse = True
    Select Case s
    Case "2", "2.", "2.5":         hv_out = 1
    Case "5", "5.", "5.0":         hv_out = 2
    Case "7", "7.", "7.5":         hv_out = 3
    Case "1", "10", "10.", "10.0": hv_out = 4
    Case Else:   HueValue_TryParse = False
    End Select
End Function

Public Function HueValue_ToStr(ByVal hv As Byte) As String
    Dim s As String
    Select Case hv
    Case 1: s = "2.5"
    Case 2: s = "5.0"
    Case 3: s = "7.5"
    Case 4: s = "10.0"
    End Select
    HueValue_ToStr = s 'Trim(Str(hv * 2.5))
End Function
' ^ ' ############################## ' ^ '     HueValue     ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '     ValValue     ' v ' ############################## ' v '
Public Function ValValue_TryParse(ByVal s As String, vv_out As Byte) As Boolean
    ValValue_TryParse = MString.Byte_TryParse(s, vv_out)
    If ValValue_TryParse Then ValValue_TryParse = 1 <= vv_out And vv_out <= 9
End Function

' ^ ' ############################## ' ^ '     ValValue     ' ^ ' ############################## ' ^ '

Public Sub EHuePrefix_ToCmb(aCmb As ComboBox)
    Dim i As Long:   aCmb.Clear
    For i = 1 To 10: aCmb.AddItem EHuePrefix_Name(i): Next
End Sub
Public Sub HueValue_ToCmb(aCmb As ComboBox)
    Dim i As Long:  aCmb.Clear
    For i = 1 To 4: aCmb.AddItem HueValue_ToStr(i): Next
End Sub

Public Sub ValValue_ToCmb(aCmb As ComboBox)
    Dim i As Long:  aCmb.Clear
    For i = 1 To 9: aCmb.AddItem CStr(i): Next
End Sub

Public Sub EHuePrefixHueValue_ToCmb(aCmb As ComboBox)
    Dim i As Long, j As Long:   aCmb.Clear
    Dim s As String
    For i = 1 To Count_HuePrefix
        s = EHuePrefix_Name(i)
        For j = 1 To Count_HueValue
            aCmb.AddItem s & " - " & HueValue_ToStr(j)
        Next
    Next
End Sub

Public Sub ValValueHueValue_ToCmb(aCmb As ComboBox)
    Dim i As Long, j As Long:   aCmb.Clear
    Dim s As String
    For i = 1 To Count_ValValue
        s = i 'ValValue_Name(i)
        For j = 1 To Count_HueValue
            aCmb.AddItem s & " - " & HueValue_ToStr(j)
        Next
    Next
End Sub

Public Sub EHuePrefixValValue_ToCmb(aCmb As ComboBox)
    Dim i As Long, j As Long:   aCmb.Clear
    Dim s As String
    For i = 1 To Count_HuePrefix
        s = EHuePrefix_Name(i)
        For j = 1 To Count_ValValue
            aCmb.AddItem s & " - " & CStr(j) 'ValValue_ToStr(j)
        Next
    Next
End Sub
'Public Sub Chroma_ToCmb(aCmb As ComboBox)
'    Dim i As Long:          aCmb.Clear
'    For i = 2 To 38 Step 2: aCmb.AddItem CStr(i): Next
'End Sub


'Public Function Byte_TryParse(ByVal s As String, Value_out As Byte) As Boolean
'Try: On Error GoTo Catch
'    Value_out = CByte(s)
'    Byte_TryParse = True
'Catch:
'End Function
'
'Public Function RGBA_TryParse(ByVal s As String, RGBA_out As RGBA) As Boolean
'Try: On Error GoTo Catch
'    Dim sa() As String: sa = Split(s, ",")
'    RGBA_out.R = CByte(sa(0))
'    RGBA_out.G = CByte(sa(1))
'    RGBA_out.B = CByte(sa(2))
'    RGBA_TryParse = True
'Catch:
'End Function
'
'Public Function RGBA_ToStr(RGBA As RGBA) As String
'    RGBA_ToStr = RGBA.R & "," & RGBA.G & "," & RGBA.B
'End Function

Public Property Get MunsellColors_ChromaValue(ByVal HuePrefix As EHuePrefix, ByVal HueValue As Byte) As HueValueCV ' TMunsellColor()
    MunsellColors_ChromaValue = m_HuePrefixesCV(HuePrefix).HueValues(HueValue)
End Property
Public Property Get MunsellColors_ChromaHuePrefix(ByVal ValValue As Byte, ByVal HueValue As Byte) As HueValueCHP 'HuePrefixCHP ' TMunsellColor()
    MunsellColors_ChromaHuePrefix = m_ValValuesCHP(ValValue).HueValues(HueValue)
End Property
Public Property Get MunsellColors_ChromaHue(ByVal HuePrefix As EHuePrefix, ByVal ValValue As Byte) As ValValueCH
    MunsellColors_ChromaHue = m_HuePrefixesCH(HuePrefix).ValValues(ValValue)
End Property

Public Property Get TMunsellColor_Key(this As TMunsellColor) As String
    'e.g. 5.0BG-5-22
    '5.0 : Hue
    ' BG : Hue-Prefix (=Blue-Green)
    '  5 : Value
    ' 22 : Chroma
    TMunsellColor_Key = MMunsell.HueValue_ToStr(this.HueValue) & MMunsell.EHuePrefix_ToStr(this.HuePrefix) & "-" & CStr(this.ValValue) & "-" & CStr(this.Chroma)
End Property

Public Function MunsellColors_ParseColorFromKey(ByVal aKey As String) As TMunsellColor
    'aKey could look like: "5.0BG-5-22"
    Dim mc As TMunsellColor
    Dim s  As String: s = Left(aKey, 1)
    Dim hv As Byte: If Not HueValue_TryParse(s, hv) Then Exit Function
    mc.HueValue = hv
    Dim hp As EHuePrefix: If Not EHuePrefixName_TryParse(aKey, hp) Then Exit Function
    mc.HuePrefix = hp
    Dim sa() As String: sa = Split(aKey, "-")
    s = sa(1)
    If IsNumeric(s) Then
        mc.ValValue = CByte(s)
    End If
    s = sa(2)
    If IsNumeric(s) Then
        mc.Chroma = CByte(s)
    End If
    MunsellColors_ParseColorFromKey = MunsellColors_GetItemByParams(mc)
    'wow we've done it really complicated, code could be so much shorter by using just a hashed Collection
End Function

Public Function MunsellColors_GetItemByParams(mc As TMunsellColor) As TMunsellColor
    Dim i As Long
    For i = 0 To UBound(m_MunsellColors)
        If TMunsellColor_EqualParams(m_MunsellColors(i), mc) Then
            MunsellColors_GetItemByParams = m_MunsellColors(i)
            Exit Function
        End If
    Next
End Function

Public Function TMunsellColor_EqualParams(mc1 As TMunsellColor, mc2 As TMunsellColor) As Boolean
    If mc1.HuePrefix <> mc2.HuePrefix Then Exit Function
    If mc1.HueValue <> mc2.HueValue Then Exit Function
    If mc1.ValValue <> mc2.ValValue Then Exit Function
    If mc1.Chroma <> mc2.Chroma Then Exit Function
    TMunsellColor_EqualParams = True
End Function

Public Function MunsellColors_ClosestColorTo(ByVal aColor As Long) As TMunsellColor
    Dim i As Long, i_minEd As Long, edi As Double
    Dim lc As LngColor: lc = MColor.LngColor(aColor)
    Dim ed0 As Double: ed0 = MColor.LngColor_EuclidRMean(LngColor((&HFFFFFF And MColor.RGBA_ToLngColor(m_MunsellColors(0).RGBA).Value)), lc)
    For i = 1 To UBound(m_MunsellColors) 'm_Count - 1
        edi = MColor.LngColor_EuclidRMean(LngColor((&HFFFFFF And MColor.RGBA_ToLngColor(m_MunsellColors(i).RGBA).Value)), lc)
        If edi < ed0 Then
            i_minEd = i
            ed0 = edi
        End If
    Next
    MunsellColors_ClosestColorTo = m_MunsellColors(i_minEd)
End Function

Public Function ReadFromMemoryStream(ba() As Byte) As Boolean
Try: On Error GoTo Catch
    'simply copy the data from the bytearray to the m_MunsellColors-Array
    Dim SizeOf_TMunsellColor As Long: SizeOf_TMunsellColor = LenB(MunsellColorZero)
    Dim c As Long: c = (UBound(ba) - LBound(ba) + 1) \ SizeOf_TMunsellColor
    'c should be 2734
    'ReDim m_MunsellColors(0 To 2733)
    ReDim m_MunsellColors(0 To c - 1)
    RtlMoveMemory m_MunsellColors(0), ba(0), c * SizeOf_TMunsellColor
    ReadFromMemoryStream = True
Catch:
End Function

Public Function SaveToFile(ByVal FNm As String) As Boolean
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    If Dir(FNm) <> "" Then Kill FNm
    Open FNm For Binary Access Write As FNr
    Put FNr, , m_MunsellColors
    SaveToFile = True
    GoTo Finally
Catch:
    MsgBox "Error in SaveToFile"
Finally:
    Close FNr
End Function

Public Function ReadFromFile(ByVal FNm As String) As Boolean
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    'If Dir(FNm) <> "" Then Kill FNm 'Blödsinn doch nicht hier?!?
    Open FNm For Binary Access Read As FNr
    Dim u As Long: u = LOF(FNr) / LenB(MunsellColorZero) - 1
    ReDim m_MunsellColors(0 To u)
    Get FNr, , m_MunsellColors
    ReadFromFile = True
    GoTo Finally
Catch:
    MsgBox "Error in ReadFromFile"
Finally:
    Close FNr
End Function

' v ############################## v '    Excel specific functions    ' v ############################## v '
'Public Sub SaveData()
'    If Not ReadData Then Exit Sub
'    If Not SaveToFile("C:\TestDir\Munsell.bin") Then Exit Sub
'    MunsellColors_ToWorkSheet Excel.ActiveWorkbook.Worksheets("Test")
'End Sub
'
'Function ReadData() As Boolean
'Try: On Error GoTo Catch
'    Dim wks As Excel.Worksheet: Set wks = Excel.ActiveWorkbook.Worksheets("Conversion Lists")
'    Dim iRow As Long, StartRow As Long: StartRow = 2: iRow = StartRow
'    Dim iCol As Long, StartCol As Long: StartCol = 3: iCol = StartCol
'    Dim c As Long
'    Dim Cell As Range, CellValue As String
'    Dim hp As EHuePrefix, v As Byte, RGBA As RGBA
'    Dim mc As TMunsellColor
'
'    ReDim m_MunsellColors(0 To 3000) 'As TMunsellR
'    Do
'
'        Set Cell = wks.Cells(iRow, iCol): iCol = iCol + 1: CellValue = Cell.Value
'
'        If CellValue = "" Then Exit Do
'
'        If EHuePrefix_TryParse(CellValue, hp) Then mc.HuePrefix = hp
'
'        Set Cell = wks.Cells(iRow, iCol): iCol = iCol + 1: CellValue = Cell.Value
'        If HueValue_TryParse(CellValue, v) Then mc.HueValue = v
'
'        Set Cell = wks.Cells(iRow, iCol): iCol = iCol + 1: CellValue = Cell.Value
'        If Byte_TryParse(CellValue, v) Then mc.ValValue = v
'
'        Set Cell = wks.Cells(iRow, iCol): iCol = iCol + 1: CellValue = Cell.Value
'        If Byte_TryParse(CellValue, v) Then mc.Chroma = v
'
'        Set Cell = wks.Cells(iRow, iCol): iCol = iCol + 1: CellValue = Cell.Value
'        If RGBA_TryParse(CellValue, RGBA) Then mc.RGBA = RGBA
'
'        If iCol >= 7 Then iCol = StartCol
'
'        m_MunsellColors(c) = mc
'        c = c + 1
'
'        iRow = iRow + 1
'    Loop
'    ReDim Preserve m_MunsellColors(0 To c - 1)
'    ReadData = True
'    Exit Function
'Catch:
'    MsgBox "Error in ReadData"
'End Function
'
'Sub MunsellColors_ToWorkSheet(wks As Worksheet)
'    Dim i As Long, u As Long: u = UBound(m_MunsellColors)
'    Dim iRow As Long: iRow = 2
'    Dim iCol As Long: iCol = 1
'    Dim Cell As Range
'    Dim mc As TMunsellColor
'    For i = 0 To u
'        mc = m_MunsellColors(i)
'        Set Cell = wks.Cells(iRow, iCol): Cell.Value = TMunsellColor_Key(mc): iCol = iCol + 1
'        Set Cell = wks.Cells(iRow, iCol): Cell.Value = "'" & RGBA_ToStr(mc.RGBA) ': iCol = iCol + 1
'        iCol = 1
'        iRow = iRow + 1
'    Next
'End Sub
'
'Sub Colorize_Area(ByVal area As String)
'    Dim Inhalt As String
'    Dim Zelle  As Range
'    Dim aRGB   As Variant
'    For Each Zelle In Range(area)
'        Inhalt = IIf(Zelle.HasFormula, Zelle.Value2, "255,255,255")
'        If Len(Inhalt) = 0 Then Inhalt = "255,255,255"
'        aRGB = Split(Inhalt, ",")
'        Zelle.Interior.Color = RGB(aRGB(0), aRGB(1), aRGB(2))
'        'End If
'        'If Zelle.HasFormula Then
'        '    Inhalt = Zelle.Value2
'        '    If Inhalt <> "" Then
'        '    Else
'        '        Zelle.Interior.Color = RGB(255, 255, 255)
'        '    End If
'        'End If
'    Next
'End Sub
' ^ ############################## ^ '    Excel specific functions    ' ^ ############################## ^ '

