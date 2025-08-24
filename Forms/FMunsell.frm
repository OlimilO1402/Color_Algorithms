VERSION 5.00
Begin VB.Form FMunsell 
   Caption         =   "Munsell-Colors"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMunsell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBColorNearest 
      Height          =   495
      Left            =   5640
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   19
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox PBColorInput 
      Height          =   495
      Left            =   5640
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox PanelBtns 
      Align           =   2  'Unten ausrichten
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8295
      TabIndex        =   12
      Top             =   7155
      Width           =   8295
      Begin VB.CommandButton BtnCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton BtnOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ComboBox CmbMunsell3 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Text            =   "CmbMunsell3"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox PBColorSelected 
      Height          =   495
      Left            =   5640
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox PBColors 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   8235
      TabIndex        =   6
      ToolTipText     =   "double-click to select, click again to unselect"
      Top             =   2040
      Width           =   8295
   End
   Begin VB.ComboBox CmbMunsell2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "CmbMunsell2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox CmbMunsell1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "CmbMunsell1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox CmbMunsell 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "CmbMunsell"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label LblColorSelectedRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   1560
      Width           =   1770
   End
   Begin VB.Label LblColorNearestRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6480
      TabIndex        =   21
      Top             =   960
      Width           =   1770
   End
   Begin VB.Label LblColorNearestKey 
      Alignment       =   2  'Zentriert
      Caption         =   "-- | --"
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label LblNearestColor 
      AutoSize        =   -1  'True
      Caption         =   "Nearest Color:"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label LblColorInputRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label LblInputColor 
      AutoSize        =   -1  'True
      Caption         =   "Previous Color:"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label LblColorSelectedKey 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "-- | --"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HuePrefix/Hue:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label LblSelectedColor 
      AutoSize        =   -1  'True
      Caption         =   "Selected Color:"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hue:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HuePrefix:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      Caption         =   "Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "FMunsell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ChromaValue     As MMunsell.HueValueCV
Private m_ChromaHuePrefix As MMunsell.HueValueCHP
Private m_ChromaHue       As MMunsell.ValValueCH

Private m_bUpdateView As Boolean
Private m_bSelected   As Boolean

Private m_Wb As Single
Private m_Hb As Single
Private m_Result As VbMsgBoxResult

Private m_ColorSelected As TMunsellColor

Private m_LastSelIndex As Long

Private Sub Form_Load()
    InitFillCmbMunsell
    m_bUpdateView = True
    UpdateView
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim t As Single: t = PBColors.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t - PanelBtns.Height
    If W > 0 And H > 0 Then
        PBColors.Move L, t, W, H
        UpdateView
    End If
    Dim b As Single: b = 8 * Screen.TwipsPerPixelX
    L = PanelBtns.Width / 2 - BtnOK.Width - b
    t = BtnOK.Top
    W = BtnOK.Width
    H = BtnOK.Height
    If W > 0 And H > 0 Then BtnOK.Move L, t, W, H
    L = L + BtnOK.Width + 2 * b
    If W > 0 And H > 0 Then BtnCancel.Move L, t, W, H
End Sub

Private Sub InitFillCmbMunsell()
    CmbMunsell.Clear
    CmbMunsell.AddItem "Chroma-Value"
    CmbMunsell.AddItem "Chroma-HuePrefix"
    CmbMunsell.AddItem "Chroma-Hue"
    CmbMunsell.ListIndex = 0
End Sub

Public Function ShowDialog(Owner As Form, Color_inout As Long) As VbMsgBoxResult
    View_Init Color_inout 'only once
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    Color_inout = MColor.RGBA_ToLngColor(m_ColorSelected.RGBA).Value
End Function

Private Sub CmbMunsell_Click()
    Dim i As Long: i = CmbMunsell.ListIndex
    Dim s1 As String: s1 = IIf(i = 1, "Value", "Hue-Prefix")
    Dim s2 As String: s2 = IIf(i = 2, "Value", "Hue")
    Label1.Caption = s1 & ":"
    Label2.Caption = s2 & ":"
    Label3.Caption = s1 & "/" & s2 & ":"
    
    Select Case i
    Case 0: MMunsell.EHuePrefix_ToCmb CmbMunsell1
            MMunsell.HueValue_ToCmb CmbMunsell2
            MMunsell.EHuePrefixHueValue_ToCmb CmbMunsell3
            
    Case 1: MMunsell.ValValue_ToCmb CmbMunsell1
            MMunsell.HueValue_ToCmb CmbMunsell2
            MMunsell.ValValueHueValue_ToCmb CmbMunsell3
            
    Case 2: MMunsell.EHuePrefix_ToCmb CmbMunsell1
            MMunsell.ValValue_ToCmb CmbMunsell2
            MMunsell.EHuePrefixValValue_ToCmb CmbMunsell3
            
    End Select
    CmbMunsell1.ListIndex = 0
    CmbMunsell2.ListIndex = 0
    UpdateView
End Sub

Private Sub CmbMunsell1_Click()
    UpdateView
End Sub

Private Sub CmbMunsell2_Click()
    UpdateView
End Sub

Private Sub CmbMunsell3_Click()
    Dim i As Long: i = CmbMunsell3.ListIndex
    m_LastSelIndex = i
    
    Dim s As String: s = CmbMunsell3.List(i) 'CmbMunsell3.SelText
    Dim sa() As String: sa = Split(s, " - ")
    Dim v As Byte 'ValValue
    Dim b As Byte
    Dim E As EHuePrefix
    
    m_bUpdateView = False
    Select Case CmbMunsell.ListIndex
    Case 0
        If MMunsell.EHuePrefixName_TryParse(sa(0), E) Then
            i = E
            i = i - 1
            If CmbMunsell1.ListIndex <> i Then
                CmbMunsell1.ListIndex = i
            End If
        End If
        If MMunsell.HueValue_TryParse(sa(1), b) Then
            i = b
            i = i - 1
            If CmbMunsell2.ListIndex <> i Then
                CmbMunsell2.ListIndex = i
            End If
        End If
    Case 1
        If MMunsell.ValValue_TryParse(sa(0), v) Then
            i = v
            i = i - 1
            If CmbMunsell1.ListIndex <> i Then
                CmbMunsell1.ListIndex = i
            End If
        End If
        If MMunsell.HueValue_TryParse(sa(1), b) Then
            i = b
            i = i - 1
            If CmbMunsell2.ListIndex <> i Then
                CmbMunsell2.ListIndex = i
            End If
        End If
    Case 2
        If MMunsell.EHuePrefixName_TryParse(sa(0), E) Then
            i = E
            i = i - 1
            If CmbMunsell1.ListIndex <> i Then
                CmbMunsell1.ListIndex = i
            End If
        End If
        If MMunsell.ValValue_TryParse(sa(1), v) Then
            i = v
            i = i - 1
            If CmbMunsell2.ListIndex <> i Then
                CmbMunsell2.ListIndex = i
            End If
        End If
    End Select

    m_bUpdateView = True
    UpdateView
End Sub

Private Sub CmbMunsell3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim li As Long: li = CmbMunsell3.ListIndex
    Dim lu As Long: lu = CmbMunsell3.ListCount - 1
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyDown:      CmbMunsell3.ListIndex = IIf(li = lu, 0, li + 1)
    Case KeyCodeConstants.vbKeyUp:        CmbMunsell3.ListIndex = IIf(li = 0, lu, li - 1)
    End Select
    KeyCode = 0
End Sub

Function GetHuePrefixCV() As EHuePrefix
    Dim i As Long: i = CmbMunsell1.ListIndex
    If i < 0 Then i = 0
    GetHuePrefixCV = i + 1
End Function
Function GetHueValueCV() As Byte
    Dim i As Long: i = CmbMunsell2.ListIndex
    If i < 0 Then i = 0
    GetHueValueCV = i + 1
End Function

Function GetValValueCHP() As Byte
    Dim i As Long: i = CmbMunsell1.ListIndex
    If i < 0 Then i = 0
    GetValValueCHP = i + 1
End Function
Function GetHueValueCHP() As Byte
    Dim i As Long: i = CmbMunsell2.ListIndex
    If i < 0 Then i = 0
    GetHueValueCHP = i + 1
End Function

Function GetHuePrefixCH() As EHuePrefix
    Dim i As Long: i = CmbMunsell1.ListIndex
    If i < 0 Then i = 0
    GetHuePrefixCH = i + 1
End Function
Function GetValValueCH() As Byte
    Dim i As Long: i = CmbMunsell2.ListIndex
    If i < 0 Then i = 0
    GetValValueCH = i + 1
End Function

Sub View_Init(ByVal aColor As Long)
    PBColorInput.BackColor = aColor
    Dim RGBA As RGBA: RGBA = MColor.LngColor_ToRGBA(LngColor(aColor))
    LblColorInputRGB.Caption = MColor.RGBA_ToStr(RGBA)
    Dim mc As TMunsellColor: mc = MMunsell.MunsellColors_ClosestColorTo(aColor)
    
    'maybe we should jump to the page with selected color?
    PBColorNearest.BackColor = MColor.RGBA_ToLngColor(mc.RGBA).Value
    LblColorNearestKey.Caption = MMunsell.TMunsellColor_Key(mc)
    LblColorNearestRGB.Caption = MColor.RGBA_ToStr(mc.RGBA)
    
    PBColorSelected.BackColor = PBColorNearest.BackColor
    LblColorSelectedKey.Caption = LblColorNearestKey.Caption
    LblColorSelectedRGB.Caption = LblColorNearestRGB.Caption
    
    'CmbMunsell1.Text = MMunsell.EHuePrefix_Name(mc.HuePrefix)
    CmbMunsell1.ListIndex = mc.HuePrefix - 1
    'CmbMunsell2.Text = MMunsell.HueValue_ToStr(mc.HueValue)
    CmbMunsell2.ListIndex = mc.HueValue - 1 '
End Sub

Sub UpdateView()
    If Not m_bUpdateView Then Exit Sub
    Dim HuePrefix As EHuePrefix
    Dim HueValue  As Byte
    Dim ValValue  As Byte
    Select Case CmbMunsell.ListIndex
    Case 0
        HuePrefix = GetHuePrefixCV
        HueValue = GetHueValueCV
        m_ChromaValue = MMunsell.MunsellColors_ChromaValue(HuePrefix, HueValue)
        ChromaValue_Draw
    Case 1
        ValValue = GetValValueCHP
        HueValue = GetHueValueCHP
        m_ChromaHuePrefix = MMunsell.MunsellColors_ChromaHuePrefix(ValValue, HueValue)
        ChromaHuePrefix_Draw
    Case 2
        HuePrefix = GetHuePrefixCH
        ValValue = GetValValueCH
        m_ChromaHue = MMunsell.MunsellColors_ChromaHue(HuePrefix, ValValue)
        ChromaHue_Draw
    End Select
End Sub

Sub ChromaValue_Draw()
    Dim i As Long, ui As Long: ui = MMunsell.Count_ValValue
    Dim j As Long, uj As Long: uj = MMunsell.Count_ChromaMax
    Dim X As Single: m_Wb = PBColors.ScaleWidth / ui
    Dim Y As Single: m_Hb = PBColors.ScaleHeight / uj
    Dim c As Long
    PBColors.Cls
    For i = 1 To ui
        uj = UBound(m_ChromaValue.ValValues(i).Chromas)
        For j = 1 To uj
            c = MColor.RGBA_ToLngColor(m_ChromaValue.ValValues(i).Chromas(j).RGBA).Value
            PBColors.Line (X, Y)-(X + m_Wb, Y + m_Hb), c, BF
            Y = Y + m_Hb
        Next
        X = X + m_Wb
        Y = 0
    Next
End Sub

Sub ChromaHuePrefix_Draw()
    Dim i As Long, ui As Long: ui = MMunsell.Count_HuePrefix
    Dim j As Long, uj As Long: uj = MMunsell.Count_ChromaMax
    Dim X As Single: m_Wb = PBColors.ScaleWidth / ui
    Dim Y As Single: m_Hb = PBColors.ScaleHeight / uj
    Dim c As Long
    PBColors.Cls
    For i = 1 To ui
        uj = UBound(m_ChromaHuePrefix.HuePrefixes(i).Chromas)
        For j = 1 To uj
            c = MColor.RGBA_ToLngColor(m_ChromaHuePrefix.HuePrefixes(i).Chromas(j).RGBA).Value
            PBColors.Line (X, Y)-(X + m_Wb, Y + m_Hb), c, BF
            Y = Y + m_Hb
        Next
        X = X + m_Wb
        Y = 0
    Next
End Sub

Sub ChromaHue_Draw()
    Dim i As Long, ui As Long: ui = MMunsell.Count_HueValue
    Dim j As Long, uj As Long: uj = MMunsell.Count_ChromaMax
    Dim X As Single: m_Wb = PBColors.ScaleWidth / ui
    Dim Y As Single: m_Hb = PBColors.ScaleHeight / uj
    Dim c As Long
    PBColors.Cls
    For i = 1 To ui
        uj = UBound(m_ChromaHue.HueValues(i).Chromas)
        For j = 1 To uj
            c = MColor.RGBA_ToLngColor(m_ChromaHue.HueValues(i).Chromas(j).RGBA).Value
            PBColors.Line (X, Y)-(X + m_Wb, Y + m_Hb), c, BF
            Y = Y + m_Hb
        Next
        X = X + m_Wb
        Y = 0
    Next
End Sub

Private Sub PBColors_Click()
    m_bSelected = False
End Sub

Private Sub PBColors_DblClick()
    m_bSelected = True
End Sub

Private Sub PBColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bSelected Then Exit Sub
    Dim i As Long: i = MMath.Floor(X / m_Wb) + 1
    Dim j As Long: j = MMath.Floor(Y / m_Hb) + 1
    Select Case CmbMunsell.ListIndex
    Case 0
        If i <= UBound(m_ChromaValue.ValValues) Then
            If j <= UBound(m_ChromaValue.ValValues(i).Chromas) Then
                m_ColorSelected = m_ChromaValue.ValValues(i).Chromas(j)
            End If
        End If
    Case 1
        If i <= UBound(m_ChromaHuePrefix.HuePrefixes) Then
            If j <= UBound(m_ChromaHuePrefix.HuePrefixes(i).Chromas) Then
                m_ColorSelected = m_ChromaHuePrefix.HuePrefixes(i).Chromas(j)
            End If
        End If
    Case 2
        If i <= UBound(m_ChromaHue.HueValues) Then
            If j <= UBound(m_ChromaHue.HueValues(i).Chromas) Then
                m_ColorSelected = m_ChromaHue.HueValues(i).Chromas(j)
            End If
        End If
    End Select
    If i <= UBound(m_ChromaValue.ValValues) Then
        If j <= UBound(m_ChromaValue.ValValues(i).Chromas) Then
            LblColorSelectedKey.Caption = MMunsell.TMunsellColor_Key(m_ColorSelected) '
            LblColorSelectedRGB.Caption = RGBA_ToStr(m_ColorSelected.RGBA)
            PBColorSelected.BackColor = MColor.RGBA_ToLngColor(m_ColorSelected.RGBA).Value
        Else
            LblColorSelectedKey.Caption = "--|--"
        End If
    Else
        LblColorSelectedKey.Caption = "--|--"
    End If
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

