VERSION 5.00
Begin VB.Form FMunsell 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Munsell-Colors"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBColorNearest 
      Height          =   615
      Left            =   5400
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   19
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox PBColorInput 
      Height          =   615
      Left            =   5400
      ScaleHeight     =   555
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
      ScaleWidth      =   8055
      TabIndex        =   12
      Top             =   7155
      Width           =   8055
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
      Left            =   1560
      TabIndex        =   10
      Text            =   "CmbMunsell3"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox PBColorSelected 
      Height          =   615
      Left            =   5400
      ScaleHeight     =   555
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
      ScaleWidth      =   7995
      TabIndex        =   6
      ToolTipText     =   "double-click to select, click again to unselect"
      Top             =   2040
      Width           =   8055
   End
   Begin VB.ComboBox CmbMunsell2 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   "CmbMunsell2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox CmbMunsell1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "CmbMunsell1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox CmbMunsell 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "CmbMunsell"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label LblColorSelectedRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   1560
      Width           =   1770
   End
   Begin VB.Label LblColorNearestRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   960
      Width           =   1770
   End
   Begin VB.Label LblColorNearestKey 
      Alignment       =   2  'Zentriert
      Caption         =   "-- | --"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label LblNearestColor 
      AutoSize        =   -1  'True
      Caption         =   "Nearest Color:"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label LblColorInputRGB 
      Alignment       =   2  'Zentriert
      Caption         =   "255, 255, 255"
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label LblInputColor 
      AutoSize        =   -1  'True
      Caption         =   "Input Color:"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label LblColorSelectedKey 
      Alignment       =   2  'Zentriert
      Caption         =   "-- | --"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HuePrefix/Hue:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label LblSelectedColor 
      AutoSize        =   -1  'True
      Caption         =   "Selected Color:"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1320
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
Private m_ChromaValue As HueValue
Private m_bUpdateView As Boolean
Private m_bSelected   As Boolean

Private m_Wb As Single
Private m_Hb As Single
Private m_Result As VbMsgBoxResult

'Private m_ColorInput    As Long
'Private m_ColorNearest  As TMunsellColor
Private m_ColorSelected As TMunsellColor

Private Sub Form_Load()
    CmbMunsell.Clear
    CmbMunsell.AddItem "Chroma-Value"
    CmbMunsell.AddItem "Chroma-HuePrefix"
    CmbMunsell.AddItem "Chroma-Hue"
    CmbMunsell.ListIndex = 0
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

Public Function ShowDialog(Owner As Form, Color_inout As Long) As VbMsgBoxResult
    View_Init Color_inout 'only once
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    Color_inout = RGBA_ToLngColor(m_ColorSelected.RGBA).Value
End Function

'Friend Property Get ColorInput() As Long
'    ColorInput = m_ColorInput
'End Property

'Friend Property Get ColorNearest() As TMunsellColor
'    ColorNearest = m_ColorNearest
'End Property
'
'Friend Property Let ColorNearest(ByVal Value As TMunsellColor)
'
'End Property
'
'Friend Property Get ColorSelected() As TMunsellColor
'    ColorSelected = m_ColorSelected
'End Property

'Friend Property Let ColorSelected(Value As TMunsellColor)
'    m_ColorSelected = Value
'
'    PBSelColor.BackColor = RGBA_ToLngColor(m_Color.RGBA).Value
'
'End Property


Private Sub CmbMunsell_Click()
    Dim i As Long: i = CmbMunsell.ListIndex
    Dim s1 As String: s1 = IIf(i = 1, "Value", "Hue-Prefix")
    Dim s2 As String: s2 = IIf(i = 2, "Value", "Hue")
    Label1.Caption = s1 & ":"
    Label2.Caption = s2 & ":"
    Label3.Caption = s1 & "/" & s2 & ":"
    
    If i = 0 Or i = 1 Then HueValue_ToCmb CmbMunsell2
    If i = 1 Or i = 2 Then ValValue_ToCmb CmbMunsell1: ValValue_ToCmb CmbMunsell2
    If i = 2 Or i = 0 Then EHuePrefix_ToCmb CmbMunsell1
    If i = 0 Then EHuePrefixHueValue_ToCmb CmbMunsell3
    
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
    Dim s As String: s = CmbMunsell3.List(i) 'CmbMunsell3.SelText
    Dim sa() As String: sa = Split(s, " - ")
    Dim e As EHuePrefix
    m_bUpdateView = False
    If EHuePrefixName_TryParse(sa(0), e) Then
        i = e
        i = i - 1
        If CmbMunsell1.ListIndex <> i Then
            CmbMunsell1.ListIndex = i
        End If
    End If
    Dim b As Byte
    If HueValue_TryParse(sa(1), b) Then
        i = b
        i = i - 1
        If CmbMunsell2.ListIndex <> i Then
            CmbMunsell2.ListIndex = i
        End If
    End If
    m_bUpdateView = True
    UpdateView
End Sub

Private Sub CmbMunsell3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyDown
        If CmbMunsell3.ListIndex = CmbMunsell3.ListCount - 1 Then
            CmbMunsell3.ListIndex = 0
        End If
    Case KeyCodeConstants.vbKeyUp
        If CmbMunsell3.ListIndex = 0 Then
            CmbMunsell3.ListIndex = CmbMunsell3.ListCount - 1
        End If
    End Select
End Sub

Function GetHuePrefix() As EHuePrefix
    Dim i As Long: i = CmbMunsell1.ListIndex
    If i < 0 Then i = 0
    GetHuePrefix = i + 1
End Function

Function GetHueValue() As Byte
    Dim i As Long: i = CmbMunsell2.ListIndex
    If i < 0 Then i = 0
    GetHueValue = i + 1
End Function

Sub View_Init(ByVal aColor As Long)
    'm_ColorInput = Value
    PBColorInput.BackColor = aColor
    Dim RGBA As RGBA: RGBA = LngColor_ToRGBA(LngColor(aColor))
    LblColorInputRGB.Caption = RGBA_ToStr(RGBA)
    Dim nc As TMunsellColor: nc = MMunsell.MunsellColors_ClosestColorTo(aColor)
    PBColorNearest.BackColor = RGBA_ToLngColor(nc.RGBA).Value
    LblColorNearestKey.Caption = MMunsell.TMunsellColor_Key(nc)
    LblColorNearestRGB.Caption = MColor.RGBA_ToStr(nc.RGBA)
    
    PBColorSelected.BackColor = PBColorNearest.BackColor
    LblColorSelectedKey.Caption = LblColorNearestKey.Caption
    LblColorSelectedRGB.Caption = LblColorNearestRGB.Caption
End Sub

Sub UpdateView()
    If Not m_bUpdateView Then Exit Sub
    Dim HuePrefix As EHuePrefix: HuePrefix = GetHuePrefix
    Dim HueValue  As Byte:        HueValue = GetHueValue
    m_ChromaValue = MMunsell.MunsellColors_ChromaValue(HuePrefix, HueValue)
    'PBSelColor.BackColor = MColor.RGBA_ToLngColor(m_ChromaValue.ValValues(3).Chromas(3).RGBA).Value
    ChromaValue_Draw
End Sub

Sub ChromaValue_Draw()
    Dim i As Long, ui As Long: ui = MMunsell.Count_ValValue ' UBound(m_ChromaValue.ValValues)
    Dim j As Long, uj As Long: uj = Count_ChromaMax
    Dim x As Single: m_Wb = PBColors.ScaleWidth / ui
    Dim Y As Single: m_Hb = PBColors.ScaleHeight / uj
    Dim c As Long
    PBColors.Cls
    For i = 1 To ui
        uj = UBound(m_ChromaValue.ValValues(i).Chromas)
        For j = 1 To uj
            c = MColor.RGBA_ToLngColor(m_ChromaValue.ValValues(i).Chromas(j).RGBA).Value
            PBColors.Line (x, Y)-(x + m_Wb, Y + m_Hb), c, BF
            Y = Y + m_Hb
        Next
        x = x + m_Wb
        Y = 0
    Next
End Sub

Private Sub PBColors_Click()
    m_bSelected = False
End Sub

Private Sub PBColors_DblClick()
    m_bSelected = True
End Sub

Private Sub PBColors_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If m_bSelected Then Exit Sub
    Dim i As Long: i = Floor(x / m_Wb) + 1
    Dim j As Long: j = Floor(Y / m_Hb) + 1
    If i <= UBound(m_ChromaValue.ValValues) Then
        If j <= UBound(m_ChromaValue.ValValues(i).Chromas) Then
            m_ColorSelected = m_ChromaValue.ValValues(i).Chromas(j)
            LblColorSelectedKey.Caption = MMunsell.TMunsellColor_Key(m_ColorSelected) & " rgb=" & RGBA_ToStr(m_ColorSelected.RGBA)
            PBColorSelected.BackColor = MColor.RGBA_ToLngColor(m_ColorSelected.RGBA).Value
        Else
            LblColorSelectedKey.Caption = "--|--"
        End If
    Else
        LblColorSelectedKey.Caption = "--|--"
    End If
End Sub

Private Sub BtnOK_Click()
    m_Result = vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel
    Unload Me
End Sub

