VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBColor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      ScaleHeight     =   167.23
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   120.012
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox PbPicture 
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   68
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnOnOff 
      Caption         =   "on/off"
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox CBValues 
      Height          =   330
      Left            =   720
      TabIndex        =   66
      Text            =   "Combo2"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox CBValuesf 
      Height          =   330
      Left            =   720
      TabIndex        =   65
      Text            =   "Combo2"
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox PnlHSV 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   9120
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   55
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetHSV 
         Caption         =   "Set  HSV"
         Height          =   375
         Left            =   0
         TabIndex        =   60
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBHSV_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   59
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBHSV_V 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   58
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBHSV_S 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   57
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBHSV_H 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   56
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   64
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "V:"
         Height          =   210
         Left            =   0
         TabIndex        =   63
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   210
         Left            =   0
         TabIndex        =   62
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         Height          =   210
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox PnlXYZ 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   10920
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   45
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetXYZ 
         Caption         =   "Set  XYZ"
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBXYZ_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   49
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_Z 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_Y 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   47
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_X 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   46
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   54
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   210
         Left            =   0
         TabIndex        =   53
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   210
         Left            =   0
         TabIndex        =   52
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   210
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox PnlHSL 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   7320
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TBHSL_H 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   40
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox TBHSL_S 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBHSL_L 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBHSL_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton BtnSetHSL 
         Caption         =   "Set  HSL"
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         Height          =   210
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   210
         Left            =   0
         TabIndex        =   43
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "L:"
         Height          =   210
         Left            =   0
         TabIndex        =   42
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   41
         Top             =   1440
         Width           =   210
      End
   End
   Begin VB.PictureBox PnlRGBAf 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   3720
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetRGBAf 
         Caption         =   "Set RGBAf"
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBRGBAf_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_B 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_G 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_R 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   35
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   210
         Left            =   0
         TabIndex        =   34
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   210
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   210
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox PnlCMYK 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1920
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TBCMYK_Y 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_K 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton BtnSetCMYK 
         Caption         =   "Set CMYKA"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBCMYK_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_M 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_C 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "C:"
         Height          =   210
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "M:"
         Height          =   210
         Left            =   0
         TabIndex        =   25
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   210
         Left            =   0
         TabIndex        =   24
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "K:"
         Height          =   210
         Left            =   0
         TabIndex        =   23
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   22
         Top             =   1440
         Width           =   210
      End
   End
   Begin VB.PictureBox PnlRGBA 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5520
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetRGBA 
         Caption         =   "Set RGBA"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBRGBA_A 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_B 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_G 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_R 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   210
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   210
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   210
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   210
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.TextBox TBLngColor 
      Alignment       =   1  'Rechts
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox CmbColorNames 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Hier wird alles von und nach CMYK konvertiert
Private m_CMYK As CMYK
Private m_TBBack  As TextBox
Private m_PnlHwnd As Long
Private m_Btn     As CommandButton
Private m_CPicker As ColorDialog
Private m_APB     As AlphaPB

Private Sub TBLngColor_LostFocus()
    Dim LC As LngColor: LC = MColor.LngColor_ParseWebHex(TBLngColor.Text)
    m_CMYK = LngColor_ToCMYK(LC)
    UpdateView
End Sub

Private Sub Timer1_Timer()
    GetCursorPos CurMousePos
    Dim C As Long: C = ColorUnderMouse(CurMousePos.X, CurMousePos.Y)
    PBColor.BackColor = C
    m_CMYK = RGBAf_ToCMYK(MColor.LngColor_ToRGBAf(LngColor(C)))
    UpdateView
End Sub

Private Function ColorUnderMouse(ByVal X As Long, ByVal Y As Long) As Long
    ColorUnderMouse = GetPixel(GetDC(0), X, Y)
End Function

Private Sub BtnOnOff_Click()
    Timer1.Enabled = Not Timer1.Enabled
    BtnOnOff.Caption = IIf(Timer1.Enabled, "on", "off")
End Sub

Private Sub Form_Load()
    Set m_CPicker = New ColorDialog
    Set m_APB = New AlphaPB: m_APB.New_ Me.PBColor, Me.PbPicture
    Timer1.Interval = 50
    Timer1.Enabled = False
    FillCmbMouseScrollf CBValuesf
    FillCmbMouseScroll CBValues
    HideCBValues
    MKnownColors.X11KnownColor_ToCB CmbColorNames
    PBColor.BackColor = vbCyan
    m_CMYK = RGBAf_ToCMYK(LngColor_ToRGBAf(LngColor(PBColor.BackColor)))
    UpdateView
End Sub

Sub UpdateView(Optional bNoUpdataColorName As Boolean = False)
    
    
    
    MColor.CMYK_ToView TBCMYK_C, TBCMYK_M, TBCMYK_Y, TBCMYK_K, TBCMYK_A, m_CMYK
    
    Dim RGBAf As RGBAf:   RGBAf = MColor.CMYK_ToRGBAf(m_CMYK)
    MColor.RGBAf_ToView TBRGBAf_R, TBRGBAf_G, TBRGBAf_B, TBRGBAf_A, RGBAf
    
    Dim RGBA  As RGBA:     RGBA = MColor.RGBAf_ToRGBA(RGBAf)
    MColor.RGBA_ToView TBRGBA_R, TBRGBA_G, TBRGBA_B, TBRGBA_A, RGBA
    
    'Todo: here reparieren!!!
    'siehe Projekt: C:\Users\olimi\OneDrive\Documents\VB\Class Color\Transparenz\alphablending
    m_APB.Alpha = 255 - RGBA.A 'vsAlphaBlend.Value
    
    Dim LCol  As LngColor: LCol = MColor.RGBA_ToLngColor(RGBA)
    TBLngColor.Text = MColor.LngColor_ToWebHex(LCol)
        
    RGBA.A = 0
    LCol = MColor.RGBA_ToLngColor(RGBA)
    PBColor.BackColor = LCol.Value
    
    Dim HSL As HSL: HSL = RGBAf_ToHSL(RGBAf)
    MColor.HSL_ToView TBHSL_H, TBHSL_S, TBHSL_L, TBHSL_A, HSL

    Dim HSV As HSV: HSV = RGBAf_ToHSV(RGBAf)
    MColor.HSV_ToView TBHSV_H, TBHSV_S, TBHSV_V, TBHSV_A, HSV
    
    Dim XYZ As XYZ: XYZ = RGBAf_ToXYZ(RGBAf)
    MColor.XYZ_ToView TBXYZ_X, TBXYZ_Y, TBXYZ_Z, TBXYZ_A, XYZ
    
    If Not bNoUpdataColorName Then
        Dim xn As String: xn = MKnownColors.NameFromColor(LCol.Value)
        If Len(xn) Then CmbColorNames.Text = xn
    End If
End Sub

Private Sub ErrMsg(sErr As String)
    MsgBox "Invalid numeric value: " & sErr & vbCrLf & "please try again"
End Sub

Private Sub BtnSetCMYK_Click()
    Dim sErr As String
    If Not MColor.CMYK_Read(m_CMYK, TBCMYK_C, TBCMYK_M, TBCMYK_Y, TBCMYK_K, TBCMYK_A, sErr) Then ErrMsg sErr: Exit Sub
    UpdateView
End Sub

Private Sub BtnSetRGBAf_Click()
    Dim RGBAf As RGBAf, sErr As String
    If Not MColor.RGBAf_Read(RGBAf, TBRGBAf_R, TBRGBAf_G, TBRGBAf_B, TBRGBAf_A, sErr) Then ErrMsg sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(RGBAf)
    UpdateView
End Sub

Private Sub BtnSetRGBA_Click()
    Dim RGBA As RGBA, sErr As String
    If Not MColor.RGBA_Read(RGBA, TBRGBA_R, TBRGBA_G, TBRGBA_B, TBRGBA_A, sErr) Then ErrMsg sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.RGBA_ToRGBAf(RGBA))
    UpdateView
End Sub

Private Sub BtnSetHSL_Click()
    Dim HSL As HSL, sErr As String
    If Not MColor.HSL_Read(HSL, TBHSL_H, TBHSL_S, TBHSL_L, TBHSL_A, sErr) Then ErrMsg sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.HSL_ToRGBAf(HSL))
    UpdateView
End Sub

Private Sub BtnSetHSV_Click()
    Dim HSV As HSV, sErr As String
    If Not MColor.HSV_Read(HSV, TBHSV_H, TBHSV_S, TBHSV_V, TBHSV_A, sErr) Then ErrMsg sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.HSV_ToRGBAf(HSV))
    UpdateView
End Sub

Private Sub BtnSetXYZ_Click()
    Dim XYZ As XYZ, sErr As String
    If Not MColor.XYZ_Read(XYZ, TBXYZ_X, TBXYZ_Y, TBXYZ_Z, TBXYZ_A, sErr) Then ErrMsg sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.XYZ_ToRGBAf(XYZ))
    UpdateView
End Sub

Private Sub CmbColorNames_Click()
    If CmbColorNames.Text = "" Then Exit Sub
    PBColor.BackColor = MKnownColors.ColorByName(CmbColorNames.Text)
    Dim LngColor As LngColor: LngColor.Value = PBColor.BackColor
    m_CMYK = MColor.RGBAf_ToCMYK(MColor.RGBA_ToRGBAf(MColor.LngColor_ToRGBA(LngColor)))
    UpdateView True
End Sub
Private Sub FillCmbMouseScrollf(Cmb As ComboBox)
    Dim i As Long
    Cmb.Clear
    Dim N As Long: N = 256
    Dim fact As Double: fact = 1 / N
    For i = N To 0 Step -1
        Cmb.AddItem Format(i * fact, "0.###")
    Next
End Sub
Private Sub FillCmbMouseScroll(Cmb As ComboBox)
    'CBValues
    Dim i As Long
    Dim N As Long: N = 255
    For i = N To 0 Step -1
        Cmb.AddItem i
    Next
End Sub

Private Sub PBColor_Click()
Try: On Error GoTo Catch
    m_CPicker.Color = PBColor.BackColor
    If m_CPicker.ShowDialog = vbCancel Then Exit Sub
    PBColor.BackColor = m_CPicker.Color
    'UpdateView
Catch:
End Sub

Private Sub HideCBValues()
    CBValues.ZOrder 1
    CBValuesf.ZOrder 1
End Sub
Private Sub PnlCMYK_DblClick():  HideCBValues: End Sub
Private Sub PnlRGBAf_DblClick(): HideCBValues: End Sub
Private Sub PnlRGBA_DblClick():  HideCBValues: End Sub
Private Sub PnlHSL_DblClick():   HideCBValues: End Sub
Private Sub PnlHSV_DblClick():   HideCBValues: End Sub
Private Sub PnlXYZ_DblClick():   HideCBValues: End Sub

Private Sub TBCMYK_C_DblClick():  SetTB TBCMYK_C, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
Private Sub TBCMYK_M_DblClick():  SetTB TBCMYK_M, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
Private Sub TBCMYK_Y_DblClick():  SetTB TBCMYK_Y, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
Private Sub TBCMYK_K_DblClick():  SetTB TBCMYK_K, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
Private Sub TBCMYK_A_DblClick():  SetTB TBCMYK_A, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub

Private Sub TBRGBAf_R_DblClick(): SetTB TBRGBAf_R, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
Private Sub TBRGBAf_G_DblClick(): SetTB TBRGBAf_G, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
Private Sub TBRGBAf_B_DblClick(): SetTB TBRGBAf_B, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
Private Sub TBRGBAf_A_DblClick(): SetTB TBRGBAf_A, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub

Private Sub TBRGBA_R_DblClick():  SetTB TBRGBA_R, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
Private Sub TBRGBA_G_DblClick():  SetTB TBRGBA_G, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
Private Sub TBRGBA_B_DblClick():  SetTB TBRGBA_B, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
Private Sub TBRGBA_A_DblClick():  SetTB TBRGBA_A, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub

Private Sub TBHSL_H_DblClick():  SetTB TBHSL_H, CBValuesf, BtnSetHSL, PnlHSL.hwnd, 256: End Sub
Private Sub TBHSL_S_DblClick():  SetTB TBHSL_S, CBValuesf, BtnSetHSL, PnlHSL.hwnd, 256: End Sub
Private Sub TBHSL_L_DblClick():  SetTB TBHSL_L, CBValuesf, BtnSetHSL, PnlHSL.hwnd, 256: End Sub
Private Sub TBHSL_A_DblClick():  SetTB TBHSL_A, CBValuesf, BtnSetHSL, PnlHSL.hwnd, 256: End Sub
'
Private Sub TBHSV_H_DblClick():  SetTB TBHSV_H, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
Private Sub TBHSV_S_DblClick():  SetTB TBHSV_S, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
Private Sub TBHSV_V_DblClick():  SetTB TBHSV_V, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
Private Sub TBHSV_A_DblClick():  SetTB TBHSV_A, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub

'Private Sub TBXYZ_X_DblClick():  SetTB TBXYZ_X, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_Y_DblClick():  SetTB TBXYZ_Y, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_Z_DblClick():  SetTB TBXYZ_Z, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_A_DblClick():  SetTB TBXYZ_A, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub


Private Sub SetTB(TB As TextBox, CB As ComboBox, Btn As CommandButton, ByVal pnlHwnd As Long, ByVal f As Single)
    Set m_TBBack = TB
    Set m_Btn = Btn
    SetParent CB.hwnd, pnlHwnd
    CB.Move m_TBBack.Left, m_TBBack.Top
    Dim N As Single: N = 256
    If f = 1 Then N = 255
    CB.ListIndex = N - (f * CSng(m_TBBack.Text))
    CB.ZOrder 0
End Sub

Private Sub CBValuesf_DblClick()
    m_TBBack.ZOrder 0
End Sub
Private Sub CBValuesf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CBValuesf_Click
        m_TBBack.ZOrder 0
    End If
End Sub
Private Sub CBValuesf_Click()
    m_TBBack.Text = CBValuesf.Text
    m_Btn.Value = True
End Sub

Private Sub CBValues_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CBValues_Click
        m_TBBack.ZOrder 0
    End If
End Sub
Private Sub CBValues_Click()
    m_TBBack.Text = CBValues.Text
    m_Btn.Value = True
End Sub

