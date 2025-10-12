VERSION 5.00
Begin VB.Form FGGUColors 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "GGU-Colors"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
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
   ScaleHeight     =   5415
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox PBNewColor 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox TxtNewIndex 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox PBPrevColor 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtPrevIndex 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox PBColors 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
      Begin VB.Shape ShpColors 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H8000000D&
         BorderWidth     =   2
         Height          =   495
         Index           =   0
         Left            =   0
         Shape           =   1  'Quadrat
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Color Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Previous Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FGGUColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PrevColor  As Long
Private m_PrevIndex  As Byte
Private m_NewColor   As Long
Private m_NewIndex   As Byte
Private m_Result     As VbMsgBoxResult
Private m_Cw         As Single
Private m_Ch         As Single
Private m_oldI       As Integer
Private m_isSelected As Boolean

Public Function ShowDialog(Owner As Form, Color_inout As Long) As VbMsgBoxResult
    m_PrevColor = Color_inout
    PBPrevColor.BackColor = m_PrevColor
    m_PrevIndex = MGGUColors.IndexFromColor(Color_inout)
    TxtPrevIndex.Text = m_PrevIndex
    'ShpColors(m_PrevIndex).BorderWidth =
    ShpColors(m_PrevIndex).BorderStyle = BorderStyleConstants.vbBSSolid
    Me.Show vbModal, Owner
    m_isSelected = False
    ShowDialog = m_Result
    If ShowDialog = VbMsgBoxResult.vbCancel Then Exit Function
    Color_inout = m_NewColor
End Function

Private Sub Form_Load()
    m_oldI = -1
    LoadShpColors
End Sub

Private Sub LoadShpColors()
    With ShpColors(0)
        Dim L0 As Single: L0 = .Left: m_Cw = PBColors.ScaleWidth / 5 ' .Width
        Dim T0 As Single: T0 = .Top:  m_Ch = PBColors.ScaleHeight / 5 '.Height
    End With
    ControlArrayOfShapes_Load ShpColors(), 5, 5
    ControlArrayOfShapes_BackColor(ShpColors()) = MGGUColors.ColorArray
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub
'
Private Sub PBColors_DblClick()
    BtnOK_Click
End Sub

Private Sub PBColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        m_isSelected = Not m_isSelected
        Dim i As Integer: i = ControlArrayOfShapes_GetIndex(ShpColors, X, Y)
        If i < 0 Then Exit Sub
        m_NewColor = ShpColors(i).BackColor
        PBNewColor.BackColor = m_NewColor
        TxtNewIndex.Text = i
    End If
End Sub

Private Sub PBColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_isSelected Then Exit Sub
    Dim i As Integer: i = ControlArrayOfShapes_GetIndex(ShpColors, X, Y)
    If -1 < m_oldI And m_oldI <> m_PrevIndex Then
        ShpColors(m_oldI).BorderStyle = BorderStyleConstants.vbTransparent
    End If
    If -1 < i Then
        PBNewColor.BackColor = ShpColors(i).BackColor
        ShpColors(i).BorderStyle = BorderStyleConstants.vbBSSolid
        TxtNewIndex.Text = i
    End If
    If i <> m_oldI Then m_oldI = i
End Sub

' v ' ############################## ' v '    Shapes    ' v ' ############################## ' v '
Private Sub ControlArrayOfShapes_Load(ControlArrayOfShapes, ByVal nw As Byte, ByVal nh As Byte)
    Dim L As Single, t As Single
    Dim W As Single: W = m_Cw ' ControlArrayOfShapes(0).Width
    Dim H As Single: H = m_Ch 'ControlArrayOfShapes(0).Height
    Dim i As Long
    For i = 0 To CLng(nw) * CLng(nh) - 1
        If i > 0 Then Load ControlArrayOfShapes(i)
        With ControlArrayOfShapes(i)
            .Move L, t, W, H
            .Visible = True
            .BorderStyle = BorderStyleConstants.vbTransparent '0
            .BorderWidth = 3
        End With
        L = L + W
        If ((i + 1) Mod nw) = 0 Then
            L = 0: t = t + H
        End If
    Next
End Sub

Private Function ControlArrayOfShapes_GetIndex(ControlArrayOfShapes, ByVal X As Single, ByVal Y As Single) As Integer
    Dim i As Long: ControlArrayOfShapes_GetIndex = -1
    Dim q As Shape
    For i = ControlArrayOfShapes.LBound To ControlArrayOfShapes.UBound
        Set q = ControlArrayOfShapes(i)
        If (q.Left < X) And (X < q.Left + q.Width) And _
           (q.Top < Y) And (Y < q.Top + q.Height) Then
            ControlArrayOfShapes_GetIndex = i
            Exit Function
        End If
    Next
End Function

Private Property Let ControlArrayOfShapes_BackColor(ControlArrayOfShapes, Color() As Long)
    Dim i As Long
    For i = ControlArrayOfShapes.LBound To ControlArrayOfShapes.UBound
        ControlArrayOfShapes(i).BackColor = Color(i)
    Next
End Property



