VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER  As Long = &H0

Private Declare Function AlphaBlend Lib "msimg32" ( _
    ByVal hDstDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, _
    ByVal blendFunct As Long) As Boolean
    
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLen As Long)

Private Type CtrlRECT
    Left   As Long
    Top    As Long
    Width  As Long
    Height As Long
End Type

Private Type BLENDFUNCTION
    BlendOp       As Byte ' = AC_SRC_OVER = 0 ' the only one defined so far
    BlendFlags    As Byte ' = 0 'must be zero
    SrcConstAlpha As Byte ' = 0 -> transparent; = 255 -> opaque -> or alpha per pixel
    AlphaFormat   As Byte ' = AC_SRC_ALPHA = 1 'only for 32 bit bitmaps with alpha-channel; otherwise = 0 -> for 24bit bitmaps;
End Type

Private Type TLng
    Value As Long
End Type

Private mForePB As PictureBox 'Vordergrundfarbe
Private mBackPB As PictureBox 'Hintergrund mit Bild

Private m_BlFnc As BLENDFUNCTION
Private m_Blend As TLng
Private m_RF    As CtrlRECT 'Rect of Fore PB
Private m_RB    As CtrlRECT 'Rect of Back PB

Friend Sub New_(ForePB As PictureBox, BackPB As PictureBox)
    
    Set mForePB = ForePB:    Set mBackPB = BackPB
    
    'With m_BlFnc
        '.BlendOp = AC_SRC_OVER
        '.BlendFlags = 0
        '.SrcConstAlpha = 0 '255
        '.AlphaFormat = AC_SRC_ALPHA
    'End With
    Prepare m_RF, mForePB
    Prepare m_RB, mBackPB
    mForePB.ZOrder 0
    
End Sub

Private Sub Prepare(Rect As CtrlRECT, aPB As PictureBox)
    With aPB
        .AutoRedraw = True
        .ScaleMode = ScaleModeConstants.vbPixels
        Rect.Left = .Left
        Rect.Top = .Top
        Rect.Width = .ScaleWidth
        Rect.Height = .ScaleHeight
    End With
End Sub

Private Sub UpdateView()
    'Dim lngBlend As Long
    'RtlMoveMemory lngBlend, m_BlFnc, 4
    mForePB.Cls
    'Debug.Print "Class: {Fore.hdc: " & mForePB.hDC & "; Back.hdc: " & mBackPB.hDC & "; AlphaB: " & Me.Alpha & "}"
    AlphaBlend mForePB.hdc, 0, 0, m_RF.Width, m_RF.Height, _
               mBackPB.hdc, 0, 0, m_RB.Width, m_RB.Height, m_Blend.Value
End Sub

Public Property Let AlphaS(ByVal Value As Single)
    m_BlFnc.SrcConstAlpha = 255 - CByte(Value * 255)
    SynchronizeBlend
    UpdateView
End Property
Public Property Get AlphaS() As Single
    AlphaS = 1 - m_BlFnc.SrcConstAlpha / 255
End Property

Public Property Let Alpha(ByVal Value As Byte)
    'wenn Alpha = 255 dann opaque = deckend
    'wenn Alpha < 255 dann transparent
    m_BlFnc.SrcConstAlpha = 255 - Value
    SynchronizeBlend
    UpdateView
End Property
Public Property Get Alpha() As Byte
    Alpha = 255 - m_BlFnc.SrcConstAlpha
    'AlphaB = m_Alpha * 255
End Property

Public Property Let BlendOp(ByVal Value As Byte)
    m_BlFnc.BlendOp = Value
    SynchronizeBlend
End Property
Public Property Get BlendOp() As Byte
    BlendOp = m_BlFnc.BlendOp
End Property

Public Property Let AlphaFormat(ByVal Value As Long)
    m_BlFnc.AlphaFormat = Value
    SynchronizeBlend
End Property
Public Property Get AlphaFormat() As Long
    AlphaFormat = m_BlFnc.AlphaFormat
End Property

Public Property Get BlendFunc() As Long
    BlendFunc = m_Blend.Value
End Property
Public Property Let BlendFunc(ByVal Value As Long)
    m_Blend.Value = Value
    SynchronizeBlFnc
End Property

'Private m_BlFnc As BLENDFUNCTION
'Private m_Blend As TLng
Private Sub SynchronizeBlFnc()
    LSet m_BlFnc = m_Blend
End Sub

Private Sub SynchronizeBlend()
    LSet m_Blend = m_BlFnc
End Sub

