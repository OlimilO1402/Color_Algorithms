VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorValueRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TBBack  As TextBox
Private m_Panel   As PictureBox
Private m_Btn     As CommandButton
Private m_Min     As Single
Private m_Max     As Single
Private m_dx      As Single
Private m_Count   As Long

Friend Sub New_(TBBack As TextBox, Panel As PictureBox, Button As CommandButton, ByVal RangeMin As Single, ByVal RangeMax As Single, ByVal CountValues As Long)
    Set m_TBBack = TBBack: Set m_Panel = Panel: Set m_Btn = Button
    m_Min = RangeMin: m_Max = RangeMax: m_Count = CountValues
End Sub

Public Property Get ValueMin() As Single
    ValueMin = m_Min
End Property
Public Property Get ValueMax() As Single
    ValueMax = m_Max
End Property

Public Property Get Count() As Long
    Count = m_Count
End Property
