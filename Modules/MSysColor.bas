Attribute VB_Name = "MSysColor"
Option Explicit

'OLE_COLORs:
'Namespace VBRUN
' System color constants (see also ColorConstants).
Public Enum SysteMColorConstants
    vbScrollBars = &H80000000           ' Scroll-bars gray area color. '
    vbDesktop = &H80000001              ' Desktop color. '
    vbActiveTitleBar = &H80000002       ' Active window caption color. '
    vbInactiveTitleBar = &H80000003     ' Inactive window caption color. '
    vbMenuBar = &H80000004              ' Menu background color. '
    vbWindowBackground = &H80000005     ' Window background color. '
    vbWindowFrame = &H80000006          ' Window frame color. '
    vbMenuText = &H80000007             ' Text color in menus. '
    vbWindowText = &H80000008           ' Text color in windows. '
    vbTitleBarText = &H80000009         ' Text color in active caption, size box, scroll-bar arrow box. '
    vbActiveTitleBarText = &H80000009   ' Text color in active caption, size box, scroll-bar arrow box. '
    vbActiveBorder = &H8000000A         ' Active window border color. '
    vbInactiveBorder = &H8000000B       ' Inactive window border color. '
    vbApplicationWorkspace = &H8000000C ' Background color of multiple document interface (MDI) applications. '
    vbHighlight = &H8000000D            ' Background color of items selected in a control. '
    vbHighlightText = &H8000000E        ' Text color of items selected in a control. '
    vbButtonFace = &H8000000F           ' Face shading on command buttons. '
    vb3DFace = &H8000000F               ' Dark shadow color for three-dimensional display elements. '
    vbButtonShadow = &H80000010         ' Edge shading on command buttons. '
    vb3DShadow = &H80000010             ' Color of automatic window shadows. '
    vbGrayText = &H80000011             ' Grayed (disabled) text. '
    vbButtonText = &H80000012           ' Text color on push buttons. '
    vbInactiveTitleBarText = &H80000013 ' Text color in inactive window caption, size box, scroll-bar arrow box. '
    vbInactiveCaptionText = &H80000013  ' Color of text in an inactive caption. '
    vb3DHighlight = &H80000014          ' Highlight color for 3D display elements. '
    vb3DDKShadow = &H80000015           ' Darkest shadow. '
    vb3DLight = &H80000016              ' Second lightest of the 3D colors after vb3DHilight. '
    vbInfoText = &H80000017             ' Color of text in ToolTips. '
    vbInfoBackground = &H80000018       ' Background color of ToolTips. '
End Enum

'SysteMColors:
Private Const COLOR_SCROLLBAR               As Long = 0
Private Const COLOR_BACKGROUND              As Long = 1
Private Const COLOR_DESKTOP                 As Long = COLOR_BACKGROUND
Private Const COLOR_ACTIVECAPTION           As Long = 2
Private Const COLOR_INACTIVECAPTION         As Long = 3
Private Const COLOR_MENU                    As Long = 4
Private Const COLOR_WINDOW                  As Long = 5
Private Const COLOR_WINDOWFRAME             As Long = 6
Private Const COLOR_MENUTEXT                As Long = 7
Private Const COLOR_WINDOWTEXT              As Long = 8
Private Const COLOR_CAPTIONTEXT             As Long = 9
Private Const COLOR_ACTIVEBORDER            As Long = 10
Private Const COLOR_INACTIVEBORDER          As Long = 11
Private Const COLOR_APPWORKSPACE            As Long = 12
Private Const COLOR_HIGHLIGHT               As Long = 13
Private Const COLOR_HIGHLIGHTTEXT           As Long = 14
Private Const COLOR_BTNFACE                 As Long = 15
Private Const COLOR_3DFACE                  As Long = COLOR_BTNFACE
Private Const COLOR_BTNSHADOW               As Long = 16
Private Const COLOR_3DSHADOW                As Long = COLOR_BTNSHADOW
Private Const COLOR_GRAYTEXT                As Long = 17
Private Const COLOR_BTNTEXT                 As Long = 18
Private Const COLOR_INACTIVECAPTIONTEXT     As Long = 19
Private Const COLOR_BTNHIGHLIGHT            As Long = 20
Private Const COLOR_BTNHILIGHT              As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHIGHLIGHT             As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHILIGHT               As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DDKSHADOW              As Long = 21
Private Const COLOR_3DLIGHT                 As Long = 22
Private Const COLOR_INFOTEXT                As Long = 23
Private Const COLOR_INFOBK                  As Long = 24
'25 ???
Private Const COLOR_HOTLIGHT                As Long = 26
Private Const COLOR_GRADIENTACTIVECAPTION   As Long = 27
Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28

'what's that???
Private Const COLOR_HUESCROLL  As Long = 700
Private Const COLOR_SATSCROLL  As Long = 701
Private Const COLOR_LUMSCROLL  As Long = 702
Private Const COLOR_HUE        As Long = 703
Private Const COLOR_SAT        As Long = 704
Private Const COLOR_LUM        As Long = 705
Private Const COLOR_RED        As Long = 706
Private Const COLOR_GREEN      As Long = 707
Private Const COLOR_BLUE       As Long = 708
Private Const COLOR_CURRENT    As Long = 709
Private Const COLOR_RAINBOW    As Long = 710
Private Const COLOR_SAVE       As Long = 711
Private Const COLOR_ADD        As Long = 712
Private Const COLOR_SOLID      As Long = 713
Private Const COLOR_TUNE       As Long = 714
Private Const COLOR_SCHEMES    As Long = 715
Private Const COLOR_ELEMENT    As Long = 716
Private Const COLOR_SAMPLES    As Long = 717
Private Const COLOR_PALETTE    As Long = 718
Private Const COLOR_MIX        As Long = 719
Private Const COLOR_BOX1       As Long = 720
Private Const COLOR_CUSTOM1    As Long = 721

Private Const COLOR_HUEACCEL   As Long = 723
Private Const COLOR_SATACCEL   As Long = 724
Private Const COLOR_LUMACCEL   As Long = 725
Private Const COLOR_REDACCEL   As Long = 726
Private Const COLOR_GREENACCEL As Long = 727
Private Const COLOR_BLUEACCEL  As Long = 728
Private Const COLOR_SOLID_LEFT As Long = 730
Private Const COLOR_SOLID_RIGHT As Long = 731

Private Const COLOR_ADJ_MAX    As Long = 100
Private Const COLOR_ADJ_MIN    As Long = -100

Private Const COLOR_MATCH_VERSION As Long = &H200
Private Const COLOR_NO_TRANSPARENT As Long = &HFFFFFFFF

'Public Enum SysteMColor
'    Desktop                 = &H80000001 '   -2147483647
'    ScrollBar               = &H80000001 '   -2147483647
'    ActiveCaption           = &H80000002 '   -2147483646
'    InactiveCaption         = &H80000003 '   -2147483645
'    Menu                    = &H80000004 '   -2147483644
'    Window                  = &H80000005 '   -2147483643
'    WindowFrame             = &H80000006 '   -2147483642
'    MenuText                = &H80000007 '   -2147483641
'    WindowText              = &H80000008 '   -2147483640
'    ActiveCaptionText       = &H80000009 '   -2147483639
'    ActiveBorder            = &H8000000A '   -2147483638
'    InactiveBorder          = &H8000000B '   -2147483637
'    AppWorkspace            = &H8000000C '   -2147483636
'    Highlight               = &H8000000D '   -2147483635
'    HighlightText           = &H8000000E '   -2147483634
'    ButtonFace              = &H8000000F '   -2147483633
'    Control                 = &H8000000F '   -2147483633
'    ButtonShadow            = &H80000010 '   -2147483632
'    ControlDark             = &H80000010 '   -2147483632
'    GrayText                = &H80000011 '   -2147483631
'    ControlText             = &H80000012 '   -2147483630
'    InactiveCaptionText     = &H80000013 '   -2147483629
'    ButtonHighlight         = &H80000014 '   -2147483628
'    ControlLightLight       = &H80000014 '   -2147483628
'    ControlDarkDark         = &H80000015 '   -2147483627
'    ControlLight            = &H80000016 '   -2147483626
'    InfoText                = &H80000017 '   -2147483625
'    Info                    = &H80000018 '   -2147483624
'    HotTrack                = &H8000001A '   -2147483622
'    GradientActiveCaption   = &H8000001B '   -2147483621
'    GradientInactiveCaption = &H8000001C '   -2147483620
'    MenuHighlight           = &H8000001D '   -2147483619
'    MenuBar                 = &H8000001E '   -2147483618
'End Enum
'

Public Enum SysteMColor 'aka VBRUN.SysteMColorConstants
    ActiveBorder = -2147483638
    ActiveCaption = -2147483646
    ActiveCaptionText = -2147483639
    AppWorkspace = -2147483636
    ButtonFace = -2147483633
    ButtonHighlight = -2147483628
    ButtonShadow = -2147483632
    Control = -2147483633
    ControlDark = -2147483632
    ControlDarkDark = -2147483627
    ControlLight = -2147483626
    ControlLightLight = -2147483628
    ControlText = -2147483630
    Desktop = -2147483647
    GradientActiveCaption = -2147483621
    GradientInactiveCaption = -2147483620
    GrayText = -2147483631
    Highlight = -2147483635
    HighlightText = -2147483634
    HotTrack = -2147483622
    InactiveBorder = -2147483637
    InactiveCaption = -2147483645
    InactiveCaptionText = -2147483629
    Info = -2147483624
    InfoText = -2147483625
    Menu = -2147483644
    MenuBar = -2147483618
    MenuHighlight = -2147483619
    MenuText = -2147483641
    ScrollBar = -2147483647
    Window = -2147483643
    WindowFrame = -2147483642
    WindowText = -2147483640
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Function SysteMColor_IsValid(ByVal aColor As Long) As Boolean
    If aColor > &H100FFFF Then Exit Function
    If aColor < 0 Then
        If (aColor And &HFF000000) <> &H80000000 Then Exit Function
        If (aColor And &HFFFF) > 28 Then Exit Function
    End If
    SysteMColor_IsValid = True
End Function

Public Function SysteMColor_ToColor(ByVal aOLEColor As Long) As Long
    'aOLEColor is a SysteMColorConstants
    Dim nIndex As Long: nIndex = &HFF& And aOLEColor
    SysteMColor_ToColor = GetSysColor(nIndex)
End Function

Public Sub SysteMColor_ToCB(CB As ComboBox)
    With CB
        .Clear
        Dim i As Long, S As String
        For i = 0 To 30
            S = SysteMColor_ToStr(i)
            If Len(S) Then .AddItem S
        Next
    End With
End Sub
'Private Const COLOR_SCROLLBAR               As Long = 0
'Private Const COLOR_BACKGROUND              As Long = 1
'Private Const COLOR_DESKTOP                 As Long = COLOR_BACKGROUND
'Private Const COLOR_ACTIVECAPTION           As Long = 2
'Private Const COLOR_INACTIVECAPTION         As Long = 3
'Private Const COLOR_MENU                    As Long = 4
'Private Const COLOR_WINDOW                  As Long = 5
'Private Const COLOR_WINDOWFRAME             As Long = 6
'Private Const COLOR_MENUTEXT                As Long = 7
'Private Const COLOR_WINDOWTEXT              As Long = 8
'Private Const COLOR_CAPTIONTEXT             As Long = 9
'Private Const COLOR_ACTIVEBORDER            As Long = 10
'Private Const COLOR_INACTIVEBORDER          As Long = 11
'Private Const COLOR_APPWORKSPACE            As Long = 12
'Private Const COLOR_HIGHLIGHT               As Long = 13
'Private Const COLOR_HIGHLIGHTTEXT           As Long = 14
'Private Const COLOR_BTNFACE                 As Long = 15
'Private Const COLOR_3DFACE                  As Long = COLOR_BTNFACE
'Private Const COLOR_BTNSHADOW               As Long = 16
'Private Const COLOR_3DSHADOW                As Long = COLOR_BTNSHADOW
'Private Const COLOR_GRAYTEXT                As Long = 17
'Private Const COLOR_BTNTEXT                 As Long = 18
'Private Const COLOR_INACTIVECAPTIONTEXT     As Long = 19
'Private Const COLOR_BTNHIGHLIGHT            As Long = 20
'Private Const COLOR_BTNHILIGHT              As Long = COLOR_BTNHIGHLIGHT
'Private Const COLOR_3DHIGHLIGHT             As Long = COLOR_BTNHIGHLIGHT
'Private Const COLOR_3DHILIGHT               As Long = COLOR_BTNHIGHLIGHT
'Private Const COLOR_3DDKSHADOW              As Long = 21
'Private Const COLOR_3DLIGHT                 As Long = 22
'Private Const COLOR_INFOTEXT                As Long = 23
'Private Const COLOR_INFOBK                  As Long = 24
''25 ???
'Private Const COLOR_HOTLIGHT                As Long = 26
'Private Const COLOR_GRADIENTACTIVECAPTION   As Long = 27
'Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28

Function SysteMColor_ToStr(C As SysteMColor) As String
    Dim S As String, nIndex As Long: nIndex = C And &HFF&
    Select Case C
    Case &H0&:  S = "ScrollBar"                ' 0   -2147483647
    Case &H1&:  S = "Desktop"                  ' 1   -2147483647
    Case &H2&:  S = "ActiveCaption"            ' 2   -2147483646
    Case &H3&:  S = "InactiveCaption"          ' 3   -2147483645
    Case &H4&:  S = "Menu"                     ' 4   -2147483644
    Case &H5&:  S = "Window"                   ' 5   -2147483643
    Case &H6&:  S = "WindowFrame"              ' 6   -2147483642
    Case &H7&:  S = "MenuText"                 ' 7   -2147483641
    Case &H8&:  S = "WindowText"               ' 8   -2147483640
    Case &H9&:  S = "ActiveCaptionText"        ' 9   -2147483639
    Case &HA&:  S = "ActiveBorder"             '10   -2147483638
    Case &HB&:  S = "InactiveBorder"           '11   -2147483637
    Case &HC&:  S = "AppWorkspace"             '12   -2147483636
    Case &HD&:  S = "Highlight"                '13   -2147483635
    Case &HE&:  S = "HighlightText"            '14   -2147483634
    'Case &HF&:  s = "ButtonFace"               '15   -2147483633
    'Case &HF&:  s = "Control"                  '15   -2147483633
    Case &HF&:  S = "ButtonFace"                  '15   -2147483633
    'Case &H10&: s = "ControlDark"              '16   -2147483632
    Case &H10&: S = "ButtonShadow"              '16   -2147483632
    Case &H11&: S = "GrayText"                 '17   -2147483631
    Case &H12&: S = "ControlText"              '18   -2147483630
    Case &H13&: S = "InactiveCaptionText"      '19   -2147483629
    'Case &H14&: s = "ButtonHighlight"          '20   -2147483628
    'Case &H14&: s = "ControlLightLight"        '20   -2147483628
    Case &H14&: S = "ButtonHighlight"        '20   -2147483628
    Case &H15&: S = "ControlDarkShadow"          '21   -2147483627
    Case &H16&: S = "ControlLight"             '22   -2147483626
    Case &H17&: S = "InfoText"                 '23   -2147483625
    Case &H18&: S = "Info"                     '24   -2147483624
    Case &H19&: S = "25?"
    Case &H1A&: S = "HotTrack"                 '26   -2147483622
    Case &H1B&: S = "GradientActiveCaption"    '27   -2147483621
    Case &H1C&: S = "GradientInactiveCaption"  '28   -2147483620
    Case &H1D&: S = "MenuHighlight"            '29   -2147483619
    Case &H1E&: S = "MenuBar"                  '30   -2147483618
    End Select
    SysteMColor_ToStr = S
End Function

'Private Function OLE_COLOR_IsValid(nColor As Long) As Boolean
'    Dim iLng As Long
'    OLE_COLOR_IsValid = True
'    If nColor > &H100FFFF Then
'        OLE_COLOR_IsValid = False
'    ElseIf nColor < 0 Then
'        If (nColor And &HFF000000) = &H80000000 Then
'            iLng = nColor And &HFFFF
'            If iLng > 18 Then
'                OLE_COLOR_IsValid = False
'            End If
'        Else
'            OLE_COLOR_IsValid = False
'        End If
'    End If
'End Function

'    Dim c As OLE_COLOR
'    c = Me.ForeColor:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = Me.FillColor:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = Me.BackColor:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbBlack:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbBlue:     Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbCyan:     Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbGreen:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbMagenta:  Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbRed:      Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbWhite:    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = ColorConstants.vbYellow:   Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'
'    c = SysteMColorConstants.vb3DDKShadow:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vb3DFace:               Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vb3DHighlight:          Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vb3DLight:              Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vb3DShadow:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbActiveBorder:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbActiveTitleBar:       Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbActiveTitleBarText:   Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbApplicationWorkspace: Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbButtonFace:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbButtonShadow:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbButtonText:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbDesktop:              Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbGrayText:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbHighlight:            Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbHighlightText:        Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInactiveBorder:       Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInactiveCaptionText:  Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInactiveTitleBar:     Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInactiveTitleBarText: Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInfoBackground:       Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbInfoText:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbMenuBar:              Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbMenuText:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbScrollBars:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbTitleBarText:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbWindowBackground:     Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbWindowFrame:          Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColorConstants.vbWindowText:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'
'    c = SysteMColor.ActiveBorder:            Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ActiveCaption:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ActiveCaptionText:       Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.AppWorkspace:            Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ButtonFace:              Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ButtonHighlight:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ButtonShadow:            Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Control:                 Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ControlDark:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ControlDarkDark:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ControlLight:            Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ControlLightLight:       Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ControlText:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Desktop:                 Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.GradientActiveCaption:   Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.GradientInactiveCaption: Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.GrayText:                Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Highlight:               Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.HighlightText:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.HotTrack:                Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.InactiveBorder:          Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.InactiveCaption:         Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.InactiveCaptionText:     Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Info:                    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.InfoText:                Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Menu:                    Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.MenuBar:                 Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.MenuHighlight:           Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.MenuText:                Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.ScrollBar:               Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.Window:                  Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.WindowFrame:             Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c
'    c = SysteMColor.WindowText:              Debug.Print OLECOLOR_IsValid(c) & " " & OLE_COLOR_IsValid(c) & " &H" & Hex(c) & " " & c


