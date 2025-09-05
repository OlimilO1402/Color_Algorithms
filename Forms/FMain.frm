VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Color Algorithms"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18495
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   18495
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfoGColor 
      Caption         =   "InfoCad Colors"
      Height          =   375
      Left            =   11400
      TabIndex        =   112
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox PBClosestMunsColor 
      Height          =   375
      Left            =   10800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   110
      ToolTipText     =   "Cick and move your mouse over your screen to view the color below. It shows the nearest color, now hit Enter to switch off. "
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton BtnMunsell 
      Caption         =   "Munsell"
      Height          =   375
      Left            =   9600
      TabIndex        =   32
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   13920
      TabIndex        =   38
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox PnlYCbCr 
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
      Left            =   16800
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   105
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetYCbCr 
         Caption         =   "Set  YCbCr"
         Height          =   375
         Left            =   0
         TabIndex        =   54
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBYCbCr_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   53
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBYCbCr_Cr 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   52
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBYCbCr_Cb 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   51
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBYCbCr_L 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   50
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   109
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Cr:"
         Height          =   225
         Left            =   0
         TabIndex        =   108
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Cb:"
         Height          =   225
         Left            =   0
         TabIndex        =   107
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   225
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.PictureBox PnlCIELab 
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
      Left            =   15000
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   100
      Top             =   120
      Width           =   1575
      Begin VB.ComboBox CmbCIELabLight 
         Height          =   345
         Left            =   360
         TabIndex        =   47
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TBCIELab_L 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   44
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox TBCIELab_aa 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBCIELab_bb 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBCIELab_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   48
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton BtnSetCIELab 
         Caption         =   "Set  CIELab"
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "L*:"
         Height          =   225
         Left            =   0
         TabIndex        =   104
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "a*:"
         Height          =   225
         Left            =   0
         TabIndex        =   103
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "b*:"
         Height          =   225
         Left            =   0
         TabIndex        =   102
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   101
         Top             =   1440
         Width           =   165
      End
   End
   Begin VB.PictureBox PBClosestRALColor 
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   99
      ToolTipText     =   "Cick and move your mouse over your screen to view the color below. It shows the nearest color, now hit Enter to switch off. "
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox PBClosestKnownColor 
      Height          =   375
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   98
      ToolTipText     =   "Cick and move your mouse over your screen to view the color below. It shows the nearest color, now hit Enter to switch off. "
      Top             =   2880
      Width           =   375
   End
   Begin VB.ComboBox CmbSysColor 
      Height          =   345
      Left            =   6840
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.PictureBox PnlHSLA 
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
      Left            =   9600
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   90
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetHSLA 
         Caption         =   "Set  HSLA"
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBHSLA_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBHSLA_L 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBHSLA_S 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBHSLA_H 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   94
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "L:"
         Height          =   225
         Left            =   0
         TabIndex        =   93
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   225
         Left            =   0
         TabIndex        =   92
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         Height          =   225
         Left            =   0
         TabIndex        =   91
         Top             =   0
         Width           =   180
      End
   End
   Begin VB.ComboBox CmbRALClassic 
      Height          =   345
      Left            =   2880
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton BtnOnOff 
      Caption         =   "on/off"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Cick and move your mouse over your screen to view the color below. It shows the nearest color, now hit Enter to switch off. "
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox CBValues 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   87
      Text            =   "Combo2"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox CBValuesf_ 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   86
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
      Left            =   11400
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   81
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetHSV 
         Caption         =   "Set  HSV"
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBHSV_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   36
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBHSV_V 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBHSV_S 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBHSV_H 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   85
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "V:"
         Height          =   225
         Left            =   0
         TabIndex        =   84
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   225
         Left            =   0
         TabIndex        =   83
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         Height          =   225
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   180
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
      Left            =   13200
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   76
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetXYZ 
         Caption         =   "Set  XYZ"
         Height          =   375
         Left            =   0
         TabIndex        =   43
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBXYZ_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   42
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_Z 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_Y 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBXYZ_X 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   39
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   80
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   225
         Left            =   0
         TabIndex        =   79
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   225
         Left            =   0
         TabIndex        =   78
         Top             =   360
         Width           =   150
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   225
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.PictureBox PnlHSLAf 
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
      Left            =   7800
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   62
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TBHSLAf_H 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox TBHSLAf_S 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBHSLAf_L 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBHSLAf_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton BtnSetHSLAf 
         Caption         =   "Set  HSLAf"
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "H:"
         Height          =   225
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "S:"
         Height          =   225
         Left            =   0
         TabIndex        =   74
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "L:"
         Height          =   225
         Left            =   0
         TabIndex        =   73
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   72
         Top             =   1440
         Width           =   165
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
      Left            =   4200
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   61
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetRGBAf 
         Caption         =   "Set RGBAf"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBRGBAf_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_B 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_G 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBRGBAf_R 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   71
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   225
         Left            =   0
         TabIndex        =   70
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   225
         Left            =   0
         TabIndex        =   69
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   225
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   150
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
      Left            =   2400
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   60
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TBCMYK_Y 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_K 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton BtnSetCMYK 
         Caption         =   "Set CMYKA"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBCMYK_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_M 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBCMYK_C 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "C:"
         Height          =   225
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "M:"
         Height          =   225
         Left            =   0
         TabIndex        =   66
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   225
         Left            =   0
         TabIndex        =   65
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "K:"
         Height          =   225
         Left            =   0
         TabIndex        =   64
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   63
         Top             =   1440
         Width           =   165
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
      Left            =   6000
      ScaleHeight     =   2295
      ScaleWidth      =   1575
      TabIndex        =   55
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton BtnSetRGBA 
         Caption         =   "Set RGBA"
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TBRGBA_A 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_B 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_G 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TBRGBA_R 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   225
         Left            =   0
         TabIndex        =   59
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   225
         Left            =   0
         TabIndex        =   58
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   225
         Left            =   0
         TabIndex        =   57
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   225
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.TextBox TBLngColor 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   1695
      Width           =   2055
   End
   Begin VB.ComboBox CmbColorNames 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2055
   End
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
      Left            =   360
      ScaleHeight     =   167.23
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   120.012
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox PbPicture 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      Picture         =   "FMain.frx":0CCA
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   88
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label LblClosestMunsColor 
      Caption         =   " "
      Height          =   375
      Left            =   9600
      TabIndex        =   111
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LblClosestRALColor 
      Caption         =   " "
      Height          =   375
      Left            =   2880
      TabIndex        =   97
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label LblClosestKnownColor 
      Caption         =   " "
      Height          =   375
      Left            =   120
      TabIndex        =   96
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label LblSysColors 
      AutoSize        =   -1  'True
      Caption         =   "SysColor:"
      Height          =   225
      Left            =   6000
      TabIndex        =   95
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LblRALColors 
      AutoSize        =   -1  'True
      Caption         =   "RAL:"
      Height          =   225
      Left            =   2400
      TabIndex        =   89
      Top             =   2520
      Width           =   360
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

'HWND SetCapture([in] HWND hWnd);
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
'[in] hWnd Type: HWND 'A handle to the window in the current thread that is to capture the mouse.

'BOOL ReleaseCapture();
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private m_OldHwnd As LongPtr

Private CurMousePos As POINTAPI

' Everything will be converted to CMYK
Private m_CMYK    As CMYK
Private m_TBBack  As TextBox
Private m_PnlHwnd As Long
Private m_Btn     As CommandButton
Private m_Max     As Single
Private m_CPicker As ColorDialog
Private m_APB     As AlphaPB
Private m_IsInit  As Boolean

'Private Sub Form_Activate()
'    Debug.Print "Activate "
'End Sub

Private Sub Form_Load()
    Set m_CPicker = New ColorDialog
    Set m_APB = MNew.AlphaPB(Me.PBColor, Me.PbPicture)
    Me.Caption = "Color Algorithms v" & Me.Version
    Timer1.Interval = 50
    Timer1.Enabled = False
    'FillCmbMouseScrollf CBValuesf
    'FillCmbMouseScroll CBValues
    HideCBValues
    MKnownColors.X11KnownColor_ToCB CmbColorNames
    MRALColors.RALClassic_ToListBox CmbRALClassic
    MSysColor.SystemColor_ToCB CmbSysColor
    CIELabLight_ToCmb CmbCIELabLight
    CmbCIELabLight.ListIndex = 1
    PBColor.BackColor = vbCyan
    m_CMYK = RGBAf_ToCMYK(LngColor_ToRGBAf(LngColor(PBColor.BackColor)))
    SetToolTipText GetControls("TextBox")
    UpdateView
End Sub

Private Sub BtnInfoGColor_Click()
    Dim Color As Long: Color = PBColor.BackColor
    If FInfoGColors.ShowDialog(Me, Color) = VbMsgBoxResult.vbCancel Then Exit Sub
    Dim lc As LngColor: lc.Value = Color
    m_CMYK = LngColor_ToCMYK(lc)
    UpdateView
End Sub


'does not work
'Private Sub Form_LostFocus()
'    Me.Caption = "Form_LostFocus "
'    If Timer1.Enabled Then BtnOnOff_Click
'End Sub
'does not work
'Private Sub Form_Deactivate()
'    Me.Caption = "Form_Deactivate "
'    If Timer1.Enabled Then BtnOnOff_Click
'End Sub

Private Sub BtnMunsell_Click()
    'FMunsell.Show
    Dim Color As Long: Color = PBColor.BackColor
    If FMunsell.ShowDialog(Me, Color) = VbMsgBoxResult.vbCancel Then Exit Sub
    Dim lc As LngColor: lc.Value = Color
    m_CMYK = LngColor_ToCMYK(lc)
    UpdateView
End Sub

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & Me.Version & vbCrLf & App.FileDescription
End Sub

Private Sub CmbCIELabLight_Click()
    UpdateView
End Sub

Private Sub LblClosestKnownColor_DblClick()
    Dim sc As String: sc = LblClosestKnownColor.Caption
    If Len(sc) = 0 Then Exit Sub
    Dim i As Long: i = X11KnownColor_IndexFromName(sc)
    CmbColorNames.ListIndex = i
End Sub

Private Sub LblClosestRALColor_DblClick()
    Dim sc As String: sc = LblClosestRALColor.Caption
    If Len(sc) = 0 Then Exit Sub
    Dim i As Long: i = MRALColors.TNamedRALColor_IndexFromName(sc)
    CmbRALClassic.ListIndex = i
End Sub

Private Sub LblClosestMunsColor_DblClick()
    Dim sc As String: sc = LblClosestMunsColor.Caption
    If Len(sc) = 0 Then Exit Sub
    Dim mc As TMunsellColor: mc = MMunsell.MunsellColors_ParseColorFromKey(sc)
    Dim lc As LngColor: lc = MColor.RGBA_ToLngColor(mc.RGBA)
    m_CMYK = MColor.LngColor_ToCMYK(lc)
    UpdateView
    'Dim i As Long: i = MMunsell.TNamedRALColor_IndexFromName(sc)
    'CmbRALClassic.ListIndex = i
End Sub
'Private Sub BtnMunsell_Click()
'    'FMunsell.Show
'    Dim Color As Long: Color = PBColor.BackColor
'    If FMunsell.ShowDialog(Me, Color) = vbCancel Then Exit Sub
'    'PBColor.BackColor = Color
'    Dim lc As LngColor: lc.Value = Color
'    m_CMYK = LngColor_ToCMYK(lc)
'    UpdateView
'End Sub

Private Sub TBLngColor_LostFocus()
    Dim lc As LngColor: lc = MColor.LngColor_ParseWebHex(TBLngColor.Text)
    m_CMYK = MColor.LngColor_ToCMYK(lc)
    UpdateView
End Sub

Private Sub Timer1_Timer()
    GetCursorPos CurMousePos
    Dim c As Long: c = ColorUnderMouse(CurMousePos.X, CurMousePos.Y)
    PBColor.BackColor = c
    m_CMYK = RGBAf_ToCMYK(MColor.LngColor_ToRGBAf(LngColor(c)))
    UpdateView
End Sub

Private Function ColorUnderMouse(ByVal X As Long, ByVal Y As Long) As Long
    ColorUnderMouse = GetPixel(GetDC(0), X, Y)
End Function

Private Sub BtnOnOff_Click()
    Timer1.Enabled = Not Timer1.Enabled
    BtnOnOff.Caption = IIf(Timer1.Enabled, "on", "off")
End Sub

Sub UpdateView(Optional bNoUpdateColorName As Boolean = False)
    
    MColor.CMYK_ToView TBCMYK_C, TBCMYK_M, TBCMYK_Y, TBCMYK_K, TBCMYK_A, m_CMYK
    
    Dim RGBAf As RGBAf:   RGBAf = MColor.CMYK_ToRGBAf(m_CMYK)
    MColor.RGBAf_ToView TBRGBAf_R, TBRGBAf_G, TBRGBAf_B, TBRGBAf_A, RGBAf
    
    Dim RGBA  As RGBA:     RGBA = MColor.RGBAf_ToRGBA(RGBAf)
    m_APB.Alpha = 255 - RGBA.a
    MColor.RGBA_ToView TBRGBA_R, TBRGBA_G, TBRGBA_B, TBRGBA_A, RGBA
    
    Dim alp As Single: alp = RGBA.a
    
    Dim LCol  As LngColor: LCol = MColor.RGBA_ToLngColor(RGBA)
    TBLngColor.Text = MColor.LngColor_ToWebHex(LCol)
    
    RGBA.a = 0
    LCol = MColor.RGBA_ToLngColor(RGBA)
    PBColor.BackColor = LCol.Value
    RGBA.a = alp
    
    Dim HSLAf As HSLAf: HSLAf = RGBAf_ToHSLAf(RGBAf)
    MColor.HSLAf_ToView TBHSLAf_H, TBHSLAf_S, TBHSLAf_L, TBHSLAf_A, HSLAf
    
    Dim HSLA As HSLA: HSLA = RGBA_ToHSLA(RGBA)
    MColor.HSLA_ToView TBHSLA_H, TBHSLA_S, TBHSLA_L, TBHSLA_A, HSLA
    
    Dim HSV As HSV: HSV = RGBAf_ToHSV(RGBAf)
    MColor.HSV_ToView TBHSV_H, TBHSV_S, TBHSV_V, TBHSV_A, HSV
    
    Dim XYZ As XYZ: XYZ = RGBAf_ToXYZ(RGBAf)
    MColor.XYZ_ToView TBXYZ_X, TBXYZ_Y, TBXYZ_Z, TBXYZ_A, XYZ
    
    Dim Lab As CIELab: Lab = XYZ_ToCIELab(XYZ, CmbCIELabLight.ListIndex)
    MColor.CIELab_ToView TBCIELab_L, TBCIELab_aa, TBCIELab_bb, TBCIELab_A, Lab
    
    Dim YCbCr As YCbCr: YCbCr = RGBA_ToYCbCr(RGBA)
    MColor.YCbCr_ToView TBYCbCr_L, TBYCbCr_Cb, TBYCbCr_Cr, TBYCbCr_A, YCbCr
    
    If Not bNoUpdateColorName Then
        Dim xn As String: xn = MKnownColors.NameFromColor(LCol.Value)
        If Len(xn) Then CmbColorNames.Text = xn
    End If
    
    m_APB.Alpha = 255 - alp
    
    Dim c As Long: c = LCol.Value
    'get closest color from knowncolors list:
    Dim nc As TNamedColor: nc = MKnownColors.X11KnownColor_ClosestColorTo(c)
    LblClosestKnownColor.Caption = nc.Name
    PBClosestKnownColor.BackColor = (&HFFFFFF And nc.X11Col)
    
    'get closest color from RAL-colors list:
    Dim rc As TNamedRALColor: rc = MRALColors.TNamedRALColor_ClosestColorTo(c)
    LblClosestRALColor.Caption = MRALColors.TNamedRALColor_ToStr(rc) ' "RAL_" & rc.RALNr & "_" & rc.Name
    PBClosestRALColor.BackColor = rc.RALCol
    
    'get closest color from munsell-colors-list
    Dim mc As TMunsellColor: mc = MMunsell.MunsellColors_ClosestColorTo(c)
    LblClosestMunsColor.Caption = MMunsell.TMunsellColor_Key(mc)
    PBClosestMunsColor.BackColor = RGBA_ToLngColor(mc.RGBA).Value
    
End Sub

'Private Sub ErrMsg(sErr As String)
'    ErrHandler "Invalid numeric value: " & sErr & vbCrLf & "please try again"
'End Sub

Private Sub BtnSetCMYK_Click()
    Dim sErr As String
    If Not MColor.CMYK_Read(m_CMYK, TBCMYK_C, TBCMYK_M, TBCMYK_Y, TBCMYK_K, TBCMYK_A, sErr) Then ErrHandler sErr: Exit Sub
    UpdateView
End Sub

Private Sub BtnSetRGBAf_Click()
    Dim RGBAf As RGBAf, sErr As String
    If Not MColor.RGBAf_Read(RGBAf, TBRGBAf_R, TBRGBAf_G, TBRGBAf_B, TBRGBAf_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(RGBAf)
    UpdateView
End Sub

Private Sub BtnSetRGBA_Click()
    Dim RGBA As RGBA, sErr As String
    If Not MColor.RGBA_Read(RGBA, TBRGBA_R, TBRGBA_G, TBRGBA_B, TBRGBA_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.RGBA_ToRGBAf(RGBA))
    UpdateView
End Sub

Private Sub BtnSetHSLAf_Click()
    Dim HSLAf As HSLAf, sErr As String
    If Not MColor.HSLAf_Read(HSLAf, TBHSLAf_H, TBHSLAf_S, TBHSLAf_L, TBHSLAf_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.HSLAf_ToRGBAf(HSLAf))
    UpdateView
End Sub

Private Sub BtnSetHSLA_Click()
    Dim HSLA As HSLA, sErr As String
    If Not MColor.HSLA_Read(HSLA, TBHSLA_H, TBHSLA_S, TBHSLA_L, TBHSLA_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(RGBA_ToRGBAf(MColor.HSLA_ToRGBA(HSLA)))
    UpdateView
End Sub

Private Sub BtnSetHSV_Click()
    Dim HSV As HSV, sErr As String
    If Not MColor.HSV_Read(HSV, TBHSV_H, TBHSV_S, TBHSV_V, TBHSV_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.HSV_ToRGBAf(HSV))
    UpdateView
End Sub

Private Sub BtnSetXYZ_Click()
    Dim XYZ As XYZ, sErr As String
    If Not MColor.XYZ_Read(XYZ, TBXYZ_X, TBXYZ_Y, TBXYZ_Z, TBXYZ_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.XYZ_ToRGBAf(XYZ))
    UpdateView
End Sub

Private Sub BtnSetCIELab_Click()
    Dim Lab As CIELab, sErr As String
    If Not MColor.CIELab_Read(Lab, TBCIELab_L, TBCIELab_aa, TBCIELab_bb, TBCIELab_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.XYZ_ToRGBAf(MColor.CIELab_ToXYZ(Lab)))
    UpdateView
End Sub

Private Sub BtnSetYCbCr_Click()
    Dim ycc As YCbCr, sErr As String
    If Not MColor.YCbCr_Read(ycc, TBYCbCr_L, TBYCbCr_Cb, TBYCbCr_Cr, TBYCbCr_A, sErr) Then ErrHandler sErr: Exit Sub
    m_CMYK = RGBAf_ToCMYK(MColor.YCbCr_ToRGBAf(ycc))
    UpdateView
End Sub

Private Sub CmbColorNames_Click()
    If CmbColorNames.Text = "" Then Exit Sub
    PBColor.BackColor = MKnownColors.ColorByName(CmbColorNames.Text)
    Dim LngColor As LngColor: LngColor.Value = PBColor.BackColor
    m_CMYK = MColor.RGBAf_ToCMYK(MColor.RGBA_ToRGBAf(MColor.LngColor_ToRGBA(LngColor)))
    UpdateView True
End Sub

Private Sub CmbRALClassic_Click()
    If CmbRALClassic.Text = "" Then Exit Sub
    Dim RALCol As Long: RALCol = MRALColors.RALClassic_Parse(CmbRALClassic.Text)
    PBColor.BackColor = RALCol
    Dim LngColor As LngColor: LngColor.Value = PBColor.BackColor
    m_CMYK = MColor.RGBAf_ToCMYK(MColor.RGBA_ToRGBAf(MColor.LngColor_ToRGBA(LngColor)))
    UpdateView True
End Sub

Private Sub CmbSysColor_Click()
    Dim i As Long: i = CmbSysColor.ListIndex
    Dim L As LngColor: L.Value = MSysColor.SystemColor_ToColor(i)
    'PBColor.BackColor = L.Value
    m_CMYK = LngColor_ToCMYK(L)
    UpdateView
End Sub

Private Sub FillCmbMouseScrollf(Cmb As ComboBox)
    Dim i As Long
    Cmb.Clear
    Dim n As Long: n = 256
    Dim Fact As Double: Fact = 1 / n
    For i = n To 0 Step -1
        Cmb.AddItem Format(i * Fact, "0.###")
    Next
End Sub

Private Sub FillCmbMouseScroll(Cmb As ComboBox)
    'CBValues
    Dim i As Long
    Dim n As Long: n = 255
    For i = n To 0 Step -1
        Cmb.AddItem i
    Next
End Sub

Private Sub PBColor_DblClick()
Try: On Error GoTo Catch
    m_CPicker.Color = PBColor.BackColor
    If m_CPicker.ShowDialog = vbCancel Then Exit Sub
    PBColor.BackColor = m_CPicker.Color
    'so OK
    'jetzt BackColor
    Dim L As LngColor: L.Value = m_CPicker.Color
    Dim RGBA As RGBA: RGBA = LngColor_ToRGBA(L)
    RGBA.a = CByte(TBRGBA_A.Text)
    m_CMYK = RGBA_ToCMYK(RGBA)
    UpdateView
Catch:
End Sub

Private Sub HideCBValues()
    CBValues.ZOrder 1
    'CBValuesf.ZOrder 1
End Sub

Private Sub PnlCMYK_DblClick():   HideCBValues:  End Sub
Private Sub PnlRGBAf_DblClick():  HideCBValues:  End Sub
Private Sub PnlRGBA_DblClick():   HideCBValues:  End Sub
Private Sub PnlHSLAf_DblClick():  HideCBValues:  End Sub
Private Sub PnlHSLA_DblClick():   HideCBValues:  End Sub
Private Sub PnlHSV_DblClick():    HideCBValues:  End Sub
Private Sub PnlXYZ_DblClick():    HideCBValues:  End Sub
Private Sub PnlCIELab_DblClick(): HideCBValues:  End Sub
Private Sub PnlYCbCr_DblClick():  HideCBValues:  End Sub

Private Sub TBCMYK_C_DblClick():  SetTB TBCMYK_C, BtnSetCMYK, PnlCMYK.hwnd, MColor.CVR_CMYK_C: End Sub
Private Sub TBCMYK_M_DblClick():  SetTB TBCMYK_M, BtnSetCMYK, PnlCMYK.hwnd, MColor.CVR_CMYK_M: End Sub
Private Sub TBCMYK_Y_DblClick():  SetTB TBCMYK_Y, BtnSetCMYK, PnlCMYK.hwnd, MColor.CVR_CMYK_Y: End Sub
Private Sub TBCMYK_K_DblClick():  SetTB TBCMYK_K, BtnSetCMYK, PnlCMYK.hwnd, MColor.CVR_CMYK_K: End Sub
Private Sub TBCMYK_A_DblClick():  SetTB TBCMYK_A, BtnSetCMYK, PnlCMYK.hwnd, MColor.CVR_CMYK_A: End Sub

Private Sub TBRGBAf_R_DblClick(): SetTB TBRGBAf_R, BtnSetRGBAf, PnlRGBAf.hwnd, MColor.CVR_RGBAf_R: End Sub
Private Sub TBRGBAf_G_DblClick(): SetTB TBRGBAf_G, BtnSetRGBAf, PnlRGBAf.hwnd, MColor.CVR_RGBAf_G: End Sub
Private Sub TBRGBAf_B_DblClick(): SetTB TBRGBAf_B, BtnSetRGBAf, PnlRGBAf.hwnd, MColor.CVR_RGBAf_B: End Sub
Private Sub TBRGBAf_A_DblClick(): SetTB TBRGBAf_A, BtnSetRGBAf, PnlRGBAf.hwnd, MColor.CVR_RGBAf_A: End Sub

Private Sub TBRGBA_R_DblClick():  SetTB TBRGBA_R, BtnSetRGBA, PnlRGBA.hwnd, MColor.CVR_RGBA_R: End Sub
Private Sub TBRGBA_G_DblClick():  SetTB TBRGBA_G, BtnSetRGBA, PnlRGBA.hwnd, MColor.CVR_RGBA_G: End Sub
Private Sub TBRGBA_B_DblClick():  SetTB TBRGBA_B, BtnSetRGBA, PnlRGBA.hwnd, MColor.CVR_RGBA_B: End Sub
Private Sub TBRGBA_A_DblClick():  SetTB TBRGBA_A, BtnSetRGBA, PnlRGBA.hwnd, MColor.CVR_RGBA_A: End Sub

Private Sub TBHSLAf_H_DblClick():  SetTB TBHSLAf_H, BtnSetHSLAf, PnlHSLAf.hwnd, MColor.CVR_HSLAf_H: End Sub
Private Sub TBHSLAf_S_DblClick():  SetTB TBHSLAf_S, BtnSetHSLAf, PnlHSLAf.hwnd, MColor.CVR_HSLAf_S: End Sub
Private Sub TBHSLAf_L_DblClick():  SetTB TBHSLAf_L, BtnSetHSLAf, PnlHSLAf.hwnd, MColor.CVR_HSLAf_L: End Sub
Private Sub TBHSLAf_A_DblClick():  SetTB TBHSLAf_A, BtnSetHSLAf, PnlHSLAf.hwnd, MColor.CVR_HSLAf_A: End Sub

Private Sub TBHSLA_H_DblClick():  SetTB TBHSLA_H, BtnSetHSLA, PnlHSLA.hwnd, MColor.CVR_HSL_H: End Sub
Private Sub TBHSLA_S_DblClick():  SetTB TBHSLA_S, BtnSetHSLA, PnlHSLA.hwnd, MColor.CVR_HSL_S: End Sub
Private Sub TBHSLA_L_DblClick():  SetTB TBHSLA_L, BtnSetHSLA, PnlHSLA.hwnd, MColor.CVR_HSL_L: End Sub
Private Sub TBHSLA_A_DblClick():  SetTB TBHSLA_A, BtnSetHSLA, PnlHSLA.hwnd, MColor.CVR_HSL_A: End Sub
'
Private Sub TBHSV_H_DblClick():  SetTB TBHSV_H, BtnSetHSV, PnlHSV.hwnd, MColor.CVR_HSV_H: End Sub
Private Sub TBHSV_S_DblClick():  SetTB TBHSV_S, BtnSetHSV, PnlHSV.hwnd, MColor.CVR_HSV_S: End Sub
Private Sub TBHSV_V_DblClick():  SetTB TBHSV_V, BtnSetHSV, PnlHSV.hwnd, MColor.CVR_HSV_V: End Sub
Private Sub TBHSV_A_DblClick():  SetTB TBHSV_A, BtnSetHSV, PnlHSV.hwnd, MColor.CVR_HSV_A: End Sub

Private Sub TBXYZ_X_DblClick():  SetTB TBXYZ_X, BtnSetXYZ, PnlXYZ.hwnd, MColor.CVR_XYZ_X: End Sub
Private Sub TBXYZ_Y_DblClick():  SetTB TBXYZ_Y, BtnSetXYZ, PnlXYZ.hwnd, MColor.CVR_XYZ_Y: End Sub
Private Sub TBXYZ_Z_DblClick():  SetTB TBXYZ_Z, BtnSetXYZ, PnlXYZ.hwnd, MColor.CVR_XYZ_Z: End Sub
Private Sub TBXYZ_A_DblClick():  SetTB TBXYZ_A, BtnSetXYZ, PnlXYZ.hwnd, MColor.CVR_XYZ_A: End Sub

Private Sub TBCIELab_L_DblClick():  SetTB TBCIELab_L, BtnSetCIELab, PnlCIELab.hwnd, MColor.CVR_CIELab_L: End Sub
Private Sub TBCIELab_aa_DblClick(): SetTB TBCIELab_aa, BtnSetCIELab, PnlCIELab.hwnd, MColor.CVR_CIELab_aa: End Sub
Private Sub TBCIELab_bb_DblClick(): SetTB TBCIELab_bb, BtnSetCIELab, PnlCIELab.hwnd, MColor.CVR_CIELab_bb: End Sub
Private Sub TBCIELab_A_DblClick():  SetTB TBCIELab_A, BtnSetCIELab, PnlCIELab.hwnd, MColor.CVR_CIELab_A: End Sub

Private Sub TBYCbCr_L_DblClick():  SetTB TBYCbCr_L, BtnSetYCbCr, PnlYCbCr.hwnd, MColor.CVR_YCbCr_Y: End Sub
Private Sub TBYCbCr_Cb_DblClick(): SetTB TBYCbCr_Cb, BtnSetYCbCr, PnlYCbCr.hwnd, MColor.CVR_YCbCr_Cb: End Sub
Private Sub TBYCbCr_Cr_DblClick(): SetTB TBYCbCr_Cr, BtnSetYCbCr, PnlYCbCr.hwnd, MColor.CVR_YCbCr_Cr: End Sub
Private Sub TBYCbCr_A_DblClick():  SetTB TBYCbCr_A, BtnSetYCbCr, PnlYCbCr.hwnd, MColor.CVR_YCbCr_A: End Sub

Private Sub SetTB(TB As TextBox, Btn As CommandButton, ByVal pnlHwnd As LongPtr, CVR As ColorValueRange)
Try: On Error GoTo Catch
    Set m_TBBack = TB
    Set m_Btn = Btn
    'm_Max = MaxVal
    SetParent Me.CBValues.hwnd, pnlHwnd
    Me.CBValues.Move m_TBBack.Left, m_TBBack.Top
    m_IsInit = True
    MColor.ColorValueRange_ToComboBox CVR, Me.CBValues, m_TBBack.Text
    m_IsInit = False
    'CBValues.Text = m_TBBack.Text
    CBValues.ZOrder 0
    'Dim n As Single: n = 256
'    If f = 1 Then n = 255
'    cb.ListIndex = n - (f * CSng(m_TBBack.Text))
'    'CBValuesf.ZOrder 1
'    'CBValues.ZOrder 1
'    HideCBValues
'    cb.ZOrder 0
    Exit Sub
Catch:
    ErrHandler "SetTB"
End Sub

'Private Sub TBCMYK_C_DblClick():  SetTB TBCMYK_C, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
'Private Sub TBCMYK_M_DblClick():  SetTB TBCMYK_M, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
'Private Sub TBCMYK_Y_DblClick():  SetTB TBCMYK_Y, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
'Private Sub TBCMYK_K_DblClick():  SetTB TBCMYK_K, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
'Private Sub TBCMYK_A_DblClick():  SetTB TBCMYK_A, CBValuesf, BtnSetCMYK, PnlCMYK.hwnd, 256: End Sub
'
'Private Sub TBRGBAf_R_DblClick(): SetTB TBRGBAf_R, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
'Private Sub TBRGBAf_G_DblClick(): SetTB TBRGBAf_G, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
'Private Sub TBRGBAf_B_DblClick(): SetTB TBRGBAf_B, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
'Private Sub TBRGBAf_A_DblClick(): SetTB TBRGBAf_A, CBValuesf, BtnSetRGBAf, PnlRGBAf.hwnd, 256: End Sub
'
'Private Sub TBRGBA_R_DblClick():  SetTB TBRGBA_R, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
'Private Sub TBRGBA_G_DblClick():  SetTB TBRGBA_G, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
'Private Sub TBRGBA_B_DblClick():  SetTB TBRGBA_B, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
'Private Sub TBRGBA_A_DblClick():  SetTB TBRGBA_A, CBValues, BtnSetRGBA, PnlRGBA.hwnd, 1: End Sub
'
'Private Sub TBHSLAf_H_DblClick():  SetTB TBHSLAf_H, CBValuesf, BtnSetHSLAf, PnlHSLAf.hwnd, 256: End Sub
'Private Sub TBHSLAf_S_DblClick():  SetTB TBHSLAf_S, CBValuesf, BtnSetHSLAf, PnlHSLAf.hwnd, 256: End Sub
'Private Sub TBHSLAf_L_DblClick():  SetTB TBHSLAf_L, CBValuesf, BtnSetHSLAf, PnlHSLAf.hwnd, 256: End Sub
'Private Sub TBHSLAf_A_DblClick():  SetTB TBHSLAf_A, CBValuesf, BtnSetHSLAf, PnlHSLAf.hwnd, 256: End Sub
'
'Private Sub TBHSLA_H_DblClick():  SetTB TBHSLA_H, CBValues, BtnSetHSLA, PnlHSLA.hwnd, 1, 239: End Sub
'Private Sub TBHSLA_S_DblClick():  SetTB TBHSLA_S, CBValues, BtnSetHSLA, PnlHSLA.hwnd, 1, 240: End Sub
'Private Sub TBHSLA_L_DblClick():  SetTB TBHSLA_L, CBValues, BtnSetHSLA, PnlHSLA.hwnd, 1, 240: End Sub
'Private Sub TBHSLA_A_DblClick():  SetTB TBHSLA_A, CBValues, BtnSetHSLA, PnlHSLA.hwnd, 1: End Sub
''
'Private Sub TBHSV_H_DblClick():  SetTB TBHSV_H, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
'Private Sub TBHSV_S_DblClick():  SetTB TBHSV_S, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
'Private Sub TBHSV_V_DblClick():  SetTB TBHSV_V, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
'Private Sub TBHSV_A_DblClick():  SetTB TBHSV_A, CBValuesf, BtnSetHSV, PnlHSV.hwnd, 256: End Sub
'
'Private Sub TBXYZ_X_DblClick():  SetTB TBXYZ_X, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_Y_DblClick():  SetTB TBXYZ_Y, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_Z_DblClick():  SetTB TBXYZ_Z, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'Private Sub TBXYZ_A_DblClick():  SetTB TBXYZ_A, CBValuesf, BtnSetXYZ, PnlXYZ.hwnd, 256: End Sub
'
'Private Sub TBCIELab_L_DblClick():  SetTB TBCIELab_L, CBValuesf, BtnSetCIELab, PnlCIELab.hwnd, 256: End Sub
'Private Sub TBCIELab_aa_DblClick(): SetTB TBCIELab_aa, CBValuesf, BtnSetCIELab, PnlCIELab.hwnd, 256: End Sub
'Private Sub TBCIELab_bb_DblClick(): SetTB TBCIELab_bb, CBValuesf, BtnSetCIELab, PnlCIELab.hwnd, 256: End Sub
'Private Sub TBCIELab_A_DblClick():  SetTB TBCIELab_A, CBValuesf, BtnSetCIELab, PnlCIELab.hwnd, 256: End Sub
'
'Private Sub TBYCbCr_L_DblClick():  SetTB TBYCbCr_L, CBValuesf, BtnSetYCbCr, PnlYCbCr.hwnd, 1, 256: End Sub
'Private Sub TBYCbCr_Cb_DblClick(): SetTB TBYCbCr_Cb, CBValuesf, BtnSetYCbCr, PnlYCbCr.hwnd, 1, 256: End Sub
'Private Sub TBYCbCr_Cr_DblClick(): SetTB TBYCbCr_Cr, CBValuesf, BtnSetYCbCr, PnlYCbCr.hwnd, 1, 256: End Sub
'Private Sub TBYCbCr_A_DblClick():  SetTB TBYCbCr_A, CBValuesf, BtnSetYCbCr, PnlYCbCr.hwnd, 256: End Sub
'
'Private Sub SetTB(TB As TextBox, cb As ComboBox, Btn As CommandButton, ByVal pnlHwnd As LongPtr, ByVal f As Single, Optional ByVal MaxVal As Single)
'    Set m_TBBack = TB
'    Set m_Btn = Btn
'    m_Max = MaxVal
'    SetParent cb.hwnd, pnlHwnd
'    cb.Move m_TBBack.Left, m_TBBack.Top
'    Dim n As Single: n = 256
'    If f = 1 Then n = 255
'    cb.ListIndex = n - (f * CSng(m_TBBack.Text))
'    'CBValuesf.ZOrder 1
'    'CBValues.ZOrder 1
'    HideCBValues
'    cb.ZOrder 0
'End Sub
'
'Private Sub CBValuesf_DblClick()
'    m_TBBack.ZOrder 0
'End Sub
'Private Sub CBValuesf_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        CBValuesf_Click
'        m_TBBack.ZOrder 0
'    End If
'End Sub
'Private Sub CBValuesf_Click()
'    m_TBBack.Text = CBValuesf.Text
'    m_Btn.Value = True
'End Sub

Private Sub CBValues_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CBValues_Click
        m_TBBack.ZOrder 0
    End If
End Sub
Private Sub CBValues_Click()
    If m_IsInit Then Exit Sub
    m_TBBack.Text = CBValues.Text
    m_Btn.Value = True

'    Dim b As Byte, s As String: s = CBValues.Text
'    If Not Byte_TryParse(s, b) Then Exit Sub
'    If m_Max > 0 Then b = MinByt(CByte(m_Max), b)
'    m_TBBack.Text = CStr(b)
'    m_Btn.Value = True
End Sub

'the following 4 functions are for creating tooltiptexts
Function GetControls(OfType As String) As Collection
    Dim ctrl: Set GetControls = New Collection
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = OfType Then GetControls.Add ctrl
    Next
End Function

Sub SetToolTipText(Ctrls As Collection)
    Dim ttt As Collection: Set ttt = ColAddText(Array("R", "Red", "G", "Green", "B", "Blue", "A", "Alpha", _
                                                      "C", "Cyan", "M", "Magenta", "YL", "Yellow", "K", "Black", _
                                                      "H", "Hue", "S", "Saturation", "L", "Luminance", "V", "Value", _
                                                      "X", "X", "Y", "Y", "Z", "Z", _
                                                      "L*", "Lum", "a*", "aa", "b*", "bb", _
                                                      "Cb", "blue-diff", "Cr", "red-diff"))
    Dim nam As String
    Dim ctrl 'As VBControlExtender
    For Each ctrl In Ctrls
        nam = ctrl.Name
        If Len(nam) < 10 Then
            ctrl.ToolTipText = "Change the " & CreateToolTipText(nam, ttt) & ". Doubleclick for using the mousewheel."
        End If
    Next
End Sub

Function ColAddText(arr) As Collection
    Set ColAddText = New Collection
    Dim i As Long
    For i = 0 To UBound(arr) Step 2
        ColAddText.Add arr(i + 1), arr(i)
    Next
End Function

Function CreateToolTipText(ByVal nam As String, ttt As Collection) As String
    'Static FncCallCounter As Long
    'FncCallCounter = FncCallCounter + 1
    nam = Mid(nam, 3) 'f.i.: "HSV_H"
    Dim sa() As String: sa = Split(nam, "_")
    Dim u As Long: u = UBound(sa)
    If u = 1 Then
        Dim S As String ': s = "Change the "
        Dim c_1 As String: c_1 = sa(0)
        Dim c_2 As String: c_2 = sa(1)
        If Len(c_1) > 3 And c_2 = "Y" Then c_2 = "YL" 'tiny optimization for CMYK-text
        S = S & ttt.Item(c_2) & "-value of "
        Dim c11 As String
        Dim c12 As String
        Dim c13 As String
        
        If c_1 = "YCbCr" Then
            c11 = "L": c12 = "Cb": c13 = "Cr"
        Else
            c11 = Mid(c_1, 1, 1): c12 = Mid(c_1, 2, 1): c13 = Mid(c_1, 3, 1)
            If Len(c_1) > 3 And c13 = "Y" Then c13 = "YL" 'tiny optimization for CMYK-text
        End If
        S = S & c_1 & " (=" & ttt.Item(c11) & ", " & ttt.Item(c12) & ", " & ttt.Item(c13)
        If c_1 <> "YCbCr" And Len(c_1) > 3 Then
            Dim c14 As String: c14 = Mid(c_1, 4, 1)
            S = S & ", " & ttt.Item(c14)
        End If
        CreateToolTipText = S & ")"
    End If
    'Debug.Print FncCallCounter
End Function

Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function


