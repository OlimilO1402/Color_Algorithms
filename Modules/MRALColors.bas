Attribute VB_Name = "MRALColors"
Option Explicit
'https://de.wikipedia.org/wiki/RAL-Farbe
Public Enum RALClassic
    'RAL-1X Gelb und Beige
    RAL_1000_Gr�nbeige = &H88BACD
    RAL_1001_Beige = &H84B0D0
    RAL_1002_Sandgelb = &H6DAAD2
    RAL_1003_Signalgelb = &HA8F9&
    RAL_1004_Goldgelb = &H7B0E2
    RAL_1005_Honiggelb = &H8ECB&
    RAL_1006_Maisgelb = &H90E2&
    RAL_1007_Narzissengelb = &H8CE8&
    RAL_1011_Braunbeige = &H548AAF
    RAL_1012_Zitronengelb = &H22C0D9
    RAL_1013_Perlwei� = &HC6D9E3
    RAL_1014_Elfenbein = &H9AC4DD
    RAL_1015_Hellelfenbein = &HB5D2E6
    RAL_1016_Schwefelgelb = &H44F0EA
    RAL_1017_Safrangelb = &H52B7F4
    RAL_1018_Zinkgelb = &H3BE0F3
    RAL_1019_Graubeige = &H7D95A4
    RAL_1020_Olivgelb = &H64949A
    RAL_1021_Rapsgelb = &HB6F6&
    RAL_1023_Verkehrsgelb = &HB5F7&
    RAL_1024_Ockergelb = &H509CB8
    RAL_1026_Leuchtgelb = &HFFFF&
    RAL_1027_Currygelb = &H158CA3
    RAL_1028_Melonengelb = &H9BFF&
    RAL_1032_Ginstergelb = &HA3E2&
    RAL_1033_Dahliengelb = &H21ABFA
    RAL_1034_Pastellgelb = &H56ABED
    RAL_1035_Perlbeige = &H8599A2
    RAL_1036_Perlgold = &H497592
    RAL_1037_Sonnengelb = &H5A2EE
    
    'RAL-2X Orange
    RAL_2000_Gelborange = &H6EDA&
    RAL_2001_Rotorange = &H1B48BA
    RAL_2002_Blutorange = &H2739C6
    RAL_2003_Pastellorange = &H2B84FA
    RAL_2004_Reinorange = &H125BE7
    RAL_2005_Leuchtorange = &H1249FF
    RAL_2007_Leuchthellorange = &H21A4FF
    RAL_2008_Hellrotorange = &H216BED
    RAL_2009_Verkehrsorange = &H155E1
    RAL_2010_Signalorange = &H2F65D4
    RAL_2011_Tieforange = &HE6EE2
    RAL_2012_Lachsorange = &H506ADB
    RAL_2013_Perlorange = &H274595
    RAL_2017_RALOrange = &H244FA
    
    'RAL-3X Rot
    RAL_3000_Feuerrot = &H2029A7
    RAL_3001_Signalrot = &H23249B
    RAL_3002_Karminrot = &H21239B
    RAL_3003_Rubinrot = &H221A86
    RAL_3004_Purpurrot = &H231C6B
    RAL_3005_Weinrot = &H1F1959
    RAL_3007_Schwarzrot = &H22203E
    RAL_3009_Oxidrot = &H2C336D
    RAL_3011_Braunrot = &H2C297E
    RAL_3012_Beigerot = &H738DCB
    RAL_3013_Tomatenrot = &H2E329C
    RAL_3014_Altrosa = &H7974D4
    RAL_3015_Hellrosa = &HA69FD7
    RAL_3016_Korallenrot = &H3440AC
    RAL_3017_Ros� = &H5F54D3
    RAL_3018_Erdbeerrot = &H5241D1
    RAL_3020_Verkehrsrot = &H101EBB
    RAL_3022_Lachsrot = &H5568CC
    RAL_3024_Leuchtrot = &H212DFF
    RAL_3026_Leuchthellrot = &HFF&
    RAL_3027_Himbeerrot = &H4120B4
    RAL_3028_Reinrot = &H242CCC
    RAL_3031_Orientrot = &H3734A6
    RAL_3032_Perlrubinrot = &H211571
    RAL_3033_Perlrosa = &H434CB2
    
    'RAL-4X Violett
    RAL_4001_Rotlila = &H835A8A
    RAL_4002_Rotviolett = &H503D93
    RAL_4003_Erikaviolett = &H8C5FC4
    RAL_4004_Bordeauxviolett = &H391669
    RAL_4005_Blaulila = &H9D6383
    RAL_4006_Verkehrspurpur = &H722599
    RAL_4007_Purpurviolett = &H3B204A
    RAL_4008_Signalviolett = &H844D88
    RAL_4009_Pastellviolett = &H9589A3
    RAL_4010_Telemagenta = &H7836C6
    RAL_4011_Perlviolett = &HA17387
    RAL_4012_Perlbrombeer = &H80686B
    
    'RAL-5X Blau
    RAL_5000_Violettblau = &H6F4E38
    RAL_5001_Gr�nblau = &H644C0F
    RAL_5002_Ultramarinblau = &H7B3800
    RAL_5003_Saphirblau = &H56372A
    RAL_5004_Schwarzblau = &H281E19
    RAL_5005_Signalblau = &H875300
    RAL_5007_Brillantblau = &H8D6741
    RAL_5008_Graublau = &H483C31
    RAL_5009_Azurblau = &H78592E
    RAL_5010_Enzianblau = &H7C4F00
    RAL_5011_Stahlblau = &H3C2B1A
    RAL_5012_Lichtblau = &HB88134
    RAL_5013_Kobaltblau = &H533119
    RAL_5014_Taubenblau = &H987C6C
    RAL_5015_Himmelblau = &HB27428
    RAL_5017_Verkehrsblau = &H8C5A00
    RAL_5018_T�rkisblau = &H8F8821
    RAL_5019_Capriblau = &H84571A
    RAL_5020_Ozeanblau = &H51410B
    RAL_5021_Wasserblau = &H7A7307
    RAL_5022_Nachtblau = &H5A2D22
    RAL_5023_Fernblau = &H8E664D
    RAL_5024_Pastellblau = &HB0936A
    RAL_5025_Perlenzian = &H786429
    RAL_5026_Perlnachtblau = &H542C10
    
    'RAL-6X Gr�n
    RAL_6000_Patinagr�n = &H60743C
    RAL_6001_Smaragdgr�n = &H356736
    RAL_6002_Laubgr�n = &H285932
    RAL_6003_Olivgr�n = &H3C5350
    RAL_6004_Blaugr�n = &H424402
    RAL_6005_Moosgr�n = &H324211
    RAL_6006_Grauoliv = &H2E393C
    RAL_6007_Flaschengr�n = &H22322C
    RAL_6008_Braungr�n = &H2A3437
    RAL_6009_Tannengr�n = &H2A3527
    RAL_6010_Grasgr�n = &H396F4D
    RAL_6011_Resedagr�n = &H597C6C
    RAL_6012_Schwarzgr�n = &H3A3D30
    RAL_6013_Schilfgr�n = &H5A767D
    RAL_6014_Gelboliv = &H354147
    RAL_6015_Schwarzoliv = &H363D3D
    RAL_6016_T�rkisgr�n = &H4C6900
    RAL_6017_Maigr�n = &H407F58
    RAL_6018_Gelbgr�n = &H3B9961
    RAL_6019_Wei�gr�n = &HACCEB9
    RAL_6020_Chromoxidgr�n = &H2F4237
    RAL_6021_Blassgr�n = &H77998A
    RAL_6022_Braunoliv = &H27333A
    RAL_6024_Verkehrsgr�n = &H518300
    RAL_6025_Farngr�n = &H3B6E5E
    RAL_6026_Opalgr�n = &H4E5F00
    RAL_6027_Lichtgr�n = &HB5BA7E
    RAL_6028_Kieferngr�n = &H425431
    RAL_6029_Minzgr�n = &H3D6F00
    RAL_6032_Signalgr�n = &H527F23
    RAL_6033_Mintt�rkis = &H7F8746
    RAL_6034_Pastellt�rkis = &HACAC7A
    RAL_6035_Perlgr�n = &H254D19
    RAL_6036_Perlopalgr�n = &H4B5704
    RAL_6037_Reingr�n = &H298B00
    RAL_6038_Leuchtgr�n = &H1AB500
    RAL_6039_Fasergr�n = &H3FC5B3
    
    'RAL-7X Grau
    RAL_7000_Fehgrau = &H8E887A
    RAL_7001_Silbergrau = &H9F998F
    RAL_7002_Olivgrau = &H637881
    RAL_7003_Moosgrau = &H69767A
    RAL_7004_Signalgrau = &H9B9B9B
    RAL_7005_Mausgrau = &H6F716B
    RAL_7006_Beigegrau = &H616F75
    RAL_7008_Khakigrau = &H3D5E74
    RAL_7009_Gr�ngrau = &H58605D
    RAL_7010_Zeltgrau = &H565C58
    RAL_7011_Eisengrau = &H615D55
    RAL_7012_Basaltgrau = &H5E5D57
    RAL_7013_Braungrau = &H445057
    RAL_7015_Schiefergrau = &H5C5651
    RAL_7016_Anthrazitgrau = &H423E38
    RAL_7021_Schwarzgrau = &H34322F
    RAL_7022_Umbragrau = &H464D4B
    RAL_7023_Betongrau = &H798481
    RAL_7024_Graphitgrau = &H504A47
    RAL_7026_Granitgrau = &H444237
    RAL_7030_Steingrau = &H889393
    RAL_7031_Blaugrau = &H6D685B
    RAL_7032_Kieselgrau = &HA1B0B5
    RAL_7033_Zementgrau = &H798981
    RAL_7034_Gelbgrau = &H6F8891
    RAL_7035_Lichtgrau = &HCCD0CB
    RAL_7036_Platingrau = &H97969A
    RAL_7037_Staubgrau = &H7A7B7A
    RAL_7038_Achatgrau = &HB0B8B4
    RAL_7039_Quarzgrau = &H5E686B
    RAL_7040_Fenstergrau = &HA6A39D
    RAL_7042_VerkehrsgrauA = &H95968F
    RAL_7043_VerkehrsgrauB = &H51544E
    RAL_7044_Seidengrau = &HB2BDBD
    RAL_7045_Telegrau1 = &H94918D
    RAL_7046_Telegrau2 = &H8E8982
    RAL_7047_Telegrau4 = &HCFD0CF
    RAL_7048_Perlmausgrau = &H758188
    
    'RAL-8X Braun
    RAL_8000_Gr�nbraun = &H3E6989
    RAL_8001_Ockerbraun = &H2D6299
    RAL_8002_Signalbraun = &H3E4D79
    RAL_8003_Lehmbraun = &H264B7E
    RAL_8004_Kupferbraun = &H354E8F
    RAL_8007_Rehbraun = &H2F4A6F
    RAL_8008_Olivbraun = &H234A6E
    RAL_8011_Nussbraun = &H293A5A
    RAL_8012_Rotbraun = &H2B3366
    RAL_8014_Sepiabraun = &H26354A
    RAL_8015_Kastanienbraun = &H343A63
    RAL_8016_Mahagonibraun = &H1F2A49
    RAL_8017_Schokoladenbraun = &H292F44
    RAL_8019_Graubraun = &H3A3A3F
    RAL_8022_Schwarzbraun = &H201F21
    RAL_8023_Orangebraun = &H2F5EA6
    RAL_8024_Beigebraun = &H385076
    RAL_8025_Blassbraun = &H495C75
    RAL_8028_Terrabraun = &H2B3B4E
    RAL_8029_Perlkupfer = &H273C77
    
    'RAL-9X Wei� und Schwarz
    RAL_9001_Cremewei� = &HD2E0E9
    RAL_9002_Grauwei� = &HCBD5D7
    RAL_9003_Signalwei� = &HF4F8F4
    RAL_9004_Signalschwarz = &H32302E
    RAL_9005_Tiefschwarz = &H100E0E
    RAL_9006_Wei�aluminium = &HA0A1A1
    RAL_9007_Graualuminium = &H838687
    RAL_9010_Reinwei� = &HEFF9F7
    RAL_9011_Graphitschwarz = &H2F2C29
    RAL_9012_Reinraumwei� = &HE6FDFF
    RAL_9016_Verkehrswei� = &HF5FBF7
    RAL_9017_Verkehrsschwarz = &H2F2D2A
    RAL_9018_Papyruswei� = &HC3CAC7
    RAL_9022_Perlhellgrau = &H9C9C9C
    RAL_9023_Perldunkelgrau = &H82817E
    RAL_9020_SeidenmattWei� = &HFDFDFD
    
    'RAL F9 Tarnfarben der Bundeswehr
    RAL_6031_Bronzegr�n = &H465748
    RAL_8027_Lederbraun = &H384950
    RAL_9021_Teerschwarz = &HE0501
    RAL_1039_Sandbeige = &H9EC1CE
    RAL_1040_Lehmbeige = &H81ACBB
    RAL_6040_Helloliv = &H587E82
    RAL_7050_Tarngrau = &H7A8882
    RAL_8031_Sandbraun = &H7B9DB4
End Enum

Public Type TNamedRALColor
    Name   As String
    RALNr  As Long
    RALCol As Long
End Type
Private m_Arr() As TNamedRALColor
Private m_Count As Long

Private Function TNamedRALColor(ByVal aName As String, ByVal aRALNr As Long, ByVal RALColor As RALClassic) As TNamedRALColor
    With TNamedRALColor
        .Name = aName
        .RALNr = aRALNr
        .RALCol = RALColor
    End With
End Function

Public Function RALClassic_ClosestColorTo(ByVal aColor As Long) As TNamedRALColor
    Dim i As Long, i_minEd As Long, edi As Double
    Dim lc As LngColor: lc = LngColor(aColor)
    Dim ed0 As Double: ed0 = LngColor_EuclidRMean(LngColor(m_Arr(0).RALCol), lc)
    For i = 1 To m_Count - 1
        edi = LngColor_EuclidRMean(LngColor(m_Arr(i).RALCol), lc)
        If edi < ed0 Then
            i_minEd = i
            ed0 = edi
        End If
    Next
    RALClassic_ClosestColorTo = m_Arr(i_minEd)
End Function

Public Sub RALClassicColor_Init()
    Dim i As Long: m_Count = 225
    ReDim m_Arr(0 To m_Count - 1)
    m_Arr(i) = TNamedRALColor("Gr�nbeige", 1000, RALClassic.RAL_1000_Gr�nbeige):             i = i + 1
    m_Arr(i) = TNamedRALColor("Beige", 1001, RALClassic.RAL_1001_Beige):                     i = i + 1
    m_Arr(i) = TNamedRALColor("Sandgelb", 1002, RALClassic.RAL_1002_Sandgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgelb", 1003, RALClassic.RAL_1003_Signalgelb):           i = i + 1
    m_Arr(i) = TNamedRALColor("Goldgelb", 1004, RALClassic.RAL_1004_Goldgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Honiggelb", 1005, RALClassic.RAL_1005_Honiggelb):             i = i + 1
    m_Arr(i) = TNamedRALColor("Maisgelb", 1006, RALClassic.RAL_1006_Maisgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Narzissengelb", 1007, RALClassic.RAL_1007_Narzissengelb):     i = i + 1
    m_Arr(i) = TNamedRALColor("Braunbeige", 1011, RALClassic.RAL_1011_Braunbeige):           i = i + 1
    m_Arr(i) = TNamedRALColor("Zitronengelb", 1012, RALClassic.RAL_1012_Zitronengelb):       i = i + 1
    m_Arr(i) = TNamedRALColor("Perlwei�", 1013, RALClassic.RAL_1013_Perlwei�):               i = i + 1
    m_Arr(i) = TNamedRALColor("Elfenbein", 1014, RALClassic.RAL_1014_Elfenbein):             i = i + 1
    m_Arr(i) = TNamedRALColor("Hellelfenbein", 1015, RALClassic.RAL_1015_Hellelfenbein):     i = i + 1
    m_Arr(i) = TNamedRALColor("Schwefelgelb", 1016, RALClassic.RAL_1016_Schwefelgelb):       i = i + 1
    m_Arr(i) = TNamedRALColor("Safrangelb", 1017, RALClassic.RAL_1017_Safrangelb):           i = i + 1
    m_Arr(i) = TNamedRALColor("Zinkgelb", 1018, RALClassic.RAL_1018_Zinkgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Graubeige", 1019, RALClassic.RAL_1019_Graubeige):             i = i + 1
    m_Arr(i) = TNamedRALColor("Olivgelb", 1020, RALClassic.RAL_1020_Olivgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Rapsgelb", 1021, RALClassic.RAL_1021_Rapsgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsgelb", 1023, RALClassic.RAL_1023_Verkehrsgelb):       i = i + 1
    m_Arr(i) = TNamedRALColor("Ockergelb", 1024, RALClassic.RAL_1024_Ockergelb):             i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchtgelb", 1026, RALClassic.RAL_1026_Leuchtgelb):           i = i + 1
    m_Arr(i) = TNamedRALColor("Currygelb", 1027, RALClassic.RAL_1027_Currygelb):             i = i + 1
    m_Arr(i) = TNamedRALColor("Melonengelb", 1028, RALClassic.RAL_1028_Melonengelb):         i = i + 1
    m_Arr(i) = TNamedRALColor("Ginstergelb", 1032, RALClassic.RAL_1032_Ginstergelb):         i = i + 1
    m_Arr(i) = TNamedRALColor("Dahliengelb", 1033, RALClassic.RAL_1033_Dahliengelb):         i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellgelb", 1034, RALClassic.RAL_1034_Pastellgelb):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perlbeige", 1035, RALClassic.RAL_1035_Perlbeige):             i = i + 1
    m_Arr(i) = TNamedRALColor("Perlgold", 1036, RALClassic.RAL_1036_Perlgold):               i = i + 1
    m_Arr(i) = TNamedRALColor("Sonnengelb", 1037, RALClassic.RAL_1037_Sonnengelb):           i = i + 1
    
    m_Arr(i) = TNamedRALColor("Gelborange", 2000, RALClassic.RAL_2000_Gelborange):             i = i + 1
    m_Arr(i) = TNamedRALColor("Rotorange", 2001, RALClassic.RAL_2001_Rotorange):               i = i + 1
    m_Arr(i) = TNamedRALColor("Blutorange", 2002, RALClassic.RAL_2002_Blutorange):             i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellorange", 2003, RALClassic.RAL_2003_Pastellorange):       i = i + 1
    m_Arr(i) = TNamedRALColor("Reinorange", 2004, RALClassic.RAL_2004_Reinorange):             i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchtorange", 2005, RALClassic.RAL_2005_Leuchtorange):         i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchthellorange", 2007, RALClassic.RAL_2007_Leuchthellorange): i = i + 1
    m_Arr(i) = TNamedRALColor("Hellrotorange", 2008, RALClassic.RAL_2008_Hellrotorange):       i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsorange", 2009, RALClassic.RAL_2009_Verkehrsorange):     i = i + 1
    m_Arr(i) = TNamedRALColor("Signalorange", 2010, RALClassic.RAL_2010_Signalorange):         i = i + 1
    m_Arr(i) = TNamedRALColor("Tieforange", 2011, RALClassic.RAL_2011_Tieforange):             i = i + 1
    m_Arr(i) = TNamedRALColor("Lachsorange", 2012, RALClassic.RAL_2012_Lachsorange):           i = i + 1
    m_Arr(i) = TNamedRALColor("Perlorange", 2013, RALClassic.RAL_2013_Perlorange):             i = i + 1
    m_Arr(i) = TNamedRALColor("RALOrange", 2017, RALClassic.RAL_2017_RALOrange):               i = i + 1
    
    m_Arr(i) = TNamedRALColor("Feuerrot", 3000, RALClassic.RAL_3000_Feuerrot):            i = i + 1
    m_Arr(i) = TNamedRALColor("Signalrot", 3001, RALClassic.RAL_3001_Signalrot):          i = i + 1
    m_Arr(i) = TNamedRALColor("Karminrot", 3002, RALClassic.RAL_3002_Karminrot):          i = i + 1
    m_Arr(i) = TNamedRALColor("Rubinrot", 3003, RALClassic.RAL_3003_Rubinrot):            i = i + 1
    m_Arr(i) = TNamedRALColor("Purpurrot", 3004, RALClassic.RAL_3004_Purpurrot):          i = i + 1
    m_Arr(i) = TNamedRALColor("Weinrot", 3005, RALClassic.RAL_3005_Weinrot):              i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzrot", 3007, RALClassic.RAL_3007_Schwarzrot):        i = i + 1
    m_Arr(i) = TNamedRALColor("Oxidrot", 3009, RALClassic.RAL_3009_Oxidrot):              i = i + 1
    m_Arr(i) = TNamedRALColor("Braunrot", 3011, RALClassic.RAL_3011_Braunrot):            i = i + 1
    m_Arr(i) = TNamedRALColor("Beigerot", 3012, RALClassic.RAL_3012_Beigerot):            i = i + 1
    m_Arr(i) = TNamedRALColor("Tomatenrot", 3013, RALClassic.RAL_3013_Tomatenrot):        i = i + 1
    m_Arr(i) = TNamedRALColor("Altrosa", 3014, RALClassic.RAL_3014_Altrosa):              i = i + 1
    m_Arr(i) = TNamedRALColor("Hellrosa", 3015, RALClassic.RAL_3015_Hellrosa):            i = i + 1
    m_Arr(i) = TNamedRALColor("Korallenrot", 3016, RALClassic.RAL_3016_Korallenrot):      i = i + 1
    m_Arr(i) = TNamedRALColor("Ros�", 3017, RALClassic.RAL_3017_Ros�):                    i = i + 1
    m_Arr(i) = TNamedRALColor("Erdbeerrot", 3018, RALClassic.RAL_3018_Erdbeerrot):        i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsrot", 3020, RALClassic.RAL_3020_Verkehrsrot):      i = i + 1
    m_Arr(i) = TNamedRALColor("Lachsrot", 3022, RALClassic.RAL_3022_Lachsrot):            i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchtrot", 3024, RALClassic.RAL_3024_Leuchtrot):          i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchthellrot", 3026, RALClassic.RAL_3026_Leuchthellrot):  i = i + 1
    m_Arr(i) = TNamedRALColor("Himbeerrot", 3027, RALClassic.RAL_3027_Himbeerrot):        i = i + 1
    m_Arr(i) = TNamedRALColor("Reinrot", 3028, RALClassic.RAL_3028_Reinrot):              i = i + 1
    m_Arr(i) = TNamedRALColor("Orientrot", 3031, RALClassic.RAL_3031_Orientrot):          i = i + 1
    m_Arr(i) = TNamedRALColor("Perlrubinrot", 3032, RALClassic.RAL_3032_Perlrubinrot):    i = i + 1
    m_Arr(i) = TNamedRALColor("Perlrosa", 3033, RALClassic.RAL_3033_Perlrosa):            i = i + 1
    
    m_Arr(i) = TNamedRALColor("Rotlila", 4001, RALClassic.RAL_4001_Rotlila):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Rotviolett", 4002, RALClassic.RAL_4002_Rotviolett):           i = i + 1
    m_Arr(i) = TNamedRALColor("Erikaviolett", 4003, RALClassic.RAL_4003_Erikaviolett):       i = i + 1
    m_Arr(i) = TNamedRALColor("Bordeauxviolett", 4004, RALClassic.RAL_4004_Bordeauxviolett): i = i + 1
    m_Arr(i) = TNamedRALColor("Blaulila", 4005, RALClassic.RAL_4005_Blaulila):               i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrspurpur", 4006, RALClassic.RAL_4006_Verkehrspurpur):   i = i + 1
    m_Arr(i) = TNamedRALColor("Purpurviolett", 4007, RALClassic.RAL_4007_Purpurviolett):     i = i + 1
    m_Arr(i) = TNamedRALColor("Signalviolett", 4008, RALClassic.RAL_4008_Signalviolett):     i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellviolett", 4009, RALClassic.RAL_4009_Pastellviolett):   i = i + 1
    m_Arr(i) = TNamedRALColor("Telemagenta", 4010, RALClassic.RAL_4010_Telemagenta):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perlviolett", 4011, RALClassic.RAL_4011_Perlviolett):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perlbrombeer", 4012, RALClassic.RAL_4012_Perlbrombeer):       i = i + 1
    
    m_Arr(i) = TNamedRALColor("Violettblau", 5000, RALClassic.RAL_5000_Violettblau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Gr�nblau", 5001, RALClassic.RAL_5001_Gr�nblau):               i = i + 1
    m_Arr(i) = TNamedRALColor("Ultramarinblau", 5002, RALClassic.RAL_5002_Ultramarinblau):   i = i + 1
    m_Arr(i) = TNamedRALColor("Saphirblau", 5003, RALClassic.RAL_5003_Saphirblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzblau", 5004, RALClassic.RAL_5004_Schwarzblau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Signalblau", 5005, RALClassic.RAL_5005_Signalblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Brillantblau", 5007, RALClassic.RAL_5007_Brillantblau):       i = i + 1
    m_Arr(i) = TNamedRALColor("Graublau", 5008, RALClassic.RAL_5008_Graublau):               i = i + 1
    m_Arr(i) = TNamedRALColor("Azurblau", 5009, RALClassic.RAL_5009_Azurblau):               i = i + 1
    m_Arr(i) = TNamedRALColor("Enzianblau", 5010, RALClassic.RAL_5010_Enzianblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Stahlblau", 5011, RALClassic.RAL_5011_Stahlblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Lichtblau", 5012, RALClassic.RAL_5012_Lichtblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Kobaltblau", 5013, RALClassic.RAL_5013_Kobaltblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Taubenblau", 5014, RALClassic.RAL_5014_Taubenblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Himmelblau", 5015, RALClassic.RAL_5015_Himmelblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsblau", 5017, RALClassic.RAL_5017_Verkehrsblau):       i = i + 1
    m_Arr(i) = TNamedRALColor("T�rkisblau", 5018, RALClassic.RAL_5018_T�rkisblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Capriblau", 5019, RALClassic.RAL_5019_Capriblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Ozeanblau", 5020, RALClassic.RAL_5020_Ozeanblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Wasserblau", 5021, RALClassic.RAL_5021_Wasserblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Nachtblau", 5022, RALClassic.RAL_5022_Nachtblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Fernblau", 5023, RALClassic.RAL_5023_Fernblau):               i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellblau", 5024, RALClassic.RAL_5024_Pastellblau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perlenzian", 5025, RALClassic.RAL_5025_Perlenzian):           i = i + 1
    m_Arr(i) = TNamedRALColor("Perlnachtblau", 5026, RALClassic.RAL_5026_Perlnachtblau):     i = i + 1
    
    m_Arr(i) = TNamedRALColor("Patinagr�n", 6000, RALClassic.RAL_6000_Patinagr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Smaragdgr�n", 6001, RALClassic.RAL_6001_Smaragdgr�n):      i = i + 1
    m_Arr(i) = TNamedRALColor("Laubgr�n", 6002, RALClassic.RAL_6002_Laubgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Olivgr�n", 6003, RALClassic.RAL_6003_Olivgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Blaugr�n", 6004, RALClassic.RAL_6004_Blaugr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Moosgr�n", 6005, RALClassic.RAL_6005_Moosgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Grauoliv", 6006, RALClassic.RAL_6006_Grauoliv):            i = i + 1
    m_Arr(i) = TNamedRALColor("Flaschengr�n", 6007, RALClassic.RAL_6007_Flaschengr�n):    i = i + 1
    m_Arr(i) = TNamedRALColor("Braungr�n", 6008, RALClassic.RAL_6008_Braungr�n):          i = i + 1
    m_Arr(i) = TNamedRALColor("Tannengr�n", 6009, RALClassic.RAL_6009_Tannengr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Grasgr�n", 6010, RALClassic.RAL_6010_Grasgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Resedagr�n", 6011, RALClassic.RAL_6011_Resedagr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzgr�n", 6012, RALClassic.RAL_6012_Schwarzgr�n):      i = i + 1
    m_Arr(i) = TNamedRALColor("Schilfgr�n", 6013, RALClassic.RAL_6013_Schilfgr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Gelboliv", 6014, RALClassic.RAL_6014_Gelboliv):            i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzoliv", 6015, RALClassic.RAL_6015_Schwarzoliv):      i = i + 1
    m_Arr(i) = TNamedRALColor("T�rkisgr�n", 6016, RALClassic.RAL_6016_T�rkisgr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Maigr�n", 6017, RALClassic.RAL_6017_Maigr�n):              i = i + 1
    m_Arr(i) = TNamedRALColor("Gelbgr�n", 6018, RALClassic.RAL_6018_Gelbgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Wei�gr�n", 6019, RALClassic.RAL_6019_Wei�gr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Chromoxidgr�n", 6020, RALClassic.RAL_6020_Chromoxidgr�n):  i = i + 1
    m_Arr(i) = TNamedRALColor("Blassgr�n", 6021, RALClassic.RAL_6021_Blassgr�n):          i = i + 1
    m_Arr(i) = TNamedRALColor("Braunoliv", 6022, RALClassic.RAL_6022_Braunoliv):          i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsgr�n", 6024, RALClassic.RAL_6024_Verkehrsgr�n):    i = i + 1
    m_Arr(i) = TNamedRALColor("Farngr�n", 6025, RALClassic.RAL_6025_Farngr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Opalgr�n", 6026, RALClassic.RAL_6026_Opalgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Lichtgr�n", 6027, RALClassic.RAL_6027_Lichtgr�n):          i = i + 1
    m_Arr(i) = TNamedRALColor("Kieferngr�n", 6028, RALClassic.RAL_6028_Kieferngr�n):      i = i + 1
    m_Arr(i) = TNamedRALColor("Minzgr�n", 6029, RALClassic.RAL_6029_Minzgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgr�n", 6032, RALClassic.RAL_6032_Signalgr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Mintt�rkis", 6033, RALClassic.RAL_6033_Mintt�rkis):        i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellt�rkis", 6034, RALClassic.RAL_6034_Pastellt�rkis):  i = i + 1
    m_Arr(i) = TNamedRALColor("Perlgr�n", 6035, RALClassic.RAL_6035_Perlgr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Perlopalgr�n", 6036, RALClassic.RAL_6036_Perlopalgr�n):    i = i + 1
    m_Arr(i) = TNamedRALColor("Reingr�n", 6037, RALClassic.RAL_6037_Reingr�n):            i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchtgr�n", 6038, RALClassic.RAL_6038_Leuchtgr�n):        i = i + 1
    m_Arr(i) = TNamedRALColor("Fasergr�n", 6039, RALClassic.RAL_6039_Fasergr�n):          i = i + 1
    
    m_Arr(i) = TNamedRALColor("Fehgrau", 7000, RALClassic.RAL_7000_Fehgrau):              i = i + 1
    m_Arr(i) = TNamedRALColor("Silbergrau", 7001, RALClassic.RAL_7001_Silbergrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Olivgrau", 7002, RALClassic.RAL_7002_Olivgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Moosgrau", 7003, RALClassic.RAL_7003_Moosgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgrau", 7004, RALClassic.RAL_7004_Signalgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Mausgrau", 7005, RALClassic.RAL_7005_Mausgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Beigegrau", 7006, RALClassic.RAL_7006_Beigegrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Khakigrau", 7008, RALClassic.RAL_7008_Khakigrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Gr�ngrau", 7009, RALClassic.RAL_7009_Gr�ngrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Zeltgrau", 7010, RALClassic.RAL_7010_Zeltgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Eisengrau", 7011, RALClassic.RAL_7011_Eisengrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Basaltgrau", 7012, RALClassic.RAL_7012_Basaltgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Braungrau", 7013, RALClassic.RAL_7013_Braungrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Schiefergrau", 7015, RALClassic.RAL_7015_Schiefergrau):    i = i + 1
    m_Arr(i) = TNamedRALColor("Anthrazitgrau", 7016, RALClassic.RAL_7016_Anthrazitgrau):  i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzgrau", 7021, RALClassic.RAL_7021_Schwarzgrau):      i = i + 1
    m_Arr(i) = TNamedRALColor("Umbragrau", 7022, RALClassic.RAL_7022_Umbragrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Betongrau", 7023, RALClassic.RAL_7023_Betongrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Graphitgrau", 7024, RALClassic.RAL_7024_Graphitgrau):      i = i + 1
    m_Arr(i) = TNamedRALColor("Granitgrau", 7026, RALClassic.RAL_7026_Granitgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Steingrau", 7030, RALClassic.RAL_7030_Steingrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Blaugrau", 7031, RALClassic.RAL_7031_Blaugrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Kieselgrau", 7032, RALClassic.RAL_7032_Kieselgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Zementgrau", 7033, RALClassic.RAL_7033_Zementgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Gelbgrau", 7034, RALClassic.RAL_7034_Gelbgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Lichtgrau", 7035, RALClassic.RAL_7035_Lichtgrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Platingrau", 7036, RALClassic.RAL_7036_Platingrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Staubgrau", 7037, RALClassic.RAL_7037_Staubgrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Achatgrau", 7038, RALClassic.RAL_7038_Achatgrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Quarzgrau", 7039, RALClassic.RAL_7039_Quarzgrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Fenstergrau", 7040, RALClassic.RAL_7040_Fenstergrau):      i = i + 1
    m_Arr(i) = TNamedRALColor("VerkehrsgrauA", 7042, RALClassic.RAL_7042_VerkehrsgrauA):  i = i + 1
    m_Arr(i) = TNamedRALColor("VerkehrsgrauB", 7043, RALClassic.RAL_7043_VerkehrsgrauB):  i = i + 1
    m_Arr(i) = TNamedRALColor("Seidengrau", 7044, RALClassic.RAL_7044_Seidengrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Telegrau1", 7045, RALClassic.RAL_7045_Telegrau1):          i = i + 1
    m_Arr(i) = TNamedRALColor("Telegrau2", 7046, RALClassic.RAL_7046_Telegrau2):          i = i + 1
    m_Arr(i) = TNamedRALColor("Telegrau4", 7047, RALClassic.RAL_7047_Telegrau4):          i = i + 1
    m_Arr(i) = TNamedRALColor("Perlmausgrau", 7048, RALClassic.RAL_7048_Perlmausgrau):    i = i + 1
    
    m_Arr(i) = TNamedRALColor("Gr�nbraun", 8000, RALClassic.RAL_8000_Gr�nbraun):               i = i + 1
    m_Arr(i) = TNamedRALColor("Ockerbraun", 8001, RALClassic.RAL_8001_Ockerbraun):             i = i + 1
    m_Arr(i) = TNamedRALColor("Signalbraun", 8002, RALClassic.RAL_8002_Signalbraun):           i = i + 1
    m_Arr(i) = TNamedRALColor("Lehmbraun", 8003, RALClassic.RAL_8003_Lehmbraun):               i = i + 1
    m_Arr(i) = TNamedRALColor("Kupferbraun", 8004, RALClassic.RAL_8004_Kupferbraun):           i = i + 1
    m_Arr(i) = TNamedRALColor("Rehbraun", 8007, RALClassic.RAL_8007_Rehbraun):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Olivbraun", 8008, RALClassic.RAL_8008_Olivbraun):               i = i + 1
    m_Arr(i) = TNamedRALColor("Nussbraun", 8011, RALClassic.RAL_8011_Nussbraun):               i = i + 1
    m_Arr(i) = TNamedRALColor("Rotbraun", 8012, RALClassic.RAL_8012_Rotbraun):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Sepiabraun", 8014, RALClassic.RAL_8014_Sepiabraun):             i = i + 1
    m_Arr(i) = TNamedRALColor("Kastanienbraun", 8015, RALClassic.RAL_8015_Kastanienbraun):     i = i + 1
    m_Arr(i) = TNamedRALColor("Mahagonibraun", 8016, RALClassic.RAL_8016_Mahagonibraun):       i = i + 1
    m_Arr(i) = TNamedRALColor("Schokoladenbraun", 8017, RALClassic.RAL_8017_Schokoladenbraun): i = i + 1
    m_Arr(i) = TNamedRALColor("Graubraun", 8019, RALClassic.RAL_8019_Graubraun):               i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzbraun", 8022, RALClassic.RAL_8022_Schwarzbraun):         i = i + 1
    m_Arr(i) = TNamedRALColor("Orangebraun", 8023, RALClassic.RAL_8023_Orangebraun):           i = i + 1
    m_Arr(i) = TNamedRALColor("Beigebraun", 8024, RALClassic.RAL_8024_Beigebraun):             i = i + 1
    m_Arr(i) = TNamedRALColor("Blassbraun", 8025, RALClassic.RAL_8025_Blassbraun):             i = i + 1
    m_Arr(i) = TNamedRALColor("Terrabraun", 8028, RALClassic.RAL_8028_Terrabraun):             i = i + 1
    m_Arr(i) = TNamedRALColor("Perlkupfer", 8029, RALClassic.RAL_8029_Perlkupfer):             i = i + 1
    
    m_Arr(i) = TNamedRALColor("Cremewei�", 9001, RALClassic.RAL_9001_Cremewei�):               i = i + 1
    m_Arr(i) = TNamedRALColor("Grauwei�", 9002, RALClassic.RAL_9002_Grauwei�):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Signalwei�", 9003, RALClassic.RAL_9003_Signalwei�):             i = i + 1
    m_Arr(i) = TNamedRALColor("Signalschwarz", 9004, RALClassic.RAL_9004_Signalschwarz):       i = i + 1
    m_Arr(i) = TNamedRALColor("Tiefschwarz", 9005, RALClassic.RAL_9005_Tiefschwarz):           i = i + 1
    m_Arr(i) = TNamedRALColor("Wei�aluminium", 9006, RALClassic.RAL_9006_Wei�aluminium):       i = i + 1
    m_Arr(i) = TNamedRALColor("Graualuminium", 9007, RALClassic.RAL_9007_Graualuminium):       i = i + 1
    m_Arr(i) = TNamedRALColor("Reinwei�", 9010, RALClassic.RAL_9010_Reinwei�):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Graphitschwarz", 9011, RALClassic.RAL_9011_Graphitschwarz):     i = i + 1
    m_Arr(i) = TNamedRALColor("Reinraumwei�", 9012, RALClassic.RAL_9012_Reinraumwei�):         i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrswei�", 9016, RALClassic.RAL_9016_Verkehrswei�):         i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsschwarz", 9017, RALClassic.RAL_9017_Verkehrsschwarz):   i = i + 1
    m_Arr(i) = TNamedRALColor("Papyruswei�", 9018, RALClassic.RAL_9018_Papyruswei�):           i = i + 1
    m_Arr(i) = TNamedRALColor("Perlhellgrau", 9022, RALClassic.RAL_9022_Perlhellgrau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perldunkelgrau", 9023, RALClassic.RAL_9023_Perldunkelgrau):     i = i + 1
    m_Arr(i) = TNamedRALColor("SeidenmattWei�", 9020, RALClassic.RAL_9020_SeidenmattWei�):     i = i + 1
    
    m_Arr(i) = TNamedRALColor("Bronzegr�n", 6031, RALClassic.RAL_6031_Bronzegr�n):    i = i + 1
    m_Arr(i) = TNamedRALColor("Lederbraun", 8027, RALClassic.RAL_8027_Lederbraun):    i = i + 1
    m_Arr(i) = TNamedRALColor("Teerschwarz", 9021, RALClassic.RAL_9021_Teerschwarz):  i = i + 1
    m_Arr(i) = TNamedRALColor("Sandbeige", 1039, RALClassic.RAL_1039_Sandbeige):      i = i + 1
    m_Arr(i) = TNamedRALColor("Lehmbeige", 1040, RALClassic.RAL_1040_Lehmbeige):      i = i + 1
    m_Arr(i) = TNamedRALColor("Helloliv", 6040, RALClassic.RAL_6040_Helloliv):        i = i + 1
    m_Arr(i) = TNamedRALColor("Tarngrau", 7050, RALClassic.RAL_7050_Tarngrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Sandbraun", 8031, RALClassic.RAL_8031_Sandbraun):      i = i + 1
    'Debug.Print i
End Sub


Public Function RALClassic_NameToStr(e As RALClassic) As String
    If m_Count = 0 Then RALClassicColor_Init
    Dim s As String
    Select Case e
    
    Case RAL_1000_Gr�nbeige:        s = "Gr�nbeige"
    Case RAL_1001_Beige:            s = "Beige"
    Case RAL_1002_Sandgelb:         s = "Sandgelb"
    Case RAL_1003_Signalgelb:       s = "Signalgelb"
    Case RAL_1004_Goldgelb:         s = "Goldgelb"
    Case RAL_1005_Honiggelb:        s = "Honiggelb"
    Case RAL_1006_Maisgelb:         s = "Maisgelb"
    Case RAL_1007_Narzissengelb:    s = "Narzissengelb"
    Case RAL_1011_Braunbeige:       s = "Braunbeige"
    Case RAL_1012_Zitronengelb:     s = "Zitronengelb"
    Case RAL_1013_Perlwei�:         s = "Perlwei�"
    Case RAL_1014_Elfenbein:        s = "Elfenbein"
    Case RAL_1015_Hellelfenbein:    s = "Hellelfenbein"
    Case RAL_1016_Schwefelgelb:     s = "Schwefelgelb"
    Case RAL_1017_Safrangelb:       s = "Safrangelb"
    Case RAL_1018_Zinkgelb:         s = "Zinkgelb"
    Case RAL_1019_Graubeige:        s = "Graubeige"
    Case RAL_1020_Olivgelb:         s = "Olivgelb"
    Case RAL_1021_Rapsgelb:         s = "Rapsgelb"
    Case RAL_1023_Verkehrsgelb:     s = "Verkehrsgelb"
    Case RAL_1024_Ockergelb:        s = "Ockergelb"
    Case RAL_1026_Leuchtgelb:       s = "Leuchtgelb"
    Case RAL_1027_Currygelb:        s = "Currygelb"
    Case RAL_1028_Melonengelb:      s = "Melonengelb"
    Case RAL_1032_Ginstergelb:      s = "Ginstergelb"
    Case RAL_1033_Dahliengelb:      s = "Dahliengelb"
    Case RAL_1034_Pastellgelb:      s = "Pastellgelb"
    Case RAL_1035_Perlbeige:        s = "Perlbeige"
    Case RAL_1036_Perlgold:         s = "Perlgold"
    Case RAL_1037_Sonnengelb:       s = "Sonnengelb"
    
    Case RAL_2000_Gelborange:       s = "Gelborange"
    Case RAL_2001_Rotorange:        s = "Rotorange"
    Case RAL_2002_Blutorange:       s = "Blutorange"
    Case RAL_2003_Pastellorange:    s = "Pastellorange"
    Case RAL_2004_Reinorange:       s = "Reinorange"
    Case RAL_2005_Leuchtorange:     s = "Leuchtorange"
    Case RAL_2007_Leuchthellorange: s = "Leuchthellorange"
    Case RAL_2008_Hellrotorange:    s = "Hellrotorange"
    Case RAL_2009_Verkehrsorange:   s = "Verkehrsorange"
    Case RAL_2010_Signalorange:     s = "Signalorange"
    Case RAL_2011_Tieforange:       s = "Tieforange"
    Case RAL_2012_Lachsorange:      s = "Lachsorange"
    Case RAL_2013_Perlorange:       s = "Perlorange"
    Case RAL_2017_RALOrange:        s = "RALOrange"
        
    Case RAL_3000_Feuerrot:         s = "Feuerrot"
    Case RAL_3001_Signalrot:        s = "Signalrot"
    Case RAL_3002_Karminrot:        s = "Karminrot"
    Case RAL_3003_Rubinrot:         s = "Rubinrot"
    Case RAL_3004_Purpurrot:        s = "Purpurrot"
    Case RAL_3005_Weinrot:          s = "Weinrot"
    Case RAL_3007_Schwarzrot:       s = "Schwarzrot"
    Case RAL_3009_Oxidrot:          s = "Oxidrot"
    Case RAL_3011_Braunrot:         s = "Braunrot"
    Case RAL_3012_Beigerot:         s = "Beigerot"
    Case RAL_3013_Tomatenrot:       s = "Tomatenrot"
    Case RAL_3014_Altrosa:          s = "Altrosa"
    Case RAL_3015_Hellrosa:         s = "Hellrosa"
    Case RAL_3016_Korallenrot:      s = "Korallenrot"
    Case RAL_3017_Ros�:             s = "Ros�"
    Case RAL_3018_Erdbeerrot:       s = "Erdbeerrot"
    Case RAL_3020_Verkehrsrot:      s = "Verkehrsrot"
    Case RAL_3022_Lachsrot:         s = "Lachsrot"
    Case RAL_3024_Leuchtrot:        s = "Leuchtrot"
    Case RAL_3026_Leuchthellrot:    s = "Leuchthellrot"
    Case RAL_3027_Himbeerrot:       s = "Himbeerrot"
    Case RAL_3028_Reinrot:          s = "Reinrot"
    Case RAL_3031_Orientrot:        s = "Orientrot"
    Case RAL_3032_Perlrubinrot:     s = "Perlrubinrot"
    Case RAL_3033_Perlrosa:         s = "Perlrosa"
    
    Case RAL_4001_Rotlila:          s = "Rotlila"
    Case RAL_4002_Rotviolett:       s = "Rotviolett"
    Case RAL_4003_Erikaviolett:     s = "Erikaviolett"
    Case RAL_4004_Bordeauxviolett:  s = "Bordeauxviolett"
    Case RAL_4005_Blaulila:         s = "Blaulila"
    Case RAL_4006_Verkehrspurpur:   s = "Verkehrspurpur"
    Case RAL_4007_Purpurviolett:    s = "Purpurviolett"
    Case RAL_4008_Signalviolett:    s = "Signalviolett"
    Case RAL_4009_Pastellviolett:   s = "Pastellviolett"
    Case RAL_4010_Telemagenta:      s = "Telemagenta"
    Case RAL_4011_Perlviolett:      s = "Perlviolett"
    Case RAL_4012_Perlbrombeer:     s = "Perlbrombeer"
    
    Case RAL_5000_Violettblau:      s = "Violettblau"
    Case RAL_5001_Gr�nblau:         s = "Gr�nblau"
    Case RAL_5002_Ultramarinblau:   s = "Ultramarinblau"
    Case RAL_5003_Saphirblau:       s = "Saphirblau"
    Case RAL_5004_Schwarzblau:      s = "Schwarzblau"
    Case RAL_5005_Signalblau:       s = "Signalblau"
    Case RAL_5007_Brillantblau:     s = "Brillantblau"
    Case RAL_5008_Graublau:         s = "Graublau"
    Case RAL_5009_Azurblau:         s = "Azurblau"
    Case RAL_5010_Enzianblau:       s = "Enzianblau"
    Case RAL_5011_Stahlblau:        s = "Stahlblau"
    Case RAL_5012_Lichtblau:        s = "Lichtblau"
    Case RAL_5013_Kobaltblau:       s = "Kobaltblau"
    Case RAL_5014_Taubenblau:       s = "Taubenblau"
    Case RAL_5015_Himmelblau:       s = "Himmelblau"
    Case RAL_5017_Verkehrsblau:     s = "Verkehrsblau"
    Case RAL_5018_T�rkisblau:       s = "T�rkisblau"
    Case RAL_5019_Capriblau:        s = "Capriblau"
    Case RAL_5020_Ozeanblau:        s = "Ozeanblau"
    Case RAL_5021_Wasserblau:       s = "Wasserblau"
    Case RAL_5022_Nachtblau:        s = "Nachtblau"
    Case RAL_5023_Fernblau:         s = "Fernblau"
    Case RAL_5024_Pastellblau:      s = "Pastellblau"
    Case RAL_5025_Perlenzian:       s = "Perlenzian"
    Case RAL_5026_Perlnachtblau:    s = "Perlnachtblau"
    
    Case RAL_6000_Patinagr�n:       s = "Patinagr�n"
    Case RAL_6001_Smaragdgr�n:      s = "Smaragdgr�n"
    Case RAL_6002_Laubgr�n:         s = "Laubgr�n"
    Case RAL_6003_Olivgr�n:         s = "Olivgr�n"
    Case RAL_6004_Blaugr�n:         s = "Blaugr�n"
    Case RAL_6005_Moosgr�n:         s = "Moosgr�n"
    Case RAL_6006_Grauoliv:         s = "Grauoliv"
    Case RAL_6007_Flaschengr�n:     s = "Flaschengr�n"
    Case RAL_6008_Braungr�n:        s = "Braungr�n"
    Case RAL_6009_Tannengr�n:       s = "Tannengr�n"
    Case RAL_6010_Grasgr�n:         s = "Grasgr�n"
    Case RAL_6011_Resedagr�n:       s = "Resedagr�n"
    Case RAL_6012_Schwarzgr�n:      s = "Schwarzgr�n"
    Case RAL_6013_Schilfgr�n:       s = "Schilfgr�n"
    Case RAL_6014_Gelboliv:         s = "Gelboliv"
    Case RAL_6015_Schwarzoliv:      s = "Schwarzoliv"
    Case RAL_6016_T�rkisgr�n:       s = "T�rkisgr�n"
    Case RAL_6017_Maigr�n:          s = "Maigr�n"
    Case RAL_6018_Gelbgr�n:         s = "Gelbgr�n"
    Case RAL_6019_Wei�gr�n:         s = "Wei�gr�n"
    Case RAL_6020_Chromoxidgr�n:    s = "Chromoxidgr�n"
    Case RAL_6021_Blassgr�n:        s = "Blassgr�n"
    Case RAL_6022_Braunoliv:        s = "Braunoliv"
    Case RAL_6024_Verkehrsgr�n:     s = "Verkehrsgr�n"
    Case RAL_6025_Farngr�n:         s = "Farngr�n"
    Case RAL_6026_Opalgr�n:         s = "Opalgr�n"
    Case RAL_6027_Lichtgr�n:        s = "Lichtgr�n"
    Case RAL_6028_Kieferngr�n:      s = "Kieferngr�n"
    Case RAL_6029_Minzgr�n:         s = "Minzgr�n"
    Case RAL_6032_Signalgr�n:       s = "Signalgr�n"
    Case RAL_6033_Mintt�rkis:       s = "Mintt�rkis"
    Case RAL_6034_Pastellt�rkis:    s = "Pastellt�rkis"
    Case RAL_6035_Perlgr�n:         s = "Perlgr�n"
    Case RAL_6036_Perlopalgr�n:     s = "Perlopalgr�n"
    Case RAL_6037_Reingr�n:         s = "Reingr�n"
    Case RAL_6038_Leuchtgr�n:       s = "Leuchtgr�n"
    Case RAL_6039_Fasergr�n:        s = "Fasergr�n"
    
    Case RAL_7000_Fehgrau:          s = "Fehgrau"
    Case RAL_7001_Silbergrau:       s = "Silbergrau"
    Case RAL_7002_Olivgrau:         s = "Olivgrau"
    Case RAL_7003_Moosgrau:         s = "Moosgrau"
    Case RAL_7004_Signalgrau:       s = "Signalgrau"
    Case RAL_7005_Mausgrau:         s = "Mausgrau"
    Case RAL_7006_Beigegrau:        s = "Beigegrau"
    Case RAL_7008_Khakigrau:        s = "Khakigrau"
    Case RAL_7009_Gr�ngrau:         s = "Gr�ngrau"
    Case RAL_7010_Zeltgrau:         s = "Zeltgrau"
    Case RAL_7011_Eisengrau:        s = "Eisengrau"
    Case RAL_7012_Basaltgrau:       s = "Basaltgrau"
    Case RAL_7013_Braungrau:        s = "Braungrau"
    Case RAL_7015_Schiefergrau:     s = "Schiefergrau"
    Case RAL_7016_Anthrazitgrau:    s = "Anthrazitgrau"
    Case RAL_7021_Schwarzgrau:      s = "Schwarzgrau"
    Case RAL_7022_Umbragrau:        s = "Umbragrau"
    Case RAL_7023_Betongrau:        s = "Betongrau"
    Case RAL_7024_Graphitgrau:      s = "Graphitgrau"
    Case RAL_7026_Granitgrau:       s = "Granitgrau"
    Case RAL_7030_Steingrau:        s = "Steingrau"
    Case RAL_7031_Blaugrau:         s = "Blaugrau"
    Case RAL_7032_Kieselgrau:       s = "Kieselgrau"
    Case RAL_7033_Zementgrau:       s = "Zementgrau"
    Case RAL_7034_Gelbgrau:         s = "Gelbgrau"
    Case RAL_7035_Lichtgrau:        s = "Lichtgrau"
    Case RAL_7036_Platingrau:       s = "Platingrau"
    Case RAL_7037_Staubgrau:        s = "Staubgrau"
    Case RAL_7038_Achatgrau:        s = "Achatgrau"
    Case RAL_7039_Quarzgrau:        s = "Quarzgrau"
    Case RAL_7040_Fenstergrau:      s = "Fenstergrau"
    Case RAL_7042_VerkehrsgrauA:    s = "VerkehrsgrauA"
    Case RAL_7043_VerkehrsgrauB:    s = "VerkehrsgrauB"
    Case RAL_7044_Seidengrau:       s = "Seidengrau"
    Case RAL_7045_Telegrau1:        s = "Telegrau1"
    Case RAL_7046_Telegrau2:        s = "Telegrau2"
    Case RAL_7047_Telegrau4:        s = "Telegrau4"
    Case RAL_7048_Perlmausgrau:     s = "Perlmausgrau"
    
    Case RAL_8000_Gr�nbraun:        s = "Gr�nbraun"
    Case RAL_8001_Ockerbraun:       s = "Ockerbraun"
    Case RAL_8002_Signalbraun:      s = "Signalbraun"
    Case RAL_8003_Lehmbraun:        s = "Lehmbraun"
    Case RAL_8004_Kupferbraun:      s = "Kupferbraun"
    Case RAL_8007_Rehbraun:         s = "Rehbraun"
    Case RAL_8008_Olivbraun:        s = "Olivbraun"
    Case RAL_8011_Nussbraun:        s = "Nussbraun"
    Case RAL_8012_Rotbraun:         s = "Rotbraun"
    Case RAL_8014_Sepiabraun:       s = "Sepiabraun"
    Case RAL_8015_Kastanienbraun:   s = "Kastanienbraun"
    Case RAL_8016_Mahagonibraun:    s = "Mahagonibraun"
    Case RAL_8017_Schokoladenbraun: s = "Schokoladenbraun"
    Case RAL_8019_Graubraun:        s = "Graubraun"
    Case RAL_8022_Schwarzbraun:     s = "Schwarzbraun"
    Case RAL_8023_Orangebraun:      s = "Orangebraun"
    Case RAL_8024_Beigebraun:       s = "Beigebraun"
    Case RAL_8025_Blassbraun:       s = "Blassbraun"
    Case RAL_8028_Terrabraun:       s = "Terrabraun"
    Case RAL_8029_Perlkupfer:       s = "Perlkupfer"
    
    Case RAL_9001_Cremewei�:        s = "Cremewei�"
    Case RAL_9002_Grauwei�:         s = "Grauwei�"
    Case RAL_9003_Signalwei�:       s = "Signalwei�"
    Case RAL_9004_Signalschwarz:    s = "Signalschwarz"
    Case RAL_9005_Tiefschwarz:      s = "Tiefschwarz"
    Case RAL_9006_Wei�aluminium:    s = "Wei�aluminium"
    Case RAL_9007_Graualuminium:    s = "Graualuminium"
    Case RAL_9010_Reinwei�:         s = "Reinwei�"
    Case RAL_9011_Graphitschwarz:   s = "Graphitschwarz"
    Case RAL_9012_Reinraumwei�:     s = "Reinraumwei�"
    Case RAL_9016_Verkehrswei�:     s = "Verkehrswei�"
    Case RAL_9017_Verkehrsschwarz:  s = "Verkehrsschwarz"
    Case RAL_9018_Papyruswei�:      s = "Papyruswei�"
    Case RAL_9022_Perlhellgrau:     s = "Perlhellgrau"
    Case RAL_9023_Perldunkelgrau:   s = "Perldunkelgrau"
    Case RAL_9020_SeidenmattWei�:   s = "SeidenmattWei�"
    
    Case RAL_6031_Bronzegr�n:       s = "Bronzegr�n"
    Case RAL_8027_Lederbraun:       s = "Lederbraun"
    Case RAL_9021_Teerschwarz:      s = "Teerschwarz"
    Case RAL_1039_Sandbeige:        s = "Sandbeige"
    Case RAL_1040_Lehmbeige:        s = "Lehmbeige"
    Case RAL_6040_Helloliv:         s = "Helloliv"
    Case RAL_7050_Tarngrau:         s = "Tarngrau"
    Case RAL_8031_Sandbraun:        s = "Sandbraun"
    
    End Select
    RALClassic_NameToStr = s
End Function

Public Function RALClassic_ToNum(e As RALClassic) As Long
    If m_Count = 0 Then RALClassicColor_Init
    Dim n As Long
    Select Case e
    
    Case RAL_1000_Gr�nbeige:        n = 1000
    Case RAL_1001_Beige:            n = 1001
    Case RAL_1002_Sandgelb:         n = 1002
    Case RAL_1003_Signalgelb:       n = 1003
    Case RAL_1004_Goldgelb:         n = 1004
    Case RAL_1005_Honiggelb:        n = 1005
    Case RAL_1006_Maisgelb:         n = 1006
    Case RAL_1007_Narzissengelb:    n = 1007
    Case RAL_1011_Braunbeige:       n = 1011
    Case RAL_1012_Zitronengelb:     n = 1012
    Case RAL_1013_Perlwei�:         n = 1013
    Case RAL_1014_Elfenbein:        n = 1014
    Case RAL_1015_Hellelfenbein:    n = 1015
    Case RAL_1016_Schwefelgelb:     n = 1016
    Case RAL_1017_Safrangelb:       n = 1017
    Case RAL_1018_Zinkgelb:         n = 1018
    Case RAL_1019_Graubeige:        n = 1019
    Case RAL_1020_Olivgelb:         n = 1020
    Case RAL_1021_Rapsgelb:         n = 1021
    Case RAL_1023_Verkehrsgelb:     n = 1023
    Case RAL_1024_Ockergelb:        n = 1024
    Case RAL_1026_Leuchtgelb:       n = 1026
    Case RAL_1027_Currygelb:        n = 1027
    Case RAL_1028_Melonengelb:      n = 1028
    Case RAL_1032_Ginstergelb:      n = 1032
    Case RAL_1033_Dahliengelb:      n = 1033
    Case RAL_1034_Pastellgelb:      n = 1034
    Case RAL_1035_Perlbeige:        n = 1035
    Case RAL_1036_Perlgold:         n = 1036
    Case RAL_1037_Sonnengelb:       n = 1037
    
    Case RAL_2000_Gelborange:       n = 2000
    Case RAL_2001_Rotorange:        n = 2001
    Case RAL_2002_Blutorange:       n = 2002
    Case RAL_2003_Pastellorange:    n = 2003
    Case RAL_2004_Reinorange:       n = 2004
    Case RAL_2005_Leuchtorange:     n = 2005
    Case RAL_2007_Leuchthellorange: n = 2007
    Case RAL_2008_Hellrotorange:    n = 2008
    Case RAL_2009_Verkehrsorange:   n = 2009
    Case RAL_2010_Signalorange:     n = 2010
    Case RAL_2011_Tieforange:       n = 2011
    Case RAL_2012_Lachsorange:      n = 2012
    Case RAL_2013_Perlorange:       n = 2013
    Case RAL_2017_RALOrange:        n = 2017
    
    Case RAL_3000_Feuerrot:         n = 3000
    Case RAL_3001_Signalrot:        n = 3001
    Case RAL_3002_Karminrot:        n = 3002
    Case RAL_3003_Rubinrot:         n = 3003
    Case RAL_3004_Purpurrot:        n = 3004
    Case RAL_3005_Weinrot:          n = 3005
    Case RAL_3007_Schwarzrot:       n = 3007
    Case RAL_3009_Oxidrot:          n = 3009
    Case RAL_3011_Braunrot:         n = 3011
    Case RAL_3012_Beigerot:         n = 3012
    Case RAL_3013_Tomatenrot:       n = 3013
    Case RAL_3014_Altrosa:          n = 3014
    Case RAL_3015_Hellrosa:         n = 3015
    Case RAL_3016_Korallenrot:      n = 3016
    Case RAL_3017_Ros�:             n = 3017
    Case RAL_3018_Erdbeerrot:       n = 3018
    Case RAL_3020_Verkehrsrot:      n = 3020
    Case RAL_3022_Lachsrot:         n = 3022
    Case RAL_3024_Leuchtrot:        n = 3024
    Case RAL_3026_Leuchthellrot:    n = 3026
    Case RAL_3027_Himbeerrot:       n = 3027
    Case RAL_3028_Reinrot:          n = 3028
    Case RAL_3031_Orientrot:        n = 3031
    Case RAL_3032_Perlrubinrot:     n = 3032
    Case RAL_3033_Perlrosa:         n = 3033
    
    Case RAL_4001_Rotlila:          n = 4001
    Case RAL_4002_Rotviolett:       n = 4002
    Case RAL_4003_Erikaviolett:     n = 4003
    Case RAL_4004_Bordeauxviolett:  n = 4004
    Case RAL_4005_Blaulila:         n = 4005
    Case RAL_4006_Verkehrspurpur:   n = 4006
    Case RAL_4007_Purpurviolett:    n = 4007
    Case RAL_4008_Signalviolett:    n = 4008
    Case RAL_4009_Pastellviolett:   n = 4009
    Case RAL_4010_Telemagenta:      n = 4010
    Case RAL_4011_Perlviolett:      n = 4011
    Case RAL_4012_Perlbrombeer:     n = 4012
    
    Case RAL_5000_Violettblau:      n = 5000
    Case RAL_5001_Gr�nblau:         n = 5001
    Case RAL_5002_Ultramarinblau:   n = 5002
    Case RAL_5003_Saphirblau:       n = 5003
    Case RAL_5004_Schwarzblau:      n = 5004
    Case RAL_5005_Signalblau:       n = 5005
    Case RAL_5007_Brillantblau:     n = 5007
    Case RAL_5008_Graublau:         n = 5008
    Case RAL_5009_Azurblau:         n = 5009
    Case RAL_5010_Enzianblau:       n = 5010
    Case RAL_5011_Stahlblau:        n = 5011
    Case RAL_5012_Lichtblau:        n = 5012
    Case RAL_5013_Kobaltblau:       n = 5013
    Case RAL_5014_Taubenblau:       n = 5014
    Case RAL_5015_Himmelblau:       n = 5015
    Case RAL_5017_Verkehrsblau:     n = 5017
    Case RAL_5018_T�rkisblau:       n = 5018
    Case RAL_5019_Capriblau:        n = 5019
    Case RAL_5020_Ozeanblau:        n = 5020
    Case RAL_5021_Wasserblau:       n = 5021
    Case RAL_5022_Nachtblau:        n = 5022
    Case RAL_5023_Fernblau:         n = 5023
    Case RAL_5024_Pastellblau:      n = 5024
    Case RAL_5025_Perlenzian:       n = 5025
    Case RAL_5026_Perlnachtblau:    n = 5026
    
    Case RAL_6000_Patinagr�n:       n = 6000
    Case RAL_6001_Smaragdgr�n:      n = 6001
    Case RAL_6002_Laubgr�n:         n = 6002
    Case RAL_6003_Olivgr�n:         n = 6003
    Case RAL_6004_Blaugr�n:         n = 6004
    Case RAL_6005_Moosgr�n:         n = 6005
    Case RAL_6006_Grauoliv:         n = 6006
    Case RAL_6007_Flaschengr�n:     n = 6007
    Case RAL_6008_Braungr�n:        n = 6008
    Case RAL_6009_Tannengr�n:       n = 6009
    Case RAL_6010_Grasgr�n:         n = 6010
    Case RAL_6011_Resedagr�n:       n = 6011
    Case RAL_6012_Schwarzgr�n:      n = 6012
    Case RAL_6013_Schilfgr�n:       n = 6013
    Case RAL_6014_Gelboliv:         n = 6014
    Case RAL_6015_Schwarzoliv:      n = 6015
    Case RAL_6016_T�rkisgr�n:       n = 6016
    Case RAL_6017_Maigr�n:          n = 6017
    Case RAL_6018_Gelbgr�n:         n = 6018
    Case RAL_6019_Wei�gr�n:         n = 6019
    Case RAL_6020_Chromoxidgr�n:    n = 6020
    Case RAL_6021_Blassgr�n:        n = 6021
    Case RAL_6022_Braunoliv:        n = 6022
    Case RAL_6024_Verkehrsgr�n:     n = 6024
    Case RAL_6025_Farngr�n:         n = 6025
    Case RAL_6026_Opalgr�n:         n = 6026
    Case RAL_6027_Lichtgr�n:        n = 6027
    Case RAL_6028_Kieferngr�n:      n = 6028
    Case RAL_6029_Minzgr�n:         n = 6029
    Case RAL_6032_Signalgr�n:       n = 6032
    Case RAL_6033_Mintt�rkis:       n = 6033
    Case RAL_6034_Pastellt�rkis:    n = 6034
    Case RAL_6035_Perlgr�n:         n = 6035
    Case RAL_6036_Perlopalgr�n:     n = 6036
    Case RAL_6037_Reingr�n:         n = 6037
    Case RAL_6038_Leuchtgr�n:       n = 6038
    Case RAL_6039_Fasergr�n:        n = 6039
    
    Case RAL_7000_Fehgrau:          n = 7000
    Case RAL_7001_Silbergrau:       n = 7001
    Case RAL_7002_Olivgrau:         n = 7002
    Case RAL_7003_Moosgrau:         n = 7003
    Case RAL_7004_Signalgrau:       n = 7004
    Case RAL_7005_Mausgrau:         n = 7005
    Case RAL_7006_Beigegrau:        n = 7006
    Case RAL_7008_Khakigrau:        n = 7008
    Case RAL_7009_Gr�ngrau:         n = 7009
    Case RAL_7010_Zeltgrau:         n = 7010
    Case RAL_7011_Eisengrau:        n = 7011
    Case RAL_7012_Basaltgrau:       n = 7012
    Case RAL_7013_Braungrau:        n = 7013
    Case RAL_7015_Schiefergrau:     n = 7015
    Case RAL_7016_Anthrazitgrau:    n = 7016
    Case RAL_7021_Schwarzgrau:      n = 7021
    Case RAL_7022_Umbragrau:        n = 7022
    Case RAL_7023_Betongrau:        n = 7023
    Case RAL_7024_Graphitgrau:      n = 7024
    Case RAL_7026_Granitgrau:       n = 7026
    Case RAL_7030_Steingrau:        n = 7030
    Case RAL_7031_Blaugrau:         n = 7031
    Case RAL_7032_Kieselgrau:       n = 7032
    Case RAL_7033_Zementgrau:       n = 7033
    Case RAL_7034_Gelbgrau:         n = 7034
    Case RAL_7035_Lichtgrau:        n = 7035
    Case RAL_7036_Platingrau:       n = 7036
    Case RAL_7037_Staubgrau:        n = 7037
    Case RAL_7038_Achatgrau:        n = 7038
    Case RAL_7039_Quarzgrau:        n = 7039
    Case RAL_7040_Fenstergrau:      n = 7040
    Case RAL_7042_VerkehrsgrauA:    n = 7042
    Case RAL_7043_VerkehrsgrauB:    n = 7043
    Case RAL_7044_Seidengrau:       n = 7044
    Case RAL_7045_Telegrau1:        n = 7045
    Case RAL_7046_Telegrau2:        n = 7046
    Case RAL_7047_Telegrau4:        n = 7047
    Case RAL_7048_Perlmausgrau:     n = 7048
    
    Case RAL_8000_Gr�nbraun:        n = 8000
    Case RAL_8001_Ockerbraun:       n = 8001
    Case RAL_8002_Signalbraun:      n = 8002
    Case RAL_8003_Lehmbraun:        n = 8003
    Case RAL_8004_Kupferbraun:      n = 8004
    Case RAL_8007_Rehbraun:         n = 8007
    Case RAL_8008_Olivbraun:        n = 8008
    Case RAL_8011_Nussbraun:        n = 8011
    Case RAL_8012_Rotbraun:         n = 8012
    Case RAL_8014_Sepiabraun:       n = 8014
    Case RAL_8015_Kastanienbraun:   n = 8015
    Case RAL_8016_Mahagonibraun:    n = 8016
    Case RAL_8017_Schokoladenbraun: n = 8017
    Case RAL_8019_Graubraun:        n = 8019
    Case RAL_8022_Schwarzbraun:     n = 8022
    Case RAL_8023_Orangebraun:      n = 8023
    Case RAL_8024_Beigebraun:       n = 8024
    Case RAL_8025_Blassbraun:       n = 8025
    Case RAL_8028_Terrabraun:       n = 8028
    Case RAL_8029_Perlkupfer:       n = 8029
    
    Case RAL_9001_Cremewei�:        n = 9001
    Case RAL_9002_Grauwei�:         n = 9002
    Case RAL_9003_Signalwei�:       n = 9003
    Case RAL_9004_Signalschwarz:    n = 9004
    Case RAL_9005_Tiefschwarz:      n = 9005
    Case RAL_9006_Wei�aluminium:    n = 9006
    Case RAL_9007_Graualuminium:    n = 9007
    Case RAL_9010_Reinwei�:         n = 9010
    Case RAL_9011_Graphitschwarz:   n = 9011
    Case RAL_9012_Reinraumwei�:     n = 9012
    Case RAL_9016_Verkehrswei�:     n = 9016
    Case RAL_9017_Verkehrsschwarz:  n = 9017
    Case RAL_9018_Papyruswei�:      n = 9018
    Case RAL_9022_Perlhellgrau:     n = 9022
    Case RAL_9023_Perldunkelgrau:   n = 9023
    Case RAL_9020_SeidenmattWei�:   n = 9020
    
    Case RAL_6031_Bronzegr�n:       n = 6031
    Case RAL_8027_Lederbraun:       n = 8027
    Case RAL_9021_Teerschwarz:      n = 9021
    Case RAL_1039_Sandbeige:        n = 1039
    Case RAL_1040_Lehmbeige:        n = 1040
    Case RAL_6040_Helloliv:         n = 6040
    Case RAL_7050_Tarngrau:         n = 7050
    Case RAL_8031_Sandbraun:        n = 8031
    
    End Select
    RALClassic_ToNum = n
End Function

Public Function RALClassic_NumToColor(ByVal num As Long) As RALClassic
    If m_Count = 0 Then RALClassicColor_Init
    Dim c As RALClassic
    Select Case num
    
    Case 1000: c = RAL_1000_Gr�nbeige
    Case 1001: c = RAL_1001_Beige
    Case 1002: c = RAL_1002_Sandgelb
    Case 1003: c = RAL_1003_Signalgelb
    Case 1004: c = RAL_1004_Goldgelb
    Case 1005: c = RAL_1005_Honiggelb
    Case 1006: c = RAL_1006_Maisgelb
    Case 1007: c = RAL_1007_Narzissengelb
    Case 1011: c = RAL_1011_Braunbeige
    Case 1012: c = RAL_1012_Zitronengelb
    Case 1013: c = RAL_1013_Perlwei�
    Case 1014: c = RAL_1014_Elfenbein
    Case 1015: c = RAL_1015_Hellelfenbein
    Case 1016: c = RAL_1016_Schwefelgelb
    Case 1017: c = RAL_1017_Safrangelb
    Case 1018: c = RAL_1018_Zinkgelb
    Case 1019: c = RAL_1019_Graubeige
    Case 1020: c = RAL_1020_Olivgelb
    Case 1021: c = RAL_1021_Rapsgelb
    Case 1023: c = RAL_1023_Verkehrsgelb
    Case 1024: c = RAL_1024_Ockergelb
    Case 1026: c = RAL_1026_Leuchtgelb
    Case 1027: c = RAL_1027_Currygelb
    Case 1028: c = RAL_1028_Melonengelb
    Case 1032: c = RAL_1032_Ginstergelb
    Case 1033: c = RAL_1033_Dahliengelb
    Case 1034: c = RAL_1034_Pastellgelb
    Case 1035: c = RAL_1035_Perlbeige
    Case 1036: c = RAL_1036_Perlgold
    Case 1037: c = RAL_1037_Sonnengelb
    
    Case 2000: c = RAL_2000_Gelborange
    Case 2001: c = RAL_2001_Rotorange
    Case 2002: c = RAL_2002_Blutorange
    Case 2003: c = RAL_2003_Pastellorange
    Case 2004: c = RAL_2004_Reinorange
    Case 2005: c = RAL_2005_Leuchtorange
    Case 2007: c = RAL_2007_Leuchthellorange
    Case 2008: c = RAL_2008_Hellrotorange
    Case 2009: c = RAL_2009_Verkehrsorange
    Case 2010: c = RAL_2010_Signalorange
    Case 2011: c = RAL_2011_Tieforange
    Case 2012: c = RAL_2012_Lachsorange
    Case 2013: c = RAL_2013_Perlorange
    Case 2017: c = RAL_2017_RALOrange
    
    Case 3000: c = RAL_3000_Feuerrot
    Case 3001: c = RAL_3001_Signalrot
    Case 3002: c = RAL_3002_Karminrot
    Case 3003: c = RAL_3003_Rubinrot
    Case 3004: c = RAL_3004_Purpurrot
    Case 3005: c = RAL_3005_Weinrot
    Case 3007: c = RAL_3007_Schwarzrot
    Case 3009: c = RAL_3009_Oxidrot
    Case 3011: c = RAL_3011_Braunrot
    Case 3012: c = RAL_3012_Beigerot
    Case 3013: c = RAL_3013_Tomatenrot
    Case 3014: c = RAL_3014_Altrosa
    Case 3015: c = RAL_3015_Hellrosa
    Case 3016: c = RAL_3016_Korallenrot
    Case 3017: c = RAL_3017_Ros�
    Case 3018: c = RAL_3018_Erdbeerrot
    Case 3020: c = RAL_3020_Verkehrsrot
    Case 3022: c = RAL_3022_Lachsrot
    Case 3024: c = RAL_3024_Leuchtrot
    Case 3026: c = RAL_3026_Leuchthellrot
    Case 3027: c = RAL_3027_Himbeerrot
    Case 3028: c = RAL_3028_Reinrot
    Case 3031: c = RAL_3031_Orientrot
    Case 3032: c = RAL_3032_Perlrubinrot
    Case 3033: c = RAL_3033_Perlrosa
    
    Case 4001: c = RAL_4001_Rotlila
    Case 4002: c = RAL_4002_Rotviolett
    Case 4003: c = RAL_4003_Erikaviolett
    Case 4004: c = RAL_4004_Bordeauxviolett
    Case 4005: c = RAL_4005_Blaulila
    Case 4006: c = RAL_4006_Verkehrspurpur
    Case 4007: c = RAL_4007_Purpurviolett
    Case 4008: c = RAL_4008_Signalviolett
    Case 4009: c = RAL_4009_Pastellviolett
    Case 4010: c = RAL_4010_Telemagenta
    Case 4011: c = RAL_4011_Perlviolett
    Case 4012: c = RAL_4012_Perlbrombeer
    
    Case 5000: c = RAL_5000_Violettblau
    Case 5001: c = RAL_5001_Gr�nblau
    Case 5002: c = RAL_5002_Ultramarinblau
    Case 5003: c = RAL_5003_Saphirblau
    Case 5004: c = RAL_5004_Schwarzblau
    Case 5005: c = RAL_5005_Signalblau
    Case 5007: c = RAL_5007_Brillantblau
    Case 5008: c = RAL_5008_Graublau
    Case 5009: c = RAL_5009_Azurblau
    Case 5010: c = RAL_5010_Enzianblau
    Case 5011: c = RAL_5011_Stahlblau
    Case 5012: c = RAL_5012_Lichtblau
    Case 5013: c = RAL_5013_Kobaltblau
    Case 5014: c = RAL_5014_Taubenblau
    Case 5015: c = RAL_5015_Himmelblau
    Case 5017: c = RAL_5017_Verkehrsblau
    Case 5018: c = RAL_5018_T�rkisblau
    Case 5019: c = RAL_5019_Capriblau
    Case 5020: c = RAL_5020_Ozeanblau
    Case 5021: c = RAL_5021_Wasserblau
    Case 5022: c = RAL_5022_Nachtblau
    Case 5023: c = RAL_5023_Fernblau
    Case 5024: c = RAL_5024_Pastellblau
    Case 5025: c = RAL_5025_Perlenzian
    Case 5026: c = RAL_5026_Perlnachtblau
    
    Case 6000: c = RAL_6000_Patinagr�n
    Case 6001: c = RAL_6001_Smaragdgr�n
    Case 6002: c = RAL_6002_Laubgr�n
    Case 6003: c = RAL_6003_Olivgr�n
    Case 6004: c = RAL_6004_Blaugr�n
    Case 6005: c = RAL_6005_Moosgr�n
    Case 6006: c = RAL_6006_Grauoliv
    Case 6007: c = RAL_6007_Flaschengr�n
    Case 6008: c = RAL_6008_Braungr�n
    Case 6009: c = RAL_6009_Tannengr�n
    Case 6010: c = RAL_6010_Grasgr�n
    Case 6011: c = RAL_6011_Resedagr�n
    Case 6012: c = RAL_6012_Schwarzgr�n
    Case 6013: c = RAL_6013_Schilfgr�n
    Case 6014: c = RAL_6014_Gelboliv
    Case 6015: c = RAL_6015_Schwarzoliv
    Case 6016: c = RAL_6016_T�rkisgr�n
    Case 6017: c = RAL_6017_Maigr�n
    Case 6018: c = RAL_6018_Gelbgr�n
    Case 6019: c = RAL_6019_Wei�gr�n
    Case 6020: c = RAL_6020_Chromoxidgr�n
    Case 6021: c = RAL_6021_Blassgr�n
    Case 6022: c = RAL_6022_Braunoliv
    Case 6024: c = RAL_6024_Verkehrsgr�n
    Case 6025: c = RAL_6025_Farngr�n
    Case 6026: c = RAL_6026_Opalgr�n
    Case 6027: c = RAL_6027_Lichtgr�n
    Case 6028: c = RAL_6028_Kieferngr�n
    Case 6029: c = RAL_6029_Minzgr�n
    Case 6032: c = RAL_6032_Signalgr�n
    Case 6033: c = RAL_6033_Mintt�rkis
    Case 6034: c = RAL_6034_Pastellt�rkis
    Case 6035: c = RAL_6035_Perlgr�n
    Case 6036: c = RAL_6036_Perlopalgr�n
    Case 6037: c = RAL_6037_Reingr�n
    Case 6038: c = RAL_6038_Leuchtgr�n
    Case 6039: c = RAL_6039_Fasergr�n
    
    Case 7000: c = RAL_7000_Fehgrau
    Case 7001: c = RAL_7001_Silbergrau
    Case 7002: c = RAL_7002_Olivgrau
    Case 7003: c = RAL_7003_Moosgrau
    Case 7004: c = RAL_7004_Signalgrau
    Case 7005: c = RAL_7005_Mausgrau
    Case 7006: c = RAL_7006_Beigegrau
    Case 7008: c = RAL_7008_Khakigrau
    Case 7009: c = RAL_7009_Gr�ngrau
    Case 7010: c = RAL_7010_Zeltgrau
    Case 7011: c = RAL_7011_Eisengrau
    Case 7012: c = RAL_7012_Basaltgrau
    Case 7013: c = RAL_7013_Braungrau
    Case 7015: c = RAL_7015_Schiefergrau
    Case 7016: c = RAL_7016_Anthrazitgrau
    Case 7021: c = RAL_7021_Schwarzgrau
    Case 7022: c = RAL_7022_Umbragrau
    Case 7023: c = RAL_7023_Betongrau
    Case 7024: c = RAL_7024_Graphitgrau
    Case 7026: c = RAL_7026_Granitgrau
    Case 7030: c = RAL_7030_Steingrau
    Case 7031: c = RAL_7031_Blaugrau
    Case 7032: c = RAL_7032_Kieselgrau
    Case 7033: c = RAL_7033_Zementgrau
    Case 7034: c = RAL_7034_Gelbgrau
    Case 7035: c = RAL_7035_Lichtgrau
    Case 7036: c = RAL_7036_Platingrau
    Case 7037: c = RAL_7037_Staubgrau
    Case 7038: c = RAL_7038_Achatgrau
    Case 7039: c = RAL_7039_Quarzgrau
    Case 7040: c = RAL_7040_Fenstergrau
    Case 7042: c = RAL_7042_VerkehrsgrauA
    Case 7043: c = RAL_7043_VerkehrsgrauB
    Case 7044: c = RAL_7044_Seidengrau
    Case 7045: c = RAL_7045_Telegrau1
    Case 7046: c = RAL_7046_Telegrau2
    Case 7047: c = RAL_7047_Telegrau4
    Case 7048: c = RAL_7048_Perlmausgrau
    
    Case 8000: c = RAL_8000_Gr�nbraun
    Case 8001: c = RAL_8001_Ockerbraun
    Case 8002: c = RAL_8002_Signalbraun
    Case 8003: c = RAL_8003_Lehmbraun
    Case 8004: c = RAL_8004_Kupferbraun
    Case 8007: c = RAL_8007_Rehbraun
    Case 8008: c = RAL_8008_Olivbraun
    Case 8011: c = RAL_8011_Nussbraun
    Case 8012: c = RAL_8012_Rotbraun
    Case 8014: c = RAL_8014_Sepiabraun
    Case 8015: c = RAL_8015_Kastanienbraun
    Case 8016: c = RAL_8016_Mahagonibraun
    Case 8017: c = RAL_8017_Schokoladenbraun
    Case 8019: c = RAL_8019_Graubraun
    Case 8022: c = RAL_8022_Schwarzbraun
    Case 8023: c = RAL_8023_Orangebraun
    Case 8024: c = RAL_8024_Beigebraun
    Case 8025: c = RAL_8025_Blassbraun
    Case 8028: c = RAL_8028_Terrabraun
    Case 8029: c = RAL_8029_Perlkupfer
    
    Case 9001: c = RAL_9001_Cremewei�
    Case 9002: c = RAL_9002_Grauwei�
    Case 9003: c = RAL_9003_Signalwei�
    Case 9004: c = RAL_9004_Signalschwarz
    Case 9005: c = RAL_9005_Tiefschwarz
    Case 9006: c = RAL_9006_Wei�aluminium
    Case 9007: c = RAL_9007_Graualuminium
    Case 9010: c = RAL_9010_Reinwei�
    Case 9011: c = RAL_9011_Graphitschwarz
    Case 9012: c = RAL_9012_Reinraumwei�
    Case 9016: c = RAL_9016_Verkehrswei�
    Case 9017: c = RAL_9017_Verkehrsschwarz
    Case 9018: c = RAL_9018_Papyruswei�
    Case 9022: c = RAL_9022_Perlhellgrau
    Case 9023: c = RAL_9023_Perldunkelgrau
    Case 9020: c = RAL_9020_SeidenmattWei�
        
    Case 6031: c = RAL_6031_Bronzegr�n
    Case 8027: c = RAL_8027_Lederbraun
    Case 9021: c = RAL_9021_Teerschwarz
    Case 1039: c = RAL_1039_Sandbeige
    Case 1040: c = RAL_1040_Lehmbeige
    Case 6040: c = RAL_6040_Helloliv
    Case 7050: c = RAL_7050_Tarngrau
    Case 8031: c = RAL_8031_Sandbraun
    
    End Select
    RALClassic_NumToColor = c
End Function

Public Function RALClassic_NumToColorname(num As Long) As String
    If m_Count = 0 Then RALClassicColor_Init
    Dim s As String: s = RALClassic_NameToStr(RALClassic_NumToColor(num))
    If Len(s) = 0 Then Exit Function
    RALClassic_NumToColorname = "RAL_" & CStr(num) & "_" & s
End Function

Public Sub RALClassic_ToListBox(aCBLB)
    If m_Count = 0 Then RALClassicColor_Init
    aCBLB.Clear
    
    Dim i As Long, j As Long
    Dim num As Long, s As String
    i = 1000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
    Loop While i <= 1037
    
    i = 2000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
    Loop While i <= 2017
    
    i = 3000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
    Loop While i <= 3033
    
    i = 4000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
    Loop While i <= 4012
    
    i = 5000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
    Loop While i <= 5026
    
    i = 6000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
        If i = 6031 Then i = i + 1
        If i = 6040 Then i = i + 1
    Loop While i <= 6039
    
    For i = 7000 To 7048
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
    Next
    
    i = 8000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
        If i = 8027 Then i = i + 1
    Loop While i <= 8029
    
    i = 9000
    Do
        s = RALClassic_NumToColorname(i): If Len(s) Then aCBLB.AddItem s
        i = i + 1
        If i = 9021 Then i = i + 1
    Loop While i <= 9023
    
    s = RALClassic_NumToColorname(6031): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(8027): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(9021): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(1039): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(1040): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(6040): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(7050): If Len(s) Then aCBLB.AddItem s
    s = RALClassic_NumToColorname(8031): If Len(s) Then aCBLB.AddItem s
    
End Sub

Public Function RALClassic_Parse(ByVal s As String) As RALClassic
    If m_Count = 0 Then RALClassicColor_Init
    'bsp: RAL_9001_Cremewei�
    'read the number
    If Left(s, 4) = "RAL_" Then
        s = Mid(s, 5)
        Dim pos As Long: pos = InStr(1, s, "_")
        If pos > 0 Then s = Left(s, pos - 1)
        If Not IsNumeric(s) Then s = Left(s, 4)
        If Not IsNumeric(s) Then Exit Function
        Dim num As Long: num = CLng(s)
        RALClassic_Parse = RALClassic_NumToColor(num)
    End If
End Function

