Attribute VB_Name = "MRALColors"
Option Explicit
'https://de.wikipedia.org/wiki/RAL-Farbe
Public Enum RALClassic
    'RAL-1X Gelb und Beige
    RAL_1000_Grünbeige = &H88BACD
    RAL_1001_Beige = &H84B0D0
    RAL_1002_Sandgelb = &H6DAAD2
    RAL_1003_Signalgelb = &HA8F9&
    RAL_1004_Goldgelb = &H7B0E2
    RAL_1005_Honiggelb = &H8ECB&
    RAL_1006_Maisgelb = &H90E2&
    RAL_1007_Narzissengelb = &H8CE8&
    RAL_1011_Braunbeige = &H548AAF
    RAL_1012_Zitronengelb = &H22C0D9
    RAL_1013_Perlweiß = &HC6D9E3
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
    RAL_3017_Rosé = &H5F54D3
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
    RAL_5001_Grünblau = &H644C0F
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
    RAL_5018_Türkisblau = &H8F8821
    RAL_5019_Capriblau = &H84571A
    RAL_5020_Ozeanblau = &H51410B
    RAL_5021_Wasserblau = &H7A7307
    RAL_5022_Nachtblau = &H5A2D22
    RAL_5023_Fernblau = &H8E664D
    RAL_5024_Pastellblau = &HB0936A
    RAL_5025_Perlenzian = &H786429
    RAL_5026_Perlnachtblau = &H542C10
    
    'RAL-6X Grün
    RAL_6000_Patinagrün = &H60743C
    RAL_6001_Smaragdgrün = &H356736
    RAL_6002_Laubgrün = &H285932
    RAL_6003_Olivgrün = &H3C5350
    RAL_6004_Blaugrün = &H424402
    RAL_6005_Moosgrün = &H324211
    RAL_6006_Grauoliv = &H2E393C
    RAL_6007_Flaschengrün = &H22322C
    RAL_6008_Braungrün = &H2A3437
    RAL_6009_Tannengrün = &H2A3527
    RAL_6010_Grasgrün = &H396F4D
    RAL_6011_Resedagrün = &H597C6C
    RAL_6012_Schwarzgrün = &H3A3D30
    RAL_6013_Schilfgrün = &H5A767D
    RAL_6014_Gelboliv = &H354147
    RAL_6015_Schwarzoliv = &H363D3D
    RAL_6016_Türkisgrün = &H4C6900
    RAL_6017_Maigrün = &H407F58
    RAL_6018_Gelbgrün = &H3B9961
    RAL_6019_Weißgrün = &HACCEB9
    RAL_6020_Chromoxidgrün = &H2F4237
    RAL_6021_Blassgrün = &H77998A
    RAL_6022_Braunoliv = &H27333A
    RAL_6024_Verkehrsgrün = &H518300
    RAL_6025_Farngrün = &H3B6E5E
    RAL_6026_Opalgrün = &H4E5F00
    RAL_6027_Lichtgrün = &HB5BA7E
    RAL_6028_Kieferngrün = &H425431
    RAL_6029_Minzgrün = &H3D6F00
    RAL_6032_Signalgrün = &H527F23
    RAL_6033_Minttürkis = &H7F8746
    RAL_6034_Pastelltürkis = &HACAC7A
    RAL_6035_Perlgrün = &H254D19
    RAL_6036_Perlopalgrün = &H4B5704
    RAL_6037_Reingrün = &H298B00
    RAL_6038_Leuchtgrün = &H1AB500
    RAL_6039_Fasergrün = &H3FC5B3
    
    'RAL-7X Grau
    RAL_7000_Fehgrau = &H8E887A
    RAL_7001_Silbergrau = &H9F998F
    RAL_7002_Olivgrau = &H637881
    RAL_7003_Moosgrau = &H69767A
    RAL_7004_Signalgrau = &H9B9B9B
    RAL_7005_Mausgrau = &H6F716B
    RAL_7006_Beigegrau = &H616F75
    RAL_7008_Khakigrau = &H3D5E74
    RAL_7009_Grüngrau = &H58605D
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
    RAL_8000_Grünbraun = &H3E6989
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
    
    'RAL-9X Weiß und Schwarz
    RAL_9001_Cremeweiß = &HD2E0E9
    RAL_9002_Grauweiß = &HCBD5D7
    RAL_9003_Signalweiß = &HF4F8F4
    RAL_9004_Signalschwarz = &H32302E
    RAL_9005_Tiefschwarz = &H100E0E
    RAL_9006_Weißaluminium = &HA0A1A1
    RAL_9007_Graualuminium = &H838687
    RAL_9010_Reinweiß = &HEFF9F7
    RAL_9011_Graphitschwarz = &H2F2C29
    RAL_9012_Reinraumweiß = &HE6FDFF
    RAL_9016_Verkehrsweiß = &HF5FBF7
    RAL_9017_Verkehrsschwarz = &H2F2D2A
    RAL_9018_Papyrusweiß = &HC3CAC7
    RAL_9022_Perlhellgrau = &H9C9C9C
    RAL_9023_Perldunkelgrau = &H82817E
    RAL_9020_SeidenmattWeiß = &HFDFDFD
    
    'RAL F9 Tarnfarben der Bundeswehr
    RAL_6031_Bronzegrün = &H465748
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
    m_Arr(i) = TNamedRALColor("Grünbeige", 1000, RALClassic.RAL_1000_Grünbeige):             i = i + 1
    m_Arr(i) = TNamedRALColor("Beige", 1001, RALClassic.RAL_1001_Beige):                     i = i + 1
    m_Arr(i) = TNamedRALColor("Sandgelb", 1002, RALClassic.RAL_1002_Sandgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgelb", 1003, RALClassic.RAL_1003_Signalgelb):           i = i + 1
    m_Arr(i) = TNamedRALColor("Goldgelb", 1004, RALClassic.RAL_1004_Goldgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Honiggelb", 1005, RALClassic.RAL_1005_Honiggelb):             i = i + 1
    m_Arr(i) = TNamedRALColor("Maisgelb", 1006, RALClassic.RAL_1006_Maisgelb):               i = i + 1
    m_Arr(i) = TNamedRALColor("Narzissengelb", 1007, RALClassic.RAL_1007_Narzissengelb):     i = i + 1
    m_Arr(i) = TNamedRALColor("Braunbeige", 1011, RALClassic.RAL_1011_Braunbeige):           i = i + 1
    m_Arr(i) = TNamedRALColor("Zitronengelb", 1012, RALClassic.RAL_1012_Zitronengelb):       i = i + 1
    m_Arr(i) = TNamedRALColor("Perlweiß", 1013, RALClassic.RAL_1013_Perlweiß):               i = i + 1
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
    m_Arr(i) = TNamedRALColor("Rosé", 3017, RALClassic.RAL_3017_Rosé):                    i = i + 1
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
    m_Arr(i) = TNamedRALColor("Grünblau", 5001, RALClassic.RAL_5001_Grünblau):               i = i + 1
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
    m_Arr(i) = TNamedRALColor("Türkisblau", 5018, RALClassic.RAL_5018_Türkisblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Capriblau", 5019, RALClassic.RAL_5019_Capriblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Ozeanblau", 5020, RALClassic.RAL_5020_Ozeanblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Wasserblau", 5021, RALClassic.RAL_5021_Wasserblau):           i = i + 1
    m_Arr(i) = TNamedRALColor("Nachtblau", 5022, RALClassic.RAL_5022_Nachtblau):             i = i + 1
    m_Arr(i) = TNamedRALColor("Fernblau", 5023, RALClassic.RAL_5023_Fernblau):               i = i + 1
    m_Arr(i) = TNamedRALColor("Pastellblau", 5024, RALClassic.RAL_5024_Pastellblau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perlenzian", 5025, RALClassic.RAL_5025_Perlenzian):           i = i + 1
    m_Arr(i) = TNamedRALColor("Perlnachtblau", 5026, RALClassic.RAL_5026_Perlnachtblau):     i = i + 1
    
    m_Arr(i) = TNamedRALColor("Patinagrün", 6000, RALClassic.RAL_6000_Patinagrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Smaragdgrün", 6001, RALClassic.RAL_6001_Smaragdgrün):      i = i + 1
    m_Arr(i) = TNamedRALColor("Laubgrün", 6002, RALClassic.RAL_6002_Laubgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Olivgrün", 6003, RALClassic.RAL_6003_Olivgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Blaugrün", 6004, RALClassic.RAL_6004_Blaugrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Moosgrün", 6005, RALClassic.RAL_6005_Moosgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Grauoliv", 6006, RALClassic.RAL_6006_Grauoliv):            i = i + 1
    m_Arr(i) = TNamedRALColor("Flaschengrün", 6007, RALClassic.RAL_6007_Flaschengrün):    i = i + 1
    m_Arr(i) = TNamedRALColor("Braungrün", 6008, RALClassic.RAL_6008_Braungrün):          i = i + 1
    m_Arr(i) = TNamedRALColor("Tannengrün", 6009, RALClassic.RAL_6009_Tannengrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Grasgrün", 6010, RALClassic.RAL_6010_Grasgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Resedagrün", 6011, RALClassic.RAL_6011_Resedagrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzgrün", 6012, RALClassic.RAL_6012_Schwarzgrün):      i = i + 1
    m_Arr(i) = TNamedRALColor("Schilfgrün", 6013, RALClassic.RAL_6013_Schilfgrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Gelboliv", 6014, RALClassic.RAL_6014_Gelboliv):            i = i + 1
    m_Arr(i) = TNamedRALColor("Schwarzoliv", 6015, RALClassic.RAL_6015_Schwarzoliv):      i = i + 1
    m_Arr(i) = TNamedRALColor("Türkisgrün", 6016, RALClassic.RAL_6016_Türkisgrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Maigrün", 6017, RALClassic.RAL_6017_Maigrün):              i = i + 1
    m_Arr(i) = TNamedRALColor("Gelbgrün", 6018, RALClassic.RAL_6018_Gelbgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Weißgrün", 6019, RALClassic.RAL_6019_Weißgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Chromoxidgrün", 6020, RALClassic.RAL_6020_Chromoxidgrün):  i = i + 1
    m_Arr(i) = TNamedRALColor("Blassgrün", 6021, RALClassic.RAL_6021_Blassgrün):          i = i + 1
    m_Arr(i) = TNamedRALColor("Braunoliv", 6022, RALClassic.RAL_6022_Braunoliv):          i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsgrün", 6024, RALClassic.RAL_6024_Verkehrsgrün):    i = i + 1
    m_Arr(i) = TNamedRALColor("Farngrün", 6025, RALClassic.RAL_6025_Farngrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Opalgrün", 6026, RALClassic.RAL_6026_Opalgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Lichtgrün", 6027, RALClassic.RAL_6027_Lichtgrün):          i = i + 1
    m_Arr(i) = TNamedRALColor("Kieferngrün", 6028, RALClassic.RAL_6028_Kieferngrün):      i = i + 1
    m_Arr(i) = TNamedRALColor("Minzgrün", 6029, RALClassic.RAL_6029_Minzgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgrün", 6032, RALClassic.RAL_6032_Signalgrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Minttürkis", 6033, RALClassic.RAL_6033_Minttürkis):        i = i + 1
    m_Arr(i) = TNamedRALColor("Pastelltürkis", 6034, RALClassic.RAL_6034_Pastelltürkis):  i = i + 1
    m_Arr(i) = TNamedRALColor("Perlgrün", 6035, RALClassic.RAL_6035_Perlgrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Perlopalgrün", 6036, RALClassic.RAL_6036_Perlopalgrün):    i = i + 1
    m_Arr(i) = TNamedRALColor("Reingrün", 6037, RALClassic.RAL_6037_Reingrün):            i = i + 1
    m_Arr(i) = TNamedRALColor("Leuchtgrün", 6038, RALClassic.RAL_6038_Leuchtgrün):        i = i + 1
    m_Arr(i) = TNamedRALColor("Fasergrün", 6039, RALClassic.RAL_6039_Fasergrün):          i = i + 1
    
    m_Arr(i) = TNamedRALColor("Fehgrau", 7000, RALClassic.RAL_7000_Fehgrau):              i = i + 1
    m_Arr(i) = TNamedRALColor("Silbergrau", 7001, RALClassic.RAL_7001_Silbergrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Olivgrau", 7002, RALClassic.RAL_7002_Olivgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Moosgrau", 7003, RALClassic.RAL_7003_Moosgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Signalgrau", 7004, RALClassic.RAL_7004_Signalgrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Mausgrau", 7005, RALClassic.RAL_7005_Mausgrau):            i = i + 1
    m_Arr(i) = TNamedRALColor("Beigegrau", 7006, RALClassic.RAL_7006_Beigegrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Khakigrau", 7008, RALClassic.RAL_7008_Khakigrau):          i = i + 1
    m_Arr(i) = TNamedRALColor("Grüngrau", 7009, RALClassic.RAL_7009_Grüngrau):            i = i + 1
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
    
    m_Arr(i) = TNamedRALColor("Grünbraun", 8000, RALClassic.RAL_8000_Grünbraun):               i = i + 1
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
    
    m_Arr(i) = TNamedRALColor("Cremeweiß", 9001, RALClassic.RAL_9001_Cremeweiß):               i = i + 1
    m_Arr(i) = TNamedRALColor("Grauweiß", 9002, RALClassic.RAL_9002_Grauweiß):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Signalweiß", 9003, RALClassic.RAL_9003_Signalweiß):             i = i + 1
    m_Arr(i) = TNamedRALColor("Signalschwarz", 9004, RALClassic.RAL_9004_Signalschwarz):       i = i + 1
    m_Arr(i) = TNamedRALColor("Tiefschwarz", 9005, RALClassic.RAL_9005_Tiefschwarz):           i = i + 1
    m_Arr(i) = TNamedRALColor("Weißaluminium", 9006, RALClassic.RAL_9006_Weißaluminium):       i = i + 1
    m_Arr(i) = TNamedRALColor("Graualuminium", 9007, RALClassic.RAL_9007_Graualuminium):       i = i + 1
    m_Arr(i) = TNamedRALColor("Reinweiß", 9010, RALClassic.RAL_9010_Reinweiß):                 i = i + 1
    m_Arr(i) = TNamedRALColor("Graphitschwarz", 9011, RALClassic.RAL_9011_Graphitschwarz):     i = i + 1
    m_Arr(i) = TNamedRALColor("Reinraumweiß", 9012, RALClassic.RAL_9012_Reinraumweiß):         i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsweiß", 9016, RALClassic.RAL_9016_Verkehrsweiß):         i = i + 1
    m_Arr(i) = TNamedRALColor("Verkehrsschwarz", 9017, RALClassic.RAL_9017_Verkehrsschwarz):   i = i + 1
    m_Arr(i) = TNamedRALColor("Papyrusweiß", 9018, RALClassic.RAL_9018_Papyrusweiß):           i = i + 1
    m_Arr(i) = TNamedRALColor("Perlhellgrau", 9022, RALClassic.RAL_9022_Perlhellgrau):         i = i + 1
    m_Arr(i) = TNamedRALColor("Perldunkelgrau", 9023, RALClassic.RAL_9023_Perldunkelgrau):     i = i + 1
    m_Arr(i) = TNamedRALColor("SeidenmattWeiß", 9020, RALClassic.RAL_9020_SeidenmattWeiß):     i = i + 1
    
    m_Arr(i) = TNamedRALColor("Bronzegrün", 6031, RALClassic.RAL_6031_Bronzegrün):    i = i + 1
    m_Arr(i) = TNamedRALColor("Lederbraun", 8027, RALClassic.RAL_8027_Lederbraun):    i = i + 1
    m_Arr(i) = TNamedRALColor("Teerschwarz", 9021, RALClassic.RAL_9021_Teerschwarz):  i = i + 1
    m_Arr(i) = TNamedRALColor("Sandbeige", 1039, RALClassic.RAL_1039_Sandbeige):      i = i + 1
    m_Arr(i) = TNamedRALColor("Lehmbeige", 1040, RALClassic.RAL_1040_Lehmbeige):      i = i + 1
    m_Arr(i) = TNamedRALColor("Helloliv", 6040, RALClassic.RAL_6040_Helloliv):        i = i + 1
    m_Arr(i) = TNamedRALColor("Tarngrau", 7050, RALClassic.RAL_7050_Tarngrau):        i = i + 1
    m_Arr(i) = TNamedRALColor("Sandbraun", 8031, RALClassic.RAL_8031_Sandbraun):      i = i + 1
    Debug.Print i
End Sub


Public Function RALClassic_NameToStr(e As RALClassic) As String
    If m_Count = 0 Then RALClassicColor_Init
    Dim S As String
    Select Case e
    
    Case RAL_1000_Grünbeige:        S = "Grünbeige"
    Case RAL_1001_Beige:            S = "Beige"
    Case RAL_1002_Sandgelb:         S = "Sandgelb"
    Case RAL_1003_Signalgelb:       S = "Signalgelb"
    Case RAL_1004_Goldgelb:         S = "Goldgelb"
    Case RAL_1005_Honiggelb:        S = "Honiggelb"
    Case RAL_1006_Maisgelb:         S = "Maisgelb"
    Case RAL_1007_Narzissengelb:    S = "Narzissengelb"
    Case RAL_1011_Braunbeige:       S = "Braunbeige"
    Case RAL_1012_Zitronengelb:     S = "Zitronengelb"
    Case RAL_1013_Perlweiß:         S = "Perlweiß"
    Case RAL_1014_Elfenbein:        S = "Elfenbein"
    Case RAL_1015_Hellelfenbein:    S = "Hellelfenbein"
    Case RAL_1016_Schwefelgelb:     S = "Schwefelgelb"
    Case RAL_1017_Safrangelb:       S = "Safrangelb"
    Case RAL_1018_Zinkgelb:         S = "Zinkgelb"
    Case RAL_1019_Graubeige:        S = "Graubeige"
    Case RAL_1020_Olivgelb:         S = "Olivgelb"
    Case RAL_1021_Rapsgelb:         S = "Rapsgelb"
    Case RAL_1023_Verkehrsgelb:     S = "Verkehrsgelb"
    Case RAL_1024_Ockergelb:        S = "Ockergelb"
    Case RAL_1026_Leuchtgelb:       S = "Leuchtgelb"
    Case RAL_1027_Currygelb:        S = "Currygelb"
    Case RAL_1028_Melonengelb:      S = "Melonengelb"
    Case RAL_1032_Ginstergelb:      S = "Ginstergelb"
    Case RAL_1033_Dahliengelb:      S = "Dahliengelb"
    Case RAL_1034_Pastellgelb:      S = "Pastellgelb"
    Case RAL_1035_Perlbeige:        S = "Perlbeige"
    Case RAL_1036_Perlgold:         S = "Perlgold"
    Case RAL_1037_Sonnengelb:       S = "Sonnengelb"
    
    Case RAL_2000_Gelborange:       S = "Gelborange"
    Case RAL_2001_Rotorange:        S = "Rotorange"
    Case RAL_2002_Blutorange:       S = "Blutorange"
    Case RAL_2003_Pastellorange:    S = "Pastellorange"
    Case RAL_2004_Reinorange:       S = "Reinorange"
    Case RAL_2005_Leuchtorange:     S = "Leuchtorange"
    Case RAL_2007_Leuchthellorange: S = "Leuchthellorange"
    Case RAL_2008_Hellrotorange:    S = "Hellrotorange"
    Case RAL_2009_Verkehrsorange:   S = "Verkehrsorange"
    Case RAL_2010_Signalorange:     S = "Signalorange"
    Case RAL_2011_Tieforange:       S = "Tieforange"
    Case RAL_2012_Lachsorange:      S = "Lachsorange"
    Case RAL_2013_Perlorange:       S = "Perlorange"
    Case RAL_2017_RALOrange:        S = "RALOrange"
        
    Case RAL_3000_Feuerrot:         S = "Feuerrot"
    Case RAL_3001_Signalrot:        S = "Signalrot"
    Case RAL_3002_Karminrot:        S = "Karminrot"
    Case RAL_3003_Rubinrot:         S = "Rubinrot"
    Case RAL_3004_Purpurrot:        S = "Purpurrot"
    Case RAL_3005_Weinrot:          S = "Weinrot"
    Case RAL_3007_Schwarzrot:       S = "Schwarzrot"
    Case RAL_3009_Oxidrot:          S = "Oxidrot"
    Case RAL_3011_Braunrot:         S = "Braunrot"
    Case RAL_3012_Beigerot:         S = "Beigerot"
    Case RAL_3013_Tomatenrot:       S = "Tomatenrot"
    Case RAL_3014_Altrosa:          S = "Altrosa"
    Case RAL_3015_Hellrosa:         S = "Hellrosa"
    Case RAL_3016_Korallenrot:      S = "Korallenrot"
    Case RAL_3017_Rosé:             S = "Rosé"
    Case RAL_3018_Erdbeerrot:       S = "Erdbeerrot"
    Case RAL_3020_Verkehrsrot:      S = "Verkehrsrot"
    Case RAL_3022_Lachsrot:         S = "Lachsrot"
    Case RAL_3024_Leuchtrot:        S = "Leuchtrot"
    Case RAL_3026_Leuchthellrot:    S = "Leuchthellrot"
    Case RAL_3027_Himbeerrot:       S = "Himbeerrot"
    Case RAL_3028_Reinrot:          S = "Reinrot"
    Case RAL_3031_Orientrot:        S = "Orientrot"
    Case RAL_3032_Perlrubinrot:     S = "Perlrubinrot"
    Case RAL_3033_Perlrosa:         S = "Perlrosa"
    
    Case RAL_4001_Rotlila:          S = "Rotlila"
    Case RAL_4002_Rotviolett:       S = "Rotviolett"
    Case RAL_4003_Erikaviolett:     S = "Erikaviolett"
    Case RAL_4004_Bordeauxviolett:  S = "Bordeauxviolett"
    Case RAL_4005_Blaulila:         S = "Blaulila"
    Case RAL_4006_Verkehrspurpur:   S = "Verkehrspurpur"
    Case RAL_4007_Purpurviolett:    S = "Purpurviolett"
    Case RAL_4008_Signalviolett:    S = "Signalviolett"
    Case RAL_4009_Pastellviolett:   S = "Pastellviolett"
    Case RAL_4010_Telemagenta:      S = "Telemagenta"
    Case RAL_4011_Perlviolett:      S = "Perlviolett"
    Case RAL_4012_Perlbrombeer:     S = "Perlbrombeer"
    
    Case RAL_5000_Violettblau:      S = "Violettblau"
    Case RAL_5001_Grünblau:         S = "Grünblau"
    Case RAL_5002_Ultramarinblau:   S = "Ultramarinblau"
    Case RAL_5003_Saphirblau:       S = "Saphirblau"
    Case RAL_5004_Schwarzblau:      S = "Schwarzblau"
    Case RAL_5005_Signalblau:       S = "Signalblau"
    Case RAL_5007_Brillantblau:     S = "Brillantblau"
    Case RAL_5008_Graublau:         S = "Graublau"
    Case RAL_5009_Azurblau:         S = "Azurblau"
    Case RAL_5010_Enzianblau:       S = "Enzianblau"
    Case RAL_5011_Stahlblau:        S = "Stahlblau"
    Case RAL_5012_Lichtblau:        S = "Lichtblau"
    Case RAL_5013_Kobaltblau:       S = "Kobaltblau"
    Case RAL_5014_Taubenblau:       S = "Taubenblau"
    Case RAL_5015_Himmelblau:       S = "Himmelblau"
    Case RAL_5017_Verkehrsblau:     S = "Verkehrsblau"
    Case RAL_5018_Türkisblau:       S = "Türkisblau"
    Case RAL_5019_Capriblau:        S = "Capriblau"
    Case RAL_5020_Ozeanblau:        S = "Ozeanblau"
    Case RAL_5021_Wasserblau:       S = "Wasserblau"
    Case RAL_5022_Nachtblau:        S = "Nachtblau"
    Case RAL_5023_Fernblau:         S = "Fernblau"
    Case RAL_5024_Pastellblau:      S = "Pastellblau"
    Case RAL_5025_Perlenzian:       S = "Perlenzian"
    Case RAL_5026_Perlnachtblau:    S = "Perlnachtblau"
    
    Case RAL_6000_Patinagrün:       S = "Patinagrün"
    Case RAL_6001_Smaragdgrün:      S = "Smaragdgrün"
    Case RAL_6002_Laubgrün:         S = "Laubgrün"
    Case RAL_6003_Olivgrün:         S = "Olivgrün"
    Case RAL_6004_Blaugrün:         S = "Blaugrün"
    Case RAL_6005_Moosgrün:         S = "Moosgrün"
    Case RAL_6006_Grauoliv:         S = "Grauoliv"
    Case RAL_6007_Flaschengrün:     S = "Flaschengrün"
    Case RAL_6008_Braungrün:        S = "Braungrün"
    Case RAL_6009_Tannengrün:       S = "Tannengrün"
    Case RAL_6010_Grasgrün:         S = "Grasgrün"
    Case RAL_6011_Resedagrün:       S = "Resedagrün"
    Case RAL_6012_Schwarzgrün:      S = "Schwarzgrün"
    Case RAL_6013_Schilfgrün:       S = "Schilfgrün"
    Case RAL_6014_Gelboliv:         S = "Gelboliv"
    Case RAL_6015_Schwarzoliv:      S = "Schwarzoliv"
    Case RAL_6016_Türkisgrün:       S = "Türkisgrün"
    Case RAL_6017_Maigrün:          S = "Maigrün"
    Case RAL_6018_Gelbgrün:         S = "Gelbgrün"
    Case RAL_6019_Weißgrün:         S = "Weißgrün"
    Case RAL_6020_Chromoxidgrün:    S = "Chromoxidgrün"
    Case RAL_6021_Blassgrün:        S = "Blassgrün"
    Case RAL_6022_Braunoliv:        S = "Braunoliv"
    Case RAL_6024_Verkehrsgrün:     S = "Verkehrsgrün"
    Case RAL_6025_Farngrün:         S = "Farngrün"
    Case RAL_6026_Opalgrün:         S = "Opalgrün"
    Case RAL_6027_Lichtgrün:        S = "Lichtgrün"
    Case RAL_6028_Kieferngrün:      S = "Kieferngrün"
    Case RAL_6029_Minzgrün:         S = "Minzgrün"
    Case RAL_6032_Signalgrün:       S = "Signalgrün"
    Case RAL_6033_Minttürkis:       S = "Minttürkis"
    Case RAL_6034_Pastelltürkis:    S = "Pastelltürkis"
    Case RAL_6035_Perlgrün:         S = "Perlgrün"
    Case RAL_6036_Perlopalgrün:     S = "Perlopalgrün"
    Case RAL_6037_Reingrün:         S = "Reingrün"
    Case RAL_6038_Leuchtgrün:       S = "Leuchtgrün"
    Case RAL_6039_Fasergrün:        S = "Fasergrün"
    
    Case RAL_7000_Fehgrau:          S = "Fehgrau"
    Case RAL_7001_Silbergrau:       S = "Silbergrau"
    Case RAL_7002_Olivgrau:         S = "Olivgrau"
    Case RAL_7003_Moosgrau:         S = "Moosgrau"
    Case RAL_7004_Signalgrau:       S = "Signalgrau"
    Case RAL_7005_Mausgrau:         S = "Mausgrau"
    Case RAL_7006_Beigegrau:        S = "Beigegrau"
    Case RAL_7008_Khakigrau:        S = "Khakigrau"
    Case RAL_7009_Grüngrau:         S = "Grüngrau"
    Case RAL_7010_Zeltgrau:         S = "Zeltgrau"
    Case RAL_7011_Eisengrau:        S = "Eisengrau"
    Case RAL_7012_Basaltgrau:       S = "Basaltgrau"
    Case RAL_7013_Braungrau:        S = "Braungrau"
    Case RAL_7015_Schiefergrau:     S = "Schiefergrau"
    Case RAL_7016_Anthrazitgrau:    S = "Anthrazitgrau"
    Case RAL_7021_Schwarzgrau:      S = "Schwarzgrau"
    Case RAL_7022_Umbragrau:        S = "Umbragrau"
    Case RAL_7023_Betongrau:        S = "Betongrau"
    Case RAL_7024_Graphitgrau:      S = "Graphitgrau"
    Case RAL_7026_Granitgrau:       S = "Granitgrau"
    Case RAL_7030_Steingrau:        S = "Steingrau"
    Case RAL_7031_Blaugrau:         S = "Blaugrau"
    Case RAL_7032_Kieselgrau:       S = "Kieselgrau"
    Case RAL_7033_Zementgrau:       S = "Zementgrau"
    Case RAL_7034_Gelbgrau:         S = "Gelbgrau"
    Case RAL_7035_Lichtgrau:        S = "Lichtgrau"
    Case RAL_7036_Platingrau:       S = "Platingrau"
    Case RAL_7037_Staubgrau:        S = "Staubgrau"
    Case RAL_7038_Achatgrau:        S = "Achatgrau"
    Case RAL_7039_Quarzgrau:        S = "Quarzgrau"
    Case RAL_7040_Fenstergrau:      S = "Fenstergrau"
    Case RAL_7042_VerkehrsgrauA:    S = "VerkehrsgrauA"
    Case RAL_7043_VerkehrsgrauB:    S = "VerkehrsgrauB"
    Case RAL_7044_Seidengrau:       S = "Seidengrau"
    Case RAL_7045_Telegrau1:        S = "Telegrau1"
    Case RAL_7046_Telegrau2:        S = "Telegrau2"
    Case RAL_7047_Telegrau4:        S = "Telegrau4"
    Case RAL_7048_Perlmausgrau:     S = "Perlmausgrau"
    
    Case RAL_8000_Grünbraun:        S = "Grünbraun"
    Case RAL_8001_Ockerbraun:       S = "Ockerbraun"
    Case RAL_8002_Signalbraun:      S = "Signalbraun"
    Case RAL_8003_Lehmbraun:        S = "Lehmbraun"
    Case RAL_8004_Kupferbraun:      S = "Kupferbraun"
    Case RAL_8007_Rehbraun:         S = "Rehbraun"
    Case RAL_8008_Olivbraun:        S = "Olivbraun"
    Case RAL_8011_Nussbraun:        S = "Nussbraun"
    Case RAL_8012_Rotbraun:         S = "Rotbraun"
    Case RAL_8014_Sepiabraun:       S = "Sepiabraun"
    Case RAL_8015_Kastanienbraun:   S = "Kastanienbraun"
    Case RAL_8016_Mahagonibraun:    S = "Mahagonibraun"
    Case RAL_8017_Schokoladenbraun: S = "Schokoladenbraun"
    Case RAL_8019_Graubraun:        S = "Graubraun"
    Case RAL_8022_Schwarzbraun:     S = "Schwarzbraun"
    Case RAL_8023_Orangebraun:      S = "Orangebraun"
    Case RAL_8024_Beigebraun:       S = "Beigebraun"
    Case RAL_8025_Blassbraun:       S = "Blassbraun"
    Case RAL_8028_Terrabraun:       S = "Terrabraun"
    Case RAL_8029_Perlkupfer:       S = "Perlkupfer"
    
    Case RAL_9001_Cremeweiß:        S = "Cremeweiß"
    Case RAL_9002_Grauweiß:         S = "Grauweiß"
    Case RAL_9003_Signalweiß:       S = "Signalweiß"
    Case RAL_9004_Signalschwarz:    S = "Signalschwarz"
    Case RAL_9005_Tiefschwarz:      S = "Tiefschwarz"
    Case RAL_9006_Weißaluminium:    S = "Weißaluminium"
    Case RAL_9007_Graualuminium:    S = "Graualuminium"
    Case RAL_9010_Reinweiß:         S = "Reinweiß"
    Case RAL_9011_Graphitschwarz:   S = "Graphitschwarz"
    Case RAL_9012_Reinraumweiß:     S = "Reinraumweiß"
    Case RAL_9016_Verkehrsweiß:     S = "Verkehrsweiß"
    Case RAL_9017_Verkehrsschwarz:  S = "Verkehrsschwarz"
    Case RAL_9018_Papyrusweiß:      S = "Papyrusweiß"
    Case RAL_9022_Perlhellgrau:     S = "Perlhellgrau"
    Case RAL_9023_Perldunkelgrau:   S = "Perldunkelgrau"
    Case RAL_9020_SeidenmattWeiß:   S = "SeidenmattWeiß"
    
    Case RAL_6031_Bronzegrün:       S = "Bronzegrün"
    Case RAL_8027_Lederbraun:       S = "Lederbraun"
    Case RAL_9021_Teerschwarz:      S = "Teerschwarz"
    Case RAL_1039_Sandbeige:        S = "Sandbeige"
    Case RAL_1040_Lehmbeige:        S = "Lehmbeige"
    Case RAL_6040_Helloliv:         S = "Helloliv"
    Case RAL_7050_Tarngrau:         S = "Tarngrau"
    Case RAL_8031_Sandbraun:        S = "Sandbraun"
    
    End Select
    RALClassic_NameToStr = S
End Function

Public Function RALClassic_ToNum(e As RALClassic) As Long
    If m_Count = 0 Then RALClassicColor_Init
    Dim n As Long
    Select Case e
    
    Case RAL_1000_Grünbeige:        n = 1000
    Case RAL_1001_Beige:            n = 1001
    Case RAL_1002_Sandgelb:         n = 1002
    Case RAL_1003_Signalgelb:       n = 1003
    Case RAL_1004_Goldgelb:         n = 1004
    Case RAL_1005_Honiggelb:        n = 1005
    Case RAL_1006_Maisgelb:         n = 1006
    Case RAL_1007_Narzissengelb:    n = 1007
    Case RAL_1011_Braunbeige:       n = 1011
    Case RAL_1012_Zitronengelb:     n = 1012
    Case RAL_1013_Perlweiß:         n = 1013
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
    Case RAL_3017_Rosé:             n = 3017
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
    Case RAL_5001_Grünblau:         n = 5001
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
    Case RAL_5018_Türkisblau:       n = 5018
    Case RAL_5019_Capriblau:        n = 5019
    Case RAL_5020_Ozeanblau:        n = 5020
    Case RAL_5021_Wasserblau:       n = 5021
    Case RAL_5022_Nachtblau:        n = 5022
    Case RAL_5023_Fernblau:         n = 5023
    Case RAL_5024_Pastellblau:      n = 5024
    Case RAL_5025_Perlenzian:       n = 5025
    Case RAL_5026_Perlnachtblau:    n = 5026
    
    Case RAL_6000_Patinagrün:       n = 6000
    Case RAL_6001_Smaragdgrün:      n = 6001
    Case RAL_6002_Laubgrün:         n = 6002
    Case RAL_6003_Olivgrün:         n = 6003
    Case RAL_6004_Blaugrün:         n = 6004
    Case RAL_6005_Moosgrün:         n = 6005
    Case RAL_6006_Grauoliv:         n = 6006
    Case RAL_6007_Flaschengrün:     n = 6007
    Case RAL_6008_Braungrün:        n = 6008
    Case RAL_6009_Tannengrün:       n = 6009
    Case RAL_6010_Grasgrün:         n = 6010
    Case RAL_6011_Resedagrün:       n = 6011
    Case RAL_6012_Schwarzgrün:      n = 6012
    Case RAL_6013_Schilfgrün:       n = 6013
    Case RAL_6014_Gelboliv:         n = 6014
    Case RAL_6015_Schwarzoliv:      n = 6015
    Case RAL_6016_Türkisgrün:       n = 6016
    Case RAL_6017_Maigrün:          n = 6017
    Case RAL_6018_Gelbgrün:         n = 6018
    Case RAL_6019_Weißgrün:         n = 6019
    Case RAL_6020_Chromoxidgrün:    n = 6020
    Case RAL_6021_Blassgrün:        n = 6021
    Case RAL_6022_Braunoliv:        n = 6022
    Case RAL_6024_Verkehrsgrün:     n = 6024
    Case RAL_6025_Farngrün:         n = 6025
    Case RAL_6026_Opalgrün:         n = 6026
    Case RAL_6027_Lichtgrün:        n = 6027
    Case RAL_6028_Kieferngrün:      n = 6028
    Case RAL_6029_Minzgrün:         n = 6029
    Case RAL_6032_Signalgrün:       n = 6032
    Case RAL_6033_Minttürkis:       n = 6033
    Case RAL_6034_Pastelltürkis:    n = 6034
    Case RAL_6035_Perlgrün:         n = 6035
    Case RAL_6036_Perlopalgrün:     n = 6036
    Case RAL_6037_Reingrün:         n = 6037
    Case RAL_6038_Leuchtgrün:       n = 6038
    Case RAL_6039_Fasergrün:        n = 6039
    
    Case RAL_7000_Fehgrau:          n = 7000
    Case RAL_7001_Silbergrau:       n = 7001
    Case RAL_7002_Olivgrau:         n = 7002
    Case RAL_7003_Moosgrau:         n = 7003
    Case RAL_7004_Signalgrau:       n = 7004
    Case RAL_7005_Mausgrau:         n = 7005
    Case RAL_7006_Beigegrau:        n = 7006
    Case RAL_7008_Khakigrau:        n = 7008
    Case RAL_7009_Grüngrau:         n = 7009
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
    
    Case RAL_8000_Grünbraun:        n = 8000
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
    
    Case RAL_9001_Cremeweiß:        n = 9001
    Case RAL_9002_Grauweiß:         n = 9002
    Case RAL_9003_Signalweiß:       n = 9003
    Case RAL_9004_Signalschwarz:    n = 9004
    Case RAL_9005_Tiefschwarz:      n = 9005
    Case RAL_9006_Weißaluminium:    n = 9006
    Case RAL_9007_Graualuminium:    n = 9007
    Case RAL_9010_Reinweiß:         n = 9010
    Case RAL_9011_Graphitschwarz:   n = 9011
    Case RAL_9012_Reinraumweiß:     n = 9012
    Case RAL_9016_Verkehrsweiß:     n = 9016
    Case RAL_9017_Verkehrsschwarz:  n = 9017
    Case RAL_9018_Papyrusweiß:      n = 9018
    Case RAL_9022_Perlhellgrau:     n = 9022
    Case RAL_9023_Perldunkelgrau:   n = 9023
    Case RAL_9020_SeidenmattWeiß:   n = 9020
    
    Case RAL_6031_Bronzegrün:       n = 6031
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
    
    Case 1000: c = RAL_1000_Grünbeige
    Case 1001: c = RAL_1001_Beige
    Case 1002: c = RAL_1002_Sandgelb
    Case 1003: c = RAL_1003_Signalgelb
    Case 1004: c = RAL_1004_Goldgelb
    Case 1005: c = RAL_1005_Honiggelb
    Case 1006: c = RAL_1006_Maisgelb
    Case 1007: c = RAL_1007_Narzissengelb
    Case 1011: c = RAL_1011_Braunbeige
    Case 1012: c = RAL_1012_Zitronengelb
    Case 1013: c = RAL_1013_Perlweiß
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
    Case 3017: c = RAL_3017_Rosé
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
    Case 5001: c = RAL_5001_Grünblau
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
    Case 5018: c = RAL_5018_Türkisblau
    Case 5019: c = RAL_5019_Capriblau
    Case 5020: c = RAL_5020_Ozeanblau
    Case 5021: c = RAL_5021_Wasserblau
    Case 5022: c = RAL_5022_Nachtblau
    Case 5023: c = RAL_5023_Fernblau
    Case 5024: c = RAL_5024_Pastellblau
    Case 5025: c = RAL_5025_Perlenzian
    Case 5026: c = RAL_5026_Perlnachtblau
    
    Case 6000: c = RAL_6000_Patinagrün
    Case 6001: c = RAL_6001_Smaragdgrün
    Case 6002: c = RAL_6002_Laubgrün
    Case 6003: c = RAL_6003_Olivgrün
    Case 6004: c = RAL_6004_Blaugrün
    Case 6005: c = RAL_6005_Moosgrün
    Case 6006: c = RAL_6006_Grauoliv
    Case 6007: c = RAL_6007_Flaschengrün
    Case 6008: c = RAL_6008_Braungrün
    Case 6009: c = RAL_6009_Tannengrün
    Case 6010: c = RAL_6010_Grasgrün
    Case 6011: c = RAL_6011_Resedagrün
    Case 6012: c = RAL_6012_Schwarzgrün
    Case 6013: c = RAL_6013_Schilfgrün
    Case 6014: c = RAL_6014_Gelboliv
    Case 6015: c = RAL_6015_Schwarzoliv
    Case 6016: c = RAL_6016_Türkisgrün
    Case 6017: c = RAL_6017_Maigrün
    Case 6018: c = RAL_6018_Gelbgrün
    Case 6019: c = RAL_6019_Weißgrün
    Case 6020: c = RAL_6020_Chromoxidgrün
    Case 6021: c = RAL_6021_Blassgrün
    Case 6022: c = RAL_6022_Braunoliv
    Case 6024: c = RAL_6024_Verkehrsgrün
    Case 6025: c = RAL_6025_Farngrün
    Case 6026: c = RAL_6026_Opalgrün
    Case 6027: c = RAL_6027_Lichtgrün
    Case 6028: c = RAL_6028_Kieferngrün
    Case 6029: c = RAL_6029_Minzgrün
    Case 6032: c = RAL_6032_Signalgrün
    Case 6033: c = RAL_6033_Minttürkis
    Case 6034: c = RAL_6034_Pastelltürkis
    Case 6035: c = RAL_6035_Perlgrün
    Case 6036: c = RAL_6036_Perlopalgrün
    Case 6037: c = RAL_6037_Reingrün
    Case 6038: c = RAL_6038_Leuchtgrün
    Case 6039: c = RAL_6039_Fasergrün
    
    Case 7000: c = RAL_7000_Fehgrau
    Case 7001: c = RAL_7001_Silbergrau
    Case 7002: c = RAL_7002_Olivgrau
    Case 7003: c = RAL_7003_Moosgrau
    Case 7004: c = RAL_7004_Signalgrau
    Case 7005: c = RAL_7005_Mausgrau
    Case 7006: c = RAL_7006_Beigegrau
    Case 7008: c = RAL_7008_Khakigrau
    Case 7009: c = RAL_7009_Grüngrau
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
    
    Case 8000: c = RAL_8000_Grünbraun
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
    
    Case 9001: c = RAL_9001_Cremeweiß
    Case 9002: c = RAL_9002_Grauweiß
    Case 9003: c = RAL_9003_Signalweiß
    Case 9004: c = RAL_9004_Signalschwarz
    Case 9005: c = RAL_9005_Tiefschwarz
    Case 9006: c = RAL_9006_Weißaluminium
    Case 9007: c = RAL_9007_Graualuminium
    Case 9010: c = RAL_9010_Reinweiß
    Case 9011: c = RAL_9011_Graphitschwarz
    Case 9012: c = RAL_9012_Reinraumweiß
    Case 9016: c = RAL_9016_Verkehrsweiß
    Case 9017: c = RAL_9017_Verkehrsschwarz
    Case 9018: c = RAL_9018_Papyrusweiß
    Case 9022: c = RAL_9022_Perlhellgrau
    Case 9023: c = RAL_9023_Perldunkelgrau
    Case 9020: c = RAL_9020_SeidenmattWeiß
    
    End Select
    RALClassic_NumToColor = c
End Function

Public Function RALClassic_NumToColorname(num As Long) As String
    If m_Count = 0 Then RALClassicColor_Init
    Dim S As String: S = RALClassic_NameToStr(RALClassic_NumToColor(num))
    If Len(S) = 0 Then Exit Function
    RALClassic_NumToColorname = "RAL_" & CStr(num) & "_" & S
End Function

Public Sub RALClassic_ToListBox(aCBLB)
    If m_Count = 0 Then RALClassicColor_Init
    aCBLB.Clear
    Dim i As Long, j As Long
    Dim num As Long, S As String
    For i = 1 To 10
        num = i * 1000
        For j = 0 To 100
            S = RALClassic_NumToColorname(num)
            If Len(S) Then aCBLB.AddItem S
            num = num + 1
        Next
    Next
End Sub

Public Function RALClassic_Parse(ByVal S As String) As RALClassic
    If m_Count = 0 Then RALClassicColor_Init
    'bsp: RAL_9001_Cremeweiß
    'read the number
    If Left(S, 4) = "RAL_" Then
        S = Mid(S, 5)
        Dim pos As Long: pos = InStr(1, S, "_")
        If pos > 0 Then S = Left(S, pos - 1)
        If Not IsNumeric(S) Then S = Left(S, 4)
        If Not IsNumeric(S) Then Exit Function
        Dim num As Long: num = CLng(S)
        RALClassic_Parse = RALClassic_NumToColor(num)
    End If
End Function

