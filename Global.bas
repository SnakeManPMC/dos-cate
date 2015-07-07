Attribute VB_Name = "Global"
Option Explicit

Public TerrainSize As Long          'Usually 64 or 128
Public Regions() As Long            'Regions (segments) array
Public UpdatedCities As Long        'Cities are a special case ...

Public Type strTileDesc             'Describe a tile
    NumTile As Integer
    Altitude As Integer
    Fog As Byte
    Unknow1 As Byte
    Unknow2 As Byte
End Type
Public Tiles() As strTileDesc       'Tiles array, when working on L2 data only

Public Type strFeatTerrain
    TileDesc As strTileDesc         'A tile
    River(1 To 4) As Byte           'Is there any river (1=N, 2=E, 3=S, 4=W)
    Road(1 To 4) As Byte            'Is there any road (1=N, 2=E, 3=S, 4=W)
    City As Byte                    'Is there a city
End Type
Public Terrain() As strFeatTerrain     'Tiles matrix when working on a terrain grid

Public Type strTileRules            'Describe a TileRule rule
    MinAlt As Long
    MaxAlt As Long
    OrigTileList() As String
    TileList() As String
    Fog As String
    Unknown1 As String
    Unknown2 As String
    AltModType As String
    AltModValue As Long
    AltModValueType As String
End Type

Public Type strFogRules             'Describes a ForceFog rule
    Fog As String
    Unknown1 As String
    Unknown2 As String
End Type

Public Type strAltRules             'Describes a RandAlt rule
    MinAlt As Long
    MaxAlt As Long
    Min As Integer
    Max As Integer
End Type

Public Type strRuleSection          'Describes a rules section
    Name As String
    R As Integer
    G As Integer
    B As Integer
    NumTileRules As Integer
    NumFogRules As Integer
    NumAltRules As Integer
    TileRules() As strTileRules
    FogRules() As strFogRules
    AltRules() As strAltRules
End Type
Public Type strRules                'Describes the rules file and rules section
    UpdateOceanTiles As Integer
    BmpFileName As String
    NumSections As Integer
    RuleSections() As strRuleSection
End Type
Public Rules As strRules            'All the rules

Public Type strHeader               'Describes the header of the BMP file
    Signature As String
    FileSize As Long
    DataOffset As Long
    HeaderSize As Long
    ImageWidth As Long
    ImageHeight As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    HorResolution As Long
    VerResolution As Long
    ColorUsed As Long
    ColorsImportant As Long
End Type
Public ImageHeader As strHeader
Public SaveImageHeader As strHeader

Public Type strColorTable           'Describe the color table of the BMP file
    R As Integer
    G As Integer
    B As Integer
End Type
Public ImageColors(0 To 255) As strColorTable
Public SaveImageColors(0 To 255) As strColorTable
Public CurrentImageColor As Integer

Public ImageDataTmp() As Byte
Public ImageData() As Byte          'BMP data (each pixel is given a reference to the color table)

Public Type strDefAirbase   'For AirBaseDef section in CATE conf file
    TypeAB() As String      'List of AB types associated to that CATE type
    XStart As Integer       'First S-W tile
    YStart As Integer
    XEnd As Integer         'Last N-E tile
    YEnd As Integer
End Type
Public DefAirbases() As strDefAirbase

Public Type strAllAirbases   'For defining airbases found in CSV file
    Name As String
    TypeAirbase As String   'Type as defined in CSV
    TypeCate As Integer     'Type as defined in CATE
    Id As Long
    x As Long               'Position
    y As Long               'Position
    z As Long               'Altitude
End Type
Public Type strAirbase      'For defining airbases found in conf file for each terrain
    TileList() As Integer   'Tiles for the AB
    Level As Integer        'Leveling type (0=none, 1=high, 2=low)
End Type

Public Type strCoord
    x As Long
    y As Long
End Type
Public Type strPath         'All the waypoints in a path
    NumCoords As Long
    Coord() As strCoord
End Type

Public Type strTransition
    Name As String
    NumTiles As Integer                 'Num tiles in tile list
    Type As Byte                        '1=standard, 2=reverse
    TileList() As Integer               'List of tiles towards which we do the translation
    TransTiles(1 To 15) As Integer      'Tile numbers used for the transition
End Type
Public Type strFeatures
    TerrainType As Integer                  'CATE terrain type
    CityBaseTile As Integer                 'Base city tile number for the terrain
    AirBases() As strAirbase                'Airbases def for this terrain
    NumTransitions As Integer
    Transitions() As strTransition          'Transitions for this terrain
End Type

Public Type strAutoFeatures
    FeatName() As String
    NumFiles As Integer         'Nb TDF files for features
    FileNames() As String       'TDF files for features
    TextureFileName As String
    NumTrnFiles As Integer
    CorrespFileName As String
    TrnFileNames() As String
    NumABFiles As Integer       'Nb CSV files for airbases
    FileABNames() As String     'CSV files for airbases
    NumRivers As Long
    RiverPaths() As strPath     'Paths found in TDF file
    NumRoads As Long
    RoadPaths() As strPath      'Paths found in TDF file
    NumCities As Long
    CityPaths() As strPath      'Paths found in TDF file
    NumAirbases As Long
    AllAirbases() As strAllAirbases
    NumTerrain As Integer
    TerrainFeatures() As strFeatures
    TRNOffsetX As Integer
    TRNOffsetY As Integer
    RiverRoadMethod As Integer
End Type
Public AutoFeatures As strAutoFeatures

Public TileToTerrain(0 To 5000) As Integer  'Each index in this array is a tile, and its value will give the associated CATE terrain
Public FeatureTiles(0 To 5000, 0 To 255) As Integer 'Gives, for a tile number, and a feature number, the new tile number to use
Public Type strTileToFeature
    Value As Integer
    River(1 To 4) As Integer
    Road(1 To 4) As Integer
End Type
Public TileToFeature(0 To 5000) As strTileToFeature 'More or less the reverse of above : gives, for a tile number, what features it has

'Used when saving a BMP of the theater
Public Type strBmpToSave
    FileName As String
    TileNum(0 To 5000) As Byte
End Type
Public BmpToSave As strBmpToSave

'The following definitions are used when using Fly terrain def files
Public TileNumToName() As String
Public Type strSuperTile
    Tile(0 To 15) As String
End Type
Public Type strGlobeTile
    st(0 To 7, 0 To 7) As strSuperTile
End Type
Public gt(0 To 9, 0 To 9) As strGlobeTile

Public Type strNameCorresp
    NumNames As String
    OrigName() As String
    DestName() As String
End Type
Public NameCorresp As strNameCorresp

Function GivePathFromName(FileName As String) As String
'Returns the path (folder) of a full file name (by searching the last \)
Dim i As Integer
    
    For i = Len(FileName) To 1 Step -1
        If Mid$(FileName, i, 1) = "\" Then Exit For
    Next i
    GivePathFromName = Left$(FileName, i)
End Function

Function GiveNameFromPath(FileName As String) As String
'Returns the name of a file from a full path (by searching the last \)
Dim i As Integer
    
    For i = Len(FileName) To 1 Step -1
        If Mid$(FileName, i, 1) = "\" Then Exit For
    Next i
    GiveNameFromPath = Right$(FileName, Len(FileName) - i)
End Function

Sub UpdatePercentBar(pic As PictureBox, Perc As Long)
'Draws on a picture box to simulate a standard percentage bar

    pic.Line (-1, -1)-(pic.Width * Perc \ 100, pic.Height), , B
End Sub

Sub LoadAltRule(Chaine As String, ByRef AltRule As strAltRules)
'Load a alt rule from a string (usually read in the conf file)
'Put here to avoid overloading the LoadRules sub in frmCate
'Needs to be in Global.bas because a user-defined type is used as parameter
Dim i As Integer, Found As Integer
Dim tmpstring  As String


    tmpstring = ""
    Found = 0
    i = 1
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            Found = Found + 1
            Select Case Found
                Case 1:
                    AltRule.MinAlt = Val(tmpstring)
                Case 2:
                    AltRule.MaxAlt = Val(tmpstring)
                Case 3:
                    AltRule.Min = Val(tmpstring)
                Case 4:
                    tmpstring = tmpstring & Right$(Chaine, 1)
                    AltRule.Max = Val(tmpstring)
                Case Else:
                    'Impossible : bad format
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop

End Sub

Sub LoadFogRule(Chaine As String, ByRef FogRule As strFogRules)
'Load a fog rule from a string (usually read in the conf file)
'Put here to avoid overloading the LoadRules sub in frmCate
'Needs to be in Global.bas because a user-defined type is used as parameter
Dim i As Integer, Found As Integer
Dim tmpstring  As String


    tmpstring = ""
    Found = 0
    i = 1
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            Found = Found + 1
            Select Case Found
                Case 1:
                    FogRule.Fog = tmpstring
                Case 2:
                    FogRule.Unknown1 = tmpstring
                Case 3:
                    tmpstring = tmpstring & Right$(Chaine, 1)
                    FogRule.Unknown2 = tmpstring
                Case Else:
                    'Impossible : bad format
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop

End Sub

Sub LoadTileRule(Chaine As String, ByRef TileRule As strTileRules)
'Load a tile rule from a string (usually read in the conf file)
'Put here to avoid overloading the LoadRules sub in frmCate
'Needs to be in Global.bas because a user-defined type is used as parameter
Dim i As Integer, j As Integer, k As Integer, l As Integer, posnext As Integer, Found As Integer
Dim tmpstring As String, tmpstring2 As String, tmpstring3 As String

    'Now we search for spaces
    tmpstring = ""
    Found = 0
    i = 1
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            Found = Found + 1
            Select Case Found
                Case 1:
                    TileRule.MinAlt = Val(Trim$(tmpstring))
                Case 2:
                    TileRule.MaxAlt = Val(Trim$(tmpstring))
                Case 3:
                    tmpstring2 = ""
                    k = 0
                    For j = 2 To Len(Trim$(tmpstring))
                        If Mid$(tmpstring, j, 1) = "-" Or Mid$(tmpstring, j, 1) = "," Or Mid$(tmpstring, j, 1) = ")" Then
                            If Mid$(tmpstring, j, 1) = "-" Then
                                posnext = InStr(j + 1, tmpstring, ",")
                                If posnext <= 0 Then
                                    posnext = InStr(j + 1, tmpstring, ")")
                                End If
                                tmpstring3 = Mid$(tmpstring, j + 1, posnext - j)
                                For l = Val(tmpstring2) To Val(tmpstring3)
                                    k = k + 1
                                    If k = 1 Then
                                        ReDim TileRule.OrigTileList(1 To 1)
                                    Else
                                        ReDim Preserve TileRule.OrigTileList(1 To k)
                                    End If
                                    TileRule.OrigTileList(k) = "" & l
                                Next l
                                j = posnext
                            Else
                                
                                k = k + 1
                                If k = 1 Then
                                    ReDim TileRule.OrigTileList(1 To 1)
                                Else
                                    ReDim Preserve TileRule.OrigTileList(1 To k)
                                End If
                                TileRule.OrigTileList(k) = tmpstring2
                            End If
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
                Case 4:
                    tmpstring2 = ""
                    k = 0
                    For j = 2 To Len(Trim$(tmpstring))
                        If Mid$(tmpstring, j, 1) = "-" Or Mid$(tmpstring, j, 1) = "," Or Mid$(tmpstring, j, 1) = ")" Then
                            If Mid$(tmpstring, j, 1) = "-" Then
                                posnext = InStr(j + 1, tmpstring, ",")
                                If posnext <= 0 Then
                                    posnext = InStr(j + 1, tmpstring, ")")
                                End If
                                tmpstring3 = Mid$(tmpstring, j + 1, posnext - j)
                                For l = Val(tmpstring2) To Val(tmpstring3)
                                    k = k + 1
                                    If k = 1 Then
                                        ReDim TileRule.TileList(1 To 1)
                                    Else
                                        ReDim Preserve TileRule.TileList(1 To k)
                                    End If
                                    TileRule.TileList(k) = "" & l
                                Next l
                                j = posnext
                            Else
                                k = k + 1
                                If k = 1 Then
                                    ReDim TileRule.TileList(1 To 1)
                                Else
                                    ReDim Preserve TileRule.TileList(1 To k)
                                End If
                                TileRule.TileList(k) = tmpstring2
                            End If
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
                Case 5:
                    TileRule.Fog = tmpstring
                Case 6:
                    TileRule.Unknown1 = tmpstring
                Case 7:
                    TileRule.Unknown2 = tmpstring
                Case 8:
                    tmpstring2 = tmpstring & Right$(Chaine, 1)
                    If tmpstring2 = "*" Then
                        TileRule.AltModType = "*"
                    Else
                        TileRule.AltModType = UCase$(Left$(tmpstring2, 1))
                        If TileRule.AltModType = "F" Then
                            TileRule.AltModValue = Val(Mid$(tmpstring2, 2, Len(tmpstring2) - 1))
                            TileRule.AltModValueType = "F"
                        Else
                            TileRule.AltModValue = Val(Mid$(tmpstring2, 2, Len(tmpstring2) - 2))
                            TileRule.AltModValueType = UCase$(Right$(tmpstring2, 1))
                        End If
                    End If
                Case Else:
                    'Impossible : bad format
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop
End Sub

Function ApplyRules(ByRef Tile As strTileDesc, Sec As Integer) As Boolean
'Apply all the known rules to a tile
'Put here to avoid overloading frmCate and because it's called in two functions (applycaterules and applycolorrules)
'Needs to be in Global.bas because a user-defined type is used as parameter
Dim j As Integer, k As Integer
Dim IsUpdated As Boolean
Dim num1 As Integer, num2 As Integer, dec As Integer, random As Integer
Dim ApplyRuleOnTile As Boolean

    IsUpdated = False
    
    If Rules.RuleSections(Sec).NumFogRules > 0 Then
        For j = 1 To Rules.RuleSections(Sec).NumFogRules 'Apply fog rules to all tiles
            If Rules.RuleSections(Sec).FogRules(j).Fog <> "*" Then Tile.Fog = Val(Rules.RuleSections(Sec).FogRules(j).Fog)
            If Rules.RuleSections(Sec).FogRules(j).Unknown1 <> "*" Then Tile.Unknow1 = Val(Rules.RuleSections(Sec).FogRules(j).Unknown1)
            If Rules.RuleSections(Sec).FogRules(j).Unknown2 <> "*" Then Tile.Unknow2 = Val(Rules.RuleSections(Sec).FogRules(j).Unknown2)
        Next j
    End If
    
    If Rules.RuleSections(Sec).NumTileRules > 0 Then
        For j = 1 To Rules.RuleSections(Sec).NumTileRules   'For each tile, read all rules
            If Tile.Altitude >= Rules.RuleSections(Sec).TileRules(j).MinAlt And Tile.Altitude <= Rules.RuleSections(Sec).TileRules(j).MaxAlt Then
                'If altitude is within parameters of rule
                If Rules.RuleSections(Sec).TileRules(j).OrigTileList(1) = "*" Then
                    'Rule : any, or any except list
                    If UBound(Rules.RuleSections(Sec).TileRules(j).OrigTileList) > 1 Then
                        'Rule : any except ...
                        ApplyRuleOnTile = True
                        For k = 2 To UBound(Rules.RuleSections(Sec).TileRules(j).OrigTileList)
                            If Tile.NumTile = Val(Rules.RuleSections(Sec).TileRules(j).OrigTileList(k)) Then
                                'Current tile found on except list
                                ApplyRuleOnTile = False
                                Exit For
                            End If
                        Next k
                    Else
                        'Rule : any
                        ApplyRuleOnTile = True
                    End If
                Else
                    'Rule : tile list
                    ApplyRuleOnTile = False
                    For k = 1 To UBound(Rules.RuleSections(Sec).TileRules(j).OrigTileList)
                        If Tile.NumTile = Val(Rules.RuleSections(Sec).TileRules(j).OrigTileList(k)) Then
                            'Current tile found on accept list
                            ApplyRuleOnTile = True
                            Exit For
                        End If
                    Next k
                End If
                If ApplyRuleOnTile Then
                    'We choose tile at random from given list
                    If Rules.RuleSections(Sec).TileRules(j).TileList(1) <> "*" Then
                        Tile.NumTile = Val(Rules.RuleSections(Sec).TileRules(j).TileList(Int(UBound(Rules.RuleSections(Sec).TileRules(j).TileList) * Rnd + 1)))
                    End If
                    If Rules.RuleSections(Sec).TileRules(j).Fog <> "*" Then Tile.Fog = Val(Rules.RuleSections(Sec).TileRules(j).Fog)
                    If Rules.RuleSections(Sec).TileRules(j).Unknown1 <> "*" Then Tile.Unknow1 = Val(Rules.RuleSections(Sec).TileRules(j).Unknown1)
                    If Rules.RuleSections(Sec).TileRules(j).Unknown2 <> "*" Then Tile.Unknow2 = Val(Rules.RuleSections(Sec).TileRules(j).Unknown2)
                    'Modify altitude if needed
                    If Rules.RuleSections(Sec).TileRules(j).AltModType = "F" Then
                        'Force altitude value
                        Tile.Altitude = Rules.RuleSections(Sec).TileRules(j).AltModValue
                    ElseIf Rules.RuleSections(Sec).TileRules(j).AltModType = "M" Then
                        If Rules.RuleSections(Sec).TileRules(j).AltModValueType = "%" Then
                            'Percent increase
                            Tile.Altitude = Tile.Altitude + (Rules.RuleSections(Sec).TileRules(j).AltModValue * Tile.Altitude) \ 100
                        Else
                            'Feet increase
                            Tile.Altitude = Tile.Altitude + Rules.RuleSections(Sec).TileRules(j).AltModValue
                            If Rules.RuleSections(Sec).TileRules(j).AltModValueType = "G" Then
                                If Tile.Altitude < 1 Then Tile.Altitude = 1
                            End If
                        End If
                    End If
                End If
                IsUpdated = True
            End If
        Next j
    End If
    
    If Rules.RuleSections(Sec).NumAltRules > 0 Then
        For j = 1 To Rules.RuleSections(Sec).NumAltRules 'Apply tile rules
            If Tile.Altitude >= Rules.RuleSections(Sec).AltRules(j).MinAlt And Tile.Altitude <= Rules.RuleSections(Sec).AltRules(j).MaxAlt Then
                If Rules.RuleSections(Sec).AltRules(j).Min < 0 Then 'We can't randomize negative number so a little trick is welcome
                    num1 = 0
                    num2 = Rules.RuleSections(Sec).AltRules(j).Max - Rules.RuleSections(Sec).AltRules(j).Min
                    dec = Rules.RuleSections(Sec).AltRules(j).Min * -1
                Else
                    num1 = Rules.RuleSections(Sec).AltRules(j).Min
                    num2 = Rules.RuleSections(Sec).AltRules(j).Max
                    dec = 0
                End If
                random = Int((num2 - num1 + 1) * Rnd + num1) - dec
                Tile.Altitude = Int(Tile.Altitude + Tile.Altitude / 100 * random)
                IsUpdated = True
            End If
        Next j
    End If
    If Tile.Altitude < 0 Then Tile.Altitude = 0
    ApplyRules = IsUpdated
End Function

Sub LoadTerrainTiles(Chaine As String, TerrainType As Integer)
'Will load the TileToTerrain array
Dim i As Integer, j As Integer, l As Integer, posnext As Integer
Dim pos As Integer
Dim tmpstring As String, tmpstring2 As String, tmpstring3 As String
Dim Found As Integer

    i = 1
    Chaine = Trim$(Mid$(Chaine, 2, Len(Chaine) - 2))
    tmpstring = ""
    Found = 0
    Do
        If Mid$(Chaine, i, 1) = "-" Then
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            Found = Found + 1
            posnext = InStr(i + 1, Chaine, ",")
            If posnext <= 0 Then posnext = Len(Chaine) + 1
            tmpstring2 = Mid$(Chaine, i + 1, posnext - i - 1)
            For l = Val(tmpstring) To Val(tmpstring2)
                TileToTerrain(l) = TerrainType
            Next l
            i = posnext
            tmpstring = ""
        ElseIf Mid$(Chaine, i, 1) = "," Or i = Len(Chaine) Then
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            TileToTerrain(Val(tmpstring)) = TerrainType
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        If i > Len(Chaine) Then Exit Do
        i = i + 1
    Loop
End Sub

Sub LoadTransition(Chaine As String, CurrentTerrain As Integer)
'Will read and load the TransitionDef for Transitions
Dim i As Integer, j As Integer, k As Integer, l As Integer, posnext As Integer, Found As Integer
Dim tmpstring As String, tmpstring2 As String, tmpstring3 As String
Dim CurrentTransition As Integer

    CurrentTransition = AutoFeatures.TerrainFeatures(CurrentTerrain).NumTransitions
    'Now we search for spaces
    tmpstring = ""
    Found = 0
    i = 1
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            Found = Found + 1
            Select Case Found
                Case 1:
                    AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).Name = Trim$(tmpstring)
                Case 2:
                    AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).Type = Val(Trim$(tmpstring))
                Case 3:
                    tmpstring2 = ""
                    k = 0
                    For j = 2 To Len(Trim$(tmpstring))
                        If Mid$(tmpstring, j, 1) = "-" Or Mid$(tmpstring, j, 1) = "," Or Mid$(tmpstring, j, 1) = ")" Then
                            If Mid$(tmpstring, j, 1) = "-" Then
                                posnext = InStr(j + 1, tmpstring, ",")
                                If posnext <= 0 Then
                                    posnext = InStr(j + 1, tmpstring, ")")
                                End If
                                tmpstring3 = Mid$(tmpstring, j + 1, posnext - j)
                                For l = Val(tmpstring2) To Val(tmpstring3)
                                    k = k + 1
                                    If k = 1 Then
                                        ReDim AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(1 To 1)
                                    Else
                                        ReDim Preserve AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(1 To k)
                                    End If
                                    AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).NumTiles = k
                                    AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(k) = l
                                Next l
                                j = posnext
                            Else
                                k = k + 1
                                If k = 1 Then
                                    ReDim AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(1 To 1)
                                Else
                                    ReDim Preserve AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(1 To k)
                                End If
                                AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).NumTiles = k
                                AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TileList(k) = Val(tmpstring2)
                            End If
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
                Case 4:
                    tmpstring2 = ""
                    k = 0
                    For j = 2 To Len(Trim$(tmpstring))
                        If Mid$(tmpstring, j, 1) = "," Or Mid$(tmpstring, j, 1) = ")" Or j = Len(Trim$(tmpstring)) Then
                            k = k + 1
                            If j = Len(Trim$(tmpstring)) Then
                                tmpstring2 = tmpstring2 & Right$(Trim$(tmpstring), 1)
                            End If
                            AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(CurrentTransition).TransTiles(k) = Val(tmpstring2)
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
                Case Else:
                    'Nothing
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop
End Sub

Function GiveTerrainFromTile(NumTile As Integer) As Integer
'Will return CATE terrain type to use for a given tile
Dim i As Integer

    If AutoFeatures.NumTerrain = 0 Then
        GiveTerrainFromTile = 0
        Exit Function
    End If
    For i = 1 To AutoFeatures.NumTerrain
        If TileToTerrain(NumTile) = AutoFeatures.TerrainFeatures(i).TerrainType Then
            GiveTerrainFromTile = i
            Exit Function
        End If
    Next i
    GiveTerrainFromTile = 0
End Function

Sub LoadABDef(Chaine As String, ABDef As strDefAirbase)
'Will load general AB defnitions of CATE conf file
Dim i As Integer, j As Integer
Dim pos As Integer
Dim tmpstring As String, tmpstring2 As String
Dim Found As Integer
Dim NBType As Integer
    i = 1
    Chaine = Trim$(Chaine)
    tmpstring = ""
    Found = 0
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            Found = Found + 1
            Select Case Found
                Case 1:
                    If Val(tmpstring) <> Found Then
                        'It's not normal !
                    End If
                Case 2:
                    ABDef.XStart = Val(tmpstring)
                Case 3:
                    ABDef.YStart = Val(tmpstring)
                Case 4:
                    ABDef.XEnd = Val(tmpstring)
                Case 5:
                    ABDef.YEnd = Val(tmpstring)
                Case 6:
                    tmpstring2 = ""
                    NBType = 0
                    For j = 2 To Len(tmpstring) - 1
                        If Mid$(tmpstring, j, 1) = "," Or j = Len(tmpstring) - 1 Then
                            If j = Len(tmpstring) - 1 Then tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                            NBType = NBType + 1
                            If NBType = 1 Then
                                ReDim ABDef.TypeAB(1 To 1)
                            Else
                                ReDim Preserve ABDef.TypeAB(1 To NBType)
                            End If
                            ABDef.TypeAB(NBType) = tmpstring2
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop
End Sub

Sub LoadAirbase(Chaine As String, Airbase As strAirbase)
'Will load terrain specific AB information from CATE conf file
Dim i As Integer, j As Integer
Dim pos As Integer
Dim tmpstring As String, tmpstring2 As String
Dim Found As Integer
Dim NBTile As Integer

    i = 1
    Chaine = Trim$(Chaine)
    tmpstring = ""
    Found = 0
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            Found = Found + 1
            Select Case Found
                Case 1:
                    ReDim Airbase.TileList(1 To GiveNumTilesInAB(DefAirbases(Val(tmpstring))))
                Case 2:
                    Select Case UCase$(tmpstring)
                        Case "NO":
                            Airbase.Level = 0
                        Case "HI":
                            Airbase.Level = 1
                        Case "LO":
                            Airbase.Level = 2
                    End Select
                Case 3:
                    tmpstring2 = ""
                    NBTile = 0
                    For j = 2 To Len(tmpstring) - 1
                        If Mid$(tmpstring, j, 1) = "," Or j = Len(tmpstring) - 1 Then
                            If j = Len(tmpstring) - 1 Then tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                            NBTile = NBTile + 1
                            Airbase.TileList(NBTile) = Val(tmpstring2)
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop
End Sub

Function GiveNumTilesInAB(ABDef As strDefAirbase) As Integer
'Just gives the number of tiles of an AB, when looking at its definition

    GiveNumTilesInAB = (ABDef.XEnd - ABDef.XStart + 1) * (ABDef.YEnd - ABDef.YStart + 1)
End Function

Sub LoadFeatureTile(Chaine As String)
'Will load tile information for features
Dim i As Integer, j As Integer, k As Integer
Dim ro As Integer, ri As Integer
Dim pos As Integer
Dim tmpstring As String, tmpstring2 As String
Dim Found As Integer
Dim NumTile As Integer
Dim NumRepTile As Integer

    i = 1
    Chaine = Trim$(Chaine)
    tmpstring = ""
    Found = 0
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            Found = Found + 1
            Select Case Found
                Case 1:
                    NumTile = Val(tmpstring)
                    FeatureTiles(NumTile, 0) = NumTile
                Case Else:
                    pos = InStr(1, tmpstring, "/")
                    If pos > 0 Then
                        'We read the defined value
                        FeatureTiles(NumTile, Val(Left$(tmpstring, pos - 1))) = Val(Right$(tmpstring, Len(tmpstring) - pos))
                        'We use the base tile number as .Value
                        TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).Value = NumTile
                        'And we define what are the features on such a tile
                        For j = 1 To 4
                            TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).Road(j) = Val(Left$(tmpstring, pos - 1)) And (2 ^ (j - 1))
                            TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).River(j) = Val(Left$(tmpstring, pos - 1)) And (2 ^ (j + 3))
                            If TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).Road(j) > 0 Then
                                TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).Road(j) = 1
                            End If
                            If TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).River(j) > 0 Then
                                TileToFeature(Val(Right$(tmpstring, Len(tmpstring) - pos))).River(j) = 1
                            End If
                        Next j
                    Else
                        'In this case, we use the value found by the index in the second part of the string
                        pos = InStr(1, tmpstring, "r", vbTextCompare)
                        FeatureTiles(NumTile, Val(Left$(tmpstring, pos - 1))) = FeatureTiles(NumTile, Val(Right$(tmpstring, Len(tmpstring) - pos)))
                    End If
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop
    'And here, we copy road tiles over road+river tiles if nothing is defined
    For ro = 1 To 15
        For ri = 1 To 15
            If FeatureTiles(NumTile, ro * 16 + ri) = -1 Then
                If FeatureTiles(NumTile, (ro * 16 + ri) And 15) <> -1 Then
                    FeatureTiles(NumTile, ro * 16 + ri) = FeatureTiles(NumTile, (ro * 16 + ri) And 15)
                End If
            End If
            DoEvents
        Next ri
        DoEvents
    Next ro
End Sub

Sub UpdatePath(ByRef Path() As strCoord, TypePath As String, Direction As Integer, x As Long, y As Long)
'This function will flag the tiles in the paths of features
'TypePath will be RI, RO or CI for rivers or roads or cities
'Direction will be 0 for N-S, 1 for S-N, 2 for W-E, 3 for E-W
'For cities, tile number will be updated here
'For roads or rivers, this is done in MergeFeatures function
Dim CityTerrain As Integer

    Select Case Direction
        Case 0:
            y = y + 1
            ReDim Preserve Path(1 To UBound(Path) + 1)
            Path(UBound(Path)).x = x
            Path(UBound(Path)).y = y
            Select Case TypePath
                Case "RI":
                    Terrain(x, y - 1).River(3) = 1
                    Terrain(x, y).River(1) = 1
                Case "RO":
                    Terrain(x, y - 1).Road(3) = 1
                    Terrain(x, y).Road(1) = 1
                Case "CI":
                    Terrain(x, y - 1).City = 1
                    Terrain(x, y).City = 1
            End Select
        Case 1:
            y = y - 1
            ReDim Preserve Path(1 To UBound(Path) + 1)
            Path(UBound(Path)).x = x
            Path(UBound(Path)).y = y
            Select Case TypePath
                Case "RI":
                    Terrain(x, y + 1).River(1) = 1
                    Terrain(x, y).River(3) = 1
                Case "RO":
                    Terrain(x, y + 1).Road(1) = 1
                    Terrain(x, y).Road(3) = 1
                Case "CI":
                    Terrain(x, y + 1).City = 1
                    Terrain(x, y).City = 1
            End Select
        Case 2:
            x = x + 1
            ReDim Preserve Path(1 To UBound(Path) + 1)
            Path(UBound(Path)).x = x
            Path(UBound(Path)).y = y
            Select Case TypePath
                Case "RI":
                    Terrain(x - 1, y).River(2) = 1
                    Terrain(x, y).River(4) = 1
                Case "RO":
                    Terrain(x - 1, y).Road(2) = 1
                    Terrain(x, y).Road(4) = 1
                Case "CI":
                    Terrain(x - 1, y).City = 1
                    Terrain(x, y).City = 1
            End Select
        Case 3:
            x = x - 1
            ReDim Preserve Path(1 To UBound(Path) + 1)
            Path(UBound(Path)).x = x
            Path(UBound(Path)).y = y
            Select Case TypePath
                Case "RI":
                    Terrain(x + 1, y).River(4) = 1
                    Terrain(x, y).River(2) = 1
                Case "RO":
                    Terrain(x + 1, y).Road(4) = 1
                    Terrain(x, y).Road(2) = 1
                Case "CI":
                    Terrain(x + 1, y).City = 1
                    Terrain(x, y).City = 1
            End Select
    End Select
    
    If TypePath = "CI" Then
        CityTerrain = GiveTerrainFromTile(Terrain(x, y).TileDesc.NumTile)
        Terrain(x, y).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CityTerrain).CityBaseTile
        UpdatedCities = UpdatedCities + 1
        Select Case Direction
            Case 0:
                CityTerrain = GiveTerrainFromTile(Terrain(x, y - 1).TileDesc.NumTile)
                Terrain(x, y - 1).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CityTerrain).CityBaseTile
            Case 1:
                CityTerrain = GiveTerrainFromTile(Terrain(x, y + 1).TileDesc.NumTile)
                Terrain(x, y + 1).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CityTerrain).CityBaseTile
            Case 2:
                CityTerrain = GiveTerrainFromTile(Terrain(x - 1, y).TileDesc.NumTile)
                Terrain(x - 1, y).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CityTerrain).CityBaseTile
            Case 3:
                CityTerrain = GiveTerrainFromTile(Terrain(x + 1, y).TileDesc.NumTile)
                Terrain(x + 1, y).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CityTerrain).CityBaseTile
        End Select
    End If
End Sub

Sub LoadBmpTile(Chaine As String)
'Will load information defining how to color a tile when saving a BMP of the theater
Dim i As Integer, j As Integer, k As Integer, l As Integer, posnext As Integer, Found As Integer
Dim tmpstring As String, tmpstring2 As String, tmpstring3 As String
Dim R As Integer, G As Integer, B As Integer
Dim pos1 As Integer, pos2 As Integer
Dim TileList As String

    i = 1
    Chaine = Trim$(Chaine)
    tmpstring = ""
    Found = 0
    Do
        If Mid$(Chaine, i, 1) = " " Or Mid$(Chaine, i, 1) = Chr$(9) Or i = Len(Chaine) Then
            If i < Len(Chaine) Then
                Do
                    If Mid$(Chaine, i + 1, 1) <> " " And Mid$(Chaine, i + 1, 1) <> Chr$(9) Or i + 1 = Len(Chaine) Then Exit Do
                    i = i + 1
                Loop
            End If
            If i = Len(Chaine) Then tmpstring = tmpstring & Right$(Chaine, 1)
            Found = Found + 1
            Select Case Found
                Case 1:
                    TileList = tmpstring
                Case 2:
                    CurrentImageColor = CurrentImageColor + 1
                    pos1 = InStr(1, tmpstring, ",")
                    pos2 = InStr(pos1 + 1, tmpstring, ",")
                    R = Val(Left$(tmpstring, pos1 - 1))
                    G = Val(Mid$(tmpstring, pos1 + 1, pos2 - pos1 - 1))
                    B = Val(Right(tmpstring, Len(tmpstring) - pos2))
                    SaveImageColors(CurrentImageColor).R = R
                    SaveImageColors(CurrentImageColor).G = G
                    SaveImageColors(CurrentImageColor).B = B
                    
                    tmpstring = TileList
                    tmpstring2 = ""
                    k = 0
                    For j = 2 To Len(Trim$(tmpstring))
                        If Mid$(tmpstring, j, 1) = "-" Or Mid$(tmpstring, j, 1) = "," Or Mid$(tmpstring, j, 1) = ")" Then
                            If Mid$(tmpstring, j, 1) = "-" Then
                                posnext = InStr(j + 1, tmpstring, ",")
                                If posnext <= 0 Then
                                    posnext = InStr(j + 1, tmpstring, ")")
                                End If
                                tmpstring3 = Mid$(tmpstring, j + 1, posnext - j)
                                For l = Val(tmpstring2) To Val(tmpstring3)
                                    k = k + 1
                                    BmpToSave.TileNum(l) = CurrentImageColor
                                    DoEvents
                                Next l
                                j = posnext
                            Else
                                l = Val(tmpstring2)
                                BmpToSave.TileNum(l) = CurrentImageColor
                            End If
                            tmpstring2 = ""
                        Else
                            tmpstring2 = tmpstring2 & Mid$(tmpstring, j, 1)
                        End If
                    Next j
            End Select
            tmpstring = ""
        Else
            tmpstring = tmpstring & Mid$(Chaine, i, 1)
        End If
        i = i + 1
        If i > Len(Chaine) Then Exit Do
    Loop

End Sub

Function CleanString(Chaine As String) As String
'Will remove tabs and spaces in a string
Dim tmp As String, char As String
Dim j As Integer

    For j = 1 To Len(Chaine)
        char = Mid$(Chaine, j, 1)
        If char <> Chr$(9) And char <> " " Then
            tmp = tmp & char
        End If
        DoEvents
    Next j
    CleanString = tmp
End Function

Sub CreateAssocFile()
'Just a quick and dirty function to create a correct assoc. file for TRN import
Dim fd1 As Integer, fd2 As Integer
Dim char1 As String
Dim char2 As String
Dim tmp As String

    fd1 = FreeFile
    Open App.Path & "\fly\org_nam.txt" For Input As #fd1
    fd2 = FreeFile
    Open App.Path & "\fly\assoc.txt" For Output As #fd2
    Do While Not EOF(fd1)
        Line Input #fd1, tmp
        If Len(tmp) > 1 Then
            char1 = Left$(tmp, InStr(1, tmp, ".") - 1)
            char2 = "H" & Mid$(char1, 2, Len(char1) - 2)
            'Note that we should write just char instead of Left$(char1, Len(char1) - 1)
            'but this would not match what is in the TRN files at this time
            Print #fd2, Left$(char1, Len(char1) - 1) & Chr$(9) & char2
        End If
    Loop
    Close #fd1
    Close #fd2
End Sub

Sub RecalculatePaths(OldPath As strPath)
'Will try to enhance the waypoints in a given path (for rivers and roads)
Dim x1 As Long, x2 As Long, x3 As Long, y1 As Long, y2 As Long, y3 As Long
Dim j As Long, j2 As Long
Dim NewPath() As strCoord
Dim NewCoords As Integer
'Dim tmp As String

    ReDim NewPath(1 To 1)
    NewCoords = 1
    j2 = -1
    x1 = OldPath.Coord(1).x
    y1 = OldPath.Coord(1).y
    NewPath(1).x = x1
    NewPath(1).y = y1
    j = 2
    Do
        x2 = OldPath.Coord(j).x
        y2 = OldPath.Coord(j).y
        If x2 <> x1 And y2 <> y1 Then
            x3 = OldPath.Coord(j - 1).x
            y3 = OldPath.Coord(j - 1).y
            If x3 <> x1 Or y3 <> y1 Then
                j = j - 1
                x2 = OldPath.Coord(j).x
                y2 = OldPath.Coord(j).y
                NewCoords = NewCoords + 1
                ReDim Preserve NewPath(1 To NewCoords)
                NewPath(NewCoords).x = x2
                NewPath(NewCoords).y = y2
                x1 = x2
                y1 = y2
            Else
                NewCoords = NewCoords + 1
                ReDim Preserve NewPath(1 To NewCoords)
                NewPath(NewCoords).x = x2
                NewPath(NewCoords).y = y2
                x1 = x2
                y1 = y2
                j2 = -1
            End If
        Else
            If j = OldPath.NumCoords Then
                If x2 = x1 And y2 = y1 Then
                    If j2 <> -1 Then
                        j = j2
                        x2 = OldPath.Coord(j).x
                        y2 = OldPath.Coord(j).y
                        NewCoords = NewCoords + 1
                        ReDim Preserve NewPath(1 To NewCoords)
                        NewPath(NewCoords).x = x2
                        NewPath(NewCoords).y = y2
                        x1 = x2
                        y1 = y2
                        j2 = -1
                    End If
                Else
                    NewCoords = NewCoords + 1
                    ReDim Preserve NewPath(1 To NewCoords)
                    NewPath(NewCoords).x = x2
                    NewPath(NewCoords).y = y2
                    x1 = x2
                    y1 = y2
                    j2 = -1
                End If
            End If
            If x2 <> x1 Or y2 <> y1 Then
                j2 = j
            End If
        End If
        j = j + 1
        If j > OldPath.NumCoords Then Exit Do
        DoEvents
    Loop
    
    'Dim tmp As String
    'tmp = ""
    'For j = 1 To OldPath.NumCoords
    '    tmp = tmp & "(" & OldPath.Coord(j).x & "," & OldPath.Coord(j).y & ") "
    'Next j
    'Debug.Print tmp
    ''tmp = tmp & " -> "
    'tmp = ""
    'For j = 1 To NewCoords
    '    tmp = tmp & "(" & NewPath(j).x & "," & NewPath(j).y & ") "
    'Next j
    'Debug.Print tmp
    'Debug.Print
    
    ReDim OldPath.Coord(1 To NewCoords)
    OldPath.NumCoords = NewCoords
    For j = 1 To NewCoords
        OldPath.Coord(j).x = NewPath(j).x
        OldPath.Coord(j).y = NewPath(j).y
        DoEvents
    Next j

End Sub
