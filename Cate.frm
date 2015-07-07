VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATE"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "Cate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Configuration file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3525
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9675
      Begin VB.TextBox txtConfFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "CATE configuration file name"
         Top             =   300
         Width           =   6315
      End
      Begin VB.CommandButton btConfFile 
         Caption         =   "&Browse ..."
         Height          =   375
         Left            =   6540
         TabIndex        =   17
         ToolTipText     =   "CATE configuration file name"
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   960
         Width           =   9435
      End
      Begin MSComDlg.CommonDialog dialConfFile 
         Left            =   8520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "List of rules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input file (Falcon4 L2 terrain file) and Output (new) file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   0
      TabIndex        =   10
      Top             =   3540
      Width           =   9675
      Begin VB.CommandButton btL2File 
         Caption         =   "&Browse ..."
         Height          =   375
         Left            =   6540
         TabIndex        =   14
         ToolTipText     =   "L2 input file name"
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtL2File 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "L2 input file name"
         Top             =   300
         Width           =   6315
      End
      Begin VB.TextBox txtOutFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "L2 output file name"
         Top             =   720
         Width           =   6315
      End
      Begin VB.CommandButton btOutFile 
         Caption         =   "&Browse ..."
         Height          =   375
         Left            =   6540
         TabIndex        =   11
         ToolTipText     =   "L2 output file name"
         Top             =   660
         Width           =   915
      End
      Begin MSComDlg.CommonDialog dialL2File 
         Left            =   8520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComDlg.CommonDialog dialOutFile 
         Left            =   8520
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   4740
      Width           =   9675
      Begin VB.CommandButton btUpdateData 
         Caption         =   "&Update data"
         Enabled         =   0   'False
         Height          =   855
         Left            =   180
         Picture         =   "Cate.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Apply rules to data into memory"
         Top             =   1200
         Width           =   1155
      End
      Begin VB.CommandButton btReadL2 
         Caption         =   "&Read data"
         Height          =   855
         Left            =   180
         Picture         =   "Cate.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Read and load L2 and O2 data"
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton btSaveL2 
         Caption         =   "&Save data"
         Enabled         =   0   'False
         Height          =   855
         Left            =   8400
         Picture         =   "Cate.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save L2 and O2 data"
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton btQuit 
         Caption         =   "&Quit"
         Height          =   855
         Left            =   8400
         Picture         =   "Cate.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "End this program"
         Top             =   1200
         Width           =   1155
      End
      Begin VB.PictureBox picAction 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2010
         ScaleHeight     =   195
         ScaleWidth      =   6225
         TabIndex        =   2
         Top             =   1020
         Visible         =   0   'False
         Width           =   6285
      End
      Begin VB.CheckBox chkBatch 
         Caption         =   "Run CATE in ""Batch"" mode"
         Height          =   255
         Left            =   1500
         TabIndex        =   1
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label lblShowAction 
         AutoSize        =   -1  'True
         Caption         =   "Updating data ..."
         Height          =   195
         Left            =   1500
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblPercAction 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 %"
         Height          =   195
         Left            =   1620
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BETA VERSION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4740
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmCate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Paths stored for dialog boxes
Dim LastConfPath As String
Dim LastL2InputPath As String
Dim LastL2OutputPath As String
'Was the data modified
Dim Modified As Boolean
'Different flags for different CATE operations
Dim RunTransitions As Boolean
Dim RunFeatures As Boolean
Dim RunBmpRules As Boolean
Dim RunCateRules As Boolean
Dim SaveBmp As Boolean
Dim RunTrnImport As Boolean
Dim FullRegionTest As Boolean
Dim UseTrefSection As Boolean
'Show information boxes between operations ?
Dim BatchMode As Boolean
'Used to know if we should or not recalculate the data
Dim TerrainCalculated As Boolean

Private Sub btConfFile_Click()
    'Let's open the configuration file
    dialConfFile.Flags = &H800& Or &H4& Or &H1000&
    dialConfFile.DefaultExt = "conf"
    dialConfFile.Filter = "CATE configuration files (*.conf)|*.conf|All files|*.*"
    dialConfFile.FilterIndex = 0
    dialConfFile.DialogTitle = "Read configuration from ..."
    If LastConfPath <> "" Then dialConfFile.InitDir = LastConfPath
    On Error Resume Next
    dialConfFile.ShowOpen
    If Err = 0 Then
        On Error GoTo 0
        txtConfFile.Text = dialConfFile.FileName
        LastConfPath = GivePathFromName(dialConfFile.FileName)
        Screen.MousePointer = vbHourglass
        DoEvents
        If LoadRules(dialConfFile.FileName) Then
            ShowRules
        Else
            txtConfFile.Text = ""
            Text1.Text = ""
        End If
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub btL2File_Click()
    'Let's open the L2 file
    dialL2File.Flags = &H800& Or &H4& Or &H1000&
    dialL2File.DefaultExt = "l2"
    dialL2File.Filter = "L2 data files (*.l2)|*.l2|All files|*.*"
    dialL2File.FilterIndex = 0
    dialL2File.DialogTitle = "Read L2 data from ..."
    If LastL2InputPath <> "" Then dialL2File.InitDir = LastL2InputPath
    On Error Resume Next
    dialL2File.ShowOpen
    If Err = 0 Then
        If Dir$(Left$(dialL2File.FileName, Len(dialL2File.FileName) - 2) & "o2", vbNormal) = "" Then
            MsgBox "Unable to find " & Left$(dialL2File.FileName, Len(dialL2File.FileName) - 2) & "o2. This file is mandatory for CATE to work.", vbCritical + vbOKOnly, "Warning"
            txtL2File.Text = ""
            Exit Sub
        Else
            txtL2File.Text = dialL2File.FileName
            LastL2InputPath = GivePathFromName(dialL2File.FileName)
        End If
    End If
End Sub

Private Sub btOutFile_Click()
    'Let's the user tell us under what file name we'll save the data
    dialOutFile.Flags = &H800& Or &H4& Or &H2&
    'dialOutFile.DefaultExt = "l2"
    dialOutFile.Filter = "L2 data files (*.l2)|*.l2|All files|*.*"
    dialOutFile.FilterIndex = 0
    dialOutFile.DialogTitle = "Save L2 data in file ..."
    If LastL2OutputPath <> "" Then dialOutFile.InitDir = LastL2OutputPath
    On Error Resume Next
    dialOutFile.ShowSave
    If Err = 0 Then
        'Verify it's a L2 (extension) file
        If UCase$(Right$(dialOutFile.FileName, 3)) <> ".L2" Then
            MsgBox "Please use a filename with a '.L2' extension for L2 data, so that CATE can create the relative O2 file as well. Thank you.", vbInformation + vbOKOnly, "Information"
            txtOutFile.Text = ""
        Else
            txtOutFile.Text = dialOutFile.FileName
            LastL2OutputPath = GivePathFromName(dialOutFile.FileName)
        End If
    End If
End Sub

Private Sub btQuit_Click()
    Unload Me
End Sub

Private Sub btReadL2_Click()
Dim LoadOK As Boolean

    If txtConfFile.Text = "" Or txtL2File.Text = "" Or txtOutFile = "" Then
        MsgBox "Please select above all three files which will be used (configuration, L2, and saved files) before starting actions. Thank you.", vbInformation + vbOKOnly, "Information"
        Exit Sub
    End If
    If chkBatch.Value = 0 Then BatchMode = False Else BatchMode = True

    'Some initializations
    Erase Tiles
    Erase Regions
    Erase Terrain
    Erase ImageData
    
    If LoadTiles(txtL2File.Text) Then
        If LoadRegions(Left$(txtL2File.Text, Len(txtL2File.Text) - 2) & "o2") Then
            If Not BatchMode Then MsgBox "L2 and O2 data are now loaded into memory", vbInformation + vbOKOnly, "Information"
            If RunBmpRules Then
                If LoadBmp(Rules.BmpFileName) Then
                    If ImageHeader.ImageHeight <> TerrainSize * 16 Then
                        MsgBox "BMP size (" & ImageHeader.ImageHeight & "x" & ImageHeader.ImageWidth & ") does not fit with Terrain size (" & TerrainSize & "x" & TerrainSize & "). Please check your files.", vbCritical + vbOKOnly, "Warning"
                        LoadOK = False
                    Else
                        If Not BatchMode Then MsgBox "Bitmap is now loaded into memory", vbInformation + vbOKOnly, "Information"
                        LoadOK = True
                    End If
                Else
                    LoadOK = False
                End If
            Else
                LoadOK = True
            End If
            If AutoFeatures.NumFiles > 0 Then
                If LoadFeaturesFiles() Then LoadOK = True Else LoadOK = False
            End If
            If AutoFeatures.NumABFiles > 0 Then
                If LoadAirbaseFiles() Then LoadOK = True Else LoadOK = False
            End If
            If RunTrnImport Then
                If AutoFeatures.CorrespFileName <> "" Then
                    If LoadCorresp(AutoFeatures.CorrespFileName) Then
                        LoadOK = True
                    Else
                        LoadOK = False
                        GoTo Fin
                    End If
                End If
                If LoadTextureIndex(AutoFeatures.TextureFileName) Then
                    If LoadTrnFiles() Then
                        If Not BatchMode Then MsgBox "Texture Index file and TRN files are now loaded into memory", vbInformation + vbOKOnly, "Information"
                        LoadOK = True
                    Else
                        LoadOK = False
                    End If
                Else
                    LoadOK = False
                End If
            End If
        Else
            LoadOK = False
        End If
    End If

Fin:
    If LoadOK Then EnableInput True Else EnableInput False
    btConfFile.Enabled = True
    btL2File.Enabled = True
    btOutFile.Enabled = True
    btReadL2.Enabled = True
    btQuit.Enabled = True
    If BatchMode And LoadOK Then btUpdateData_Click

End Sub

Private Sub btSaveL2_Click()
    If txtConfFile.Text = "" Or txtL2File.Text = "" Or txtOutFile = "" Then
        MsgBox "Please select above all three files which will be used (configuration, L2, and saved files) before starting actions. Thank you.", vbInformation + vbOKOnly, "Information"
        Exit Sub
    End If
    If SaveFile(txtOutFile.Text) = True Then Modified = False

End Sub

Private Sub btUpdateData_Click()
Dim nb As Long

    If txtConfFile.Text = "" Or txtL2File.Text = "" Or txtOutFile = "" Then
        MsgBox "Please select above all three files which will be used (configuration, L2, and saved files) before starting actions. Thank you.", vbInformation + vbOKOnly, "Information"
        Exit Sub
    End If
    
    TerrainCalculated = False
    If RunTrnImport Then
        nb = UpdateTrnTiles()
        If Not BatchMode Then MsgBox "TRN tiles imported, " & nb & " tiles updated (" & Format$((nb * 100) / (TerrainSize * TerrainSize * 256), "0.00") & " %)", vbInformation, "Information"
    End If
    If RunFeatures Or RunTransitions Then ApplyFeatures
    If RunBmpRules Then ApplyBmpRules
    If SaveBmp Then SaveTerrainAsBmp
    If RunCateRules Then ApplyCateRules
    If BatchMode Then btSaveL2_Click

End Sub

Private Sub Form_Load()
    Me.Caption = "Configurable Auto Tiling Enhancer - Version " & App.Major & "." & App.Minor & App.Revision
    Modified = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Rep As Integer

    If btQuit.Enabled = False Then Cancel = True
    If Modified = True Then
        Rep = MsgBox("You have modified data, but file has not been saved. Are you sure you want to quit without saving ?", vbYesNo + vbQuestion, "Warning")
        If Rep = vbNo Then Cancel = True
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub EnableInput(Value As Boolean)
'Enables/disables user input while working

    btConfFile.Enabled = Value
    btL2File.Enabled = Value
    btOutFile.Enabled = Value
    btReadL2.Enabled = Value
    btUpdateData.Enabled = Value
    btSaveL2.Enabled = Value
    btQuit.Enabled = Value
    chkBatch.Enabled = Value
End Sub

Function LoadRules(FileName As String) As Boolean
'Loads the rules file
Dim fd As String
Dim i As Integer, k As Integer
Dim pos1 As Integer, pos2 As Integer, Pos3 As Integer
Dim tmp As String
Dim NumStarts As Integer, NumEnds As Integer
Dim NumFeatStarts As Integer, NumFeatEnds As Integer, NumABDefs As Integer
Dim Message As String
'Dim nb As Long
Dim CurrentTerrain As Integer

    'Some initializations
    RunTransitions = False
    RunFeatures = False
    RunBmpRules = False
    RunCateRules = False
    RunTrnImport = False
    FullRegionTest = True
    UseTrefSection = True
    SaveBmp = False
    Rules.NumSections = 0
    Rules.BmpFileName = ""
    Rules.UpdateOceanTiles = 0
    AutoFeatures.NumTerrain = -1
    AutoFeatures.NumFiles = 0
    AutoFeatures.NumABFiles = 0
    AutoFeatures.NumTrnFiles = 0
    AutoFeatures.RiverRoadMethod = 0
    AutoFeatures.CorrespFileName = ""
    AutoFeatures.TRNOffsetX = 0
    AutoFeatures.TRNOffsetY = 0
    BmpToSave.FileName = ""
    AutoFeatures.TextureFileName = ""
    ReDim AutoFeatures.TrnFileNames(0 To 0)
    ReDim AutoFeatures.TerrainFeatures(0 To 0)
    ReDim AutoFeatures.FeatName(0 To 0)
    ReDim AutoFeatures.FileNames(0 To 0)
    ReDim AutoFeatures.FileABNames(0 To 0)
    'ReDim gt(0 To 9, 0 To 9)
    Erase TileToTerrain
    Erase FeatureTiles
    Erase TileToFeature
    Erase FeatureTiles
    Erase BmpToSave.TileNum
    'Erase SaveImageColors
    CurrentImageColor = -1
    For i = 0 To 255
        SaveImageColors(i).R = 0
        SaveImageColors(i).G = 0
        SaveImageColors(i).B = 0
    Next i
    For i = 0 To 5000
        For k = 0 To 255
            FeatureTiles(i, k) = -1
        Next k
        TileToFeature(i).Value = -1
        BmpToSave.TileNum(i) = 255
    Next i
    NumStarts = 0
    NumEnds = 0
    CurrentTerrain = -1

    fd = FreeFile
    On Error Resume Next
    'We open the conf file
    Open FileName For Input As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted", vbCritical, "Warning"
        LoadRules = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    On Error GoTo Problem
    Do While Not EOF(fd)
        Line Input #fd, tmp                 'we read a full line in tmp string
        If Len(tmp) > 7 Then                '"Empty" lines are ignored
            If Left$(tmp, 1) <> "#" Then    'Comments are ignored
                If UCase$(Left$(tmp, 4)) = "[END" Then  'this is the end of a rules section
                    If UCase$(Left$(tmp, 9)) = "[ENDRULES" Then 'end of rules section
                        NumEnds = NumEnds + 1
                        If Rules.RuleSections(Rules.NumSections).R <= -2 Or Rules.RuleSections(Rules.NumSections).G <= -2 Or Rules.RuleSections(Rules.NumSections).B <= -2 Then
                            Screen.MousePointer = vbNormal
                            MsgBox "Incorrect format of configuration file (incorrect or absent color rule). Check the file " & FileName, vbCritical, "Warning"
                            LoadRules = False
                            Close #fd
                            Exit Function
                        End If
                    ElseIf UCase$(Left$(tmp, 11)) = "[ENDTERRAIN" Then 'end of features section
                        NumFeatEnds = NumFeatEnds + 1
                    ElseIf UCase$(Left$(tmp, 12)) = "[ENDAIRBASES" Then 'end of airbases section
                        'Nothing special to do
                        NumEnds = NumEnds + 1
                    ElseIf UCase$(Left$(tmp, 16)) = "[ENDFEATURESDESC" Then 'end of airbases section
                        'Nothing special to do
                        NumEnds = NumEnds + 1
                    ElseIf UCase$(Left$(tmp, 11)) = "[ENDSAVEBMP" Then
                        NumEnds = NumEnds + 1
                    ElseIf UCase$(Left$(tmp, 13)) = "[ENDIMPORTTRN" Then
                        NumEnds = NumEnds + 1
                    End If
                ElseIf UCase$(Left$(tmp, 1)) = "[" Then 'this is the beginnning of a section
                    If UCase$(Left$(tmp, 7)) = "[RULES:" Then   'Rules section
                        'We add on section to rules array
                        NumStarts = NumStarts + 1
                        Rules.NumSections = Rules.NumSections + 1
                        If Rules.NumSections = 1 Then
                            ReDim Rules.RuleSections(1 To 1)
                        Else
                            ReDim Preserve Rules.RuleSections(1 To Rules.NumSections)
                        End If
                        'Some initializations
                        Rules.RuleSections(Rules.NumSections).R = -2
                        Rules.RuleSections(Rules.NumSections).G = -2
                        Rules.RuleSections(Rules.NumSections).B = -2
                        Rules.RuleSections(Rules.NumSections).NumAltRules = 0
                        Rules.RuleSections(Rules.NumSections).NumTileRules = 0
                        Rules.RuleSections(Rules.NumSections).NumFogRules = 0
                        Rules.RuleSections(Rules.NumSections).Name = Trim$(Mid$(tmp, 8, Len(tmp) - 8))
                    ElseIf UCase$(Left$(tmp, 8)) = "[TERRAIN" Then   'Feats section
                        'We add on section to features array
                        CurrentTerrain = CurrentTerrain + 1
                        If AutoFeatures.NumTerrain = -1 Then AutoFeatures.NumTerrain = 0
                        If RunTransitions And CurrentTerrain = 0 And AutoFeatures.NumFiles <= 0 And AutoFeatures.NumABFiles <= 0 Then CurrentTerrain = 1
                        NumFeatStarts = NumFeatStarts + 1
                        'Some initializations
                        ReDim Preserve AutoFeatures.FeatName(0 To CurrentTerrain)
                        If CurrentTerrain = 0 And NumABDefs > 0 Then
                            'We prepare some things for airbases
                            ReDim AutoFeatures.TerrainFeatures(0).AirBases(1 To NumABDefs)
                            AutoFeatures.TerrainFeatures(0).AirBases(1).Level = -1
                        End If
                        AutoFeatures.FeatName(CurrentTerrain) = Trim$(Mid$(tmp, 18, Len(tmp) - 18))
                    ElseIf UCase$(Left$(tmp, 12)) = "[AIRBASESDEF" Then   'Airbases section
                        'Nothing special to do
                        NumStarts = NumStarts + 1
                    ElseIf UCase$(Left$(tmp, 13)) = "[FEATURESDESC" Then   'Airbases section
                        If AutoFeatures.NumFiles = 0 Then
                            Screen.MousePointer = vbNormal
                            MsgBox "Incorrect format of configuration file (FeaturesDesc section found, but no TDF file defined). Check the file " & FileName, vbCritical, "Warning"
                            LoadRules = False
                            Close #fd
                            Exit Function
                        End If
                        NumStarts = NumStarts + 1
                    ElseIf UCase$(Left$(tmp, 8)) = "[SAVEBMP" Then
                        NumStarts = NumStarts + 1
                        SaveBmp = True
                    ElseIf UCase$(Left$(tmp, 10)) = "[IMPORTTRN" Then
                        NumStarts = NumStarts + 1
                        RunTrnImport = True
                    End If
                Else                                    'this a rule
                    'We search for =
                    pos1 = InStr(1, tmp, "=")
                    If pos1 <= 0 Then
                        'no = : there's a problem
                        Screen.MousePointer = vbNormal
                        MsgBox "Incorrect format of configuration file (rule without a '='). Check the file " & FileName, vbCritical, "Warning"
                        LoadRules = False
                        Close #fd
                        Exit Function
                    Else
                        '= found, we now search which type of rule we have found
                        Select Case LCase$(Left$(tmp, pos1 - 1))
                            Case "bmpfilename":     'BMP, out of any section
                                Rules.BmpFileName = Right$(tmp, Len(tmp) - pos1)
                            Case "abfilename":
                                AutoFeatures.NumABFiles = AutoFeatures.NumABFiles + 1
                                If AutoFeatures.NumABFiles = 1 Then
                                    ReDim AutoFeatures.FileABNames(1 To 1)
                                Else
                                    ReDim Preserve AutoFeatures.FileABNames(1 To AutoFeatures.NumABFiles)
                                End If
                                AutoFeatures.FileABNames(AutoFeatures.NumABFiles) = Right$(tmp, Len(tmp) - pos1)
                            Case "tdffilename":
                                AutoFeatures.NumFiles = AutoFeatures.NumFiles + 1
                                If AutoFeatures.NumFiles = 1 Then
                                    ReDim AutoFeatures.FileNames(1 To 1)
                                Else
                                    ReDim Preserve AutoFeatures.FileNames(1 To AutoFeatures.NumFiles)
                                End If
                                AutoFeatures.FileNames(AutoFeatures.NumFiles) = Right$(tmp, Len(tmp) - pos1)
                            Case "updateoceantiles":
                                If UCase(Right$(tmp, Len(tmp) - pos1)) = "OK" Then
                                    Rules.UpdateOceanTiles = 1
                                Else
                                    Rules.UpdateOceanTiles = 0
                                End If
                            Case "riverroadmethod":
                                AutoFeatures.RiverRoadMethod = Val(Right$(tmp, Len(tmp) - pos1))
                            Case "fullregiontest":
                                If UCase$(Right$(tmp, Len(tmp) - pos1)) = "NO" Then
                                    FullRegionTest = False
                                Else
                                    FullRegionTest = True
                                End If
                            Case "color":           'Color rule
                                pos2 = InStr(pos1 + 1, tmp, ",")
                                Pos3 = InStr(pos2 + 1, tmp, ",")
                                Rules.RuleSections(Rules.NumSections).R = Val(Mid$(tmp, pos1 + 1, pos2 - pos1 - 1))
                                Rules.RuleSections(Rules.NumSections).G = Val(Mid$(tmp, pos2 + 1, Pos3 - pos2 - 1))
                                Rules.RuleSections(Rules.NumSections).B = Val(Mid$(tmp, Pos3 + 1, Len(tmp) - Pos3))
                            Case "forcefog":        'ForceFog rule
                                Rules.RuleSections(Rules.NumSections).NumFogRules = Rules.RuleSections(Rules.NumSections).NumFogRules + 1
                                If Rules.RuleSections(Rules.NumSections).NumFogRules = 1 Then
                                    ReDim Rules.RuleSections(Rules.NumSections).FogRules(1 To Rules.RuleSections(Rules.NumSections).NumFogRules)
                                Else
                                    ReDim Preserve Rules.RuleSections(Rules.NumSections).FogRules(1 To Rules.RuleSections(Rules.NumSections).NumFogRules)
                                End If
                                LoadFogRule Right$(tmp, Len(tmp) - pos1), Rules.RuleSections(Rules.NumSections).FogRules(Rules.RuleSections(Rules.NumSections).NumFogRules)
                            Case "tilerule":        'Tile Rule
                                Rules.RuleSections(Rules.NumSections).NumTileRules = Rules.RuleSections(Rules.NumSections).NumTileRules + 1
                                If Rules.RuleSections(Rules.NumSections).NumTileRules = 1 Then
                                    ReDim Rules.RuleSections(Rules.NumSections).TileRules(1 To Rules.RuleSections(Rules.NumSections).NumTileRules)
                                Else
                                    ReDim Preserve Rules.RuleSections(Rules.NumSections).TileRules(1 To Rules.RuleSections(Rules.NumSections).NumTileRules)
                                End If
                                LoadTileRule Right$(tmp, Len(tmp) - pos1), Rules.RuleSections(Rules.NumSections).TileRules(Rules.RuleSections(Rules.NumSections).NumTileRules)
                            Case "randalt":         'Random altitude rule
                                Rules.RuleSections(Rules.NumSections).NumAltRules = Rules.RuleSections(Rules.NumSections).NumAltRules + 1
                                If Rules.RuleSections(Rules.NumSections).NumAltRules = 1 Then
                                    ReDim Rules.RuleSections(Rules.NumSections).AltRules(1 To Rules.RuleSections(Rules.NumSections).NumAltRules)
                                Else
                                    ReDim Preserve Rules.RuleSections(Rules.NumSections).AltRules(1 To Rules.RuleSections(Rules.NumSections).NumAltRules)
                                End If
                                LoadAltRule Right$(tmp, Len(tmp) - pos1), Rules.RuleSections(Rules.NumSections).AltRules(Rules.RuleSections(Rules.NumSections).NumAltRules)
                            Case "dotransitions":
                                If UCase(Right$(tmp, Len(tmp) - pos1)) = "OK" Then
                                    RunTransitions = True
                                Else
                                    RunTransitions = False
                                End If
                            Case "terraintype":
                                If RunTransitions And CurrentTerrain = 0 And AutoFeatures.NumFiles <= 0 And AutoFeatures.NumABFiles <= 0 Then
                                    CurrentTerrain = 1
                                    AutoFeatures.NumTerrain = 1
                                    ReDim AutoFeatures.TerrainFeatures(0 To CurrentTerrain)
                                Else
                                    AutoFeatures.NumTerrain = AutoFeatures.NumTerrain + 1
                                    ReDim Preserve AutoFeatures.TerrainFeatures(0 To CurrentTerrain)
                                End If
                                AutoFeatures.TerrainFeatures(CurrentTerrain).TerrainType = Val(Right$(tmp, Len(tmp) - pos1))
                                If AutoFeatures.NumFiles > 0 Then
                                    AutoFeatures.TerrainFeatures(CurrentTerrain).CityBaseTile = AutoFeatures.TerrainFeatures(0).CityBaseTile
                                End If
                                If NumABDefs > 0 Then
                                    ReDim AutoFeatures.TerrainFeatures(CurrentTerrain).AirBases(1 To NumABDefs)
                                    'For each new terrain, we assign to it all the default airbases definitions
                                    AutoFeatures.TerrainFeatures(CurrentTerrain).AirBases = AutoFeatures.TerrainFeatures(0).AirBases()
                                End If
                            Case "terraintiles":
                                LoadTerrainTiles Right$(tmp, Len(tmp) - pos1), AutoFeatures.TerrainFeatures(CurrentTerrain).TerrainType
                            Case "transitiondef":
                                AutoFeatures.TerrainFeatures(CurrentTerrain).NumTransitions = AutoFeatures.TerrainFeatures(CurrentTerrain).NumTransitions + 1
                                If AutoFeatures.TerrainFeatures(CurrentTerrain).NumTransitions = 1 Then
                                    ReDim AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(1 To 1)
                                Else
                                    ReDim Preserve AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(1 To AutoFeatures.TerrainFeatures(CurrentTerrain).NumTransitions)
                                End If
                                LoadTransition Right$(tmp, Len(tmp) - pos1), CurrentTerrain
                            Case "airbasedef":
                                NumABDefs = NumABDefs + 1
                                If NumABDefs = 1 Then
                                    ReDim DefAirbases(1 To 1)
                                Else
                                    ReDim Preserve DefAirbases(1 To NumABDefs)
                                End If
                                LoadABDef Right$(tmp, Len(tmp) - pos1), DefAirbases(NumABDefs)
                            Case "airbase":
                                LoadAirbase Right$(tmp, Len(tmp) - pos1), AutoFeatures.TerrainFeatures(CurrentTerrain).AirBases(Val(Right$(tmp, Len(tmp) - pos1)))
                            Case "city":
                                AutoFeatures.TerrainFeatures(CurrentTerrain).CityBaseTile = Val(Right$(tmp, Len(tmp) - pos1))
                            Case "tile":
                                LoadFeatureTile Right$(tmp, Len(tmp) - pos1)
                            Case "savebmp":
                                BmpToSave.FileName = Right$(tmp, Len(tmp) - pos1)
                            Case "bmptile":
                                LoadBmpTile Right$(tmp, Len(tmp) - pos1)
                            Case "textureindexfile":
                                AutoFeatures.TextureFileName = Right$(tmp, Len(tmp) - pos1)
                            Case "correspfilename":
                                AutoFeatures.CorrespFileName = Right$(tmp, Len(tmp) - pos1)
                            Case "trnfilename":
                                AutoFeatures.NumTrnFiles = AutoFeatures.NumTrnFiles + 1
                                If AutoFeatures.NumTrnFiles = 1 Then
                                    ReDim AutoFeatures.TrnFileNames(1 To 1)
                                Else
                                    ReDim Preserve AutoFeatures.TrnFileNames(1 To AutoFeatures.NumTrnFiles)
                                End If
                                AutoFeatures.TrnFileNames(AutoFeatures.NumTrnFiles) = Right$(tmp, Len(tmp) - pos1)
                            Case "trnoffset":
                                pos2 = InStr(1, Right$(tmp, Len(tmp) - pos1), ",")
                                AutoFeatures.TRNOffsetX = Left$(Right$(tmp, Len(tmp) - pos1), pos2 - 1)
                                AutoFeatures.TRNOffsetY = Right$(Right$(tmp, Len(tmp) - pos1), Len(Right$(tmp, Len(tmp) - pos1)) - pos2)
                            Case "usetrefsection":
                                If LCase$(Right$(tmp, Len(tmp) - pos1)) = "no" Then
                                    UseTrefSection = False
                                Else
                                    UseTrefSection = True
                                End If
                            Case Else:
                                'Unknow rule : problem
                                Screen.MousePointer = vbNormal
                                MsgBox "Incorrect format of configuration file (unknown rule " & Left$(tmp, pos1 - 1) & "). Check the file " & FileName, vbCritical, "Warning"
                                LoadRules = False
                                Close #fd
                                Exit Function
                        End Select
                    End If
                End If
            End If
        End If
        DoEvents
    Loop
    Close #fd

    'Some verifications on read format
    Message = ""
    If NumStarts <> NumEnds Then
        Message = "Incorrect format of configuration file (" & NumStarts & " rules section starts, " & NumEnds & " rules section ends). Check the file " & FileName
        GoTo Fin
    End If
    If Rules.BmpFileName = "" And Rules.NumSections > 1 Then
        Message = "Incorrect format of configuration file (no BMP file defined, but several sections detected). Check the file " & FileName
        GoTo Fin
    End If
    If RunTransitions Then
        If AutoFeatures.NumTerrain = 0 Then
            Message = "Incorrect format of configuration file (DoTransitions detected, but no Features section). Check the file " & FileName
            GoTo Fin
        End If
    End If
    If AutoFeatures.NumABFiles > 0 Then
        If DefAirbases(1).XEnd = 0 And DefAirbases(1).YEnd = 0 And DefAirbases(1).XStart = 0 And DefAirbases(1).YStart = 0 Then
            Message = "Incorrect format of configuration file (No Airbase definitions found). Check the file " & FileName
            GoTo Fin
        End If
        If AutoFeatures.TerrainFeatures(0).AirBases(1).Level < 0 Then
            Message = "Incorrect format of configuration file (no correct airbase descriptions found in terrain default section). Check the file " & FileName
            GoTo Fin
        End If
    End If
    If RunTrnImport Then
        If AutoFeatures.TextureFileName = "" Then
            Message = "Incorrect format of configuration file (ImportTRN section detected, but no Texture Index file found). Check the file " & FileName
            GoTo Fin
        End If
        If AutoFeatures.NumTrnFiles = 0 Then
            Message = "Incorrect format of configuration file (ImportTRN section detected, but no TRN files found). Check the file " & FileName
            GoTo Fin
        End If
    End If
    If Rules.NumSections > 0 Then
        For i = 1 To Rules.NumSections
            If Rules.RuleSections(i).NumAltRules = 0 And Rules.RuleSections(i).NumFogRules = 0 And Rules.RuleSections(i).NumTileRules = 0 Then
                Message = "Incorrect format of configuration file (at least one section has no rule at all). Check the file " & FileName
                GoTo Fin
            End If
        Next i
        If Rules.BmpFileName = "" Then
            If Rules.RuleSections(1).R >= 0 Or Rules.RuleSections(1).G >= 0 Or Rules.RuleSections(1).B >= 0 Then
                Message = "Incorrect format of configuration file (no BMP file defined, but color rule is not equal to -1,-1,-1). Check the file " & FileName
                GoTo Fin
            End If
        Else
            For i = 1 To Rules.NumSections
                If Rules.RuleSections(i).R < 0 Or Rules.RuleSections(i).G < 0 Or Rules.RuleSections(i).B < 0 Then
                    Message = "Incorrect format of configuration file (BMP file defined, but one or several color rules have negative colors or no colors defined). Check the file " & FileName
                    Exit For
                End If
            Next i
        End If
    End If
    
Fin:
    Screen.MousePointer = vbNormal
    If Message = "" Then
        If Rules.BmpFileName <> "" Then
            RunCateRules = False
            RunBmpRules = True
            If AutoFeatures.NumFiles > 0 Or AutoFeatures.NumABFiles > 0 Then RunFeatures = True Else RunFeatures = False
        Else
            If Rules.NumSections > 0 Then
                RunCateRules = True
                RunBmpRules = False
                If AutoFeatures.NumFiles > 0 Or AutoFeatures.NumABFiles > 0 Then RunFeatures = True Else RunFeatures = False
            Else
                RunCateRules = False
                RunBmpRules = False
                If AutoFeatures.NumFiles > 0 Or AutoFeatures.NumABFiles > 0 Then RunFeatures = True Else RunFeatures = False
            End If
        End If
        LoadRules = True
    Else
        MsgBox Message, vbCritical, "Warning"
        LoadRules = False
    End If
    Exit Function

Problem:
    Screen.MousePointer = vbNormal
    MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reading conf file " & FileName & ". Please check your conf file format. Action aborted.", vbCritical, "Warning"
    Close #fd
    LoadRules = False
    Exit Function
End Function

Sub ShowRules()
Dim i As Integer, j As Integer, k As Integer
Dim tmp As String
Dim NL As String

    'We show all the rules in a text box
    Screen.MousePointer = vbHourglass
    NL = Chr$(13) & Chr$(10)
    Text1.Text = ""

    If FullRegionTest Then
        Text1.Text = Text1.Text & "Full Region Building when saving theater : YES" & NL & NL
    Else
        Text1.Text = Text1.Text & "Full Region Building when saving theater : NO" & NL & NL
    End If
    
    If RunTransitions Then
        Text1.Text = Text1.Text & "TRANSITIONS DEFINITIONS" & NL & NL
        For i = 1 To AutoFeatures.NumTerrain
            If AutoFeatures.TerrainFeatures(i).NumTransitions > 0 Then
                Text1.Text = Text1.Text & "  - For terrain type " & AutoFeatures.TerrainFeatures(i).TerrainType & " (Features section " & AutoFeatures.FeatName(i) & ")" & NL
                For j = 1 To AutoFeatures.TerrainFeatures(i).NumTransitions
                    Text1.Text = Text1.Text & "    * Transition " & AutoFeatures.TerrainFeatures(i).Transitions(j).Name & " - Type " & AutoFeatures.TerrainFeatures(i).Transitions(j).Type & " : for tiles ("
                    tmp = ""
                    For k = 1 To AutoFeatures.TerrainFeatures(i).Transitions(j).NumTiles
                        tmp = tmp & AutoFeatures.TerrainFeatures(i).Transitions(j).TileList(k) & ","
                    Next k
                    tmp = Left$(tmp, Len(tmp) - 1) & "), use tiles : ("
                    For k = 1 To 15
                        tmp = tmp & AutoFeatures.TerrainFeatures(i).Transitions(j).TransTiles(k) & ","
                    Next k
                    tmp = Left$(tmp, Len(tmp) - 1) & ")"
                    Text1.Text = Text1.Text & tmp & NL
                Next j
                Text1.Text = Text1.Text & NL
            End If
        Next i
    End If
    
    If AutoFeatures.NumFiles > 0 Or AutoFeatures.NumABFiles > 0 Then
        Text1.Text = Text1.Text & "FEATURES AUTOTILING" & NL & NL
        If AutoFeatures.NumFiles > 0 Then
            Text1.Text = Text1.Text & "  - Rivers and Roads tiling method : " & AutoFeatures.RiverRoadMethod & NL
            Text1.Text = Text1.Text & "  - TDF Files" & NL
            For k = 1 To AutoFeatures.NumFiles
                Text1.Text = Text1.Text & "    * " & AutoFeatures.FileNames(k) & NL
            Next k
            Text1.Text = Text1.Text & NL
        End If
        If AutoFeatures.NumABFiles > 0 Then
            Text1.Text = Text1.Text & "  - CSV Airbase Files" & NL
            For k = 1 To AutoFeatures.NumABFiles
                Text1.Text = Text1.Text & "    * " & AutoFeatures.FileABNames(k) & NL
            Next k
            Text1.Text = Text1.Text & NL
        End If
        If AutoFeatures.NumABFiles > 0 Then
            Text1.Text = Text1.Text & "  - Airbase definitions" & NL
            For k = 1 To UBound(DefAirbases)
                Text1.Text = Text1.Text & "    * Airbase type " & k & " : Coords=(" & DefAirbases(k).XStart & "," & DefAirbases(k).YStart & "-" & DefAirbases(k).XEnd & "," & DefAirbases(k).YEnd & ") / Type list=("
                tmp = ""
                For j = 1 To UBound(DefAirbases(k).TypeAB)
                    tmp = tmp & DefAirbases(k).TypeAB(j) & ", "
                Next j
                tmp = Left$(tmp, Len(tmp) - 2)
                Text1.Text = Text1.Text & tmp & ")" & NL
            Next k
        End If
        Text1.Text = Text1.Text & NL
        
        If AutoFeatures.NumTerrain >= 0 Then
            For j = 0 To AutoFeatures.NumTerrain
                Text1.Text = Text1.Text & "  - TerrainFeatures Section: " & AutoFeatures.FeatName(j) & NL & NL
                If AutoFeatures.NumABFiles > 0 Then
                    Text1.Text = Text1.Text & "    * City Base Tile = " & AutoFeatures.TerrainFeatures(j).CityBaseTile & NL
                End If
                
                If AutoFeatures.NumABFiles > 0 Then
                    Text1.Text = Text1.Text & "    * Airbases : " & NL
                    For i = 1 To UBound(AutoFeatures.TerrainFeatures(j).AirBases)
                        Text1.Text = Text1.Text & "      Type " & i & " : Leveling="
                        Select Case AutoFeatures.TerrainFeatures(j).AirBases(i).Level
                            Case 0:
                                Text1.Text = Text1.Text & "None"
                            Case 1:
                                Text1.Text = Text1.Text & "High"
                            Case 2:
                                Text1.Text = Text1.Text & "Low"
                        End Select
                        Text1.Text = Text1.Text & ", Tiles=("
                        tmp = ""
                        For k = 1 To UBound(AutoFeatures.TerrainFeatures(j).AirBases(i).TileList)
                            tmp = tmp & AutoFeatures.TerrainFeatures(j).AirBases(i).TileList(k) & ", "
                        Next k
                        Text1.Text = Text1.Text & Left$(tmp, Len(tmp) - 2) & ")" & NL
                    Next i
                End If
                Text1.Text = Text1.Text & NL & NL
            Next j
        End If
        
        If AutoFeatures.NumFiles > 0 Then
            k = 0
            For i = 0 To 5000
                If FeatureTiles(i, 0) <> -1 Then
                    k = k + 1
                    Text1.Text = Text1.Text & "  - Tile " & i & " -> "
                    tmp = ""
                    For j = 0 To 255
                        If FeatureTiles(i, j) <> -1 Then tmp = tmp & j & "/" & FeatureTiles(i, j) & " "
                    Next j
                    DoEvents
                    Text1.Text = Text1.Text & tmp & NL & NL
                    If k = 5 Then Exit For
                End If
            Next i
            Text1.Text = Text1.Text & "  - etc." & NL & NL
            DoEvents
        End If
    End If
    
    If RunBmpRules Or RunCateRules Then
        If Rules.BmpFileName <> "" Then
            Text1.Text = Text1.Text & "BmpFileName : " & Rules.BmpFileName & " (" & IIf(Rules.UpdateOceanTiles = 1, "update ocean tiles)", "ignore ocean tiles)") & NL
            Text1.Text = Text1.Text & NL
        End If
        For k = 1 To Rules.NumSections
            Text1.Text = Text1.Text & "RULES SECTION " & k & " : " & Rules.RuleSections(k).Name & NL
            Text1.Text = Text1.Text & NL
            
            If Rules.RuleSections(k).R >= 0 Then
                Text1.Text = Text1.Text & "  - Color rule (R,G,B) : " & Rules.RuleSections(k).R & "," & Rules.RuleSections(k).G & "," & Rules.RuleSections(k).B & NL
                Text1.Text = Text1.Text & NL
            End If
            
            Text1.Text = Text1.Text & "  - Force Fog rules" & NL
            If Rules.RuleSections(k).NumFogRules > 0 Then
                For i = 1 To Rules.RuleSections(k).NumFogRules
                    Text1.Text = Text1.Text & "    * Force fog values (X1-X2-X3) to : " & IIf(Rules.RuleSections(k).FogRules(i).Fog = "*", "Original", Rules.RuleSections(k).FogRules(i).Fog) & "-" & IIf(Rules.RuleSections(k).FogRules(i).Unknown1 = "*", "Original", Rules.RuleSections(k).FogRules(i).Unknown1) & "-" & IIf(Rules.RuleSections(k).FogRules(i).Unknown2 = "*", "Original", Rules.RuleSections(k).FogRules(i).Unknown2) & NL
                Next i
            Else
                Text1.Text = Text1.Text & "    * None" & NL
            End If
            Text1.Text = Text1.Text & NL
            
            Text1.Text = Text1.Text & "  - Tile Rules" & NL
            If Rules.RuleSections(k).NumTileRules > 0 Then
                For i = 1 To Rules.RuleSections(k).NumTileRules
                    tmp = "    * If " & Rules.RuleSections(k).TileRules(i).MinAlt & " < Alt < " & Rules.RuleSections(k).TileRules(i).MaxAlt & " and tile is "
                    If Rules.RuleSections(k).TileRules(i).OrigTileList(1) = "*" Then
                        If UBound(Rules.RuleSections(k).TileRules(i).OrigTileList) > 1 Then
                            tmp = tmp & "(Any, except "
                            For j = 2 To UBound(Rules.RuleSections(k).TileRules(i).OrigTileList)
                                tmp = tmp & Rules.RuleSections(k).TileRules(i).OrigTileList(j) & ","
                            Next j
                            tmp = Left$(tmp, Len(tmp) - 1) & ")"
                        Else
                            tmp = tmp & "-Any-"
                        End If
                    Else
                        tmp = tmp & " in list ("
                        For j = 1 To UBound(Rules.RuleSections(k).TileRules(i).OrigTileList)
                            tmp = tmp & Rules.RuleSections(k).TileRules(i).OrigTileList(j) & ","
                        Next j
                        tmp = Left$(tmp, Len(tmp) - 1) & ")"
                    End If
                    If Rules.RuleSections(k).TileRules(i).TileList(1) = "*" Then
                        tmp = tmp & ", " & NL & "      then do not replace tile, and " & IIf(Rules.RuleSections(k).TileRules(i).Fog = "*", "original X1", "X1=" & Rules.RuleSections(k).TileRules(i).Fog) & ", " & IIf(Rules.RuleSections(k).TileRules(i).Unknown1 = "*", "original X2", "X2=" & Rules.RuleSections(k).TileRules(i).Unknown1) & ", " & IIf(Rules.RuleSections(k).TileRules(i).Unknown2 = "*", "original X3", "X3=" & Rules.RuleSections(k).TileRules(i).Unknown2)
                    Else
                        tmp = tmp & ", " & NL & "      then use tiles ("
                        For j = 1 To UBound(Rules.RuleSections(k).TileRules(i).TileList)
                            tmp = tmp & Rules.RuleSections(k).TileRules(i).TileList(j) & ","
                        Next j
                        tmp = Left$(tmp, Len(tmp) - 1) & ")" & ", with " & IIf(Rules.RuleSections(k).TileRules(i).Fog = "*", "original X1", "X1=" & Rules.RuleSections(k).TileRules(i).Fog) & ", " & IIf(Rules.RuleSections(k).TileRules(i).Unknown1 = "*", "original X2", "X2=" & Rules.RuleSections(k).TileRules(i).Unknown1) & ", " & IIf(Rules.RuleSections(k).TileRules(i).Unknown2 = "*", "original X3", "X3=" & Rules.RuleSections(k).TileRules(i).Unknown2)
                    End If
                    Text1.Text = Text1.Text & tmp & NL
                    If Rules.RuleSections(k).TileRules(i).AltModType <> "*" Then
                        tmp = "      and modify altitude by "
                        If Rules.RuleSections(k).TileRules(i).AltModType = "F" Then
                            tmp = tmp & "forcing its value to " & Rules.RuleSections(k).TileRules(i).AltModValue & " feet"
                        Else
                            tmp = tmp & " relatively increasing its value by " & Rules.RuleSections(k).TileRules(i).AltModValue & IIf(Rules.RuleSections(k).TileRules(i).AltModValueType = "%", " %", " feet") & IIf(Rules.RuleSections(k).TileRules(i).AltModValueType = "G", " keeping altitude above 1", "")
                        End If
                        Text1.Text = Text1.Text & tmp & NL
                    End If
                Next i
            Else
                Text1.Text = Text1.Text & "    * None" & NL
            End If
            Text1.Text = Text1.Text & NL
            
            Text1.Text = Text1.Text & "  - Rand Altitude Rules" & NL
            If Rules.RuleSections(k).NumAltRules > 0 Then
                For i = 1 To Rules.RuleSections(k).NumAltRules
                    Text1.Text = Text1.Text & "    * If " & Rules.RuleSections(k).AltRules(i).MinAlt & " < Alt < " & Rules.RuleSections(k).AltRules(i).MaxAlt & ", then randomize altitude from " & Rules.RuleSections(k).AltRules(i).Min & " % to " & Rules.RuleSections(k).AltRules(i).Max & " %" & NL
                Next i
            Else
                Text1.Text = Text1.Text & "    * None" & NL
            End If
            Text1.Text = Text1.Text & NL
        Next k
    End If
    
    If SaveBmp Then
        Text1.Text = Text1.Text & "SAVING TERRAIN AS A BMP :" & NL & NL
        Text1.Text = Text1.Text & "  - Under file name " & BmpToSave.FileName & NL & NL
        For i = 0 To 500
            If SaveImageColors(BmpToSave.TileNum(i)).R <> 0 Or SaveImageColors(BmpToSave.TileNum(i)).G <> 0 Or SaveImageColors(BmpToSave.TileNum(i)).B <> 0 Then
                Text1.Text = Text1.Text & "  - Tile " & i & " -> " & SaveImageColors(BmpToSave.TileNum(i)).R & "," & SaveImageColors(BmpToSave.TileNum(i)).G & "," & SaveImageColors(BmpToSave.TileNum(i)).B & NL
            End If
            DoEvents
        Next i
        Text1.Text = Text1.Text & "  - etc." & NL
    End If
    
    If RunTrnImport Then
        Text1.Text = Text1.Text & "IMPORT TRN FILES SECTION :" & NL & NL
        Text1.Text = Text1.Text & "  - Use <tref> section : " & IIf(UseTrefSection, "Yes", "No") & NL
        If AutoFeatures.CorrespFileName <> "" Then
            Text1.Text = Text1.Text & "  - Correspondance file : " & AutoFeatures.CorrespFileName & NL
        End If
        Text1.Text = Text1.Text & "  - Texture Index file : " & AutoFeatures.TextureFileName & NL
        Text1.Text = Text1.Text & "  - Offset X : " & AutoFeatures.TRNOffsetX & NL
        Text1.Text = Text1.Text & "  - Offset Y : " & AutoFeatures.TRNOffsetY & NL
        For i = 1 To AutoFeatures.NumTrnFiles
            Text1.Text = Text1.Text & "  - TRN File : " & AutoFeatures.TrnFileNames(i) & NL
            DoEvents
        Next i
    End If
    Screen.MousePointer = vbNormal
End Sub

Function LoadTiles(FileName As String) As Boolean
'Loads the L2 file
Dim fdl As Integer          'File descriptors
Dim FileSizel As Long       'Size of the files

    'We size the Tiles array
    FileSizel = FileLen(FileName)
    ReDim Tiles(0 To FileSizel \ 7 - 1)
    
    'Then we open the l2 terrain file
    fdl = FreeFile
    On Error Resume Next
    Open FileName For Binary Access Read As #fdl
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted.", vbCritical + vbOKOnly, "Warning"
        LoadTiles = False
        Exit Function
    End If
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    lblShowAction.Caption = "Loading L2 (" & FileName & ") data ..."
    lblShowAction.Visible = True
    DoEvents
    EnableInput False
    'Now we load all l2 data in Tiles structure
    Get #fdl, , Tiles
    Close #fdl
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    LoadTiles = True
    Screen.MousePointer = vbNormal

End Function

Function LoadRegions(FileName As String) As Boolean
'Loads the O2 file
Dim fdo As Integer                              'File descriptors
Dim FileSizeo As Long                 'Size of the files
Dim i As Long

    'We size the Regions array
    FileSizeo = FileLen(FileName)
    TerrainSize = Sqr(FileSizeo \ 4)
    ReDim Regions(0 To TerrainSize * TerrainSize - 1)
    
    On Error Resume Next
    'We first open the .o2 reference file
    fdo = FreeFile
    Open FileName For Binary Access Read As #fdo
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted", vbCritical, "Warning"
        LoadRegions = False
        Exit Function
    End If
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    lblShowAction.Caption = "Loading O2 (" & FileName & ") data ..."
    lblShowAction.Visible = True
    DoEvents
    EnableInput False
    Get #fdo, , Regions
    Close #fdo
    For i = 0 To UBound(Regions)
        'Each tile is 7 bytes, so to match with Tiles data, we do :
        Regions(i) = Regions(i) \ (7)
        DoEvents
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    LoadRegions = True
    Screen.MousePointer = vbNormal

End Function

Function LoadBmp(FileName As String) As Boolean
'Loads a BMP file
Dim fd As Integer
Dim Bytes() As Byte
Dim NumLong As Long
Dim NumByte As Byte
Dim NumInteger As Integer
Dim Chaine As String
Dim i As Integer, j As Integer
Dim Percent As Long
Dim FalseImage As Boolean
Dim Message As String

    FalseImage = False
    fd = FreeFile
    On Error Resume Next
    Open FileName For Binary As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted.", vbCritical + vbOKOnly, "Warning"
        LoadBmp = False
        Exit Function
    End If
    On Error GoTo 0
    lblShowAction.Caption = "Loading bitmap (" & FileName & ") ..."
    lblPercAction.Caption = "0 %"
    picAction.Cls
    UpdatePercentBar picAction, 0
    DoEvents
    picAction.FillColor = &HFF00&
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.Visible = True
    EnableInput False
    Screen.MousePointer = vbHourglass
    'Read the Header
    Chaine = String$(2, " ")
    Get #fd, , Chaine
    ImageHeader.Signature = Chaine
    Get #fd, , NumLong
    ImageHeader.FileSize = NumLong
    Get #fd, , NumLong
    Get #fd, , NumLong
    ImageHeader.DataOffset = NumLong
    'Read info header
    Get #fd, , NumLong
    ImageHeader.HeaderSize = NumLong
    Get #fd, , NumLong
    ImageHeader.ImageWidth = NumLong
    Get #fd, , NumLong
    ImageHeader.ImageHeight = NumLong
    Get #fd, , NumInteger
    ImageHeader.Planes = NumInteger
    Get #fd, , NumInteger
    ImageHeader.BitCount = NumInteger
    Get #fd, , NumLong
    ImageHeader.Compression = NumLong
    Get #fd, , NumLong
    ImageHeader.ImageSize = NumLong
    Get #fd, , NumLong
    ImageHeader.HorResolution = NumLong
    Get #fd, , NumLong
    ImageHeader.VerResolution = NumLong
    Get #fd, , NumLong
    ImageHeader.ColorUsed = NumLong
    Get #fd, , NumLong
    ImageHeader.ColorsImportant = NumLong
    'Do some verifications
    If ImageHeader.Signature <> "BM" Then FalseImage = True
    If ImageHeader.ImageHeight <> ImageHeader.ImageWidth Then FalseImage = True
    If ImageHeader.ImageHeight <> 1024 And ImageHeader.ImageHeight <> 2048 And ImageHeader.ImageHeight <> 4096 Then FalseImage = True
    If ImageHeader.BitCount <> 8 Then FalseImage = True
    If ImageHeader.Compression <> 0 Then FalseImage = True
    If FalseImage = True Then
        Screen.MousePointer = vbNormal
        Message = "Your bitmap is not of the expected format. Action canceled. Details :" & Chr$(13) & Chr$(10)
        Message = Message & "Signature = " & ImageHeader.Signature & " (waiting for BM)" & Chr$(13) & Chr$(10)
        Message = Message & "Height = " & ImageHeader.ImageHeight & " (waiting for 1024, 2048 or 4096)" & Chr$(13) & Chr$(10)
        Message = Message & "Width = " & ImageHeader.ImageWidth & " (waiting for 1024, 2048 or 4096)" & Chr$(13) & Chr$(10)
        Message = Message & "Height and width should be the same" & Chr$(13) & Chr$(10)
        Message = Message & "Bitcount = " & ImageHeader.BitCount & " (waiting for 8)" & Chr$(13) & Chr$(10)
        Message = Message & "Compression = " & ImageHeader.Compression & " (waiting for 0)"
        MsgBox Message, vbCritical + vbOKOnly, "Warning"
        Close #fd
        EnableInput True
        LoadBmp = False
        Exit Function
    End If
    DoEvents
    'Read Colors Table
    ReDim Bytes(0 To 3)
    For i = 0 To 255
        Get #fd, , Bytes
        ImageColors(i).B = Bytes(0)
        ImageColors(i).G = Bytes(1)
        ImageColors(i).R = Bytes(2)
        DoEvents
    Next i
    'Read Image Data
    lblShowAction.Caption = "Loading bitmap (" & FileName & ") : reversing lines ..."
    DoEvents
    ReDim ImageDataTmp(0 To ImageHeader.ImageWidth - 1, 0 To ImageHeader.ImageHeight - 1)
    ReDim ImageData(0 To ImageHeader.ImageWidth - 1, 0 To ImageHeader.ImageHeight - 1)
    Get #fd, , ImageDataTmp
    'Reverse the image from bottom to top
    For j = 0 To ImageHeader.ImageHeight - 1
        For i = 0 To ImageHeader.ImageWidth - 1
            ImageData(i, j) = ImageDataTmp(i, ImageHeader.ImageHeight - j - 1)
        Next i
        Percent = j / (ImageHeader.ImageHeight - 1) * 100
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = "" & Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next j
    'Free some memory
    Erase ImageDataTmp
    Close #fd
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    LoadBmp = True
End Function

Function SaveFile(FileName As String) As Boolean  'Save the l2 data into a new file
'Saves data (L2 and O2)
Dim fd As Integer
Dim i As Long, Max As Long, Percent As Long
Dim char1 As String, char2 As String
Dim O2Created As Boolean

    If TerrainCalculated Then
        TranslateTerrainIntoTiles
        TerrainCalculated = False
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    Max = UBound(Tiles)
    fd = FreeFile
    On Error Resume Next
    If Dir$(FileName, vbNormal) <> "" Then Kill FileName
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reseting " & FileName & ". Action aborted", vbCritical, "Warning"
        SaveFile = False
        Exit Function
    End If
    Open FileName For Binary As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted", vbCritical, "Warning"
        SaveFile = False
        Exit Function
    End If
    On Error GoTo 0
    'We here try to avoid that the user starts something else while we're writing the file
    EnableInput False
    lblShowAction.Caption = "Saving L2 data (" & FileName & ") ..."
    lblShowAction.Visible = True
    DoEvents
    Put #fd, , Tiles
    Close #fd
    
    'Now we want to create the O2 file
    O2Created = True
    fd = FreeFile
    char1 = Left$(FileName, Len(FileName) - 2) & "o2"
    On Error Resume Next
    If Dir$(char1, vbNormal) <> "" Then Kill char1
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reseting " & char1 & ". Action aborted", vbCritical, "Warning"
        SaveFile = False
        Exit Function
    End If
    Open char1 For Binary As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & char1 & ". Action aborted", vbCritical, "Warning"
        SaveFile = False
        Exit Function
    End If
    On Error GoTo 0
    'Regions was pointing to an array
    'we make it back pointing to a file and as each tile is 7 bytes ...
    lblShowAction.Caption = "Saving O2 data (" & char1 & ") ..."
    DoEvents
    For i = 0 To UBound(Regions)
        Regions(i) = Regions(i) * 7
        DoEvents
    Next i
    DoEvents
    Put #fd, , Regions
    Close #fd
    O2Created = True
    'And we do the reverse in case user start updating again
    For i = 0 To UBound(Regions)
        Regions(i) = Regions(i) \ 7
        DoEvents
    Next i
    
    'Back to normal state
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    MsgBox "L2 modified data is now saved under " & FileName & " file" & IIf(O2Created, ". O2 file created as " & char1, ""), vbInformation, "Information"
    SaveFile = True
    Exit Function
    
Problem:
    Screen.MousePointer = vbNormal
    MsgBox "Error #" & Err & " (" & Error$(Err) & ") occurred while saving data under filename " & FileName & ". Action aborted", vbCritical + vbOKOnly, "Warning"
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    On Error Resume Next
    Close #fd
    SaveFile = False
End Function

Sub ApplyCateRules()    'Will apply tile rules to data loaded into memory
Dim i As Long
Dim j As Integer, k As Integer
Dim updated As Long, Percent As Long
Dim IsUpdated As Boolean
Dim MessageString As String
Dim num1 As Integer, num2 As Integer, dec As Integer, random As Integer
Dim Max As Long

    Screen.MousePointer = vbHourglass
    If TerrainCalculated Then
        TranslateTerrainIntoTiles
        TerrainCalculated = False
    End If
    Randomize (Timer)
    lblShowAction.Caption = "Applying CATE rules on raw L2 data ..."
    lblPercAction.Caption = "0 %"
    UpdatePercentBar picAction, 0
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.Cls
    picAction.FillColor = &HFF0000
    picAction.Visible = True
    DoEvents
    EnableInput False
    updated = 0
    Max = UBound(Tiles)
    
    For i = 0 To Max  'Read all tiles
        IsUpdated = ApplyRules(Tiles(i), 1)
        If IsUpdated Then updated = updated + 1
        'Give some feedback
        If i Mod 100 = 0 Then
            Percent = (i * 100) \ Max
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
        End If
        DoEvents
    Next i
    Screen.MousePointer = vbNormal
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    'Final message
    If Rules.RuleSections(1).NumFogRules > 0 Then
        Modified = True
        MessageString = "Fog values applied to all tiles, and "
        If updated > 0 Then
            MessageString = MessageString & updated & " tiles (" & Format$((updated * 100) \ UBound(Tiles), "0.00") & " %) have been updated by other rules"
            If Not BatchMode Then MsgBox MessageString, vbInformation + vbOKOnly, "Information"
        Else
            MessageString = MessageString & "tiles have not been more updated"
            If Not BatchMode Then MsgBox MessageString, vbInformation + vbOKOnly, "Information"
        End If
    Else
        If updated > 0 Then
            Modified = True
            MessageString = "" & updated & " tiles (" & Format$((updated * 100) / UBound(Tiles), "0.00") & " %) have been updated"
            If Not BatchMode Then MsgBox MessageString, vbInformation + vbOKOnly, "Information"
        Else
            Modified = False
            If Not BatchMode Then MsgBox "No tiles have been updated", vbInformation + vbOKOnly, "Information"
        End If
    End If
    EnableInput True
End Sub

Function ApplyBmpRules() As Boolean
'In this function, we need to translate O2 and L2 data in a format
'that fits with the BMP data
'then apply the rules
Dim nb As Long

    Screen.MousePointer = vbHourglass
    If Not TerrainCalculated Then
        TranslateTilesIntoTerrain
        TerrainCalculated = True
    End If
    nb = ApplyColorRules()
    EnableInput True
    Screen.MousePointer = vbNormal
    If Not BatchMode Then MsgBox "BMP rules have been applied. " & nb & " (" & Format$(nb / (TerrainSize * TerrainSize * 2.56), "0.00") & " %) tiles updated", vbInformation + vbOKOnly, "Information"
    If nb > 0 Then Modified = True
    ApplyBmpRules = True
End Function

Sub TranslateTilesIntoTerrain()
'Will transform L2/O2 raw data in a full grid of tiles
Dim i As Long, j As Long, k As Long
Dim xr As Long, yr As Long, xt As Long, yt As Long
Dim Index As Long
Dim Percent As Long
Dim Max As Long

'Note : regions, and tiles within regions, start at South West corner of map/region
'then go on Eastwards and Northwards, up to North East corner

    lblShowAction.Caption = "Translation of L2/O2 raw data into a terrain grid of tiles ..."
    lblPercAction.Caption = "0 %"
    picAction.FillColor = &HC0C000
    UpdatePercentBar picAction, 0
    picAction.Cls
    lblPercAction.Visible = True
    lblShowAction.Visible = True
    picAction.Visible = True
    EnableInput False
    Screen.MousePointer = vbHourglass
    DoEvents
    ReDim Terrain(0 To TerrainSize * 16 - 1, 0 To TerrainSize * 16 - 1)
    Max = UBound(Regions)
    For i = 0 To Max
        'xr and yr are the coordinates of the region in terrain
        xr = i Mod TerrainSize
        yr = TerrainSize - i \ TerrainSize - 1
        'Now we must load region info (16x16 tiles) into Terrain structure
        For j = 0 To 255
            'xt and yt are coordinates of tile in terrain
            xt = xr * 16 + j Mod 16
            yt = (yr + 1) * 16 - j \ 16 - 1
            'Index is the position of current tile in Tiles structure array
            Index = Regions(i) + j
            'Now we transform a list of tiles into a terrain grid
            Terrain(xt, yt).TileDesc.NumTile = Tiles(Index).NumTile
            Terrain(xt, yt).TileDesc.Altitude = Tiles(Index).Altitude
            Terrain(xt, yt).TileDesc.Fog = Tiles(Index).Fog
            Terrain(xt, yt).TileDesc.Unknow1 = Tiles(Index).Unknow1
            Terrain(xt, yt).TileDesc.Unknow2 = Tiles(Index).Unknow2
        Next j
        'Feedback to user
        If i Mod 20 = 0 Then
            Percent = i * 100 / Max
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = "" & Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
        End If
        DoEvents
    Next i
    lblPercAction.Visible = False
    lblShowAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
End Sub

Sub TranslateTerrainIntoTiles()
'Will revert the full grid data into L2/O2 data
Dim ByteRegions() As Integer
Dim IncludeRegions() As Long
Dim IncludeRegionsBool() As Boolean
Dim i As Long, j As Long, jm As Long, k As Long, Index As Long
Dim xr As Long, yr As Long, xt As Long, yt As Long
Dim Equal As Boolean
Dim Percent As Long
Dim Max As Long

    Screen.MousePointer = vbHourglass
    If FullRegionTest Then
        lblShowAction.Caption = "Full region building ..."
    Else
        lblShowAction.Caption = "Preparing regions ..."
    End If
    lblPercAction.Caption = "0 %"
    UpdatePercentBar picAction, 0
    picAction.Cls
    lblPercAction.Visible = True
    lblShowAction.Visible = True
    picAction.Visible = True
    EnableInput False
    DoEvents

    'Here we build arrays of full regions (i.e. 16x16 tiles of 7 bytes)
    Max = UBound(Regions)
    ReDim ByteRegions(0 To TerrainSize * TerrainSize - 1, 0 To 5 * 256 - 1)
    ReDim IncludeRegions(0 To TerrainSize * TerrainSize - 1)
    ReDim IncludeRegionsBool(0 To TerrainSize * TerrainSize - 1)
    If FullRegionTest Then
        For i = 0 To Max
            xr = i Mod TerrainSize
            yr = TerrainSize - i \ TerrainSize - 1
            For j = 0 To 255
                xt = xr * 16 + j Mod 16
                yt = (yr + 1) * 16 - j \ 16 - 1
                jm = j * 5
                ByteRegions(i, jm) = Terrain(xt, yt).TileDesc.NumTile
                ByteRegions(i, jm + 1) = Terrain(xt, yt).TileDesc.Altitude
                ByteRegions(i, jm + 2) = Terrain(xt, yt).TileDesc.Fog
                ByteRegions(i, jm + 3) = Terrain(xt, yt).TileDesc.Unknow1
                ByteRegions(i, jm + 4) = Terrain(xt, yt).TileDesc.Unknow2
                DoEvents
            Next j
            If i Mod 20 = 0 Then
                Percent = i * 100 / Max
                If Percent Mod 2 = 0 Then
                    lblPercAction.Caption = Percent & " %"
                    UpdatePercentBar picAction, Percent
                End If
            End If
            DoEvents
        Next i
    End If
    
    'Here we compare each full region to others to find duplicate
    If FullRegionTest Then
        lblShowAction.Caption = "Searching duplicate regions ..."
    Else
        lblShowAction.Caption = "Preparing regions ..."
    End If
    lblPercAction.Caption = "0 %"
    UpdatePercentBar picAction, 0
    picAction.Cls
    DoEvents
    Index = 1
    jm = 5 * 256 - 1
    For i = 0 To Max - 1
        If IncludeRegions(i) = 0 Then
            IncludeRegions(i) = Index
            IncludeRegionsBool(i) = True
            If FullRegionTest Then
                For j = i + 1 To Max
                    'We suppose the two regions are equal
                    Equal = True
                    For k = 0 To jm
                        If ByteRegions(j, k) <> ByteRegions(i, k) Then
                            'As soon as 1 byte differ, we get out of the loop
                            Equal = False
                            Exit For
                        End If
                    Next k
                    If Equal Then
                        IncludeRegions(j) = Index
                        IncludeRegionsBool(j) = False
                    End If
                Next j
            End If
            Index = Index + 1
            DoEvents
        End If
        If i Mod 20 = 0 Then
            Percent = i * 100 / Max
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
        End If
    Next i
    'Little test for final region, excluded of the loop
    If FullRegionTest Then
        If IncludeRegions(UBound(IncludeRegions)) = 0 Then
            IncludeRegions(UBound(IncludeRegions)) = Index
            IncludeRegionsBool(UBound(IncludeRegions)) = True
            Index = Index + 1
        Else
            IncludeRegionsBool(UBound(IncludeRegions)) = False
        End If
    Else
        IncludeRegions(UBound(IncludeRegions)) = Index
        IncludeRegionsBool(UBound(IncludeRegions)) = True
        Index = Index + 1
    End If
    
    'And finally we rewrite O2 and L2 as they will be in final file
    lblShowAction.Caption = "Translating terrain grid into L2/O2 raw data ..."
    lblPercAction.Caption = "0 %"
    UpdatePercentBar picAction, 0
    picAction.Cls
    DoEvents
    ReDim Tiles(0 To (Index - 1) * 256 - 1)
    k = 0
    For i = 0 To Max
        'xr and yr are the coordinates of the region in terrain
        xr = i Mod TerrainSize
        yr = TerrainSize - i \ TerrainSize - 1
        'Now we must load region info (16x16 tiles) into Terrain structure
        If IncludeRegionsBool(i) Then
            For j = 0 To 255
                'xt and yt are coordinates of tile in terrain
                xt = xr * 16 + j Mod 16
                yt = (yr + 1) * 16 - j \ 16 - 1
                'We reverse the terrain grid to list of tiles (without duplicates)
                Tiles(k).NumTile = Terrain(xt, yt).TileDesc.NumTile
                Tiles(k).Altitude = Terrain(xt, yt).TileDesc.Altitude
                Tiles(k).Fog = Terrain(xt, yt).TileDesc.Fog
                Tiles(k).Unknow1 = Terrain(xt, yt).TileDesc.Unknow1
                Tiles(k).Unknow2 = Terrain(xt, yt).TileDesc.Unknow2
                k = k + 1
                DoEvents
            Next j
        End If
        If i Mod 20 = 0 Then
            Percent = i * 100 / Max
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
        End If
        'Regions array is updated, ready to be saved
        '(each region is 16x16 tiles , hence the multiplication)
        Regions(i) = (IncludeRegions(i) - 1) * 256
        DoEvents
    Next i

    lblPercAction.Visible = False
    lblShowAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
End Sub

Function ApplyColorRules() As Long
Dim i As Long, j As Long
Dim k As Integer
Dim Percent As Long
Dim updated As Long
Dim IsUpdated As Long

    Randomize (Timer)
    lblShowAction.Caption = "Applying CATE rules on BMP defined areas ..."
    lblPercAction.Caption = "0 %"
    UpdatePercentBar picAction, 0
    picAction.Cls
    lblPercAction.Visible = True
    lblShowAction.Visible = True
    picAction.Visible = True
    DoEvents
    'We match BMP data with color rules, and update tile if necessary
    For j = 0 To TerrainSize * 16 - 1
        For i = 0 To TerrainSize * 16 - 1
            If Terrain(i, j).TileDesc.NumTile <> 0 Or Rules.UpdateOceanTiles = 1 Then 'We skip ocean tiles if needed
                For k = 1 To Rules.NumSections
                    If Rules.RuleSections(k).R = ImageColors(ImageData(i, j)).R Then
                        If Rules.RuleSections(k).G = ImageColors(ImageData(i, j)).G Then
                            If Rules.RuleSections(k).B = ImageColors(ImageData(i, j)).B Then
                                IsUpdated = ApplyRules(Terrain(i, j).TileDesc, k)
                                If IsUpdated Then updated = updated + 1
                                Exit For    'We found a rule matching, no need to look others
                            End If
                        End If
                    End If
                    DoEvents
                Next k
            End If
        Next i
        If j Mod 10 = 0 Then
            Percent = j * 100 / (TerrainSize * 16)
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
        End If
        DoEvents
    Next j
    lblPercAction.Visible = False
    lblShowAction.Visible = False
    picAction.Visible = False
    ApplyColorRules = updated
End Function

Function ApplyFeatures() As Boolean
Dim nb1 As Long, nb2 As Long, nb3 As Long

    Screen.MousePointer = vbHourglass
    Randomize (Timer)
    If Not TerrainCalculated Then
        TranslateTilesIntoTerrain
        TerrainCalculated = True
    End If
    Screen.MousePointer = vbHourglass
    EnableInput False
    DoEvents
    If RunTransitions Then nb1 = ApplyTransitions()
    
    If RunFeatures Then
        If AutoFeatures.NumRoads > 0 Then ApplyRoads
        If AutoFeatures.NumRivers > 0 Then ApplyRivers
        UpdatedCities = 0
        If AutoFeatures.NumCities > 0 Then ApplyCities
        If AutoFeatures.NumRoads > 0 Or AutoFeatures.NumRivers > 0 Then nb2 = MergeFeatures()
    End If
    If AutoFeatures.NumABFiles > 0 Then nb3 = ApplyAirbases()
    
    EnableInput True
    Screen.MousePointer = vbNormal
    If Not BatchMode Then MsgBox "Features applied : " & (nb1 + nb2 + nb3 + UpdatedCities) & " (" & Format$(((nb1 + nb2 + nb3 + UpdatedCities) * 100) / (TerrainSize * TerrainSize * 256), "0.00") & " %) tiles updated", vbInformation + vbOKOnly, "Information"
    If nb1 + nb2 + nb3 + UpdatedCities > 0 Then Modified = True
    ApplyFeatures = True
End Function

Function ApplyTransitions() As Long
Dim i As Long, j As Long
Dim k As Integer, l As Integer, m As Integer
Dim Percent As Long, TerTileSize As Long
Dim NE As Integer, NW As Integer, SE As Integer, SW As Integer
Dim updated As Long
Dim NumFeat As Integer

    lblShowAction.Caption = "Applying transitions to terrain ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    TerTileSize = TerrainSize * 16 - 1
    For i = 0 To TerTileSize
        For j = 0 To TerTileSize
            k = GiveTerrainFromTile(Terrain(i, j).TileDesc.NumTile)
            If AutoFeatures.TerrainFeatures(k).NumTransitions > 0 Then
                For l = 1 To AutoFeatures.TerrainFeatures(k).NumTransitions
                    If IsTileToCheck(k, l, i, j) Then   'We check if should make a transition on that tile
                        'We then search each corner value
                        NW = GiveCorner(1, i, j, k, TerTileSize, l)
                        NE = GiveCorner(2, i, j, k, TerTileSize, l)
                        SW = GiveCorner(3, i, j, k, TerTileSize, l)
                        SE = GiveCorner(4, i, j, k, TerTileSize, l)
                        
                        'Two different methods of calculations
                        If AutoFeatures.TerrainFeatures(k).Transitions(l).Type = 1 Then
                            NumFeat = 15 - (NW + NE * 2 + SW * 4 + SE * 8)
                        Else
                            NumFeat = NW + NE * 2 + SW * 4 + SE * 8
                        End If
                        If NumFeat <> 0 Then
                            Terrain(i, j).TileDesc.NumTile = AutoFeatures.TerrainFeatures(k).Transitions(l).TransTiles(NumFeat)
                            updated = updated + 1
                        Else
                            'In type 2, we clear standalone tiles
                            If AutoFeatures.TerrainFeatures(k).Transitions(l).Type = 2 Then
                                Terrain(i, j).TileDesc.NumTile = AutoFeatures.TerrainFeatures(k).Transitions(l).TileList(1)
                                updated = updated + 1
                            End If
                        End If
                    End If
                Next l
            End If
            DoEvents
        Next j
        Percent = (i * 100) / TerTileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    ApplyTransitions = updated
End Function

Function IsTileToCheck(CurTerrain, CurTransition, i As Long, j As Long) As Boolean
'When doing transitions, search if a given tile should be checked for that transition, based on the surrounding terrain types
Dim Terrain1 As Integer, Terrain2 As Integer, Terrain3 As Integer, Terrain4 As Integer
Dim Terrain5 As Integer, Terrain6 As Integer, Terrain7 As Integer, Terrain8 As Integer
Dim IsInTrans As Boolean
Dim k As Integer
Dim TerTileSize As Long

    TerTileSize = TerrainSize * 16 - 1
    Terrain1 = -1
    Terrain2 = -1
    Terrain3 = -1
    Terrain4 = -1
    Terrain5 = -1
    Terrain6 = -1
    Terrain7 = -1
    Terrain8 = -1
    If j > 0 Then Terrain1 = GiveTerrainFromTile(Terrain(i, j - 1).TileDesc.NumTile)
    If i < TerTileSize Then Terrain2 = GiveTerrainFromTile(Terrain(i + 1, j).TileDesc.NumTile)
    If j < TerTileSize Then Terrain3 = GiveTerrainFromTile(Terrain(i, j + 1).TileDesc.NumTile)
    If i > 0 Then Terrain4 = GiveTerrainFromTile(Terrain(i - 1, j).TileDesc.NumTile)
    If j > 0 And i < TerTileSize Then Terrain5 = GiveTerrainFromTile(Terrain(i + 1, j - 1).TileDesc.NumTile)
    If i < TerTileSize And j < TerTileSize Then Terrain6 = GiveTerrainFromTile(Terrain(i + 1, j + 1).TileDesc.NumTile)
    If i > 0 And j < TerTileSize Then Terrain7 = GiveTerrainFromTile(Terrain(i - 1, j + 1).TileDesc.NumTile)
    If j > 0 And i > 0 Then Terrain8 = GiveTerrainFromTile(Terrain(i - 1, j - 1).TileDesc.NumTile)
    If Terrain1 = CurTerrain Then
        If Terrain2 = CurTerrain Then
            If Terrain3 = CurTerrain Then
                If Terrain4 = CurTerrain Then
                    If Terrain5 = CurTerrain Then
                        If Terrain6 = CurTerrain Then
                            If Terrain7 = CurTerrain Then
                                If Terrain8 = CurTerrain Then
                                    'If all tiles around are of the same terrain type
                                    'No need to do a transition
                                    IsTileToCheck = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'We then check if we find a tile around (among the 8 surrounding tiles) which is part of the transition definition
    If j > 0 Then
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i, j - 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If
    
    If i < TerTileSize Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i + 1, j).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If

    If j < TerTileSize Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i, j + 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If
    
    If i > 0 Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i - 1, j).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If

    If j > 0 And i < TerTileSize Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i + 1, j - 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If
    
    If i < TerTileSize And j < TerTileSize Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i + 1, j + 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If

    If i > 0 And j < TerTileSize Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i - 1, j + 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If

    If j > 0 And i > 0 Then
        IsInTrans = False
        For k = 1 To AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).NumTiles
            If AutoFeatures.TerrainFeatures(CurTerrain).Transitions(CurTransition).TileList(k) = Terrain(i - 1, j - 1).TileDesc.NumTile Then
                IsTileToCheck = True
                Exit Function
            End If
        Next k
    End If
    
    IsTileToCheck = False
End Function

Function LoadAirbaseFiles() As Boolean
'Will load all CSV files describing airbases
Dim fd As Integer
Dim i As Integer, j As Integer, k As Integer, l As Integer, Rep As Integer
Dim ReadSize As Long, TotalSize As Long, TempRead As Long
Dim tmp As String, tmpstring As String, tmpname As String
Dim Percent As Long
Dim pos1 As Long, pos2 As Long, Pos3 As Long
Dim AllOK As Boolean

    For i = 1 To AutoFeatures.NumABFiles
        If Dir$(AutoFeatures.FileABNames(i)) = "" Then
            MsgBox "CSV File " & AutoFeatures.FileABNames(i) & " does not exist. Please check your conf file.", vbCritical, "Warning"
            LoadAirbaseFiles = False
            Exit Function
        End If
    Next i
    For i = 1 To AutoFeatures.NumABFiles
        TotalSize = TotalSize + FileLen(AutoFeatures.FileABNames(i))
    Next i
    ReadSize = 0
    AutoFeatures.NumAirbases = 0
    Erase AutoFeatures.AllAirbases
    Screen.MousePointer = vbHourglass
    UpdatePercentBar picAction, 0
    picAction.Cls
    picAction.FillColor = &HFF00&
    picAction.Visible = True
    lblShowAction.Caption = ""
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    DoEvents
    EnableInput False
    For i = 1 To AutoFeatures.NumABFiles
        lblShowAction.Caption = "Loading Airbase information from " & AutoFeatures.FileABNames(i)
        DoEvents
        fd = FreeFile
        On Error Resume Next
        Open AutoFeatures.FileABNames(i) For Input As #fd
        If Err <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Error #" & Err & " (" & Error$(Err) & ") while reading TDF file " & AutoFeatures.FileNames(i) & ". Do you wish to go on anyway ?", vbCritical + vbYesNo, "Warning"
            If Rep = vbNo Then
                LoadAirbaseFiles = False
                Exit Function
            End If
        End If
        On Error GoTo 0
        TempRead = 0
        Line Input #fd, tmp
        Line Input #fd, tmp
        Do While Not EOF(fd)
            Line Input #fd, tmp
            pos1 = 0
            For j = 1 To 8
                pos2 = InStr(pos1 + 1, tmp, ",")
                If pos2 <= 0 Then
                    pos2 = InStr(pos1 + 1, tmp, ";")
                End If
                tmpstring = UCase$(Trim$(Mid$(tmp, pos1 + 1, pos2 - pos1 - 1)))
                Select Case j
                    Case 1: 'Name : to keep, in case
                        tmpname = tmpstring
                    Case 2: 'Type : only airbases
                        If tmpstring <> "AIRBASE" Then
                            Exit For
                        Else
                            AutoFeatures.NumAirbases = AutoFeatures.NumAirbases + 1
                            If AutoFeatures.NumAirbases = 1 Then
                                ReDim AutoFeatures.AllAirbases(1 To 1)
                            Else
                                ReDim Preserve AutoFeatures.AllAirbases(1 To AutoFeatures.NumAirbases)
                            End If
                            AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).Name = tmpname
                        End If
                    Case 3: 'SubType : to keep
                        AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).TypeAirbase = Left$(tmpstring, InStr(1, tmpstring, " ") - 1)
                        For k = 1 To UBound(DefAirbases)
                            For l = 1 To UBound(DefAirbases(k).TypeAB)
                                If DefAirbases(k).TypeAB(l) = AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).TypeAirbase Then
                                    AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).TypeCate = k
                                    Exit For
                                End If
                            Next l
                            If AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).TypeCate > 0 Then Exit For
                        Next k
                    Case 4: 'ID : to keep, in case
                        AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).Id = Val(tmpstring)
                    Case 5: 'Objective : does not matter
                    Case 6: 'X : to keep
                        AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).x = Val(tmpstring)
                    Case 7: 'Y : to keep
                        AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).y = TerrainSize * 16 - Val(tmpstring) - 1
                        'AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).y = Val(tmpstring)
                    Case 8: 'Z : to keep, in case
                        AutoFeatures.AllAirbases(AutoFeatures.NumAirbases).z = Val(tmpstring)
                End Select
                pos1 = pos2
            Next j
            TempRead = Seek(fd)
            Percent = (ReadSize + TempRead) * 100 \ TotalSize
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = "" & Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
            DoEvents
        Loop
        On Error Resume Next
        Close #fd
        ReadSize = ReadSize & FileLen(AutoFeatures.FileABNames(i))
    Next i
    AllOK = True
    For i = 1 To AutoFeatures.NumAirbases
        'Debug.Print "Type=" & AutoFeatures.AllAirbases(i).TypeAirbase & " - ID=" & AutoFeatures.AllAirbases(i).Id & " - X=" & AutoFeatures.AllAirbases(i).X & " - Y=" & AutoFeatures.AllAirbases(i).Y & " - Z=" & AutoFeatures.AllAirbases(i).Z
        If AutoFeatures.AllAirbases(i).TypeCate = 0 Then
            AllOK = False
            'MsgBox "Airbase " & AutoFeatures.AllAirbases(i).Name & " of type " & AutoFeatures.AllAirbases(i).TypeAirbase & " has no link in CATE definitions", vbInformation, "Warning"
        End If
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    If Not AllOK And Not BatchMode Then MsgBox "Some airbases in CSV file(s) have no associated definition in CATE conf file", vbInformation, "Warning"
    If Not BatchMode Then MsgBox "Airbase data is now loaded into memory", vbInformation, "Information"
    LoadAirbaseFiles = True
End Function

Function ApplyAirbases() As Long
'Will apply airbase, and level terrain around
Dim i As Long, IndexTile As Integer
Dim x As Long, y As Long
Dim NumTerrain As Integer
Dim High As Long, Low As Long
Dim Percent As Long
Dim updated As Long

    lblShowAction.Caption = "Calculating and leveling airbases ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    For i = 1 To AutoFeatures.NumAirbases
        If AutoFeatures.AllAirbases(i).TypeCate <> 0 Then   'If the AB has a defined type
            'We first search on which terrain the AB is
            NumTerrain = GiveTerrainFromTile(Terrain(AutoFeatures.AllAirbases(i).x, AutoFeatures.AllAirbases(i).y).TileDesc.NumTile)
            'Then we search for highest and lowest terrain around the AB
            If AutoFeatures.TerrainFeatures(NumTerrain).AirBases(AutoFeatures.AllAirbases(i).TypeCate).Level <> 0 Then
                High = 0
                Low = 100000
                For y = AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YStart - 1 To AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YEnd + 1
                    For x = AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XStart - 1 To AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XEnd + 1
                        If x >= 0 And x < TerrainSize * 16 And y >= 0 And y < TerrainSize * 16 Then
                            If Terrain(x, y).TileDesc.Altitude > High Then High = Terrain(x, y).TileDesc.Altitude
                            If Terrain(x, y).TileDesc.Altitude < Low Then Low = Terrain(x, y).TileDesc.Altitude
                        End If
                    Next x
                Next y
                'And here, we assign that altitude to the tiles around the AB, depending on the leveling type
                For y = AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YStart - 1 To AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YEnd + 1
                    For x = AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XStart - 1 To AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XEnd + 1
                        If x >= 0 And x < TerrainSize * 16 And y >= 0 And y < TerrainSize * 16 Then
                            If AutoFeatures.TerrainFeatures(NumTerrain).AirBases(AutoFeatures.AllAirbases(i).TypeCate).Level = 1 Then
                                Terrain(x, y).TileDesc.Altitude = High
                                updated = updated + 1
                            ElseIf AutoFeatures.TerrainFeatures(NumTerrain).AirBases(AutoFeatures.AllAirbases(i).TypeCate).Level = 2 Then
                                Terrain(x, y).TileDesc.Altitude = Low
                                updated = updated + 1
                            End If
                        End If
                    Next x
                Next y
            End If
            'Now, we put the AB tiles into place, from S-W to N-E
            IndexTile = 0
            For y = AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YEnd To AutoFeatures.AllAirbases(i).y + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).YStart Step -1
                For x = AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XStart To AutoFeatures.AllAirbases(i).x + DefAirbases(AutoFeatures.AllAirbases(i).TypeCate).XEnd
                    IndexTile = IndexTile + 1
                    If x >= 0 And x < TerrainSize * 16 And y >= 0 And y < TerrainSize * 16 Then
                        Terrain(x, y).TileDesc.NumTile = AutoFeatures.TerrainFeatures(NumTerrain).AirBases(AutoFeatures.AllAirbases(i).TypeCate).TileList(IndexTile)
                        If AutoFeatures.TerrainFeatures(NumTerrain).AirBases(AutoFeatures.AllAirbases(i).TypeCate).Level = 0 Then
                            updated = updated + 1
                        End If
                    End If
                Next x
            Next y
        End If
        Percent = (i * 100) / AutoFeatures.NumAirbases
        lblShowAction.Caption = Percent & " %"
        UpdatePercentBar picAction, Percent
        DoEvents
    Next i
        
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    ApplyAirbases = updated
End Function

Function LoadFeaturesFiles() As Boolean
'Will load all TDF files for features definitions
Dim fd As Integer
Dim i As Integer, j As Integer
Dim ReadSize As Long, TotalSize As Long, TempRead As Long
Dim NumPaths As Integer
Dim Rep As Integer
Dim tmp As String, tmpstring As String
Dim Percent As Long
Dim pos1 As Long, pos2 As Long, Pos3 As Long
Dim NumCoords As Integer
Dim TerTileSize As Long

    TerTileSize = TerrainSize * 16
    NumPaths = 0
    AutoFeatures.NumCities = 0
    AutoFeatures.NumRoads = 0
    AutoFeatures.NumRivers = 0
    For i = 1 To AutoFeatures.NumFiles
        If Dir$(AutoFeatures.FileNames(i)) = "" Then
            MsgBox "TDF File " & AutoFeatures.FileNames(i) & " does not exist. Please check your conf file.", vbCritical, "Warning"
            LoadFeaturesFiles = False
            Exit Function
        End If
    Next i
    For i = 1 To AutoFeatures.NumFiles
        TotalSize = TotalSize + FileLen(AutoFeatures.FileNames(i))
    Next i
    ReadSize = 0
    Screen.MousePointer = vbHourglass
    UpdatePercentBar picAction, 0
    picAction.Cls
    picAction.FillColor = &HFF00&
    picAction.Visible = True
    lblShowAction.Caption = ""
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    DoEvents
    EnableInput False
    For i = 1 To AutoFeatures.NumFiles
        lblShowAction.Caption = "Loading TDF features from " & AutoFeatures.FileNames(i)
        DoEvents
        fd = FreeFile
        On Error Resume Next
        Open AutoFeatures.FileNames(i) For Input As #fd
        If Err <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Error #" & Err & " (" & Error$(Err) & ") while reading TDF file " & AutoFeatures.FileNames(i) & ". Do you wish to go on anyway ?", vbCritical + vbYesNo, "Warning"
            If Rep = vbNo Then
                LoadFeaturesFiles = False
                Exit Function
            End If
        End If
        On Error GoTo 0
        TempRead = 0
        Do While Not EOF(fd)
            Line Input #fd, tmp
            tmp = Trim$(tmp)
            If Len(tmp) > 4 Then
                pos1 = InStr(1, tmp, " ")
                If pos1 > 0 Then
                    Select Case UCase$(Left$(tmp, pos1 - 1))
                        Case "CITY":
                            AutoFeatures.NumCities = AutoFeatures.NumCities + 1
                            NumCoords = 0
                            If AutoFeatures.NumCities = 1 Then
                                ReDim AutoFeatures.CityPaths(1 To 1)
                                AutoFeatures.CityPaths(1).NumCoords = 0
                            Else
                                ReDim Preserve AutoFeatures.CityPaths(1 To AutoFeatures.NumCities)
                                AutoFeatures.CityPaths(AutoFeatures.NumCities).NumCoords = 0
                            End If
                            tmpstring = ""
                            For j = pos1 + 1 To Len(tmp)
                                If Mid$(tmp, j, 1) = " " Or j = Len(tmp) Then
                                    If j = Len(tmp) Then tmpstring = tmpstring & Right$(tmp, 1)
                                    pos2 = InStr(1, tmpstring, ",")
                                    If pos2 <= 0 Then
                                        'Number of coords
                                        AutoFeatures.CityPaths(AutoFeatures.NumCities).NumCoords = Val(tmpstring)
                                        ReDim AutoFeatures.CityPaths(AutoFeatures.NumCities).Coord(1 To Val(tmpstring))
                                    Else
                                        'A coord
                                        NumCoords = NumCoords + 1
                                        AutoFeatures.CityPaths(AutoFeatures.NumCities).Coord(NumCoords).x = Val(Left$(tmpstring, pos2 - 1))
                                        AutoFeatures.CityPaths(AutoFeatures.NumCities).Coord(NumCoords).y = TerTileSize - Val(Right$(tmpstring, Len(tmpstring) - pos2)) - 1
                                    End If
                                    tmpstring = ""
                                Else
                                    tmpstring = tmpstring & Mid$(tmp, j, 1)
                                End If
                            Next j
                        Case "RIVER":
                            AutoFeatures.NumRivers = AutoFeatures.NumRivers + 1
                            NumCoords = 0
                            If AutoFeatures.NumRivers = 1 Then
                                ReDim AutoFeatures.RiverPaths(1 To 1)
                                AutoFeatures.RiverPaths(1).NumCoords = 0
                            Else
                                ReDim Preserve AutoFeatures.RiverPaths(1 To AutoFeatures.NumRivers)
                                AutoFeatures.RiverPaths(AutoFeatures.NumRivers).NumCoords = 0
                            End If
                            tmpstring = ""
                            For j = pos1 + 1 To Len(tmp)
                                If Mid$(tmp, j, 1) = " " Or j = Len(tmp) Then
                                    If j = Len(tmp) Then tmpstring = tmpstring & Right$(tmp, 1)
                                    pos2 = InStr(1, tmpstring, ",")
                                    If pos2 <= 0 Then
                                        'Number of coords
                                        AutoFeatures.RiverPaths(AutoFeatures.NumRivers).NumCoords = Val(tmpstring)
                                        ReDim AutoFeatures.RiverPaths(AutoFeatures.NumRivers).Coord(1 To Val(tmpstring))
                                    Else
                                        'A coord
                                        NumCoords = NumCoords + 1
                                        AutoFeatures.RiverPaths(AutoFeatures.NumRivers).Coord(NumCoords).x = Val(Left$(tmpstring, pos2 - 1))
                                        AutoFeatures.RiverPaths(AutoFeatures.NumRivers).Coord(NumCoords).y = TerTileSize - Val(Right$(tmpstring, Len(tmpstring) - pos2)) - 1
                                    End If
                                    tmpstring = ""
                                Else
                                    tmpstring = tmpstring & Mid$(tmp, j, 1)
                                End If
                            Next j
                        Case "ROAD":
                            AutoFeatures.NumRoads = AutoFeatures.NumRoads + 1
                            NumCoords = 0
                            If AutoFeatures.NumRoads = 1 Then
                                ReDim AutoFeatures.RoadPaths(1 To 1)
                                AutoFeatures.RoadPaths(1).NumCoords = 0
                            Else
                                ReDim Preserve AutoFeatures.RoadPaths(1 To AutoFeatures.NumRoads)
                                AutoFeatures.RoadPaths(AutoFeatures.NumRoads).NumCoords = 0
                            End If
                            tmpstring = ""
                            For j = pos1 + 1 To Len(tmp)
                                If Mid$(tmp, j, 1) = " " Or j = Len(tmp) Then
                                    If j = Len(tmp) Then tmpstring = tmpstring & Right$(tmp, 1)
                                    pos2 = InStr(1, tmpstring, ",")
                                    If pos2 <= 0 Then
                                        'Number of coords
                                        AutoFeatures.RoadPaths(AutoFeatures.NumRoads).NumCoords = Val(tmpstring)
                                        ReDim AutoFeatures.RoadPaths(AutoFeatures.NumRoads).Coord(1 To Val(tmpstring))
                                    Else
                                        'A coord
                                        NumCoords = NumCoords + 1
                                        AutoFeatures.RoadPaths(AutoFeatures.NumRoads).Coord(NumCoords).x = Val(Left$(tmpstring, pos2 - 1))
                                        AutoFeatures.RoadPaths(AutoFeatures.NumRoads).Coord(NumCoords).y = TerTileSize - Val(Right$(tmpstring, Len(tmpstring) - pos2)) - 1
                                    End If
                                    tmpstring = ""
                                Else
                                    tmpstring = tmpstring & Mid$(tmp, j, 1)
                                End If
                            Next j
                        Case Else:
                            'Impossible : do nothing for now
                    End Select
                Else
                    Screen.MousePointer = vbNormal
                    MsgBox "Format error in TDF file " & AutoFeatures.FileNames(i) & ". Do you wish to go on anyway ?", vbCritical + vbYesNo, "Warning"
                    If Rep = vbNo Then
                        LoadFeaturesFiles = False
                        Close #fd
                        Exit Function
                    Else
                        GoTo EndLoop
                    End If
                End If
            End If
            TempRead = Seek(fd)
            Percent = (ReadSize + TempRead) * 100 \ TotalSize
            If Percent Mod 2 = 0 Then
                lblPercAction.Caption = "" & Percent & " %"
                UpdatePercentBar picAction, Percent
            End If
            DoEvents
        Loop
EndLoop:
        On Error Resume Next
        Close #fd
        ReadSize = ReadSize & FileLen(AutoFeatures.FileNames(i))
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    If Not BatchMode Then MsgBox "Features are now loaded into memory", vbInformation, "Information"
    LoadFeaturesFiles = True
End Function

Sub ApplyRoads()
'Will calculate paths for roads, and update terrain info
'but will not update tile number yet
Dim i As Long, j As Long
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
Dim Percent As Long

    lblShowAction.Caption = "Calculating roads ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    For i = 1 To AutoFeatures.NumRoads
        If AutoFeatures.RiverRoadMethod = 1 Then RecalculatePaths AutoFeatures.RoadPaths(i)
        For j = 1 To AutoFeatures.RoadPaths(i).NumCoords - 1
            x1 = AutoFeatures.RoadPaths(i).Coord(j).x
            y1 = AutoFeatures.RoadPaths(i).Coord(j).y
            x2 = AutoFeatures.RoadPaths(i).Coord(j + 1).x
            y2 = AutoFeatures.RoadPaths(i).Coord(j + 1).y
            CalculatePath x1, y1, x2, y2, "RO"
            DoEvents
        Next j
        
        Percent = (i * 100) / AutoFeatures.NumRoads
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False

End Sub

Sub ApplyRivers()
'Will calculate paths for roads, and update terrain info
'but will not update tile number yet
Dim i As Long, j As Long
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
Dim Percent As Long

    lblShowAction.Caption = "Calculating rivers ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    For i = 1 To AutoFeatures.NumRivers
        If AutoFeatures.RiverRoadMethod = 1 Then RecalculatePaths AutoFeatures.RiverPaths(i)
        For j = 1 To AutoFeatures.RiverPaths(i).NumCoords - 1
            x1 = AutoFeatures.RiverPaths(i).Coord(j).x
            y1 = AutoFeatures.RiverPaths(i).Coord(j).y
            x2 = AutoFeatures.RiverPaths(i).Coord(j + 1).x
            y2 = AutoFeatures.RiverPaths(i).Coord(j + 1).y
            CalculatePath x1, y1, x2, y2, "RI"
            DoEvents
        Next j
        
        Percent = (i * 100) / AutoFeatures.NumRivers
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
End Sub

Sub ApplyCities()
'Will calculate paths for for cities, tile number will be updated now (in UpdatePath function)
Dim i As Long, j As Long
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
Dim Percent As Long

    lblShowAction.Caption = "Calculating cities ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    For i = 1 To AutoFeatures.NumCities
        For j = 1 To AutoFeatures.CityPaths(i).NumCoords - 1
            x1 = AutoFeatures.CityPaths(i).Coord(j).x
            y1 = AutoFeatures.CityPaths(i).Coord(j).y
            x2 = AutoFeatures.CityPaths(i).Coord(j + 1).x
            y2 = AutoFeatures.CityPaths(i).Coord(j + 1).y
            CalculatePath x1, y1, x2, y2, "CI"
            DoEvents
        Next j
        Percent = (i * 100) / AutoFeatures.NumCities
        'If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        'End If
        DoEvents
    Next i
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
End Sub

Sub CalculatePath(xb As Long, yb As Long, xe As Long, ye As Long, Feat As String)
'Will calculate all tiles in a path
'From a beginning and ending tile numbers
'Rivers will try to avoid roads if possible
Dim xc As Long, yc As Long
Dim Path() As strCoord
Dim i As Integer, nb As Integer
Dim TileRoad1 As Integer, TileRoad2 As Integer
Dim TileRiver1 As Integer, TileRiver2 As Integer
Dim TileNum1 As Integer, TileNum2 As Integer
Dim DoRandom As Boolean, UseOldAlgo As Boolean

    ReDim Path(1 To 1)
    Path(1).x = xb
    Path(1).y = yb
    xc = xb
    yc = yb
    Do While xc <> xe Or yc <> ye
        If xc = xe Then 'X is the same
            If ye > yc Then 'We go from north to south
                Do While yc <> ye
                    UpdatePath Path, Feat, 0, xc, yc
                Loop
            Else    'We go from south to north
                Do While yc <> ye
                    UpdatePath Path, Feat, 1, xc, yc
                Loop
            End If
        ElseIf yc = ye Then 'Y is the same
            If xe > xc Then 'We go from west to east
                Do While xc <> xe
                    UpdatePath Path, Feat, 2, xc, yc
                Loop
            Else    'We go from east to west
                Do While xc <> xe
                    UpdatePath Path, Feat, 3, xc, yc
                Loop
            End If
        Else
            If Feat = "RI" Or Feat = "RO" Then
                'We will try both directions and see if one would result in a defined tile
                'We get infirmation from tile N or S as if we were to tile there
                If ye > yc Then
                    If Feat = "RI" Then
                        TileRoad1 = Terrain(xc, yc + 1).Road(1) + Terrain(xc, yc + 1).Road(2) * 2 + Terrain(xc, yc + 1).Road(3) * 4 + Terrain(xc, yc + 1).Road(4) * 8
                        TileRiver1 = 1 * 16 + Terrain(xc, yc + 1).River(2) * 32 + Terrain(xc, yc + 1).River(3) * 64 + Terrain(xc, yc + 1).River(4) * 128
                    Else
                        TileRoad1 = 1 + Terrain(xc, yc + 1).Road(2) * 2 + Terrain(xc, yc + 1).Road(3) * 4 + Terrain(xc, yc + 1).Road(4) * 8
                        TileRiver1 = Terrain(xc, yc + 1).River(1) * 16 + Terrain(xc, yc + 1).River(2) * 32 + Terrain(xc, yc + 1).River(3) * 64 + Terrain(xc, yc + 1).River(4) * 128
                    End If
                    TileNum1 = Terrain(xc, yc + 1).TileDesc.NumTile
                Else
                    If Feat = "RI" Then
                        TileRoad1 = Terrain(xc, yc - 1).Road(1) + Terrain(xc, yc - 1).Road(2) * 2 + Terrain(xc, yc - 1).Road(3) * 4 + Terrain(xc, yc - 1).Road(4) * 8
                        TileRiver1 = Terrain(xc, yc - 1).River(1) * 16 + Terrain(xc, yc - 1).River(2) * 32 + 1 * 64 + Terrain(xc, yc - 1).River(4) * 128
                    Else
                        TileRoad1 = Terrain(xc, yc - 1).Road(1) + Terrain(xc, yc - 1).Road(2) * 2 + 1 * 4 + Terrain(xc, yc - 1).Road(4) * 8
                        TileRiver1 = Terrain(xc, yc - 1).River(1) * 16 + Terrain(xc, yc - 1).River(2) * 32 + Terrain(xc, yc - 1).River(3) * 64 + Terrain(xc, yc - 1).River(4) * 128
                    End If
                    TileNum1 = Terrain(xc, yc - 1).TileDesc.NumTile
                End If
                'We get infirmation from tile E or W as if we were to tile there
                If xe > xc Then
                    If Feat = "RI" Then
                        TileRoad2 = Terrain(xc + 1, yc).Road(1) + Terrain(xc + 1, yc).Road(2) * 2 + Terrain(xc + 1, yc).Road(3) * 4 + Terrain(xc + 1, yc).Road(4) * 8
                        TileRiver2 = Terrain(xc + 1, yc).River(1) * 16 + Terrain(xc + 1, yc).River(2) * 32 + Terrain(xc + 1, yc).River(3) * 64 + 1 * 128
                    Else
                        TileRoad2 = Terrain(xc + 1, yc).Road(1) + Terrain(xc + 1, yc).Road(2) * 2 + Terrain(xc + 1, yc).Road(3) * 4 + 1 * 8
                        TileRiver2 = Terrain(xc + 1, yc).River(1) * 16 + Terrain(xc + 1, yc).River(2) * 32 + Terrain(xc + 1, yc).River(3) * 64 + Terrain(xc + 1, yc).River(4) * 128
                    End If
                    TileNum2 = Terrain(xc + 1, yc).TileDesc.NumTile
                Else
                    If Feat = "RI" Then
                        TileRoad2 = Terrain(xc - 1, yc).Road(1) + Terrain(xc - 1, yc).Road(2) * 2 + Terrain(xc - 1, yc).Road(3) * 4 + Terrain(xc - 1, yc).Road(4) * 8
                        TileRiver2 = Terrain(xc - 1, yc).River(1) * 16 + 1 * 32 + Terrain(xc - 1, yc).River(3) * 64 + Terrain(xc - 1, yc).River(4) * 128
                    Else
                        TileRoad2 = Terrain(xc - 1, yc).Road(1) + 1 * 2 + Terrain(xc - 1, yc).Road(3) * 4 + Terrain(xc - 1, yc).Road(4) * 8
                        TileRiver2 = Terrain(xc - 1, yc).River(1) * 16 + Terrain(xc - 1, yc).River(2) * 32 + Terrain(xc - 1, yc).River(3) * 64 + Terrain(xc - 1, yc).River(4) * 128
                    End If
                    TileNum2 = Terrain(xc - 1, yc).TileDesc.NumTile
                End If
                If FeatureTiles(TileNum1, TileRoad1 + TileRiver1) >= 0 Then
                    If FeatureTiles(TileNum2, TileRoad2 + TileRiver2) >= 0 Then
                        'Both tiles exist
                        UseOldAlgo = True
                    Else
                        'Only the N or S tile exists
                        UseOldAlgo = False
                        If ye > yc Then
                            UpdatePath Path, Feat, 0, xc, yc
                        Else
                            UpdatePath Path, Feat, 1, xc, yc
                        End If
                    End If
                Else
                    If FeatureTiles(TileNum2, TileRoad2 + TileRiver2) >= 0 Then
                        'Only the W or E tile exists
                        UseOldAlgo = False
                        If xe > xc Then
                            UpdatePath Path, Feat, 2, xc, yc
                        Else
                            UpdatePath Path, Feat, 3, xc, yc
                        End If
                    Else
                        'Neither tile exists
                        UseOldAlgo = True
                    End If
                End If
            Else
                'It's a city
                UseOldAlgo = True
            End If
            
            If UseOldAlgo Then
                DoRandom = True
                If AutoFeatures.NumRoads > 0 And Feat = "RI" Then   'Special case for river
                    'Let's try N or S
                    If ye > yc Then
                        TileRoad1 = Terrain(xc, yc + 1).Road(1) + Terrain(xc, yc + 1).Road(2) + Terrain(xc, yc + 1).Road(3) + Terrain(xc, yc + 1).Road(4)
                    Else
                        TileRoad1 = Terrain(xc, yc - 1).Road(1) + Terrain(xc, yc - 1).Road(2) + Terrain(xc, yc - 1).Road(3) + Terrain(xc, yc - 1).Road(4)
                    End If
                    'Let's try E or W
                    If xe > xc Then
                        TileRoad2 = Terrain(xc + 1, yc).Road(1) + Terrain(xc + 1, yc).Road(2) + Terrain(xc + 1, yc).Road(3) + Terrain(xc + 1, yc).Road(4)
                    Else
                        TileRoad2 = Terrain(xc - 1, yc).Road(1) + Terrain(xc - 1, yc).Road(2) + Terrain(xc - 1, yc).Road(3) + Terrain(xc - 1, yc).Road(4)
                    End If
                    'What we do here is try to choose the path that has the less roads
                    If TileRoad1 = TileRoad2 Then
                        DoRandom = True
                    ElseIf TileRoad1 > TileRoad2 Then
                        DoRandom = False
                        If xe > xc Then 'We go from west to east
                            UpdatePath Path, Feat, 2, xc, yc
                        Else    'We go from east to west
                            UpdatePath Path, Feat, 3, xc, yc
                        End If
                    ElseIf TileRoad1 < TileRoad2 Then
                        DoRandom = False
                        If ye > yc Then 'We go from north to south
                            UpdatePath Path, Feat, 0, xc, yc
                        Else    'We go from south to north
                            UpdatePath Path, Feat, 1, xc, yc
                        End If
                    End If
                Else
                    DoRandom = True
                End If
                
                If DoRandom Then
                    If Abs(xe - xc) > Abs(ye - yc) Then 'we will go east or west
                        If xe > xc Then 'We go from west to east
                            UpdatePath Path, Feat, 2, xc, yc
                        Else    'We go from east to west
                            UpdatePath Path, Feat, 3, xc, yc
                        End If
                    ElseIf Abs(xe - xc) > Abs(ye - yc) Then 'we will go north or south
                        If ye > yc Then 'We go from north to south
                            UpdatePath Path, Feat, 0, xc, yc
                        Else    'We go from south to north
                            UpdatePath Path, Feat, 1, xc, yc
                        End If
                    Else    'we choose at random
                        nb = Int(2 * Rnd + 1)
                        If nb = 1 Then  'We'll go north or south
                            If ye > yc Then 'We go from north to south
                                UpdatePath Path, Feat, 0, xc, yc
                            Else    'We go from south to north
                                UpdatePath Path, Feat, 1, xc, yc
                            End If
                        Else    'we'll go east or west
                            If xe > xc Then 'We go from west to east
                                UpdatePath Path, Feat, 2, xc, yc
                            Else    'We go from east to west
                                UpdatePath Path, Feat, 3, xc, yc
                            End If
                        End If
                    End If
                End If
            
            End If
        End If
    Loop
End Sub

Function MergeFeatures() As Long
'Will update tile numbers after river/road calculations
Dim i As Long, j As Long, TerTileSize As Long, updated As Long
Dim k As Integer
Dim FeatureIndex As Integer
Dim Percent As Long
Dim ToUpdate As Boolean

    lblShowAction.Caption = "Merging all features ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    TerTileSize = TerrainSize * 16 - 1
    For i = 0 To TerTileSize
        For j = 0 To TerTileSize
            'We calculate which features should be here
            FeatureIndex = 0
            For k = 1 To 4
                FeatureIndex = FeatureIndex + Terrain(i, j).River(k) * (2 ^ (k + 3)) + Terrain(i, j).Road(k) * (2 ^ (k - 1))
            Next k
            If FeatureIndex > 0 Then
                If FeatureTiles(Terrain(i, j).TileDesc.NumTile, FeatureIndex) >= 0 Then
                    'We update tile number with the one given in the conf file
                    Terrain(i, j).TileDesc.NumTile = FeatureTiles(Terrain(i, j).TileDesc.NumTile, FeatureIndex)
                    updated = updated + 1
                    'And we update the number of features in the tiles (the conf file may replace some river+road tiles by road only ones for example)
                    For k = 1 To 4
                        Terrain(i, j).River(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).River(k)
                        Terrain(i, j).Road(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).Road(k)
                    Next k
                    'Debug.Print "Apres1 Tile(" & i & "," & j & ")=" & Terrain(i, j).TileDesc.NumTile & " " & Terrain(i, j).River(4) & Terrain(i, j).River(3) & Terrain(i, j).River(2) & Terrain(i, j).River(1) & " " & Terrain(i, j).Road(4) & Terrain(i, j).Road(3) & Terrain(i, j).Road(2) & Terrain(i, j).Road(1)
                End If
            End If
            DoEvents
        Next j
        Percent = (i * 100) / TerTileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    
    
    lblShowAction.Caption = "Gathering terrain informations ..."
    'Note : there must a bug somewhere, because apparently in some cases
    'CATE does not have the correct information about features on tiles
    'Hence the need for this new pass
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    For i = 0 To TerTileSize
        For j = 0 To TerTileSize
            For k = 1 To 4
                Terrain(i, j).Road(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).Road(k)
                Terrain(i, j).River(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).River(k)
            Next k
            DoEvents
        Next j
        Percent = (i * 100) / TerTileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    
    lblShowAction.Caption = "Second pass, trying to clean up rivers and roads ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    DoEvents
    'Now we will try to end rivers and roads correctly
    For i = 0 To TerTileSize
        For j = 0 To TerTileSize
            ToUpdate = False
            'We check tiles to the N, S, E and W, and check if there are features connections problems
            If j > 0 Then
                If Terrain(i, j - 1).River(3) = 0 And Terrain(i, j).River(1) = 1 Then
                    Terrain(i, j).River(1) = 0
                    ToUpdate = True
                End If
                If Terrain(i, j - 1).Road(3) = 0 And Terrain(i, j).Road(1) = 1 Then
                    Terrain(i, j).Road(1) = 0
                    ToUpdate = True
                End If
            End If
            If i > 0 Then
                If Terrain(i - 1, j).River(2) = 0 And Terrain(i, j).River(4) = 1 Then
                    Terrain(i, j).River(4) = 0
                    ToUpdate = True
                End If
                If Terrain(i - 1, j).Road(2) = 0 And Terrain(i, j).Road(4) = 1 Then
                    Terrain(i, j).Road(4) = 0
                    ToUpdate = True
                End If
            End If
            If j < TerTileSize Then
                If Terrain(i, j + 1).River(1) = 0 And Terrain(i, j).River(3) = 1 Then
                    Terrain(i, j).River(3) = 0
                    ToUpdate = True
                End If
                If Terrain(i, j + 1).Road(1) = 0 And Terrain(i, j).Road(3) = 1 Then
                    Terrain(i, j).Road(3) = 0
                    ToUpdate = True
                End If
            End If
            If i < TerTileSize Then
                If Terrain(i + 1, j).River(4) = 0 And Terrain(i, j).River(2) = 1 Then
                    Terrain(i, j).River(2) = 0
                    ToUpdate = True
                End If
                If Terrain(i + 1, j).Road(4) = 0 And Terrain(i, j).Road(2) = 1 Then
                    Terrain(i, j).Road(2) = 0
                    ToUpdate = True
                End If
            End If
            If ToUpdate Then
                FeatureIndex = 0
                For k = 1 To 4
                    FeatureIndex = FeatureIndex + Terrain(i, j).River(k) * (2 ^ (k + 3)) + Terrain(i, j).Road(k) * (2 ^ (k - 1))
                Next k
                'We update tile number and features info
                If TileToFeature(Terrain(i, j).TileDesc.NumTile).Value >= 0 Then
                    If FeatureTiles(TileToFeature(Terrain(i, j).TileDesc.NumTile).Value, FeatureIndex) >= 0 Then
                        Terrain(i, j).TileDesc.NumTile = FeatureTiles(TileToFeature(Terrain(i, j).TileDesc.NumTile).Value, FeatureIndex)
                        For k = 1 To 4
                            Terrain(i, j).River(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).River(k)
                            Terrain(i, j).Road(k) = TileToFeature(Terrain(i, j).TileDesc.NumTile).Road(k)
                        Next k
                    End If
                End If
            End If
            DoEvents
        Next j
        Percent = (i * 100) / TerTileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next i
    
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    MergeFeatures = updated
End Function

Function GiveCorner(Corner As Integer, x As Long, y As Long, CurrentTerrain As Integer, TerTileSize As Long, Trans As Integer) As Integer
'Will return 0 or 1 when calculating transitions
'1 if the three tiles around a corner are belonging to tiles towards which we translate
'e.g. : in a city to ocean transition, 1 returned if these 3 tiles are not ocean
'0 otherwise
Dim x1 As Long, x2 As Long, x3 As Long
Dim y1 As Long, y2 As Long, y3 As Long
Dim Terrain1 As Integer, Terrain2 As Integer, Terrain3 As Integer
Dim i As Integer
Dim Found As Boolean

    Select Case Corner
        Case 1:
            'NW corner
            x1 = x - 1
            y1 = y
            x2 = x - 1
            y2 = y - 1
            x3 = x
            y3 = y - 1
        Case 2:
            'NE corner
            x1 = x
            y1 = y - 1
            x2 = x + 1
            y2 = y - 1
            x3 = x + 1
            y3 = y
        Case 3:
            'SW corner
            x1 = x - 1
            y1 = y
            x2 = x - 1
            y2 = y + 1
            x3 = x
            y3 = y + 1
        Case 4:
            'SE corner
            x1 = x
            y1 = y + 1
            x2 = x + 1
            y2 = y + 1
            x3 = x + 1
            y3 = y
    End Select
    
    If x1 < 0 Or x1 > TerTileSize Or y1 < 0 Or y1 > TerTileSize Then
        'If out of map, we suppose same terrain
        Terrain1 = CurrentTerrain
    Else
        Terrain1 = GiveTerrainFromTile(Terrain(x1, y1).TileDesc.NumTile)
        If Terrain1 <> CurrentTerrain Then
            'if terrain is different, we check if the related tile is part or not of the tiles towards which we translate
            Found = False
            For i = 1 To AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).NumTiles
                If Terrain(x1, y1).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).TileList(i) Then
                    Found = True
                    Exit For
                End If
            Next i
            'if the tile is of another type (e.g. farm when translating to ocean)
            'we do as if it were a tile of the current terrain type (e.g. city with the same example)
            If Not Found Then Terrain1 = CurrentTerrain
        End If
    End If
    
    'Same thing for second tile
    If x2 < 0 Or x2 > TerTileSize Or y2 < 0 Or y2 > TerTileSize Then
        Terrain2 = CurrentTerrain
    Else
        Terrain2 = GiveTerrainFromTile(Terrain(x2, y2).TileDesc.NumTile)
        If Terrain2 <> CurrentTerrain Then
            Found = False
            For i = 1 To AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).NumTiles
                If Terrain(x2, y2).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).TileList(i) Then
                    Found = True
                    Exit For
                End If
            Next i
            If Not Found Then Terrain2 = CurrentTerrain
        End If
    End If
    
    'Same thing for third tile
    If x3 < 0 Or x3 > TerTileSize Or y3 < 0 Or y3 > TerTileSize Then
        Terrain3 = CurrentTerrain
    Else
        Terrain3 = GiveTerrainFromTile(Terrain(x3, y3).TileDesc.NumTile)
        If Terrain3 <> CurrentTerrain Then
            Found = False
            For i = 1 To AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).NumTiles
                If Terrain(x3, y3).TileDesc.NumTile = AutoFeatures.TerrainFeatures(CurrentTerrain).Transitions(Trans).TileList(i) Then
                    Found = True
                    Exit For
                End If
            Next i
            If Not Found Then Terrain3 = CurrentTerrain
        End If
    End If
    
    'Results
    GiveCorner = 0
    If Terrain1 = CurrentTerrain Then
        If Terrain2 = CurrentTerrain Then
            If Terrain3 = CurrentTerrain Then
                GiveCorner = 1
            End If
        End If
    End If

End Function

Function SaveTerrainAsBmp() As Boolean
'Will save a BMP picture of the theater
Dim fd As Integer
Dim Rien As Long
Dim i As Long, j As Long
Dim TerTileSize As Long
Dim Percent As Long

    Screen.MousePointer = vbHourglass
    If Not TerrainCalculated Then
        TranslateTilesIntoTerrain
        TerrainCalculated = True
    End If
    
    TerTileSize = TerrainSize * 16
    fd = FreeFile
    On Error Resume Next
    If Dir$(BmpToSave.FileName, vbNormal) <> "" Then Kill BmpToSave.FileName
    If Err <> 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reseting " & BmpToSave.FileName & ". Action aborted", vbCritical, "Warning"
        SaveTerrainAsBmp = False
        Exit Function
    End If
    Open BmpToSave.FileName For Binary As #fd
    If Err <> 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & BmpToSave.FileName & ". Action aborted.", vbCritical + vbOKOnly, "Warning"
        SaveTerrainAsBmp = False
        Exit Function
    End If
    On Error GoTo 0
    
    lblShowAction.Caption = "Saving bitmap (" & BmpToSave.FileName & ") ..."
    lblPercAction.Caption = "0 %"
    picAction.Cls
    UpdatePercentBar picAction, 0
    DoEvents
    picAction.FillColor = &HC0C000
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.Visible = True
    EnableInput False
    Screen.MousePointer = vbHourglass
    
    SaveImageHeader.Signature = "BM"
    SaveImageHeader.FileSize = TerTileSize * TerTileSize + 54 + 256 * 4
    Rien = 0
    SaveImageHeader.DataOffset = 1078
    SaveImageHeader.HeaderSize = 40
    SaveImageHeader.ImageWidth = TerTileSize
    SaveImageHeader.ImageHeight = TerTileSize
    SaveImageHeader.Planes = 1
    SaveImageHeader.BitCount = 8
    SaveImageHeader.Compression = 0
    SaveImageHeader.ImageSize = TerTileSize * TerTileSize
    SaveImageHeader.HorResolution = 2880
    SaveImageHeader.VerResolution = 2880
    'SaveImageHeader.ColorUsed = CurrentImageColor + 1
    'SaveImageHeader.ColorsImportant = CurrentImageColor + 1
    SaveImageHeader.ColorUsed = 256
    SaveImageHeader.ColorsImportant = 256
    With SaveImageHeader
        Put #fd, , .Signature
        Put #fd, , .FileSize
        Put #fd, , Rien
        Put #fd, , .DataOffset
        Put #fd, , .HeaderSize
        Put #fd, , .ImageWidth
        Put #fd, , .ImageHeight
        Put #fd, , .Planes
        Put #fd, , .BitCount
        Put #fd, , .Compression
        Put #fd, , .ImageSize
        Put #fd, , .HorResolution
        Put #fd, , .VerResolution
        Put #fd, , .ColorUsed
        Put #fd, , .ColorsImportant
    End With
    'color table
    For i = 0 To 255
        Put #fd, , Chr$(SaveImageColors(i).B) & Chr$(SaveImageColors(i).G) & Chr$(SaveImageColors(i).R) & Chr$(0)
    Next i
    'data
    For j = TerTileSize - 1 To 0 Step -1
        For i = 0 To TerTileSize - 1
            Put #fd, , BmpToSave.TileNum(Terrain(i, j).TileDesc.NumTile)
            DoEvents
        Next i
        Percent = ((TerTileSize - 1 - j) * 100) / TerTileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Next j
    
    Close #fd
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    SaveTerrainAsBmp = True

End Function

Function LoadTextureIndex(FileName As String) As Boolean
'Will load the txt format of a texture.bin file
Dim fd As Integer
Dim tmp As String
Dim FileSize As Long, Percent As Long
Dim CurrentSet As Integer
Dim CurrentIndex As Integer
Dim pos As Integer

    fd = FreeFile
    On Error Resume Next
    FileSize = FileLen(FileName)
    Open FileName For Input As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted", vbCritical, "Warning"
        LoadTextureIndex = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    ReDim TileNumToName(0 To 0)
    UpdatePercentBar picAction, 0
    picAction.Cls
    picAction.FillColor = &HFF00&
    picAction.Visible = True
    lblShowAction.Caption = "Loading Texture index file " & FileName & " ..."
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    DoEvents
    EnableInput False
    'On Error GoTo Problem
    CurrentSet = -1
    Do While Not EOF(fd)
        Line Input #fd, tmp                 'we read a full line in tmp string
        If Len(tmp) > 4 Then                '"Empty" lines are ignored
            If UCase$(Left$(tmp, 1)) <> "#" Then
                If UCase$(Left$(tmp, 3)) = "SET" Then
                    'We count the number of Sets seen
                    CurrentSet = CurrentSet + 1
                    CurrentIndex = -1
                    ReDim Preserve TileNumToName(0 To CurrentSet * 16 + 15)
                Else
                    pos = InStr(1, tmp, "pcx", vbTextCompare)
                    If pos > 0 Then
                        'We keep track of the texture name (without the .pcx part) and index
                        CurrentIndex = CurrentIndex + 1
                        TileNumToName(CurrentSet * 16 + CurrentIndex) = Left$(tmp, pos - 2)
                    End If
                End If
            End If
        End If
        Percent = Seek(fd) * 100 \ FileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Loop
    Close #fd
    
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    DoEvents
    LoadTextureIndex = True
    Exit Function
    
Problem:
    Screen.MousePointer = vbNormal
    MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reading file " & FileName & ". Please check your file format. Action aborted.", vbCritical, "Warning"
    Close #fd
    LoadTextureIndex = False
    Exit Function
End Function

Function LoadTrnFiles() As Boolean
'Will load all trn files
Dim i As Integer, fd As Integer, Rep As Integer, pos As Integer, pos2 As Integer
Dim TotalSize As Long
Dim ReadSize As Long, TempRead As Long
Dim tmp As String, tmpclean As String, tmp2 As String
Dim Percent As Long
Dim gtx As Integer, gty As Integer
Dim stx As Integer, sty As Integer
Dim j As Integer
Dim char As String
Dim Textures(0 To 15) As String
Dim CurrentTexture As String
Dim NumTextures As String

    'We check if all files exist
    For i = 1 To AutoFeatures.NumTrnFiles
        If Dir$(AutoFeatures.TrnFileNames(i)) = "" Then
            MsgBox "TRN File " & AutoFeatures.TrnFileNames(i) & " does not exist. Please check your conf file.", vbCritical, "Warning"
            LoadTrnFiles = False
            Exit Function
        End If
    Next i
    Erase gt
    For i = 1 To AutoFeatures.NumTrnFiles
        TotalSize = TotalSize + FileLen(AutoFeatures.TrnFileNames(i))
    Next i
    ReadSize = 0
    Screen.MousePointer = vbHourglass
    UpdatePercentBar picAction, 0
    picAction.Cls
    picAction.FillColor = &HFF00&
    picAction.Visible = True
    lblShowAction.Caption = ""
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    DoEvents
    EnableInput False
    For i = 1 To AutoFeatures.NumTrnFiles
        lblShowAction.Caption = "Loading TRN information from " & AutoFeatures.TrnFileNames(i)
        DoEvents
        fd = FreeFile
        On Error Resume Next
        Open AutoFeatures.TrnFileNames(i) For Input As #fd
        If Err <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Error #" & Err & " (" & Error$(Err) & ") while reading TRN file " & AutoFeatures.TrnFileNames(i) & ". Do you wish to go on anyway ?", vbCritical + vbYesNo, "Warning"
            If Rep = vbNo Then
                LoadTrnFiles = False
                Exit Function
            End If
        End If
        'We get the current global tile coords. from the name of the file
        gtx = Val(Mid$(GiveNameFromPath(AutoFeatures.TrnFileNames(i)), 2, 1))
        gty = Val(Mid$(GiveNameFromPath(AutoFeatures.TrnFileNames(i)), 3, 1))
        'This test may have to change of we have more global tiles than 4
        'If (gtx = 0 Or gtx = 1) And (gty = 0 Or gty = 1) Then
            TempRead = 0
            On Error GoTo 0
            Do While Not EOF(fd)
                Line Input #fd, tmp
                If Len(tmp) > 5 Then    'we skip empty lines
                    tmpclean = CleanString(tmp)
                    If LCase$(Left$(tmpclean, 6)) = "<supr>" Then
                        'Here we will have the super tiles coords on two different lines
                        Line Input #fd, tmp
                        stx = Val(CleanString(tmp))
                        Line Input #fd, tmp
                        sty = Val(CleanString(tmp))
                        Erase Textures
                        CurrentTexture = -1
                    End If
                    
                    If LCase$(Left$(tmpclean, 6)) = "<txtl>" Then
                        Line Input #fd, tmp
                        NumTextures = CleanString(tmp)
                        'There are txo different <txtl> sections, the following test make sure
                        'we are in the correct one
                        If LCase(Left$(NumTextures, 2)) <> "<b" Then
                            'j = 1
                            Do
                            'For j = 1 To Val(NumTextures)
                                'Now we will read all NumTextures names (hoping this number is correctly set ! : well no, it is wrong, so we read up to <endo>)
                                'and skip those :
                                '   - who do not have raw in their name
                                '   - who have an \ in their name (probably because of a bad format
                                '   - whose name end with a W
                                'and record the others in Textures array, after translating them
                                'following what we previously found in the association file if necessary
                                Line Input #fd, tmp
                                tmp2 = CleanString(tmp)
                                If LCase$(Left$(tmp2, 6)) = "<endo>" Then Exit Do
                                pos = InStr(1, tmp2, "raw", vbTextCompare)
                                pos2 = InStr(1, tmp2, "\")
                                If pos > 0 And pos2 <= 0 Then
                                    CurrentTexture = CurrentTexture + 1
                                    tmp2 = Left$(tmp2, InStr(1, tmp2, ".") - 1)
                                    If UCase$(Right$(tmp2, 1)) <> "W" Then
                                        If AutoFeatures.CorrespFileName <> "" Then
                                            Textures(CurrentTexture) = FindCorrespName(tmp2)
                                        Else
                                            Textures(CurrentTexture) = tmp2
                                        End If
                                    End If
                                End If
                                'j = j = 1
                                DoEvents
                            Loop
                            'Next j
                        End If
                    End If
                    
                    If LCase$(Left$(tmpclean, 6)) = "<tref>" Then
                        For j = 1 To 16
                            'Here we will read all textures index (i.e. the position in the supertile)
                            'The index should be ranging from 0 to 15, but listed as 16 to 31
                            'Due to some bad format this is not always the case
                            'Then we update the "TRN terrain" with the new name
                            Line Input #fd, tmp
                            pos2 = Val(CleanString(tmp))
                            If pos2 >= 16 Then pos2 = pos2 - 16
                            If Textures(j - 1) <> "" Then
                                If UseTrefSection Then
                                    gt(gtx, gty).st(stx, sty).Tile(pos2) = Textures(j - 1)
                                Else
                                    gt(gtx, gty).st(stx, sty).Tile(j - 1) = Textures(j - 1)
                                End If
                            End If
                            DoEvents
                        Next j
                    End If

                End If
                
                
                TempRead = Seek(fd)
                Percent = (ReadSize + TempRead) * 100 / TotalSize
                If Percent Mod 2 = 0 Then
                    lblPercAction.Caption = "" & Percent & " %"
                    UpdatePercentBar picAction, Percent
                End If
                DoEvents
            Loop
            On Error Resume Next
            Close #fd
        'Else
        'End If
        ReadSize = ReadSize + FileLen(AutoFeatures.TrnFileNames(i))
    Next i

    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    DoEvents
    LoadTrnFiles = True
End Function

Function LoadCorresp(FileName As String) As Boolean
'Will load the file associating the names of trn files to those in texture.bin file
Dim fd As Integer
Dim tmp As String
Dim FileSize As Long, Percent As Long
Dim pos As Integer, pos2 As Integer
Dim Name1 As String, Name2 As String

    fd = FreeFile
    On Error Resume Next
    FileSize = FileLen(FileName)
    Open FileName For Input As #fd
    If Err <> 0 Then
        MsgBox "Error #" & Err & " (" & Error$(Err) & ") when opening " & FileName & ". Action aborted", vbCritical, "Warning"
        LoadCorresp = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    UpdatePercentBar picAction, 0
    picAction.Cls
    picAction.FillColor = &HFF00&
    picAction.Visible = True
    lblShowAction.Caption = "Loading correspondance file " & FileName & " ..."
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    DoEvents
    NameCorresp.NumNames = 0
    ReDim NameCorresp.OrigName(0 To 0)
    ReDim NameCorresp.DestName(0 To 0)
    EnableInput False
    On Error GoTo Problem
    Do While Not EOF(fd)
        Line Input #fd, tmp                 'we read a full line in tmp string
        If Len(tmp) > 2 Then                '"Empty" lines are ignored
            'We search for tab, space, "," or ";"
            pos = InStr(1, tmp, Chr$(9))
            If pos <= 0 Then pos = InStr(1, tmp, " ")
            If pos <= 0 Then pos = InStr(1, tmp, ",")
            If pos <= 0 Then pos = InStr(1, tmp, ";")
            If pos > 0 Then
                Name1 = Left$(tmp, pos - 1)
                Name2 = Right$(tmp, Len(tmp) - pos)
                pos2 = InStr(1, Name1, ".")
                If pos2 > 0 Then Name1 = Left$(Name1, pos2 - 1)
                pos2 = InStr(1, Name2, ".")
                If pos2 > 0 Then Name2 = Left$(Name2, pos2 - 1)
                NameCorresp.NumNames = NameCorresp.NumNames + 1
                If NameCorresp.NumNames = 1 Then
                    ReDim NameCorresp.OrigName(1 To 1)
                    ReDim NameCorresp.DestName(1 To 1)
                Else
                    ReDim Preserve NameCorresp.OrigName(1 To NameCorresp.NumNames)
                    ReDim Preserve NameCorresp.DestName(1 To NameCorresp.NumNames)
                End If
                NameCorresp.OrigName(NameCorresp.NumNames) = Name1
                NameCorresp.DestName(NameCorresp.NumNames) = Name2
            End If
        End If
        Percent = Seek(fd) * 100 \ FileSize
        If Percent Mod 2 = 0 Then
            lblPercAction.Caption = Percent & " %"
            UpdatePercentBar picAction, Percent
        End If
        DoEvents
    Loop
    Close #fd
    
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    DoEvents
    LoadCorresp = True
    Exit Function
    
Problem:
    Screen.MousePointer = vbNormal
    MsgBox "Error #" & Err & " (" & Error$(Err) & ") when reading file " & FileName & ". Please check your file format. Action aborted.", vbCritical, "Warning"
    Close #fd
    LoadCorresp = False
    Exit Function

End Function

Function UpdateTrnTiles() As Long
'Will update F4 terrain with imported TRN new tiles
Dim gtx As Long, gty As Long, stx As Long, sty As Long, i As Integer
Dim x As Integer
Dim y As Integer
Dim NewTile As Integer
Dim updated As Long
Dim fd As Integer
Dim Trace As Boolean
Dim Percent As Long

    Trace = True
    If Trace Then
        fd = FreeFile
        Open App.Path & "\UpdatedTiles.txt" For Output As #fd
    End If
    Screen.MousePointer = vbHourglass
    'If terrain info is not yet loaded, we do it now
    If Not TerrainCalculated Then
        EnableInput False
        DoEvents
        TranslateTilesIntoTerrain
        TerrainCalculated = True
    End If
    
    lblShowAction.Caption = "Importing tiles from TRN into F4 terrain ..."
    picAction.Cls
    lblPercAction.Caption = "0 %"
    lblShowAction.Visible = True
    lblPercAction.Visible = True
    picAction.FillColor = &HC0C000
    picAction.Visible = True
    EnableInput False
    Screen.MousePointer = vbHourglass
    DoEvents
    For gtx = 0 To 9
        For gty = 0 To 9
            For stx = 0 To 7
                For sty = 0 To 7
                    For i = 0 To 15
                        'if the current trn tile is defined
                        If gt(gtx, gty).st(stx, sty).Tile(i) <> "" Then
                            NewTile = FindTileNumber(gt(gtx, gty).st(stx, sty).Tile(i))
                            'if we have found the new tile number
                            If NewTile >= 0 Then
                                'Conversion to CATE coords
                                x = gtx * 32 + stx * 4 + i Mod 4
                                y = TerrainSize * 16 - 1 - (gty * 32 + sty * 4 + i \ 4)
                                Terrain(x + AutoFeatures.TRNOffsetX, y - AutoFeatures.TRNOffsetY).TileDesc.NumTile = NewTile
                                updated = updated + 1
                                If Trace Then
                                    Print #fd, "Updated tile " & x + AutoFeatures.TRNOffsetX & ","; 1023 - y + AutoFeatures.TRNOffsetY & " with tile number " & NewTile & " (tile name : " & gt(gtx, gty).st(stx, sty).Tile(i) & ", coords in TRN : " & gtx & "," & gty & "," & stx & "," & sty & "," & i & ")"
                                End If
                            Else
                                If Trace Then
                                    Print #fd, "Found no tile number for tile name " & gt(gtx, gty).st(stx, sty).Tile(i)
                                End If
                            End If
                        End If
                        DoEvents
                    Next i
                    'Percent = (gtx * 128 + gty * 64 + stx * 8 + (sty + 1)) * 100 / 256
                    Percent = (gtx * 640 + gty * 64 + stx * 8 + (sty + 1)) * 100 / 6400
                    If Percent Mod 2 = 0 Then
                        lblPercAction.Caption = Percent & " %"
                        UpdatePercentBar picAction, Percent
                    End If
                    DoEvents
                Next sty
                DoEvents
            Next stx
            DoEvents
        Next gty
        DoEvents
    Next gtx
    If Trace Then
        Close #fd
    End If
    lblShowAction.Visible = False
    lblPercAction.Visible = False
    picAction.Visible = False
    EnableInput True
    Screen.MousePointer = vbNormal
    UpdateTrnTiles = updated
End Function

Function FindCorrespName(Name1 As String) As String
'Will return for a given tile name (from a TRN file) the correct name (for texture.bin / .txt)
Dim i As Integer

    FindCorrespName = ""
    For i = 1 To NameCorresp.NumNames
        If NameCorresp.OrigName(i) = Name1 Then
            FindCorrespName = NameCorresp.DestName(i)
            Exit For
        End If
        DoEvents
    Next i
End Function

Function FindTileNumber(TileName As String) As Integer
'Will return, for a given tile name its tile number defined in texture.bin / .txt
Dim i As Integer

    FindTileNumber = -1
    For i = 0 To UBound(TileNumToName)
        If TileNumToName(i) = TileName Then
            FindTileNumber = i
            Exit Function
        End If
    Next i
End Function
