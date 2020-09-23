Attribute VB_Name = "modPBEM"
Option Explicit

Global Const GOLD_PER_MINE_PER_TURN = 50

Public Enum eEmpire
    Black = 0
    Red = 1
    Blue = 2
    Yellow = 3
    Purple = 4
End Enum

Public Enum eHexPicID
    Town1ID = 0
    Town2ID = 1
    Town3ID = 2
    CapitalID = 3
    GoldID = 4
    
    MountainsID = 5
    ForestID = 6
    BlankID = 7
    
    'Mercenary = 8
End Enum


Public Enum eHexPictures
    
    Town1Black = 0
    Town2Black = 1
    Town3Black = 2
    CapitalBlack = 3
    GoldBlack = 4
    
    Town1Red = 5
    Town2Red = 6
    Town3Red = 7
    CapitalRed = 8
    GoldRed = 9
    
    Town1Blue = 10
    Town2Blue = 11
    Town3Blue = 12
    CapitalBlue = 13
    GoldBlue = 14
    
    Town1Yellow = 15
    Town2Yellow = 16
    Town3Yellow = 17
    CapitalYellow = 18
    GoldYellow = 19
    
    Town1Purple = 20
    Town2Purple = 21
    Town3Purple = 22
    CapitalPurple = 23
    GoldPurple = 24
    
    ' Non-Empire affiliated pictures
    Mountains = 25
    Forest = 26
    Blank = 27
    
End Enum


Global goHexPictures(Town1Black To Blank) As clsVDC
Global goHexMasks(Town1Black To Blank) As clsVDC


Public Function DetermineHexPictureID(iEmpire As eEmpire, HexPicID As eHexPicID) As eHexPictures
    Select Case HexPicID
        Case Town1ID To GoldID
            ' Those value that can have an empire tied to them
            DetermineHexPictureID = (iEmpire * 5) + HexPicID
        Case Else
            ' Those that have no empire have a difference of 20 between
            ' their eHexPicID and their eHexPictures values
            DetermineHexPictureID = HexPicID + 20
    End Select
End Function


Public Function GetRandomNumber(iMin As Integer, iMax As Integer) As Integer
    Randomize Timer
    GetRandomNumber = Int((iMax * Rnd) + iMin)
End Function




Public Sub Main()
On Error GoTo ErrHandler
    
    Randomize Timer
    frmMain.Show
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
    Exit Sub
    Resume
End Sub


Public Function ScreenTest(frm As Form) As Boolean
Dim lScreenWidth As Long
Dim lScreenHeight As Long
    lScreenWidth = frm.ScaleX(Screen.Width, vbTwips, vbPixels)
    lScreenHeight = frm.ScaleY(Screen.Height, vbTwips, vbPixels)
    If (lScreenWidth < 1024) Or (lScreenHeight < 768) Then
        If (vbNo = MsgBox("The program has detected you are running a screen resolution of " & lScreenWidth & " x " & lScreenHeight & ", however this program requires a minimum of 1024 x 768. It is recommended you quit, change your screen resolution, and then return." & vbCrLf & vbCrLf & "Do you still wish to continue with the program?", vbQuestion + vbYesNo + vbDefaultButton2, "Invalid Screen Resolution")) Then
            Exit Function
        End If
    End If
    ScreenTest = True
End Function


Public Sub PaintGraphicID(ID As eHexPictures, oCanvas As Object, Optional x As Long, Optional y As Long, Optional vPercent As Variant)
    
    If (IsMissing(vPercent)) Then
    
        goHexMasks(ID).Paint oCanvas, x, y, , , , , , , vbSrcAnd
        goHexPictures(ID).Paint oCanvas, x, y, , , , , , , vbSrcPaint
        
    Else
        
        goHexMasks(ID).PaintType = vbSrcAnd
        goHexMasks(ID).PaintPercent oCanvas, 4
        goHexMasks(ID).PaintType = vbSrcCopy
        
        goHexPictures(ID).PaintType = vbSrcPaint
        goHexPictures(ID).PaintPercent oCanvas, 4
        goHexPictures(ID).PaintType = vbSrcCopy
        
    End If
    
End Sub



Public Sub PaintGraphicFromStats(vEmpire As eEmpire, vCurrentHexPic As eHexPicID, oCanvas As Object, Optional x As Long, Optional y As Long, Optional vPercent As Variant)
Dim ID As eHexPictures
    ID = DetermineHexPictureID(vEmpire, vCurrentHexPic)
    PaintGraphicID ID, oCanvas, x, y, vPercent
End Sub



Public Function LoadGraphic(HexPicID As eHexPictures, szFile As String) As Boolean
    
    Set goHexPictures(HexPicID) = New clsVDC
    LoadGraphic = goHexPictures(HexPicID).CreateFromFile(szFile)
    DoEvents
    DoEvents
    
    ' If false, then file doesn't exist. Rollback error to calling proc
    If (Not LoadGraphic) Then Err.Raise vbObjectError, , "File """ & szFile & """ does not exist. Cannot load graphics."
    
End Function

Public Function LoadMask(HexPicID As eHexPictures, szFile As String) As Boolean

    Set goHexMasks(HexPicID) = New clsVDC
    LoadMask = goHexMasks(HexPicID).CreateFromFile(szFile)
    DoEvents
    DoEvents
    
    ' If false, then file doesn't exist. Rollback error to calling proc
    If (Not LoadMask) Then Err.Raise vbObjectError, , "File """ & szFile & """ does not exist. Cannot load graphics."
    
End Function


Public Function TownIDToName(iHexPicID As Integer)
    Select Case iHexPicID
        Case Town1ID
            TownIDToName = "Level 1"
        Case Town2ID
            TownIDToName = "Level 2"
        Case Town3ID
            TownIDToName = "Level 3"
        Case CapitalID
            TownIDToName = "Capital"
    End Select
End Function



Public Function ColorIDToName(iEmpire As Integer) As String
    Select Case iEmpire
        Case Black
            ColorIDToName = "Black"
        Case Red
            ColorIDToName = "Red"
        Case Blue
            ColorIDToName = "Blue"
        Case Yellow
            ColorIDToName = "Yellow"
        Case Purple
            ColorIDToName = "Purple"
    End Select
End Function

Public Function ColorNameToID(szColor As String) As Integer
    Select Case UCase$(szColor)
        Case "BLACK"
            ColorNameToID = eEmpire.Black
        Case "RED"
            ColorNameToID = eEmpire.Red
        Case "BLUE"
            ColorNameToID = eEmpire.Blue
        Case "YELLOW"
            ColorNameToID = eEmpire.Yellow
        Case "PURPLE"
            ColorNameToID = eEmpire.Purple
    End Select
End Function


Public Function LevelNameToID(ByVal szLevel As String) As Integer
    Select Case UCase$(szLevel)
        Case "LEVEL 1"
            LevelNameToID = 0
        Case "LEVEL 2"
            LevelNameToID = 1
        Case "LEVEL 3"
            LevelNameToID = 2
        Case "CAPITAL"
            LevelNameToID = 3
    End Select
End Function

Public Sub UnloadGraphics()
Dim v As eHexPictures

    For v = Town1Black To Blank
        Set goHexPictures(v) = Nothing
        Set goHexMasks(v) = Nothing
    Next v
    
End Sub


Public Sub LoadAllGraphics()

    ' ************** Load Hex Unit Bitmaps for Map **************
    LoadGraphic Town1Black, App.Path & "\Pics\Hex\Town1Black.bmp"
    LoadGraphic Town2Black, App.Path & "\Pics\Hex\Town2Black.bmp"
    LoadGraphic Town3Black, App.Path & "\Pics\Hex\Town3Black.bmp"
    LoadGraphic CapitalBlack, App.Path & "\Pics\Hex\CapitalBlack.bmp"
    LoadGraphic GoldBlack, App.Path & "\Pics\Hex\GoldBlack.bmp"
    
    LoadGraphic Town1Red, App.Path & "\Pics\Hex\Town1Red.bmp"
    LoadGraphic Town2Red, App.Path & "\Pics\Hex\Town2Red.bmp"
    LoadGraphic Town3Red, App.Path & "\Pics\Hex\Town3Red.bmp"
    LoadGraphic CapitalRed, App.Path & "\Pics\Hex\CapitalRed.bmp"
    LoadGraphic GoldRed, App.Path & "\Pics\Hex\GoldRed.bmp"
    
    LoadGraphic Town1Blue, App.Path & "\Pics\Hex\Town1Blue.bmp"
    LoadGraphic Town2Blue, App.Path & "\Pics\Hex\Town2Blue.bmp"
    LoadGraphic Town3Blue, App.Path & "\Pics\Hex\Town3Blue.bmp"
    LoadGraphic CapitalBlue, App.Path & "\Pics\Hex\CapitalBlue.bmp"
    LoadGraphic GoldBlue, App.Path & "\Pics\Hex\GoldBlue.bmp"
    
    LoadGraphic Town1Yellow, App.Path & "\Pics\Hex\Town1Yellow.bmp"
    LoadGraphic Town2Yellow, App.Path & "\Pics\Hex\Town2Yellow.bmp"
    LoadGraphic Town3Yellow, App.Path & "\Pics\Hex\Town3Yellow.bmp"
    LoadGraphic CapitalYellow, App.Path & "\Pics\Hex\CapitalYellow.bmp"
    LoadGraphic GoldYellow, App.Path & "\Pics\Hex\GoldYellow.bmp"
    
    LoadGraphic Town1Purple, App.Path & "\Pics\Hex\Town1Purple.bmp"
    LoadGraphic Town2Purple, App.Path & "\Pics\Hex\Town2Purple.bmp"
    LoadGraphic Town3Purple, App.Path & "\Pics\Hex\Town3Purple.bmp"
    LoadGraphic CapitalPurple, App.Path & "\Pics\Hex\CapitalPurple.bmp"
    LoadGraphic GoldPurple, App.Path & "\Pics\Hex\GoldPurple.bmp"
    
    LoadGraphic Mountains, App.Path & "\Pics\Hex\Mountains.bmp"
    LoadGraphic Forest, App.Path & "\Pics\Hex\Forest.bmp"
    LoadGraphic Blank, App.Path & "\Pics\Hex\Blank.bmp"
    
    
    
    
    ' ************** Load Hex Unit Masks for Map **************
    LoadMask Town1Black, App.Path & "\Pics\HexMasks\Town1Mask.bmp"
    LoadMask Town2Black, App.Path & "\Pics\HexMasks\Town2Mask.bmp"
    LoadMask Town3Black, App.Path & "\Pics\HexMasks\Town3Mask.bmp"
    LoadMask CapitalBlack, App.Path & "\Pics\HexMasks\CapitalMask.bmp"
    LoadMask GoldBlack, App.Path & "\Pics\HexMasks\GoldMask.bmp"
    
    LoadMask Town1Red, App.Path & "\Pics\HexMasks\Town1Mask.bmp"
    LoadMask Town2Red, App.Path & "\Pics\HexMasks\Town2Mask.bmp"
    LoadMask Town3Red, App.Path & "\Pics\HexMasks\Town3Mask.bmp"
    LoadMask CapitalRed, App.Path & "\Pics\HexMasks\CapitalMask.bmp"
    LoadMask GoldRed, App.Path & "\Pics\HexMasks\GoldMask.bmp"
    
    LoadMask Town1Blue, App.Path & "\Pics\HexMasks\Town1Mask.bmp"
    LoadMask Town2Blue, App.Path & "\Pics\HexMasks\Town2Mask.bmp"
    LoadMask Town3Blue, App.Path & "\Pics\HexMasks\Town3Mask.bmp"
    LoadMask CapitalBlue, App.Path & "\Pics\HexMasks\CapitalMask.bmp"
    LoadMask GoldBlue, App.Path & "\Pics\HexMasks\GoldMask.bmp"
    
    LoadMask Town1Yellow, App.Path & "\Pics\HexMasks\Town1Mask.bmp"
    LoadMask Town2Yellow, App.Path & "\Pics\HexMasks\Town2Mask.bmp"
    LoadMask Town3Yellow, App.Path & "\Pics\HexMasks\Town3Mask.bmp"
    LoadMask CapitalYellow, App.Path & "\Pics\HexMasks\CapitalMask.bmp"
    LoadMask GoldYellow, App.Path & "\Pics\HexMasks\GoldMask.bmp"
    
    LoadMask Town1Purple, App.Path & "\Pics\HexMasks\Town1Mask.bmp"
    LoadMask Town2Purple, App.Path & "\Pics\HexMasks\Town2Mask.bmp"
    LoadMask Town3Purple, App.Path & "\Pics\HexMasks\Town3Mask.bmp"
    LoadMask CapitalPurple, App.Path & "\Pics\HexMasks\CapitalMask.bmp"
    LoadMask GoldPurple, App.Path & "\Pics\HexMasks\GoldMask.bmp"
    
    LoadMask Mountains, App.Path & "\Pics\HexMasks\MountainsMask.bmp"
    LoadMask Forest, App.Path & "\Pics\HexMasks\ForestMask.bmp"
    LoadMask Blank, App.Path & "\Pics\HexMasks\BlankMask.bmp"

End Sub


'   Returns True if successful, False otherwise.
Public Function DoesDirectoryExist(ByVal szPath As String) As Boolean
On Error GoTo DoesDirectoryExistErr
    
    If (Trim$(szPath) = "") Then Exit Function
    
    'Make sure the path ends with a : or \
    Dim szLast As String
    szLast = Right$(szPath, 1)
    Select Case (szLast)
        Case ":"
        Case "\"
        Case Else
            szPath = szPath & "\"
    End Select
    
    'if the given subdirectory contains any files or subdirectories then it exists
    If ("" <> Dir$(szPath, vbNormal + vbHidden + vbSystem + vbDirectory)) Then
        DoesDirectoryExist = True
    End If
    
DoesDirectoryExistErr:
    Exit Function
End Function
