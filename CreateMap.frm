VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Map Creator"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   -722
   ScaleMode       =   0  'User
   ScaleTop        =   500
   ScaleWidth      =   848
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   2880
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7440
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraControls 
      ClipControls    =   0   'False
      Height          =   10695
      Left            =   8400
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton pbExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   10200
         Width           =   615
      End
      Begin VB.CommandButton pbHelp 
         Height          =   855
         Left            =   1320
         Picture         =   "CreateMap.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   9240
         Width           =   495
      End
      Begin VB.CommandButton pbClearMap 
         Caption         =   "&New Map"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   10200
         Width           =   975
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   6
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3840
         Width           =   615
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   0
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton pbSave 
         Height          =   855
         Left            =   720
         Picture         =   "CreateMap.frx":12D2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   9240
         Width           =   495
      End
      Begin VB.CommandButton pbLoad 
         Height          =   855
         Left            =   120
         Picture         =   "CreateMap.frx":25A4
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   9240
         Width           =   495
      End
      Begin VB.CheckBox chkDeletions 
         Caption         =   "Verify Deletions"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   8880
         Width           =   1395
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   3
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   2
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   1
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   5
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   4
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2640
         Width           =   615
      End
      Begin VB.Frame fraEmpire 
         BackColor       =   &H0000C000&
         Caption         =   "Empire Color"
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   5760
         Width           =   1455
         Begin VB.PictureBox picEmpire 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C000C0&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   250
            Index           =   4
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1800
            Width           =   250
         End
         Begin VB.PictureBox picEmpire 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   250
            Index           =   2
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1080
            Width           =   250
         End
         Begin VB.PictureBox picEmpire 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   250
            Index           =   3
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1440
            Width           =   250
         End
         Begin VB.PictureBox picEmpire 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   250
            Index           =   1
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   720
            Width           =   250
         End
         Begin VB.PictureBox picEmpire 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   250
            Index           =   0
            Left            =   240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   250
         End
         Begin VB.Shape shpEmpire 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            Height          =   495
            Left            =   960
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblEmpire 
            BackStyle       =   0  'Transparent
            Caption         =   "Black"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblEmpire 
            BackStyle       =   0  'Transparent
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblEmpire 
            BackStyle       =   0  'Transparent
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   10
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblEmpire 
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   9
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblEmpire 
            BackStyle       =   0  'Transparent
            Caption         =   "Purple"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   8
            Top             =   1800
            Width           =   615
         End
      End
      Begin VB.PictureBox picUnit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   540
         Index           =   7
         Left            =   120
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Forest"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   33
         Top             =   3960
         Width           =   975
      End
      Begin VB.Shape shpUnit 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         Height          =   495
         Left            =   1200
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 1 Town"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 2 Town"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 3 Town"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Capital"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Gold Mine"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Mountains"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   19
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Blank/Delete"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   18
         Top             =   4560
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Check for duplicate capitals and don't allow before saving"
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Enable Move Somehow"
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmCreateMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The settings in the Scale properties for this form make the lower
'left corner 0,0 and incrementing X moves right, while Y moves up.

Dim mCoor() As clsHex   ' Will contain vertices of all the hexes
Dim r As Long           ' This is radius of circle which would surround the hexes
Dim mxColumns As Long   ' # of columns of hexes to generate
Dim myRows As Long      ' # of rows of hexes to generate
Dim xmStart As Integer  ' Pixel X point to start at
Dim ymStart As Integer  ' Pixel Y point to start at
Dim miCurrentHexPic As eHexPicID
Dim miCurrentEmpire As eEmpire
Dim mxLastHex As Integer
Dim myLastHex As Integer
Dim mbCalledFromMouseMove As Boolean
Dim mbMovingHex As Boolean
Dim mxMovingHex As Integer
Dim myMovingHex As Integer
Dim mbMouseDown As Boolean
Dim mbMouseDisabled As Boolean
Dim mlDefaultGold As Long


Private Sub SetupGrid()
Dim x As Integer
Dim y As Integer
Dim xIncr As Double
Dim yIncr As Double
Dim xFactor As Integer
Dim yFactor As Integer
    
    ReDim mCoor(1 To mxColumns, 1 To myRows)
    
    xIncr = r * 1.5
    yIncr = (r * Sqr(3))
    
    For x = 1 To mxColumns
        For y = 1 To myRows
            
            Set mCoor(x, y) = New clsHex
            
            xFactor = x - 1
            yFactor = y - 1
            
            mCoor(x, y).xCoor = xmStart + (xFactor * xIncr)
            mCoor(x, y).yCoor = ymStart + (yFactor * yIncr)
            
            If (x / 2) = (x \ 2) Then ' if it's an even column
                ' Up more by half of yIncr cuz it's an even column
                mCoor(x, y).yCoor = mCoor(x, y).yCoor + (0.5 * yIncr)
            End If
            
            ' p1 = point at the 1 o'clock position, then move around clockwise
            mCoor(x, y).p1x = mCoor(x, y).xCoor + (r * 0.5)
            mCoor(x, y).p1y = mCoor(x, y).yCoor + (yIncr * 0.5)
            
            mCoor(x, y).p2x = mCoor(x, y).xCoor + r
            mCoor(x, y).p2y = mCoor(x, y).yCoor
            
            mCoor(x, y).p3x = mCoor(x, y).xCoor + (r * 0.5)
            mCoor(x, y).p3y = mCoor(x, y).yCoor - (yIncr * 0.5)
            
            mCoor(x, y).p4x = mCoor(x, y).xCoor - (r * 0.5)
            mCoor(x, y).p4y = mCoor(x, y).yCoor - (yIncr * 0.5)
            
            mCoor(x, y).p5x = mCoor(x, y).xCoor - r
            mCoor(x, y).p5y = mCoor(x, y).yCoor
            
            mCoor(x, y).p6x = mCoor(x, y).xCoor - (r * 0.5)
            mCoor(x, y).p6y = mCoor(x, y).yCoor + (yIncr * 0.5)
            
            mCoor(x, y).HexPicID = BlankID
            
        Next y
    Next x
    
End Sub

Private Sub DrawHex(xGrid As Integer, yGrid As Integer, Optional vColor As Variant)
Dim lColor As Long
    
    If (IsMissing(vColor)) Then lColor = vbBlack Else lColor = vColor
    
    ' Draw single hex
    With mCoor(xGrid, yGrid)
        
        ' Draws a hex
        Line (.p1x, .p1y)-(.p2x, .p2y), lColor
        Line (.p2x, .p2y)-(.p3x, .p3y), lColor
        Line (.p3x, .p3y)-(.p4x, .p4y), lColor
        Line (.p4x, .p4y)-(.p5x, .p5y), lColor
        Line (.p5x, .p5y)-(.p6x, .p6y), lColor
        Line (.p6x, .p6y)-(.p1x, .p1y), lColor
        
        'PSet (.xCoor, .yCoor)
        
    End With
    
End Sub

Private Sub DrawGrid()
Dim x As Integer
Dim y As Integer

    Cls
    For x = 1 To mxColumns
        For y = 1 To myRows
            DrawHex x, y
        Next y
    Next x
    
    With mCoor(mxColumns, myRows)
        fraControls.Left = .xCoor + 25
        fraControls.Top = .yCoor + (r * 2) + 3
        fraControls.Height = .yCoor - mCoor(mxColumns, 1).yCoor + (r * 3) '+ 3
    End With
    
End Sub

Private Sub PaintHex(x As Integer, y As Integer, Optional bRefreshHex As Boolean = False, Optional bRefreshForm As Boolean = True)
Dim xPaint As Long
Dim yPaint As Long
Dim lWhite As Long
    
    If (Not bRefreshHex) Then
        If (mCoor(x, y).HexPicID <> BlankID) And (chkDeletions.Value = vbChecked) Then
            If (vbNo = MsgBox("This Hex already contains an object. Are you sure you want to replace it with the current object? (To not receive this message in the future, uncheck the ""Verify Deletions"" option)." & vbCrLf & vbCrLf & "NOTE: If you wish to edit any existing objects on the map, right-click on them and change their values.", vbYesNo + vbQuestion + vbDefaultButton2, "Replace Object?")) Then
                mbMouseDown = False
                Exit Sub
            End If
        End If
    End If
    
    With mCoor(x, y)
        xPaint = .xCoor
        yPaint = .yCoor
    End With
    
    ' Necessary since coor system on form is diff than BitBlt thinks
    ' BitBlt stills thinks coor system starts from upper left corner
    yPaint = -Me.ScaleHeight - yPaint
    
    ' At this point the upper left corner of pic will go to center of hex,
    ' so we subtract half the dimensions of the pic to get it to center in hex
    Dim vHexPicture As eHexPictures
    vHexPicture = DetermineHexPictureID(miCurrentEmpire, miCurrentHexPic)
    
    xPaint = xPaint - (goHexPictures(vHexPicture).Width \ 2)
    yPaint = yPaint - (goHexPictures(vHexPicture).Height \ 2)
    
    ' Blank it out first
    PaintGraphicID Blank, Me, xPaint, yPaint
    
    If (Not bRefreshHex) Then
        ' Paint correct image from current object select in key
        PaintGraphicID vHexPicture, Me, xPaint, yPaint
        mCoor(x, y).Empire = miCurrentEmpire
        mCoor(x, y).HexPicID = miCurrentHexPic
        mCoor(x, y).TotalGold = mlDefaultGold
    Else
        PaintGraphicFromStats mCoor(x, y).Empire, mCoor(x, y).HexPicID, Me, xPaint, yPaint
    End If
    
    If (bRefreshForm) Then Me.Refresh
    
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    LoadAllGraphics
    Me.DrawWidth = 1
    
    xmStart = 25
    ymStart = 35
    mxColumns = 29
    myRows = 20
    r = 20
    mlDefaultGold = -1  ' means unlimited
    
    mbMouseDown = False
    SetupGrid
    
    ' Create unit legend
    Dim i As Integer
    fraControls.BackColor = Me.BackColor
    fraEmpire.BackColor = Me.BackColor
    chkDeletions.BackColor = Me.BackColor
    chkDeletions.Value = vbChecked
    
    For i = Town1ID To ForestID
        picUnit(i).BackColor = Me.BackColor
    Next i
    
    ' These are just for display on the Control Key
    PaintGraphicID Town1Black, picUnit(Town1ID)
    PaintGraphicID Town2Black, picUnit(Town2ID)
    PaintGraphicID Town3Black, picUnit(Town3ID)
    PaintGraphicID CapitalBlack, picUnit(CapitalID)
    PaintGraphicID GoldBlack, picUnit(GoldID)
    PaintGraphicID Mountains, picUnit(MountainsID)
    PaintGraphicID Forest, picUnit(ForestID)
    
    ' We don't display this one since would look goofy,
    ' so just leave it white for display purposes
    'PaintGraphicID Blank, picUnit(BlankID)
    
    ' Select defaults
    picUnit_MouseDown GoldID, 1, 0, 0, 0
    picEmpire_MouseDown Black, 1, 0, 0, 0
    Screen.MousePointer = vbDefault
    
    ' Since won't have "Save As..." always display dialog when saving
    ' Check if folder "Maps" exists and if doesn't then make it
    If (False = DoesDirectoryExist(App.Path & "\Maps")) Then
        MkDir App.Path & "\Maps"
    End If
    dlg.InitDir = App.Path & "\Maps"
    dlg.Filter = "Map (*.map)|*.map|All Files (*.*)|*.*"
    dlg.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    dlg.CancelError = True
    
    Me.Show
    MsgBox "Click on the Help button (big question mark in lower right corner) to get a short description on working with my Hex Grid Map Creator. Note that I will be adding additional functionality to this Map Creator, such as creating Squads of units on the map and in the towns." & vbCrLf & vbCrLf & "Thanks for giving this a try!", vbInformation, "Additional Note"
    
    Exit Sub
    
ErrHandler:
    MsgBox "CreateMap Form_Load - " & Err.Description, vbExclamation
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
End Sub


' Calculate distance between two points
Private Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Distance = Sqr(((x2 - x1) ^ 2) + ((y2 - y1) ^ 2))
End Function



Private Function FindClosetHex(ByRef x As Single, ByRef y As Single, xHex As Integer, yHex As Integer) As Boolean

Dim xx As Integer
Dim yy As Integer
Dim dDist As Double
Dim dClosest As Double
    
    FindClosetHex = True
    
    ' This proc will detect which center of which hex is closest to the point
    ' on the form you clicked which will be whatever hex you clicked on
    dClosest = 10000
    
    For xx = 1 To mxColumns
        For yy = 1 To myRows
            dDist = Distance(x, y, mCoor(xx, yy).xCoor, mCoor(xx, yy).yCoor)
            If dDist < dClosest Then
                xHex = xx
                yHex = yy
                dClosest = dDist
            End If
        Next yy
    Next xx
    
    ' This is outside the grid coordinates, so return false
    If (dClosest > r) Then FindClosetHex = False
    
End Function


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xHex As Integer
Dim yHex As Integer
    
    If mbMouseDisabled Then Exit Sub
    
    If (False = FindClosetHex(x, y, xHex, yHex)) Then Exit Sub
    
    If (Button = 2) Then    ' Right-click
        Select Case mCoor(xHex, yHex).HexPicID
            Case Town1ID To CapitalID
                DrawHex xHex, yHex, vbWhite      ' Highlight Hex on map
                If (frmEditTown.ShowDlg(mCoor(xHex, yHex))) Then
                    PaintHex xHex, yHex, True
                End If
                DrawHex xHex, yHex, vbBlack       ' Unhighlight Hex on map
            Case GoldID
                DrawHex xHex, yHex, vbWhite      ' Highlight Hex on map
                Dim lTempGold As Long
                If (frmEditGold.ShowDlg(mCoor(xHex, yHex), lTempGold)) Then
                    PaintHex xHex, yHex, True
                    If (lTempGold <> -999) Then ChangeDefaultGold lTempGold
                End If
                DrawHex xHex, yHex, vbBlack       ' Unhighlight Hex on map
        End Select
        Exit Sub
    End If
    
    
    
    ' ************ If have reached here, then left-mouse click ************
    
    
    ' If are moving within the same hex, then do nothing
    If (mbCalledFromMouseMove) And (mxLastHex = xHex) And (myLastHex = yHex) Then Exit Sub
    
    
    ' Need to keep track of last Hex clicked on
    mxLastHex = xHex
    myLastHex = yHex
    
    
    ' Check if hex is occupied and Current control hex isn't blank, then
    '   get ready to move the hex by checking move flag and other values
    ' Also check if mouse is already down from before, in which case don't try to move hex
    If (mCoor(xHex, yHex).HexPicID <> BlankID) And (miCurrentHexPic <> BlankID) _
        And (mbMouseDown = False) Then
        mbMovingHex = True
        mxMovingHex = xHex
        myMovingHex = yHex
        DrawHex xHex, yHex, vbWhite  ' Highlight hex to move
        Exit Sub
    End If
    
    mbMouseDown = True
    
    ' This now prevents overwriting entirely unless it's the delete object
    If (mCoor(xHex, yHex).HexPicID = BlankID) Or (miCurrentHexPic = BlankID) Then
        PaintHex xHex, yHex
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mbMouseDisabled Then Exit Sub
    If mbMovingHex Then Exit Sub
    If (Button <> 1) Then Exit Sub
    mbCalledFromMouseMove = True
    Form_MouseDown Button, Shift, x, y
    mbCalledFromMouseMove = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xHold As Integer
Dim yHold As Integer
    
    mbMouseDown = False
    If (Not mbMovingHex) Then Exit Sub
    If (False = FindClosetHex(x, y, xHold, yHold)) Then Exit Sub
    
    ' If hex is empty, then move the hex by transfering info and blank out previous!!
    If (mCoor(xHold, yHold).HexPicID = BlankID) Then
        mCoor(xHold, yHold).Empire = mCoor(mxMovingHex, myMovingHex).Empire
        mCoor(xHold, yHold).HexPicID = mCoor(mxMovingHex, myMovingHex).HexPicID
        mCoor(xHold, yHold).Name = mCoor(mxMovingHex, myMovingHex).Name
        mCoor(xHold, yHold).TotalGold = mCoor(mxMovingHex, myMovingHex).TotalGold
        mCoor(mxMovingHex, myMovingHex).HexPicID = BlankID
        mCoor(mxMovingHex, myMovingHex).Name = ""
        mCoor(mxMovingHex, myMovingHex).TotalGold = mlDefaultGold
        PaintHex xHold, yHold, True
        PaintHex mxMovingHex, myMovingHex, True
        DrawHex mxMovingHex, myMovingHex, vbBlack
    Else
        ' Tried moving to an occupied hex (could even be itself) so undo
        DrawHex mxMovingHex, myMovingHex
    End If
    
    ' Check off move flag
    mbMovingHex = False
    
End Sub


Private Sub ChangeDefaultGold(lNewGold As Long)
Dim x As Integer
Dim y As Integer
    For x = 1 To mxColumns
        For y = 1 To myRows
            With mCoor(x, y)
                If (.HexPicID = GoldID) Then
                    .TotalGold = lNewGold
                End If
            End With
        Next y
    Next x
    mlDefaultGold = lNewGold
End Sub

Private Sub Form_Resize()
Dim lPixelHeight As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lPixelHeight = ScaleY(Height, vbTwips, vbPixels) - 27
    Me.ScaleTop = lPixelHeight
    Me.ScaleHeight = -lPixelHeight
    
    Me.ScaleWidth = ScaleX(Width, vbTwips, vbPixels) - 8
    
    DrawGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim szMsg As String
    SaveMap True
    UnloadGraphics
    szMsg = "THANK YOU for trying the Map Creator portion of my game which I believe I will eventually entitle ""Battle Squads.""" & vbCrLf & vbCrLf & "Although the Map Creator is the main aspect nearing completion at this point (I will be adding the ability to garrison troops in cities and to create other squads on the map), it will eventually be a strategy play-by-email game with some RPG elements which will allow units to level up, cast spells, etc. The current roster of units planned are..." & vbCrLf & vbCrLf & "Leaders" & vbCrLf & "Soldiers" & vbCrLf & "Archers" & vbCrLf & "Wizards" & vbCrLf & "Clerics" & vbCrLf & "Cavalry" & vbCrLf & "Catapults" & vbCrLf & "Spies (special ability units)" & vbCrLf & vbCrLf & "Along with some other neutral units which you'll be able to either fight or hire. Within your Capital City you will also be able to create builings which will enhance your units abilities on the battlefield." & _
        vbCrLf & vbCrLf & "Again, thanks for checking out my game, and let me know your thoughts!"
    MsgBox szMsg, vbInformation, "Thanks!"
    frmMain.Show
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picUnit_MouseDown Index, Button, 0, 0, 0
End Sub

Private Sub lblEmpire_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picEmpire_MouseDown Index, Button, 0, 0, 0
End Sub

Private Function ClearMap(bPrompt As Boolean)
    If (bPrompt) Then
        If (vbNo = MsgBox("Are you sure you want to clear the current map?", vbYesNo + vbQuestion + vbDefaultButton2, "Verify New Map")) Then
            Exit Function
        End If
    End If
    Screen.MousePointer = vbHourglass
    Me.Cls
    SetupGrid
    DrawGrid
    Screen.MousePointer = vbDefault
    ClearMap = True
End Function

Private Sub pbClearMap_Click()
On Error GoTo ErrHandler
Dim x As Integer
Dim y As Integer
    ClearMap True
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub pbExit_Click()
    Unload Me
End Sub

Private Sub pbHelp_Click()
Dim clc As Collection
Dim frm As frmHelp
    
    Set clc = New Collection
    clc.Add "The Legend along the right hand side is where you select what to place on the Hex Grid Map. Select both the object and the Empire Color you wish it to be associated with by clicking on them. Then left-click on the Hex where you wish the object to be placed. Not all objects (such as Mountains and Forest) are altered by the Empire Color selection." '& vbCrLf & vbCrLf & "Note that the default selected object is a Gold Mine and the default Empire Color is Black (Neutral)."
    clc.Add "EDITING OBJECTS" & vbCrLf & vbCrLf & "To change the settings of an individual object already on the map, simply right-click the object and an ""Object Settings Screen"" will come up. Some objects (such as Mountains and Forests) have no settings."
    clc.Add "MOVING OBJECTS" & vbCrLf & vbCrLf & "Click and hold the left mouse button over an existing object. Then move to the new hex you wish it to be located and release the left mouse button." & vbCrLf & vbCrLf & "NOTE: When moving, be sure you don't have the ""Blank/Delete"" in the Legend selected or the program will think you are deleting objects rather than trying to move them."
    clc.Add "LAND OBJECTS INFO" & vbCrLf & vbCrLf & "Mountains" & vbCrLf & "          - Impassable" & vbCrLf & vbCrLf & "Forest" & vbCrLf & "          - Cuts Movement in Half" & vbCrLf & "          - Hides Squads from other players"
    clc.Add "You may hold down the left mouse button and simply move over the Hex Map to place the same object continually (or delete multiple objects if ""Blank/Delete"" is selected)." & vbCrLf & vbCrLf & "This is so you may quickly create such things as a Mountain range or large Forest. Note that this will NOT overwrite existing objects. (Again, unless ""Blank/Delete"" is selected)."
    clc.Add "NOTE: When later setting up an actual game with a map, you will assign players to certain colored Empires. (Max Human Players will be number of uniquely colored Capitals)." & vbCrLf & vbCrLf & "At that point, you may also set up Empires to be Neutral, meaning no human player controls them, and the forces they are given in the map at beginning of game will be only forces they ever have. Neutrals never attack or move, but will defend if attacked."
    clc.Add "You may Save and Load maps with the buttons near the bottom left corner. There is also a New Map button to start a new map." & vbCrLf & vbCrLf & "That's it!"
    
    Set frm = New frmHelp
    frm.ShowDlg clc, "Map Creator Help"
    Set frm = Nothing
    
End Sub

Private Sub RefreshMap()
Dim x As Integer
Dim y As Integer
    For x = 1 To mxColumns
        For y = 1 To myRows
            PaintHex x, y, True, False
        Next y
    Next x
    Me.Refresh
End Sub

Private Sub pbLoad_Click()
On Error GoTo ErrHandler

Dim szFile As String
Dim iFile As Integer
Dim szMapName As String
Dim szMapDesc As String
Dim x As Integer
Dim y As Integer
Dim iHexPicType As Integer
Dim iEmpire As Integer
Dim szName As String
Dim iTotalGold As Integer
    
    mbMouseDisabled = True
    
    dlg.DialogTitle = "Load Map"
    dlg.ShowOpen
    szFile = dlg.FileName
    
    If (Not ClearMap(True)) Then Exit Sub
    
    iFile = FreeFile
    Open dlg.FileName For Input As iFile
    Me.MousePointer = vbHourglass
    
    Input #iFile, szMapName, mlDefaultGold ', MaxX, MaxY
    Input #iFile, szMapDesc
    
    While Not EOF(iFile)
        Input #iFile, x, y, iHexPicType, iEmpire, szName, iTotalGold
        With mCoor(x, y)
            .HexPicID = iHexPicType
            .Empire = iEmpire
            .Name = szName
            .TotalGold = iTotalGold
        End With
    Wend
    
    RefreshMap
    
    Me.MousePointer = vbDefault
    Close iFile
    tmrMouse.Enabled = True
    Exit Sub
    
ErrHandler:
    Close   ' Close all files
    tmrMouse.Enabled = True
    Me.MousePointer = vbDefault
    If Err = 32755 Then
        Exit Sub    ' Canceled dialog
    End If
    ClearMap False
    MsgBox Err.Description
End Sub


Private Function SaveMap(bPrompt As Boolean) As Boolean
On Error GoTo ErrHandler

Dim iFile As Integer
Dim szMapDesc As String
Dim x As Integer
Dim y As Integer
    
    If (bPrompt) Then
        If (vbNo = MsgBox("Would you like to save the existing map?", vbYesNo + vbQuestion, "Save Map?")) Then
            Exit Function
        End If
    End If
    
    dlg.DialogTitle = "Save Map"
    dlg.ShowSave
    Me.MousePointer = vbHourglass
    
    iFile = FreeFile
    Open dlg.FileName For Output As #iFile
    Write #iFile, "Map Name", mlDefaultGold ', MaxX, MaxY
    szMapDesc = "Map Description"
    Write #iFile, szMapDesc
    
    For x = 1 To mxColumns
        For y = 1 To myRows
            With mCoor(x, y)
                If (mCoor(x, y).HexPicID <> BlankID) Then
                    Write #iFile, x, y, .HexPicID, .Empire, .Name, .TotalGold
                End If
            End With
        Next y
    Next x
    
    Me.MousePointer = vbDefault
    Close iFile
    SaveMap = True
    Exit Function
    
ErrHandler:
    Close   ' Close all files
    Me.MousePointer = vbDefault
    If Err = 32755 Then Exit Function ' Canceled dialog
    MsgBox Err.Description
End Function


Private Sub pbSave_Click()
    SaveMap False
End Sub


Private Sub picEmpire_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim vEmpireID As eEmpire
Dim HexPicID As eHexPicID
    
    If Button <> 1 Then Exit Sub
    
    ' Position Highlighter box around pic
    shpEmpire.Left = picEmpire(Index).Left - 35
    shpEmpire.Top = picEmpire(Index).Top - 35
    shpEmpire.Width = picEmpire(Index).Width + 85
    shpEmpire.Height = picEmpire(Index).Height + 85
    
    ' Change color of towns and gold mine to paint in control box
    vEmpireID = Index
    For HexPicID = Town1ID To GoldID
        PaintGraphicFromStats vEmpireID, HexPicID, picUnit(HexPicID)
        picUnit(HexPicID).Refresh
    Next HexPicID
    
    shpEmpire.Visible = True
    miCurrentEmpire = Index
End Sub

Private Sub picUnit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    ' Position Highlighter box around pic
    shpUnit.Left = picUnit(Index).Left - 35
    shpUnit.Top = picUnit(Index).Top - 35
    shpUnit.Width = picUnit(Index).Width + 75
    shpUnit.Height = picUnit(Index).Height + 75
    
    shpUnit.Visible = True
    miCurrentHexPic = Index
    
End Sub

Private Sub tmrMouse_Timer()
    ' This is necessary because when double-clicking on a file in the Load File
    ' dialog it will also kick off the MouseMove event on the form once the file
    ' has loaded and the dialog is no longer there between the mouse and the form.
    ' So we disable mouse movement or clicking for a 1/10 of a second.
    mbMouseDisabled = False
    tmrMouse.Enabled = False
End Sub
