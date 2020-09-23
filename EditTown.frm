VERSION 5.00
Begin VB.Form frmEditTown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Town Settings"
   ClientHeight    =   3630
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbLevel 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cbEmpire 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   360
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   2460
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton pbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Level"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Empire"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmEditTown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbOK As Boolean
Dim mszName As String
Dim miEmpire As Integer
Dim miLevel As Integer

Private Sub ChangePicture()
    pic.Picture = Nothing   ' Reset and then paint new pic
    PaintGraphicFromStats ColorNameToID(cbEmpire), LevelNameToID(cbLevel), pic, , , 4
End Sub

Private Sub cbEmpire_Click()
    ChangePicture
End Sub

Private Sub cbLevel_Click()
    ChangePicture
End Sub

Private Sub Form_Load()
Dim i As Integer

    With cbEmpire
        For i = Black To Purple
            .AddItem ColorIDToName(i)
        Next i
    End With
    
    With cbLevel
        For i = Town1ID To CapitalID
            .AddItem TownIDToName(i)
        Next i
    End With
    
End Sub

Private Sub pbCancel_Click()
    Unload Me
End Sub

Private Sub pbOK_Click()
    
    If (Len(txtName) = 0) Then
        MsgBox "Must have a city name.", vbInformation
        txtName.SetFocus
        Exit Sub
    End If
    
    mszName = Trim$(txtName)
    miEmpire = ColorNameToID(cbEmpire)
    miLevel = LevelNameToID(cbLevel)
    
    mbOK = True
    Unload Me
    
End Sub


Public Function ShowDlg(ByRef oHex As clsHex) As Boolean
    
    txtName = oHex.Name
    cbEmpire = ColorIDToName(oHex.Empire)
    cbLevel = TownIDToName(oHex.HexPicID)
    PaintGraphicFromStats oHex.Empire, oHex.HexPicID, pic, , , 4
    
    Me.Show vbModal
    
    If (mbOK) Then
        oHex.Name = mszName
        oHex.Empire = miEmpire
        oHex.HexPicID = miLevel
        ShowDlg = mbOK
    End If
    
End Function
