VERSION 5.00
Begin VB.Form frmEditGold 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Town Settings"
   ClientHeight    =   3630
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pbDefault 
      Caption         =   "Apply Gold to all mines on map"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox cbTotalGold 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox cbEmpire 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   2460
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton pbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblNote 
      Caption         =   "NOTE:"
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Total Gold"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Empire"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmEditGold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbOK As Boolean
Dim miEmpire As Integer
Dim miTotalGold As Long
Dim mlDefaultGold As Long

Private Sub ChangePicture()
    pic.Picture = Nothing   ' Reset and then paint new pic
    PaintGraphicFromStats ColorNameToID(cbEmpire), GoldID, pic, , , 4
End Sub

Private Sub cbEmpire_Click()
    ChangePicture
End Sub


Private Sub Form_Load()
Dim i As Integer

    With cbEmpire
        For i = Black To Purple
            .AddItem ColorIDToName(i)
        Next i
    End With
    cbTotalGold.AddItem "(unlimited)"
    lblNote.Caption = "Gold Mines always generate " & GOLD_PER_MINE_PER_TURN & " gold per turn. Total Gold is the amount of gold this mine begins with at the start of the game."
    
End Sub

Private Sub pbCancel_Click()
    Unload Me
End Sub

Private Sub pbDefault_Click()
    If (Not Validate()) Then
        Exit Sub
    End If
    If (vbYes = MsgBox("This will change all mines on the current map to have a Total Gold amount of " & cbTotalGold & " as well as any new mines you later create on this map. Are you sure you want to do this?", vbYesNo + vbQuestion + vbDefaultButton2, "Change Default Gold Per Mine?")) Then
        mlDefaultGold = IIf(cbTotalGold = "(unlimited)", -1, cbTotalGold)
        pbOK_Click
    End If
End Sub

Private Function Validate() As Boolean
    
    If (Len(cbTotalGold) = 0) Then
        MsgBox "Must have a gold amount.", vbInformation
        cbTotalGold.SetFocus
        Exit Function
    End If
    
    If (Not IsNumeric(cbTotalGold)) Then
        If (cbTotalGold = "(unlimited)") Then
            GoTo GOLD_OK
        End If
    Else    ' Is numeric
        If (cbTotalGold > 0) Then
            GoTo GOLD_OK
        End If
    End If
    
    MsgBox "The amount of gold must either be ""(unlimited)"" or a number amount greater than 0."
    cbTotalGold.SetFocus
    Exit Function
    
GOLD_OK:
    Validate = True
End Function

Private Sub pbOK_Click()
    
    If (Not Validate()) Then
        Exit Sub
    End If
    
    miTotalGold = IIf(cbTotalGold = "(unlimited)", -1, cbTotalGold)
    miEmpire = ColorNameToID(cbEmpire)
    
    mbOK = True
    Unload Me
    
End Sub


Public Function ShowDlg(ByRef oHex As clsHex, ByRef lGoldDefault As Long) As Boolean
    
    cbEmpire = ColorIDToName(oHex.Empire)
    cbTotalGold = IIf(oHex.TotalGold = -1, "(unlimited)", oHex.TotalGold)
    lGoldDefault = -999
    mlDefaultGold = -999
    mbOK = False
    
    PaintGraphicFromStats oHex.Empire, oHex.HexPicID, pic, , , 4
    
    Me.Show vbModal
    
    If (mbOK) Then
        oHex.TotalGold = miTotalGold
        oHex.Empire = miEmpire
        lGoldDefault = mlDefaultGold
        ShowDlg = mbOK
    End If
    
End Function
