VERSION 5.00
Begin VB.Form frmEMails 
   Caption         =   "Add/Edit Players"
   ClientHeight    =   3945
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pbDelete 
      Caption         =   "&Delete Player..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton pbEdit 
      Caption         =   "&Edit Player..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton pbNew 
      Caption         =   "&New Player..."
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox lstPlayers 
      Height          =   2205
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton pbClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Available Players"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmEMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' I'm writing to an ini file in the event the user may need to move to
' another machine. If I wrote to the registry, then that would not be possible

Dim moIni As clsIniFile

Private Sub LoadNames()
Dim i As Integer
Dim oNames As Collection
Dim oEMails As Collection
    lstPlayers.Clear
    moIni.GetSection "EMails", oNames, oEMails
    For i = 1 To oNames.Count
        lstPlayers.AddItem oNames(i)
    Next i
End Sub

Public Sub ShowDlg()
    
    Set moIni = New clsIniFile
    moIni.Init App.Path & "\Players.ini"
    LoadNames
    
    Me.Show 'vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moIni = Nothing
    frmMain.Visible = True
End Sub

Private Sub lstPlayers_Click()
    If lstPlayers.ListIndex = -1 Then Exit Sub
    pbEdit.Enabled = True
    pbDelete.Enabled = True
End Sub

Private Sub lstPlayers_DblClick()
    pbEdit_Click
End Sub

Private Sub pbClose_Click()
Dim i As Integer
    Unload Me
End Sub

Private Sub pbDelete_Click()
    If (vbNo = MsgBox("Are you sure you want to delete player """ & lstPlayers.Text & "?""", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Player?")) Then
        Exit Sub
    End If
    moIni.DeleteKey "EMails", lstPlayers.Text
    moIni.DeleteKey "Notes", lstPlayers.Text
    lstPlayers.RemoveItem lstPlayers.ListIndex
    lstPlayers.ListIndex = -1
    pbDelete.Enabled = False
    pbEdit.Enabled = False
End Sub

Private Sub pbEdit_Click()
Dim f As frmEMail
Dim szNewName As String

    Set f = New frmEMail
    If f.ShowDlg(lstPlayers.Text, moIni, szNewName) Then
        ' Reload all names
        LoadNames
    End If
    Set f = Nothing
    
    pbEdit.Enabled = False
    pbDelete.Enabled = False
    lstPlayers.ListIndex = -1
    
End Sub

Private Sub pbNew_Click()
Dim f As frmEMail
Dim szNewName As String
    Set f = New frmEMail
    If f.ShowDlg("", moIni, szNewName) Then
        lstPlayers.AddItem szNewName
    End If
    Set f = Nothing
    
    pbEdit.Enabled = False
    pbDelete.Enabled = False
    lstPlayers.ListIndex = -1
    
End Sub
