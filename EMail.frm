VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Player"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEMail 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox txtNotes 
      Height          =   1095
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton pbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Notes"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "EMail"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbOK As Boolean
Dim mszName As String
Dim mszEMail As String
Dim mszNotes As String
Dim moIni As clsIniFile

Public Function ShowDlg(szName As String, oIni As clsIniFile, szNewName As String) As Boolean
    
    txtName = szName
    txtEMail = oIni.GetSetting("EMails", szName)
    txtNotes = oIni.GetSetting("Notes", szName)
    Set moIni = oIni
    Me.Show vbModal
    
    If mbOK Then
        
        ' Delete old name first in case changed the name
        oIni.DeleteKey "EMails", szName
        oIni.DeleteKey "Notes", szName
        
        oIni.SaveSetting "EMails", mszName, mszEMail
        oIni.SaveSetting "Notes", mszName, mszNotes
        szNewName = mszName
        ShowDlg = True
        
    End If
    
End Function


Private Sub pbCancel_Click()
    Unload Me
End Sub


Private Sub pbOK_Click()
Dim szCheckName As String
    
    txtName = Trim$(txtName)
    txtEMail = Trim$(txtEMail)
    txtNotes = Trim$(txtNotes)
    
    If (Len(txtName) = 0) Then
        txtName.SetFocus
        MsgBox "Name is required.", vbExclamation, "Required Field"
        Exit Sub
    End If
    If (Len(txtEMail) = 0) Then
        txtEMail.SetFocus
        MsgBox "EMail is required.", vbExclamation, "Required Field"
        Exit Sub
    End If
    
    szCheckName = moIni.GetSetting("EMails", txtName, "")
    If (szCheckName <> "") Then
        txtName.SetFocus
        MsgBox "Name already exists. Cannot save.", vbExclamation
        Exit Sub
    End If
    
    mbOK = True
    mszName = txtName
    mszEMail = txtEMail
    mszNotes = txtNotes
    Unload Me
    
End Sub
