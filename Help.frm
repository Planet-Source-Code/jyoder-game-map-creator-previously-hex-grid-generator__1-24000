VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help Screen"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pbPrevious 
      Caption         =   "<< &Previous"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton pbNext 
      Caption         =   "&Next >>"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblPage 
      Alignment       =   2  'Center
      Caption         =   "Page"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim miNext As Integer
Dim mo As Collection

Public Function ShowDlg(clc As Collection, Optional szCaption As String = "")
Dim i As Integer
    
    If (szCaption <> "") Then Me.Caption = szCaption
    
    miNext = 1
    Set mo = New Collection
    For i = 1 To clc.Count
        mo.Add clc(i)
    Next i
    NextScreen
    
    Me.Show vbModal
    
    Set mo = Nothing
    Unload Me
    
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub pbNext_Click()
    miNext = miNext + 1
    If miNext > mo.Count Then Unload Me: Exit Sub
    NextScreen
End Sub

Private Sub NextScreen()
    pbNext.Enabled = True
    pbPrevious.Enabled = True
    lblHelp.Caption = mo(miNext)
    lblPage = "Page " & miNext & " of " & mo.Count
    If (miNext = 1) Then pbPrevious.Enabled = False
    If (miNext = mo.Count) Then pbNext.Enabled = False
End Sub

Private Sub pbPrevious_Click()
    If miNext = 1 Then Exit Sub
    miNext = miNext - 1
    NextScreen
End Sub
