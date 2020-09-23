VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Title Screen"
   ClientHeight    =   3465
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pbExit 
      Caption         =   "E&xit Game"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton pbPlay 
      Caption         =   "&Load Game (not done yet)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton pbNewGame 
      Caption         =   "&New Game (not done yet)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton pbMap 
      Caption         =   "&Create/Edit Maps"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton pbEMails 
      Caption         =   "&Setup Player Emails"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "<-- The Good Part you need to see"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "<-- Simple Address Book (check if you want)."
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If (False = ScreenTest(Me)) Then
        UnloadGraphics
        Unload Me
    End If
    Me.Show
    pbMap.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadGraphics
End Sub

Private Sub pbEMails_Click()
    Me.Hide
    frmEMails.ShowDlg
End Sub

Private Sub pbExit_Click()
    Unload Me
End Sub

Private Sub pbMap_Click()
    Me.Hide
    Screen.MousePointer = vbHourglass
    DoEvents
    DoEvents
    frmCreateMap.Show
End Sub
