VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarEmpire As eEmpire
Private mvarHexPicID As eHexPicID
Private mvarxCoor As Double
Private mvaryCoor As Double
Private mvarp1x As Double
Private mvarp1y As Double
Private mvarp2x As Double
Private mvarp2y As Double
Private mvarp3x As Double
Private mvarp3y As Double
Private mvarp4x As Double
Private mvarp4y As Double
Private mvarp5x As Double
Private mvarp5y As Double
Private mvarp6x As Double
Private mvarp6y As Double
Private mvarName As String
Private mvarTotalGold As Long


Private Sub Class_Initialize()
    HexPicID = BlankID
    TotalGold = -1      ' Means it's set to unlimited
End Sub



Public Property Let TotalGold(ByVal vData As Long)
    mvarTotalGold = vData
End Property

Public Property Get TotalGold() As Long
    TotalGold = mvarTotalGold
End Property



Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property



Public Property Let p6y(ByVal vData As Double)
    mvarp6y = vData
End Property

Public Property Get p6y() As Double
    p6y = mvarp6y
End Property




Public Property Let p6x(ByVal vData As Double)
    mvarp6x = vData
End Property

Public Property Get p6x() As Double
    p6x = mvarp6x
End Property



Public Property Let p5y(ByVal vData As Double)
    mvarp5y = vData
End Property

Public Property Get p5y() As Double
    p5y = mvarp5y
End Property



Public Property Let p5x(ByVal vData As Double)
    mvarp5x = vData
End Property

Public Property Get p5x() As Double
    p5x = mvarp5x
End Property



Public Property Let p4y(ByVal vData As Double)
    mvarp4y = vData
End Property

Public Property Get p4y() As Double
    p4y = mvarp4y
End Property



Public Property Let p4x(ByVal vData As Double)
    mvarp4x = vData
End Property

Public Property Get p4x() As Double
    p4x = mvarp4x
End Property



Public Property Let p3y(ByVal vData As Double)
    mvarp3y = vData
End Property

Public Property Get p3y() As Double
    p3y = mvarp3y
End Property



Public Property Let p3x(ByVal vData As Double)
    mvarp3x = vData
End Property

Public Property Get p3x() As Double
    p3x = mvarp3x
End Property



Public Property Let p2y(ByVal vData As Double)
    mvarp2y = vData
End Property

Public Property Get p2y() As Double
    p2y = mvarp2y
End Property



Public Property Let p2x(ByVal vData As Double)
    mvarp2x = vData
End Property

Public Property Get p2x() As Double
    p2x = mvarp2x
End Property



Public Property Let p1y(ByVal vData As Double)
    mvarp1y = vData
End Property

Public Property Get p1y() As Double
    p1y = mvarp1y
End Property



Public Property Let p1x(ByVal vData As Double)
    mvarp1x = vData
End Property

Public Property Get p1x() As Double
    p1x = mvarp1x
End Property



Public Property Let yCoor(ByVal vData As Double)
    mvaryCoor = vData
End Property

Public Property Get yCoor() As Double
    yCoor = mvaryCoor
End Property



Public Property Let xCoor(ByVal vData As Double)
    mvarxCoor = vData
End Property

Public Property Get xCoor() As Double
    xCoor = mvarxCoor
End Property



Private Function IsCity(vHexPicID As eHexPicID) As Boolean
    Select Case vHexPicID
        Case Town1ID To CapitalID
            IsCity = True
    End Select
End Function


Public Property Let HexPicID(ByVal vData As eHexPicID)
Dim oIni As clsIniFile
Dim szSection As String
    
    ' If wasn't a city before, but now will be, then assign random name
    If (Not IsCity(HexPicID)) And (IsCity(vData)) Then
        ' Assign a random name
        Set oIni = New clsIniFile
        oIni.Init App.Path & "\Names.ini"
        szSection = "Town"
        Name = oIni.GetRandomEntry(szSection, szSection)
        Set oIni = Nothing
    End If
    
    mvarHexPicID = vData
    
End Property

Public Property Get HexPicID() As eHexPicID
    HexPicID = mvarHexPicID
End Property



Public Property Let Empire(ByVal vData As eEmpire)
    mvarEmpire = vData
End Property

Public Property Get Empire() As eEmpire
    Empire = mvarEmpire
End Property
