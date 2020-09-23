VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' StretchBlt takes more time than BitBlt, but if StretchBlt is called and won't
'       stretch anything, it's just as fast as BitBlt, so always use StretchBlt


' ****************************************
' USUALLY MAKE AUTODRAW TRUE!!! However...
' ****************************************


' REQUIREMENTS FOR THIS CLASS: Make sure...
' For the forms to print it on...
'       ScaleMode = Pixels
'       AutoDraw = It depends
'                   if set to true, then must use Me.Refresh to display (or can min and max)
'                           PRO: It's always there, even if minimize or another window comes over it
'                           CON: SLOWER!!
'                           PRO: For big images, can't "watch" it paint. If paint, there's a pause, but when Me.Refresh called, image appears instantly
'                           PRO: For moving, there's NO FLICKER!!! (but use me.refresh)
'                   if set to false, then appears instantly without Me.Refresh
'                           CON: Disappears if min and maxed, or another windows comes over, or use Me.Refresh
'                           CON: Disappears if Me.Refresh is called
'                           PRO: FASTER!!
'                   So if Screen Saver, make false with no worries
'                   Otherwise, should probably set to true and use Me.Refresh so can't disappear

' If source is from picture control, then it's properties should be...
'       BorderStyle = 0 (None)
'       AutoSize = True (Makes picture fit "tight" into pic control
'       ScaleMode = Pixels
'       Clipcontrols = FALSE (i THINK that's best, but can experiment)
'       Appearance = 0 (flat, though doesn't matter, i believe, since borderstyle = 0 takes care of this)

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Const IMAGE_BITMAP As Long = 0

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

' LoadImage Constants
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBHeader = &H2000
Private Const LR_CREATEDIBSECTION As Long = 2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000

Private mhBMPSource As Long
Private mhSourceDC As Long
Private mxStart As Long
Private myStart As Long
Private mxWidth As Long
Private myHeight As Long
Private mlPaintType As RasterOpConstants

Private Sub Class_Initialize()
    mhBMPSource = -1
    mlPaintType = vbSrcCopy
End Sub

Private Sub Class_Terminate()
    ' Destroy Source
    If (mhBMPSource <> -1) Then Destroy
End Sub

Public Function Destroy()
    Call DeleteDC(mhSourceDC)
    Call DeleteObject(mhBMPSource)
    mhBMPSource = -1
End Function

Public Property Get StartX() As Long
    StartX = mxStart
End Property
Public Property Get StartY() As Long
    StartY = myStart
End Property
Public Property Get Width() As Long
    Width = mxWidth
End Property
Public Property Get Height() As Long
    Height = myHeight
End Property

' Reference to the virtual hDC if want to use it for anything
Public Property Get SourceDC() As Long
    SourceDC = IIf(mhBMPSource <> -1, mhSourceDC, -1)
End Property

' Checks if file exists
Private Function DoesFileExist(szFile As String) As Boolean
    If ("" = szFile) Then Exit Function
    If ("" = Dir(szFile)) Then Exit Function
    DoesFileExist = True
End Function

Public Function CreateFromFile(szFile As String) As Boolean
On Error GoTo ErrHandler
Dim udtBitMap As BITMAP
    
    If (Not DoesFileExist(szFile)) Then Exit Function
    If (mhBMPSource <> -1) Then Destroy
    
    ' Setup up source bitmap
    mhBMPSource = LoadImage(0, szFile, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE)
    
    'Failure in loading bitmap
    If mhBMPSource = 0 Then Exit Function
    
    'Throw the Bitmap into the Device Context
    mhSourceDC = CreateCompatibleDC(0)
    SelectObject mhSourceDC, mhBMPSource
    
    ' Retrieve size
    Call GetObject(mhBMPSource, Len(udtBitMap), udtBitMap)
    mxWidth = Trim(Str(udtBitMap.bmWidth))
    myHeight = Trim(Str(udtBitMap.bmHeight))
    
    DeleteObject mhBMPSource
    CreateFromFile = True
    Exit Function
    
ErrHandler:
    Exit Function
End Function


' Assumes you want the whole thing if no optional params passed
Public Sub CreateFromControl(ctl As Control, Optional xStart As Long = 0, Optional yStart As Long = 0, Optional xWidth As Long = -1, Optional yHeight As Long = -1)
    ' If not include w & h, then calc from difference
    If (xWidth = -1) Then xWidth = ctl.Width - xStart
    If (yHeight = -1) Then yHeight = ctl.Height - yStart
    CreateFromhDC ctl.hdc, xStart, yStart, xWidth, yHeight
End Sub


' Your standard, run-of-the-mill setup from any old object's hdc
Public Function CreateFromhDC(hdc As Long, xStart As Long, yStart As Long, xWidth As Long, yHeight As Long) As Boolean
    
    If (mhBMPSource <> -1) Then Destroy
    
    ' Assign class variables
    mxStart = xStart
    myStart = yStart
    mxWidth = xWidth
    myHeight = yHeight
    
    ' Set up source bitmap & device context
    mhBMPSource = CreateCompatibleBitmap(hdc, mxWidth, myHeight)
    mhSourceDC = CreateCompatibleDC(hdc)
    SelectObject mhSourceDC, mhBMPSource
    
    ' Copy source to memory, which we will reference from now on
    BitBlt mhSourceDC, 0, 0, mxWidth, myHeight, hdc, mxStart, myStart, vbSrcCopy
    
    DeleteObject mhBMPSource
    CreateFromhDC = True
    
End Function


' ************************ Variable Descriptions *******************************
' hDC = handle of Destination (required)
' xDestStart = x coor start of dest (if blank, then 0)
' yDestStart = y coor start of dest (if blank, then 0)
' xDestWidth = width of dest
'               (if blank, then assume orig width unless overwritten by xSourceWidth)
' yDestHeight = height of dest
'               (if blank, then assume orig height unless overwritten by ySourceHeight)
' xSourceStart = x coor start of source (if blank, then 0)
' ySourceStart = y coor start of source (if blank, then 0)
' xSourceWidth = width of source area (the offset width from xSourceStart)
'               (if blank, then assume want whole thing)
' ySourceHeight = height of source area (the offset width from ySourceStart)
'               (if blank, then assume want whole thing)
' dPercent = if this is passed ALL parameters except xDestStart and yDestStart are
'               ignored. It takes the original bitmap w and h, and stretches to the
'               percent passed.
' ************************ Variable Descriptions *******************************

Public Sub PaintPercent(objCanvas As Object, ByVal dPercent As Double, Optional ByVal xDestStart As Long = 0, Optional ByVal yDestStart As Long = 0)
    PaintForReal objCanvas, xDestStart, yDestStart, Width * dPercent, Height * dPercent, mhSourceDC, 0, 0, Width, Height, mlPaintType
End Sub

Public Property Let PaintType(lPaintType As RasterOpConstants)
    mlPaintType = lPaintType
End Property


Public Sub Paint(objCanvas As Object, Optional ByVal xDestStart As Long = 0, Optional ByVal yDestStart As Long = 0, Optional ByVal xDestWidth As Long = -1, Optional ByVal yDestHeight As Long = -1, Optional ByVal xSourceStart As Long = 0, Optional ByVal ySourceStart As Long = 0, Optional ByVal xSourceWidth As Long = -1, Optional ByVal ySourceHeight As Long = -1, Optional ByVal lPaintType As RasterOpConstants = -1)
    
    ' If no PaintType passed to overwrite, then use current value
    If (lPaintType = -1) Then lPaintType = mlPaintType
    
    If (xDestWidth = -1) Then
        If (xSourceWidth = -1) Then
            ' set to default width of orig
            xDestWidth = Width
            xSourceWidth = Width
        Else
            ' Get New Offset Width to edge of orig pic
            If (xSourceWidth + xSourceStart > Width) Then xSourceWidth = Width - xSourceStart
            xDestWidth = xSourceWidth
        End If
    Else
        If (xSourceWidth = -1) Then
            xSourceWidth = Width
        End If
    End If
    
    If (yDestHeight = -1) Then
        If (ySourceHeight = -1) Then
            ' set to default width of orig
            yDestHeight = Height
            ySourceHeight = Height
        Else
            ' Get New Offset Width to edge of orig pic
            If (ySourceHeight + ySourceHeight > Height) Then ySourceHeight = Height - ySourceStart
            yDestHeight = ySourceHeight
        End If
    Else
        If (ySourceHeight = -1) Then
            ySourceHeight = Height
        End If
    End If
    
    ' If xSourceWidth or ySourceHeight is too great, then make so will
    '       go to end of bitmap for either one
    ' Questionable. If this is called with xSourceWidth and ySourceHeight
    '       greater than Width or Height, normally won't paint at all. But
    '       this will make so paints right to edge of their one. In an ideal
    '       world, this code would never be called since would be a mistake.
    If (xSourceWidth + xSourceStart > Width) Then xSourceWidth = Width - xSourceStart
    If (ySourceHeight + ySourceStart > Height) Then ySourceHeight = Height - ySourceStart
    
    PaintForReal objCanvas, xDestStart, yDestStart, xDestWidth, yDestHeight, mhSourceDC, xSourceStart, ySourceStart, xSourceWidth, ySourceHeight, lPaintType
    
End Sub


Private Sub PaintForReal(objCanvas As Object, xDestStart As Long, yDestStart As Long, xDestWidth As Long, yDestHeight As Long, hSourceDC As Long, xSourceStart As Long, ySourceStart As Long, xSourceWidth As Long, ySourceHeight As Long, lPaintType As Long)
    
    StretchBlt objCanvas.hdc, xDestStart, yDestStart, xDestWidth, yDestHeight, hSourceDC, xSourceStart, ySourceStart, xSourceWidth, ySourceHeight, lPaintType
    
    ' Call refresh only If Canvas (Form or Pic) property AutoRedraw = TRUE
    'If (objCanvas.AutoRedraw) Then objCanvas.Refresh
    
End Sub
