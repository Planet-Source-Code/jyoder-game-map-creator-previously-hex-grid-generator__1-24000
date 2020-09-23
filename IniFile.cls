VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Property for file to read
Public m_szFile$

' API to read/write ini's
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal szAppname As String, ByVal lpKeyName As Any, ByVal szDefault As String, ByVal szReturnedString As String, ByVal iSize As Integer, ByVal szFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal szNewString As Any, ByVal szFileName As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal szAppname As String, ByVal lpKeyName As Any, ByVal szDefault As String, ByVal szReturnedString As String, ByVal iSize As Integer, ByVal szFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer
#End If

Public Sub Init(ByVal szFile$)
    m_szFile = szFile$
End Sub

Public Function SaveSetting(ByVal szSection As String, ByVal szKey As String, ByVal szValue As String) As Integer
   SaveSetting = WritePrivateProfileString(szSection, szKey, szValue, m_szFile)
End Function

Public Function GetSetting(ByVal szSection As String, ByVal szKey As String, Optional ByVal vDefault As Variant) As String
    Dim iRetValue As Integer
    Dim szTemp As String * 255
    
    ' Get the value
    iRetValue = GetPrivateProfileString(szSection, szKey, "", szTemp, Len(szTemp), m_szFile)
    
    ' If there is one, return it
    If (iRetValue > 0) Then
        GetSetting = Left$(szTemp, iRetValue)
    Else
        ' Unknown section or key
        If IsMissing(vDefault) Then
            GetSetting = ""
        Else
            GetSetting = vDefault
        End If
    End If
    
End Function

Public Sub GetSection(ByVal szSection As String, oKeyClc As Collection, oEntryClc As Collection)
    
    Dim iRetValue As Integer
    Dim szTemp As String * 2500
    Dim szValue As String
    Dim iLastPointer As Integer
    Dim iNullPointer As Integer
    Dim szKey As String
    
    Set oKeyClc = New Collection
    Set oEntryClc = New Collection
    
    ' Get the value
    iRetValue = GetPrivateProfileString(szSection, 0&, "", szTemp, Len(szTemp), m_szFile)
    
    ' If there is one, return it
    If (iRetValue > 0) Then
       
       ' Separate the keys
       iNullPointer = InStr(szTemp, Chr$(0))
       iLastPointer = 1
       
       Do While (iNullPointer <> 0 And iNullPointer >= iLastPointer + 1)
          
            ' Extract key string
            szKey = Mid$(szTemp, iLastPointer, iNullPointer - iLastPointer)
            
            ' Now add to collection
            szValue = GetSetting(szSection, szKey)
            oEntryClc.Add szValue, szKey
            oKeyClc.Add szKey
            
            ' Find next null
            iLastPointer = iNullPointer + 1
            iNullPointer = InStr(iNullPointer + 1, szTemp, Chr$(0))
            
       Loop
    End If
    
End Sub

Public Sub DeleteSection(ByVal Section$)
    Dim iRetCode As Integer
    iRetCode = WritePrivateProfileString(Section$, 0&, "", m_szFile)
End Sub

Public Sub DeleteKey(ByVal szSection As String, ByVal szKey As String)
    Dim iRetCode As Integer
    iRetCode = WritePrivateProfileString(szSection, szKey, 0&, m_szFile)
End Sub

Public Function TotalKeys(ByVal szSection As String) As Integer
    Dim oTempClc As Collection
    Dim oTempClc2 As Collection
    Set oTempClc = New Collection
    Set oTempClc2 = New Collection
    GetSection szSection, oTempClc, oTempClc2
    TotalKeys = oTempClc.Count
End Function

Public Sub GetAllSections(oSectionsClc As Collection)

    Dim szTextLine As String
    Dim szSection As String
    
    Open m_szFile For Input As #1 ' Open file.
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, szTextLine ' Read line into variable.
        szTextLine = Trim$(szTextLine)
        If (Left$(szTextLine, 1) = "[") And (Right$(szTextLine, 1) = "]") Then
            ' It's a section header
            ' Trim off "[" and "]"
            szSection = Mid$(szTextLine, 2, Len(szTextLine) - 2)
            oSectionsClc.Add szSection, szSection
        End If
    Loop
    Close #1

End Sub


Public Function GetRandomEntry(ByVal szSection As String, ByVal szDefault As String) As String
Dim oKeys As Collection
Dim oEntries As Collection
Dim iEntry As Integer
    GetSection szSection, oKeys, oEntries
    If (oKeys.Count = 0) Then
        GetRandomEntry = szDefault
    Else
        iEntry = GetRandomNumber(1, oEntries.Count)
        GetRandomEntry = oEntries(iEntry)
    End If
    Set oKeys = Nothing
    Set oEntries = Nothing
End Function
