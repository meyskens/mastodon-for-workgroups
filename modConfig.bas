Attribute VB_Name = "modConfig"
'--------for INI file read/write
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'-------------------
    
Private Function GetPath() As String
    Dim path As String

    path = App.path
    If Right$(path, 1) = "\" Then ' fix for A:\ path
        path = Left(path, Len(path) - 1)
    End If
    
    GetPath = path & "\mastodon.ini"
End Function
    
'reads ini string
Public Function ReadIni(Section As String, Key As String) As String
    Dim RetVal As String * 255, v As Long
    v = GetPrivateProfileString(Section, Key, "", RetVal, 255, GetPath())
    ReadIni = Left(RetVal, v)
End Function
    
'reads ini section
Public Function ReadIniSection(Section As String) As String
    Dim RetVal As String * 255, v As Long
    v = GetPrivateProfileSection(Section, RetVal, 255, GetPath())
    ReadIniSection = Left(RetVal, v - 1)
End Function
    
'writes ini
Public Sub WriteIni(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, GetPath()
End Sub
    
'writes ini section
Public Sub WriteIniSection(Section As String, Value As String)
    WritePrivateProfileSection Section, Value, GetPath()
End Sub

Public Function GetInstance() As String
    GetInstance = ReadIni("auth", "instance")
End Function

Public Function GetToken() As String
    GetToken = ReadIni("auth", "token")
End Function
