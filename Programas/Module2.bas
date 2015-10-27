Attribute VB_Name = "Module2"
    'declarations for working with Ini files
    'Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias         "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String,         ByVal nSize As Long, ByVal lpFileName As String) As Long
     
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
     
    'Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias         "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String,         ByVal lpFileName As String) As Long
     
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
     
    '// INI CONTROLLING PROCEDURES
    'reads an Ini string
    Public Function ReadIni(Filename As String, Section As String, Key As String, Optional Default As String = "") As String
    Dim RetVal As String * 255, v As Long
      v = GetPrivateProfileString(Section, Key, Default, RetVal, 255, Filename)
        If v <= 0 Then
            ReadIni = Default
        Else
            ReadIni = Left$(RetVal, v)
        End If
    End Function
     
    'reads an Ini section
    'Public Function ReadIniSection(Filename As String, Section As String) As String
    'Dim RetVal As String * 255, v As Long
    '  v = GetPrivateProfileSection(Section, RetVal, 255, Filename)
    '  ReadIniSection = Left(RetVal, v - 1)
    'End Function
     
    'writes an Ini string
    Public Sub WriteIni(Filename As String, Section As String, Key As String, Value As String)
      WritePrivateProfileString Section, Key, Value, Filename
    End Sub
     
    ''writes an Ini section
    'Public Sub WriteIniSection(Filename As String, Section As String, Value As String)
    '  WritePrivateProfileSection Section, Value, Filename
    'End Sub

