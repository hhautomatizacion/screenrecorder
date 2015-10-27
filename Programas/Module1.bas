Attribute VB_Name = "Module1"
Public Type LargeInt
  lngLower As Long
  lngUpper As Long
End Type
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As LargeInt, lpTotalNumberOfBytes As LargeInt, lpTotalNumberofFreeBytes As LargeInt) As Long

'Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Function FreeDiskSpace(ByVal sDriveLetter As String) As Double
'---------------------------------------------------------------------------------------
' Procedure : FreeDiskSpace
' DateTime  : 02/01/2004 12:20
' Author    : Luke
' Purpose   : Returns Free Disk Space using LARGE INT method (Over 2Gb Drives)
'---------------------------------------------------------------------------------------
'


Dim udtFreeBytesAvail As LargeInt, udtTtlBytes As LargeInt
Dim udtTTlFree As LargeInt
Dim dblFreeSpace As Double

If GetDiskFreeSpaceEx(sDriveLetter, udtFreeBytesAvail, udtTtlBytes, udtTTlFree) Then
       
        If udtFreeBytesAvail.lngLower < 0 Then
           dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower + 4294967296#
        Else
           dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower
        End If

End If

FreeDiskSpace = dblFreeSpace

End Function
