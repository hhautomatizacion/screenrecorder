Attribute VB_Name = "modCursorPos"


Option Explicit

Private Type POINTAPI
    x              As Long
    y              As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
                                      (lpPoint As POINTAPI) As Long


Public Sub GetCursorXY(ByRef x As Long, ByRef y As Long)
    Dim pt         As POINTAPI
    GetCursorPos pt
    x = pt.x
    y = pt.y

End Sub

Public Function GetXCursorPos() As Long
    Dim pt         As POINTAPI
    GetCursorPos pt
    GetXCursorPos = pt.x
End Function

Public Function GetYCursorPos() As Long
    Dim pt         As POINTAPI
    GetCursorPos pt
    GetYCursorPos = pt.y
End Function
