Attribute VB_Name = "modGhost"


Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub SetGhost(ByVal hWnd As Long, ByVal Opacity As Byte)

    ' defines the window as click-through, makes it semi-transparent, and sets it to always on top

    Const GWL_EXSTYLE = -20, WS_EX_TRANSPARENT = &H20&, WS_EX_LAYERED = &H80000
    Const LWA_ALPHA = &H2&
    Const SWP_NOMOVE = 2, SWP_NOSIZE = 1, HWND_TOPMOST = -1

    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    SetLayeredWindowAttributes hWnd, 0, Opacity, LWA_ALPHA
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub


Public Sub ResetGhost(ByVal hWnd As Long)

    ' turns ghosting off

    Const GWL_EXSTYLE = -20, WS_EX_TRANSPARENT = &H20&, WS_EX_LAYERED = &H80000
    Const SWP_NOMOVE = 2, SWP_NOSIZE = 1, HWND_NOTOPMOST = -2

    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not (WS_EX_LAYERED Or WS_EX_TRANSPARENT)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
