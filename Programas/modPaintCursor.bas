Attribute VB_Name = "modPaintCursor"
'**************************************
'Windows API/Global Declarations for :Display Current Mouse Pointer Image
'**************************************
' Get the handle of the window the mouse is over
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
' Retrieves the handle of the current cursor
Private Declare Function GetCursor Lib "user32" () As Long
' Gets the coordinates of the mouse pointer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Gets the PID of the window specified
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
' Gets the PID of the current program
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
' This attaches our program to whichever thread "owns" the cursor at the moment
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
' The next function draws the cursor to picCursor
' Note: If you want to display it in an Image control, use the GetDc API call
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' The POINTAPI type hold the (X,Y) for GetCursorPos()
Private Type POINTAPI
    x              As Long
    y              As Long
End Type
' The following are used for keeping the window always on top. This is optional.
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_TOPMOST = -1
Private Const SWP_NOTOPMOST = -2

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'**************************************
' Name: Display Current Mouse Pointer Image
' Description:This code displays a picture of the current mouse pointer in a PictureBox control. This could be useful for doing screen captures that include the mouse pointer.
' By: Will Brendel
'
' Assumes:Create a Form (frmMain), a PictureBox (picCursor), a Timer (tmrCursor), and a Command Button (cmdExit). Set tmrCursor's interval to 10.
'
' Side Effects:It seems to prevent double-clicking.
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=8252&lngWId=1'for details.'**************************************


' Paints the cursor image to the picturebox
Public Sub PaintCursor()
    Dim pt         As POINTAPI
    Dim hWnd       As Long
    Dim ThreadID   As Long
    Dim CurrentThreadID As Long
    Dim hCursor

'This sub Negative affect the mouse double click event
'Don't know how to solve it


    ' Get the position of the cursor
    GetCursorPos pt
    ' Then get the handle of the window the cursor is over
    hWnd = WindowFromPoint(pt.x, pt.y)

    ' Get the PID of the thread
    ThreadID = GetWindowThreadProcessId(hWnd, vbNull)

    ' Get the thread of our program
    CurrentThreadID = App.ThreadID

    ' If the cursor is "owned" by a thread other than ours, attach to that thread and get the cursor
    If CurrentThreadID <> ThreadID Then
        DoEvents
        AttachThreadInput CurrentThreadID, ThreadID, True
        hCursor = GetCursor()
        AttachThreadInput CurrentThreadID, ThreadID, False
        DoEvents
        DoEvents
        DoEvents
  ' If the cursor is owned by our thread, use GetCursor() normally
    Else
        hCursor = GetCursor()
    End If

    ' Use DrawIcon to draw the cursor to picM
    
    'clear picM
    BitBlt Form1.PicM.hDC, 0, 0, 32, 32, Form1.PicM.hDC, 0, 0, vbWhite
    'DrawIcon
    DrawIcon Form1.PicM.hDC, 0, 0, hCursor
    Form1.PicM.Refresh

End Sub


