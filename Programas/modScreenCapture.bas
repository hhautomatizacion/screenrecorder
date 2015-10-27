Attribute VB_Name = "modScreenCapture"
Option Explicit

'Public Enum tRecState
'    sStop = 0
'    sREC = 1
'    sPause = 2
'End Enum
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
                                                 ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, _
                                                        ByVal nStretchMode As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" _
                                          (ByVal vKey As Long) As Integer

Public ScreenX     As Long
Public ScreenY     As Long

'Public GrabWindowW          As Long
'Public GrabWindowH          As Long

Public goX         As Single
Public goY         As Single

Public CurrX       As Single
Public CurrY       As Single


'Public FullScreen  As Boolean

Public CNT         As Long

'Public RecState    As tRecState

Public TargethDC As Long

Public MouseScreenX As Long
Public MouseScreenY As Long

Public STPPixelX As Long
Public STPPixelY As Long

Public GrabScreenXStart      As Long
Public GrabScreenYStart      As Long

Public Function ScreenCapture(ByVal filename As String, ByRef toPICbox As PictureBox, Optional ByVal picQuality As Byte = 32)
    'PIC.scalemode must be Pixel
    'FORM must be Pixel
   
    Dim Ret        As Long

    Ret = GetDC(0)
    DoEvents

    TargethDC = toPICbox.hdc

    If Ret Then

        'SetStretchBltMode targethdc, vbPaletteModeNone
        'With Screen
        '    StretchBlt targethdc, 0, 0, toPICbox.Width, toPICbox.Height, Ret, _
             '              goX - GrabWindowW \ 2, goY - GrabWindowH \ 2, GrabWindowW, GrabWindowH, vbSrcCopy
        'End With

        'If GrabScreenXStart < 0 Or GrabScreenYStart < 0 Or _
           GrabScreenXStart + GrabWindowW > ScreenX Or _
           GrabScreenYStart + GrabWindowH > ScreenY Then
        '    BitBlt TargethDC, 0, 0, GrabWindowW, GrabWindowH, TargethDC, _
                   0, 0, vbBlack

        'End If

        BitBlt TargethDC, 0, 0, ScreenX, ScreenY, Ret, _
               GrabScreenXStart, GrabScreenYStart, vbSrcCopy


'        BitBlt TargethDC, -2, GrabWindowH - 1 - Form1.picL.Height, Form1.picL.Width, Form1.picL.Height, _
               Form1.picL.hDC, 0, 0, vbSrcAnd


       ' If Form1.chMouse Then

            'If FullScreen Then
                'Yellow
                'BitBlt TargethDC, -16 + MouseScreenX, -16 + MouseScreenY, _
                       32, 32, Form1.PicYellow.hDC, 0, 0, vbSrcAnd
                'Cursor
                'PaintCursor
                BitBlt TargethDC, -10 + MouseScreenX, -10 + MouseScreenY, _
                       32, 32, Form1.PicM.hdc, 0, 0, vbSrcAnd
                'Click event
                If GetAsyncKeyState(vbLeftButton) Or GetAsyncKeyState(vbRightButton) Then
                    BitBlt TargethDC, -10 + MouseScreenX, -10 + MouseScreenY, _
                           32, 32, Form1.PicMk.hdc, 0, 0, vbSrcAnd
                End If
            'Else
             '   'gox and goY
             '   BitBlt TargethDC, -16 + goX - frmGrabWindow.Left \ STPPixelX, -16 + goY - frmGrabWindow.Top \ STPPixelY, _
             '          32, 32, Form1.PicYellow.hDC, 0, 0, vbSrcAnd
             '   PaintCursor
             '   BitBlt TargethDC, -10 + goX - frmGrabWindow.Left \ STPPixelX, -10 + goY - frmGrabWindow.Top \ STPPixelY, _
             '          32, 32, Form1.PicM.hDC, 0, 0, vbSrcAnd'
'                If GetAsyncKeyState(vbLeftButton) Or GetAsyncKeyState(vbRightButton) Then
'                    BitBlt TargethDC, -10 + goX - frmGrabWindow.Left \ STPPixelX, -10 + goY - frmGrabWindow.Top \ STPPixelY, _
'                           32, 32, Form1.PicMk.hDC, 0, 0, vbSrcAnd
 '               End If
            'End If

    '    End If

        SaveJPG toPICbox.Image, filename, picQuality

    End If
    Ret = ReleaseDC(0, Ret)
    

End Function
