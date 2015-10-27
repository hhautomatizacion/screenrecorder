Attribute VB_Name = "modAVIwrite"
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszLongPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long    'BOOL

Const STRETCHMODE = vbPaletteModeNone    'You can find other modes in the "PaletteModeConstants" section of your Object Browser
Private Currframe  As Long

Private fNames()   As String
Private NF         As Long

Public Sub CreateFrameList(ByVal SourcePath As String, Optional Extension As String = "jpg")
    Dim S          As String
    Dim SearchPath As String

    NF = 0

    If right$(SourcePath, 1) <> "\" Then SourcePath = SourcePath & "\"
    SearchPath = SourcePath & "*." & Extension

    S = Dir(SearchPath)
    If S <> "" Then
        ReDim fNames(1)
        fNames(1) = SourcePath & S
        NF = 1
        Do
            S = Dir
            If S <> "" Then
                NF = NF + 1
                ReDim Preserve fNames(NF)
                fNames(NF) = SourcePath & S
            End If
        Loop While S <> ""
    End If


End Sub

Public Sub WriteAVI(ByVal filename As String, ByVal FrameRate As Long, SourcePath As String, Optional Extension As String = "jpg")
    Dim S$
    Dim InitDir    As String
    Dim szOutputAVIFile As String
    Dim res        As Long
    Dim pfile      As Long        'ptr PAVIFILE
    Dim bmp        As cDIB
    Dim ps         As Long        'ptr PAVISTREAM
    Dim psCompressed As Long      'ptr PAVISTREAM
    Dim strhdr     As AVI_STREAM_INFO
    Dim BI         As BITMAPINFOHEADER
    Dim opts       As AVI_COMPRESS_OPTIONS
    Dim pOpts      As Long
    Dim I          As Long
    Dim TEMPO
    Dim Doubled    As Integer
    Dim RealFrame  As Long
    Dim lHwnd      As Long

    Dim CurFile    As String

    lHwnd = Form1.hWnd


    CreateFrameList SourcePath, Extension


    'get an avi filename from user
    szOutputAVIFile = filename$
    '    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB

    S$ = App.Path & IIf(right$(App.Path, 1) <> "\", "\", "") & "temp.bmp"

    Set Form1.picTemp.Picture = LoadPicture(fNames(1))
    '
    Call SetStretchBltMode(Form1.PicSave.hDC, STRETCHMODE)
    Call StretchBlt(Form1.PIC.hDC, 0, 0, Form1.PIC.ScaleWidth - 1, Form1.PIC.ScaleHeight - 1, _
                    Form1.picTemp.hDC, 0, 0, Form1.picTemp.ScaleWidth, Form1.picTemp.ScaleHeight, vbSrcCopy)
    Form1.PicSave.Refresh
    Set Form1.PicSave.Picture = Form1.PicSave.Image

    Call SavePicture(Form1.PicSave.Picture, S$)

    If Form1.PicSave.Width = Form1.picTemp.Width Then
        If Form1.PicSave.Height = Form1.picTemp.Height Then
            Form1.PicSave.Visible = False
        Else
            Form1.PicSave.Visible = True
        End If
    Else
        Form1.PicSave.Visible = True
    End If


    If bmp.CreateFromFile(S$) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

    '   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&          '// default AVI handler
        .dwScale = 1
        .dwRate = Val(FrameRate)  '// fps
        .dwSuggestedBufferSize = bmp.SizeImage    '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)    '// rectangle for stream
    End With

    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

    '   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error


    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(lHwnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then              'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If


    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error

    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With

    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error
    TEMPO = Timer

    S$ = App.Path & IIf(right$(App.Path, 1) <> "\", "\", "") & "temp.bmp"

    For Currframe = 1 To NF

        Set Form1.picTemp.Picture = LoadPicture(fNames(Currframe))

        If (Form1.PicSave.Width = Form1.picTemp.Width) And _
           (Form1.PicSave.Height = Form1.picTemp.Height) Then

            Set Form1.PicSave.Picture = Form1.picTemp.Image
            Call SavePicture(Form1.PicSave.Picture, S$)
            bmp.CreateFromFile (S$)    'load the bitmap (ignore errors)
            'End If
            Form1.PicSave.Visible = True

        Else
            Call SetStretchBltMode(Form1.PicSave.hDC, STRETCHMODE)
            Call StretchBlt(Form1.PicSave.hDC, 0, 0, Form1.PicSave.ScaleWidth, Form1.PicSave.ScaleHeight, Form1.picTemp.hDC, 0, 0, Form1.picTemp.ScaleWidth, Form1.picTemp.ScaleHeight, vbSrcCopy)
            Form1.PicSave.Refresh
            Set Form1.PicSave.Picture = Form1.PicSave.Image
            Call SavePicture(Form1.PicSave.Picture, S$)
            bmp.CreateFromFile (S$)    'load the bitmap (ignore errors)
        End If


        '        For Doubled = 0 To Val(txtEXTRA)
        '        res = AVIStreamWrite(psCompressed, currFrame - 1, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        res = AVIStreamWrite(psCompressed, RealFrame, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
        RealFrame = RealFrame + 1
        '        Next


        Form1.LabelFrame = Currframe & "/" & NF & "  " & Format(((NF - Currframe) / Currframe) * (TEMPO - Timer) / 86400, "hh:nn:ss")
        DoEvents

        'txtINFO = Currframe & " / " & NF & vbCrLf
        'txtINFO = txtINFO & Format((TEMPO - Timer) / 86400, "hh:mm:ss") & "   " & Format(((lvwFiles.ListItems.Count - Currframe) / Currframe) * (TEMPO - Timer) / 86400, "hh:mm:ss")
        'txtINFO = txtINFO & vbCrLf & "Video Lenght " & Format(Currframe \ FrameRate, "##") & "  (Complete " & Format(lvwFiles.ListItems.Count \ FrameRate, "##") & ")"

        DoEvents
    Next


error:
    '   Now close the file
    Set bmp = Nothing

    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
End Sub

