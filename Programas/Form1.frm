VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "  Screen Recorder"
   ClientHeight    =   1755
   ClientLeft      =   360
   ClientTop       =   1035
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   6360
      Top             =   6240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   8280
      Top             =   600
   End
   Begin VB.PictureBox PicMk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VB.PictureBox picL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   360
      Picture         =   "Form1.frx":0C42
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim lMinutosLog As Long
Dim lMinutosLogRestantes As Long
Dim lIntervaloCaptura As Long
Dim lFolder As Long
Dim lEspacioLibre As Long
Dim sMaquina As String
Dim sFolder As String
Dim lArchivosEnCarpeta As Long
Dim lEspacioLibreMinMb As Long

Dim fso

Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Ya se esta ejecuntado el programa."
        End
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    STPPixelX = Screen.TwipsPerPixelX
    STPPixelY = Screen.TwipsPerPixelY
    
    
    ScreenX = Screen.Width \ STPPixelX
    ScreenY = Screen.Height \ STPPixelY

    PIC.Cls
    PIC.Width = ScreenX
    PIC.Height = ScreenY
    WriteLog ("Inicio")
    sMaquina = ReadIni(App.Path & "\" & App.EXEName & ".ini", "Main", "Maquina", "Lav73")
    lArchivosEnCarpeta = Val(ReadIni(App.Path & "\" & App.EXEName & ".ini", "Main", "ArchivosEnCarpeta", "1000"))
    lEspacioLibreMinMb = Val(ReadIni(App.Path & "\" & App.EXEName & ".ini", "Main", "EspacioLibreMinMB", "10"))
    lMinutosLog = Val(ReadIni(App.Path & "\" & App.EXEName & ".ini", "Main", "MinutosLog", "60"))
    lIntervaloCaptura = Val(ReadIni(App.Path & "\" & App.EXEName & ".ini", "Main", "mSegudosCaptura", "1000"))
    
    
    If lArchivosEnCarpeta = 0 Then lArchivosEnCarpeta = 1000
    If lEspacioLibreMinMb = 0 Then lEspacioLibreMinMb = 10
    If lMinutosLog = 0 Then lMinutosLog = 60
    If lIntervaloCaptura = 0 Then lIntervaloCaptura = 1000
    
    lMinutosLogRestantes = lMinutosLog
    
    sFolder = App.Path & "\" & sMaquina & " " & Format$(Date, "yyyy-MM-dd") & "\"
    CreaFolder sFolder
    CreaFolder (sFolder & Format$(lFolder, "00000") & "\")
    
    PaintCursor
    Timer1.Interval = lIntervaloCaptura
    Timer2.Interval = 100
    Timer3.Interval = 60000

    VerificarEspacioLibre

    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True




End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    WriteLog "Finalizado"
    WriteIni App.Path & "\" & App.EXEName & ".ini", "Main", "Maquina", sMaquina
    WriteIni App.Path & "\" & App.EXEName & ".ini", "Main", "ArchivosEnCarpeta", Format$(lArchivosEnCarpeta)
    WriteIni App.Path & "\" & App.EXEName & ".ini", "Main", "EspacioLibreMinMB", Format$(lEspacioLibreMinMb)
    WriteIni App.Path & "\" & App.EXEName & ".ini", "Main", "MinutosLog", Format$(lMinutosLog)
    WriteIni App.Path & "\" & App.EXEName & ".ini", "Main", "mSegudosCaptura", Format$(lIntervaloCaptura)
    
    On Error GoTo 0

End Sub



Private Sub Timer1_Timer()

    Dim wLeft      As Long
    Dim wTop       As Long


    On Error GoTo ManejoError
        GetCursorXY MouseScreenX, MouseScreenY
        goX = ScreenX \ 2
        goY = ScreenY \ 2
        ScreenCapture sFolder & "\" & Format$(lFolder, "00000") & "\" & Format$(CNT, "00000000") & ".jpg", PIC, 80
        PIC.Picture = Nothing
        If CNT Mod lArchivosEnCarpeta = 0 Then
            lFolder = lFolder + 1
            CreaFolder (sFolder & Format$(lFolder, "00000") & "\")
            DoEvents
        End If

GoTo Fin
ManejoError:
    WriteLog "Error: " & Err.Number & " " & Err.Description
Fin:
    On Error GoTo 0


End Sub

Private Sub Timer2_Timer()
    Dim XrWindow   As Long
    Dim YrWindow   As Long

    CurrX = goX
    CurrY = goY

End Sub
Private Sub VerificarEspacioLibre()
    lEspacioLibre = FreeDiskSpace(App.Path) / 1048576
    Me.Caption = sMaquina & " " & Format$(lFolder, "000000") & " " & Format$(CNT, "00000000") & " " & Format$(lEspacioLibre) & "Mb"
    If lEspacioLibre <= lEspacioLibreMinMb Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
    
        WriteLog ("FreeDisk:" & vbTab & Format$(lEspacioLibre) & "Mb")
        WriteLog "Fin"
        End
    End If
End Sub
Private Sub Timer3_Timer()
    lMinutosLogRestantes = lMinutosLogRestantes - 1
    If lMinutosLogRestantes <= 0 Then
        WriteLog ("FreeDisk:" & vbTab & Format$(lEspacioLibre) & "Mb")
        lMinutosLogRestantes = lMinutosLog
    End If
    VerificarEspacioLibre
End Sub
Public Sub WriteLog(sLogEntry As String)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim sLogFile As String, sLogPath As String, iLogSize As Long
   Dim f
   
On Error GoTo ErrHandler

   sLogPath = App.Path & "\" & App.EXEName
   sLogFile = sLogPath & ".log"
   
   Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   Debug.Print sLogEntry
   f.WriteLine Now() & vbTab & sLogEntry
   f = Nothing
ErrHandler:
    On Error GoTo 0
    Exit Sub
End Sub
Sub CreaFolder(Folder As String)
    On Error Resume Next
    If Len(Dir(Folder)) = 0 Then
        VerificarEspacioLibre
        MkDir Folder
    Else
        WriteLog ("Folder encontrado:" & lFolder)
        lFolder = lFolder + 1
        CNT = lFolder * lArchivosEnCarpeta
        CreaFolder (sFolder & Format$(lFolder, "00000") & "\")
    End If
    On Error GoTo 0
End Sub

