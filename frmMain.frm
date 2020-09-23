VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB MPlayer3"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer tmrFile 
      Interval        =   100
      Left            =   5640
      Top             =   240
   End
   Begin VB.Timer tmrMPlayer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
      Begin VB.ListBox lstNames 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FFFF&
         Height          =   2625
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0BC2
         Left            =   30
         List            =   "frmMain.frx":0BC4
         TabIndex        =   0
         Top             =   120
         Width           =   4400
      End
      Begin VB.ListBox lstPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0BC6
         Left            =   480
         List            =   "frmMain.frx":0BC8
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.PictureBox picMP3 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1260
      TabIndex        =   22
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgSlider 
      Height          =   180
      Left            =   360
      Picture         =   "frmMain.frx":0BCA
      Top             =   990
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MPlayer3 by Frederico Machado"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblbitrate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "kbps"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   630
      Width           =   375
   End
   Begin VB.Label lblKhz 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "kHz"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   630
      Width           =   375
   End
   Begin VB.Label lblmode 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   630
      Width           =   975
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   0
      Left            =   3120
      Picture         =   "frmMain.frx":115C
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Index           =   0
      Left            =   2640
      Picture         =   "frmMain.frx":11F0
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Index           =   0
      Left            =   2160
      Picture         =   "frmMain.frx":1285
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image imgPause 
      Height          =   300
      Index           =   0
      Left            =   1680
      Picture         =   "frmMain.frx":130A
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":1398
      Top             =   1320
      Width           =   1050
   End
   Begin VB.Image imgEject 
      Height          =   300
      Index           =   0
      Left            =   3840
      Picture         =   "frmMain.frx":1441
      Top             =   1320
      Width           =   330
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   150
      Left            =   360
      Top             =   1005
      Width           =   3975
   End
   Begin VB.Label lblRate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   900
      TabIndex        =   14
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblAddDir 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   5355
      Width           =   375
   End
   Begin VB.Label lblDelFile 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   630
      TabIndex        =   11
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label lblNewList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1170
      TabIndex        =   10
      Top             =   5340
      Width           =   225
   End
   Begin VB.Label lblListUP 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblListDOWN 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   5475
      Width           =   255
   End
   Begin VB.Label lblAddFile 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3375
      TabIndex        =   7
      Top             =   5370
      Width           =   375
   End
   Begin VB.Label lblOpenPList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1485
      TabIndex        =   6
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label lblSavePList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1845
      TabIndex        =   5
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image imgPlist 
      Height          =   195
      Index           =   2
      Left            =   8280
      Picture         =   "frmMain.frx":14C7
      Top             =   1680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgPlist 
      Height          =   195
      Index           =   1
      Left            =   7800
      Picture         =   "frmMain.frx":1A0D
      Top             =   1680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblVol 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   510
      TabIndex        =   2
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   220
      TabIndex        =   1
      Top             =   1880
      Width           =   230
   End
   Begin VB.Image imgEject 
      Height          =   300
      Index           =   2
      Left            =   8040
      Picture         =   "frmMain.frx":1D5B
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgEject 
      Height          =   300
      Index           =   1
      Left            =   8040
      Picture         =   "frmMain.frx":1DB6
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   312
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Index           =   2
      Left            =   4920
      Picture         =   "frmMain.frx":1E3C
      Top             =   660
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgPause 
      Height          =   300
      Index           =   2
      Left            =   6120
      Picture         =   "frmMain.frx":1EAB
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Index           =   2
      Left            =   6600
      Picture         =   "frmMain.frx":1F0C
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Index           =   2
      Left            =   7080
      Picture         =   "frmMain.frx":1F66
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   2
      Left            =   7560
      Picture         =   "frmMain.frx":1FD0
      Top             =   660
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Index           =   1
      Left            =   4920
      Picture         =   "frmMain.frx":203C
      Top             =   1020
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgPause 
      Height          =   300
      Index           =   1
      Left            =   6120
      Picture         =   "frmMain.frx":20E5
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Index           =   1
      Left            =   6600
      Picture         =   "frmMain.frx":2173
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Index           =   1
      Left            =   7080
      Picture         =   "frmMain.frx":21F8
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgNext 
      Height          =   300
      Index           =   1
      Left            =   7560
      Picture         =   "frmMain.frx":228D
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlist 
      Height          =   195
      Index           =   0
      Left            =   4200
      Picture         =   "frmMain.frx":2321
      Top             =   1920
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":266F
      Top             =   1920
      Width           =   525
   End
   Begin VB.Image imgPListBar 
      Height          =   300
      Left            =   360
      Picture         =   "frmMain.frx":2D71
      Top             =   5325
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   840
      Picture         =   "frmMain.frx":2F83
      Top             =   1920
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currpos As String
Dim CurrentTime As String
Dim TotalFrames As String
Dim TotalTime As String
Dim FramesPerSecond As String
Dim Paused As Boolean
Dim MP3Path As String
Dim strLOpenPath As String
Dim strLSavePath As String
Dim PLVisible As Boolean
Dim SlideFlag As Boolean
Dim IX, IY, TX, TY, FX, FY

Public Sub CloseMP3()
  Dim Result As String
  Result = CloseMultimedia(AliasName)
  If Result = "Success" Then
    tmrMPlayer3.Enabled = False
    currpos = 0
  End If
End Sub

Public Sub OpenMP3(FileName As String)
  Dim typeDevice As String
  Dim Result As String
  typeDevice = "MPEGVideo"
  Result = OpenMultimedia(picMP3.hwnd, AliasName, FileName, typeDevice)
  If Result = "Success" Then
    On Error Resume Next
    lblName = lstNames.List(lstNames.ListIndex)
    ReadMP3Header FileName
    If frmRate.Visible = True Then frmRate.txtRate = GetRate(AliasName)
    FramesPerSecond = GetFramesPerSecond(AliasName)
    TotalFrames = GetTotalframes(AliasName)  'Get total frames
    TotalTime = GetTotalTimeByMS(AliasName) / 1000   'Get Total Time
    tmrMPlayer3.Enabled = True
    strFilePath = FileName
  End If
End Sub

Public Sub PauseMP3()
  Dim Result As String
  Result = PauseMultimedia(AliasName)
End Sub

Public Sub PlayMP3()
  Dim Result As String
  imgSlider.Move 24, 66: imgSlider.Visible = True
  Result = PlayMultimedia(AliasName, 0, 0)
End Sub

Public Sub ResumeMP3()
  Dim Result As String
  Result = ResumeMultimedia(AliasName)
End Sub

Public Sub StopMP3()
  Dim Result As String
  If SlideFlag = True Then
    Result = StopMultimedia(AliasName)
    Exit Sub
  End If
  imgSlider.Visible = False
  lblName = "MPlayer3 by Frederico Machado"
  lblbitrate = ""
  lblKhz = ""
  lblmode = ""
  lblTime = ":"
  Result = StopMultimedia(AliasName)
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
    Dim SaveTitle As String
    If Command <> "" Then
      SaveSetting App.Title, "Config", "File", Replace(Command, Chr$(34), "")
    End If
    SaveTitle = App.Title
    App.Title = "": Caption = ""
    AppActivate SaveTitle
    End
  End If
  SaveSetting App.Title, "Config", "File", ""
  If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
    SetDefaultDevice "MPEGVideo", "mciqtz.drv"
  End If
  XLeft = GetSetting(App.Title, "Config", "X", Me.left)
  YTop = GetSetting(App.Title, "Config", "Y", Me.top)
  Me.Move XLeft, YTop
  Show
  If Dir$(App.Path & "\plist.m3u") <> "" Then
    OpenPList App.Path & "\plist.m3u"
    lstIndex = GetSetting(App.Title, "Config", "Index", "0")
    If lstIndex = 0 Then
      lstNames.ListIndex = 0
    Else
      lstNames.ListIndex = Val(lstIndex)
    End If
  End If
  If GetSetting(App.Title, "Config", "SPLS", "True") = "True" Then
    imgPlist_Click 0
  End If
  If GetSetting(App.Title, "Config", "SRS", "False") = "True" Then
    frmRate.Show , Me
  End If
  strLOpenPath = GetSetting(App.Title, "Config", "LOpenPath", App.Path)
  strLSavePath = GetSetting(App.Title, "Config", "LSavePath", App.Path)
  strFilePath = Command
  strFilePath = Replace(strFilePath, Chr$(34), "")
  If strFilePath <> "" Then
    If LCase(Right$(strFilePath, 4)) = ".m3u" Then
      OpenPList strFilePath
      DoEvents
      lstNames.ListIndex = 0: lstNames_DblClick
    Else
      SOpenPlayFile
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  tmrMPlayer3.Enabled = False
  tmrFile.Enabled = False
  Dim Result As String
  Result = CloseAll()
  SaveSetting App.Title, "Config", "X", Str$(Me.left)
  SaveSetting App.Title, "Config", "Y", Str$(Me.top)
  If lstNames.ListCount >= 2 Then
    SaveSetting App.Title, "Config", "Index", lstNames.ListIndex
  Else
    SaveSetting App.Title, "Config", "Index", "0"
  End If
  If lstPath.ListCount > 0 Then
    file = LTrim$(App.Path & "\plist.m3u")
    Open file For Output As #1
      For i = 0 To lstPath.ListCount - 1
        Print #1, lstPath.List(i)
      Next
    Close #1
  ElseIf lstPath.ListCount = 0 Then
    On Error Resume Next
    Kill App.Path & "\plist.m3u"
  End If
  End
End Sub

Private Sub imgEject_Click(Index As Integer)
  Dim CD As New clsDialog
  temp = CD.OpenDialog(Me, "MP3 Files (*.MP3) |*.mp3|", "Open File", strLOpenPath)
  If temp = "" Then Exit Sub
  lstPath.Clear
  lstNames.Clear
  file = LTrim$(temp)
  lstPath.AddItem file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: P$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
  lstNames.AddItem b$
  SaveSetting App.Title, "Config", "LOpenPath", P$
  strLOpenPath = P$
  imgPlay_Click (0)
End Sub

Private Sub imgEject_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgEject(0) = imgEject(2)
End Sub

Private Sub imgEject_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgEject(0) = imgEject(1)
End Sub

Private Sub imgNext_Click(Index As Integer)
  If lstNames.ListIndex = lstNames.ListCount - 1 Then Exit Sub
  lstNames.ListIndex = lstNames.ListIndex + 1
  lstNames_DblClick
End Sub

Private Sub imgNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext(0) = imgNext(2)
End Sub

Private Sub imgNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext(0) = imgNext(1)
End Sub

Private Sub imgPause_Click(Index As Integer)
  PauseMP3
  Paused = True
End Sub

Private Sub imgPause_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause(0) = imgPause(2)
End Sub

Private Sub imgPause_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause(0) = imgPause(1)
End Sub

Private Sub imgPlay_Click(Index As Integer)
  If Paused = True Then
    ResumeMP3
    Paused = False
    Exit Sub
  End If
  If strFilePath <> "" Then
    If strFilePath = lstPath.List(lstNames.ListIndex) Then
      PlayMP3
    Else
      lstNames_DblClick
    End If
  Else
    If lstPath.ListCount = 0 Then Exit Sub
    If lstNames.SelCount = 0 Then
      lstNames.ListIndex = 0
      lstNames_DblClick
    End If
  End If
End Sub

Private Sub imgPlay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay(0) = imgPlay(2)
End Sub

Private Sub imgPlay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay(0) = imgPlay(1)
End Sub

Private Sub imgPlist_Click(Index As Integer)
  If PLVisible = False Then
    Height = 6105
    PLVisible = True
    imgPlist(0) = imgPlist(2)
  Else
    Height = 2655
    PLVisible = False
    imgPlist(0) = imgPlist(1)
  End If
End Sub

Private Sub imgPrev_Click(Index As Integer)
  If lstNames.ListIndex <= 0 Then Exit Sub
  lstNames.ListIndex = lstNames.ListIndex - 1
  lstNames_DblClick
End Sub

Private Sub imgPrev_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev(0) = imgPrev(2)
End Sub

Private Sub imgPrev_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev(0) = imgPrev(1)
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SlideFlag = False Then
    IX = X: FX = imgSlider.left
    TX = Screen.TwipsPerPixelX
    SlideFlag = True
  End If
End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SlideFlag = True Then
    pos = FX + (X - IX) / TX
    If pos < 24 Then pos = 24
    If pos > 260 Then pos = 260
    FX = pos: imgSlider.left = pos
  End If
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim From As String
  From = Int(((imgSlider.left - 24) / 229) * TotalFrames)
  StopMP3
  Result = PlayMultimedia(AliasName, From, 0)
  SlideFlag = False
End Sub

Private Sub imgStop_Click(Index As Integer)
  StopMP3
  CloseMP3
  SetFocus
End Sub

Private Sub imgStop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop(0) = imgStop(2)
End Sub

Private Sub imgStop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop(0) = imgStop(1)
End Sub

Private Sub lblAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub lblAddDir_Click()
  MP3Path = BrowseForDirectory
  temp = Dir$(MP3Path & "\*.mp3")
  While Len(temp) > 0
    lstPath.AddItem MP3Path & "\" & temp
    i = InStr(temp, ".")
    If i > 0 Then temp = left$(temp, i - 1)
    lstNames.AddItem temp
    temp = Dir$
  Wend
End Sub

Private Sub lblAddFile_Click()
  Dim CD As New clsDialog
  temp = CD.OpenDialog(Me, "MP3 Files (*.MP3) |*.mp3|", "Add File", strLOpenPath)
  If temp = "" Then Exit Sub
  file = LTrim$(temp)
  lstPath.AddItem file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: P$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
  lstNames.AddItem b$
  SaveSetting App.Title, "Config", "LOpenPath", P$
  strLOpenPath = P$
End Sub

Private Sub lblDelFile_Click()
  If lstNames.ListCount = 0 Then Exit Sub
  If lstNames.SelCount = 0 Then Exit Sub
  If lstNames.ListIndex = lstNames.ListCount - 1 Then
    temp = lstNames.ListIndex - 1
  Else
    temp = lstNames.ListIndex
  End If
  lstPath.RemoveItem lstNames.ListIndex
  lstNames.RemoveItem lstNames.ListIndex
  lstNames.ListIndex = temp
End Sub

Private Sub lblListDOWN_Click()
  ListMove 1
End Sub

Private Sub lblListUP_Click()
  ListMove -1
End Sub

Private Sub lblNewList_Click()
  lstNames.Clear
  lstPath.Clear
End Sub

Private Sub lblOpenPList_Click()
  Dim CD As New clsDialog
  Dim temp As String
  temp = CD.OpenDialog(Me, "M3U Playlist (*.M3U) |*.m3u|", "Open Playlist", strLOpenPath)
  If temp = "" Then Exit Sub
  OpenPList temp
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P$ = left$(file, j)
  SaveSetting App.Title, "Config", "LOpenPath", P$
  strLOpenPath = P$
End Sub

Private Sub lblOptions_Click()
  frmOptions.Show 1
End Sub

Private Sub lblRate_Click()
  frmRate.Show , Me
End Sub

Private Sub lblSavePList_Click()
  If lstPath.ListCount = 0 Then Exit Sub
  Dim CD As New clsDialog
  temp = CD.SaveDialog(Me, "M3U Playlist (*.M3U) |*.m3u|", "Save Playlist", strLSavePath)
  If temp = "" Then Exit Sub
  If LCase(Right$(temp, 4)) <> ".m3u" Then temp = temp & ".m3u"
  file = LTrim$(temp)
  Open file For Output As #1
    For i = 0 To lstPath.ListCount - 1
      Print #1, lstPath.List(i)
    Next
  Close #1
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P$ = left$(file, j)
  SaveSetting App.Title, "Config", "LSavePath", P$
  strLSavePath = P$
End Sub

Private Sub lblVol_Click()
  Shell "sndvol32.exe", vbNormalFocus
End Sub

Public Sub lstNames_DblClick()
  If lstNames.ListCount = 0 Then Exit Sub
  StopMP3
  CloseMP3
  OpenMP3 lstPath.List(lstNames.ListIndex)
  DoEvents
  PlayMP3
End Sub

Private Sub lstNames_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then lstNames_DblClick
  If KeyCode = 46 Then lblDelFile_Click
End Sub

Private Sub tmrFile_Timer()
  strFilePath = GetSetting(App.Title, "Config", "File", "")
  If strFilePath = "" Then Exit Sub
  SaveSetting App.Title, "Config", "File", ""
  If LCase(Right$(strFilePath, 4)) = ".m3u" Then
    OpenPList strFilePath
    DoEvents
    lstNames.ListIndex = 0: lstNames_DblClick
  Else
    SOpenPlayFile
  End If
  WindowState = 0
End Sub

Private Sub tmrMPlayer3_Timer()
  Dim Percent As Long
  Dim min As Integer
  Dim sec As Integer
  currpos = GetCurrentMultimediaPos(AliasName)
  CurrentTime = Val(currpos) / Val(FramesPerSecond)
  If SlideFlag = False Then
    imgSlider.left = 26 + Int((currpos / TotalFrames) * 235)
  End If
  min = CurrentTime \ 60
  sec = CurrentTime - (min * 60)
  If sec = "-1" Then sec = "0"
  lblTime = Format$(min, "00") & ":" & Format$(sec, "00")
  If AreMultimediaAtEnd(AliasName, 0) = True Then
    PlayNext
  End If
End Sub

Public Sub PlayNext()
  StopMP3
  CloseMP3
  If lstNames.ListIndex = lstNames.ListCount - 1 Then Exit Sub
  OpenMP3 lstPath.List(lstNames.ListIndex + 1)
  PlayMP3
  lstNames.ListIndex = lstNames.ListIndex + 1
End Sub

Public Sub ListMove(D)
  N = lstNames.ListIndex
  If (N + D) > 0 And (N + D) < lstNames.ListCount Then
    T1$ = lstNames.List(N): T2$ = lstPath.List(N)
    lstNames.List(N) = lstNames.List(N + D)
    lstPath.List(N) = lstPath.List(N + D)
    lstNames.List(N + D) = T1$
    lstPath.List(N + D) = T2$
    lstNames.ListIndex = N + D
  End If
End Sub

Public Sub OpenPList(file As String)
  Open file For Input As 1
    lstNames.Clear: lstPath.Clear
    GoSub LoadM3U
  Close 1
  Exit Sub
LoadM3U:
    While Not EOF(1)
        Line Input #1, AA$: a$ = LTrim$(AA$)
        GoSub AddIt
    Wend
    Return

AddIt:
    GoSub SplitPF: N = N + 1
    lstNames.AddItem b$
    lstPath.AddItem P2$ + a$
    Return

SplitPF:
    For j = Len(a$) To 1 Step -1
        If Mid$(a$, j, 1) = "\" Then Exit For
    Next
    P2$ = P1$: P$ = left$(a$, j): X$ = Mid$(a$, j + 1)
    If left$(a$, 1) = "\" Then P2$ = left$(P1$, 2)
    b$ = X$: i = InStr(X$, ".")
    If i > 0 Then b$ = left$(X$, i - 1)
    Return
End Sub

Sub SOpenPlayFile()
  lstPath.Clear
  lstNames.Clear
  file = LTrim$(strFilePath)
  lstPath.AddItem file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: P$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
  lstNames.AddItem b$
  lstNames.ListIndex = 0
  lstNames_DblClick
End Sub
