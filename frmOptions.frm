VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFFF&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Other:"
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
      Begin VB.CheckBox chkSRS 
         BackColor       =   &H00000000&
         Caption         =   "Show Rate at Startup"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkSPLS 
         BackColor       =   &H00000000&
         Caption         =   "Show Playlist at Startup"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "File Type:"
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "Select All"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelNone 
         Caption         =   "Select None"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox lstFTypes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00C0FFFF&
         Height          =   450
         ItemData        =   "frmOptions.frx":000C
         Left            =   240
         List            =   "frmOptions.frx":0016
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   320
         Width           =   735
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register File Type"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If chkSPLS.Value = 1 Then
    SaveSetting App.Title, "Config", "SPLS", "True"
  Else
    SaveSetting App.Title, "Config", "SPLS", "False"
  End If
  If chkSRS.Value = 1 Then
    SaveSetting App.Title, "Config", "SRS", "True"
  Else
    SaveSetting App.Title, "Config", "SRS", "False"
  End If
  Unload Me
End Sub

Private Sub cmdRegister_Click()
  Dim MyFileType As filetype
  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  Path = Path & App.EXEName & ".exe"
  If lstFTypes.Selected(0) = True Then
    MyFileType.ProperName = "MPlayer3"
    MyFileType.FullName = "MPlayer3 MP3 File"
    MyFileType.ContentType = "audio/mp3"
    MyFileType.extension = ".mp3"
    MyFileType.Commands.Captions.Add "Open"
    MyFileType.Commands.Commands.Add Chr$(34) & Path & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)
    MyFileType.IconPath = Path
    MyFileType.IconIndex = 0
    CreateExtension MyFileType
  End If
  If lstFTypes.Selected(1) = True Then
    MyFileType.ProperName = "MPlayer3"
    MyFileType.FullName = "MPlayer3 Playlist"
    MyFileType.ContentType = "audio/m3u"
    MyFileType.extension = ".m3u"
    MyFileType.Commands.Captions.Add "Open"
    MyFileType.Commands.Commands.Add Chr$(34) & Path & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)
    MyFileType.IconPath = Path
    MyFileType.IconIndex = 0
    CreateExtension MyFileType
  End If
End Sub

Private Sub cmdSelAll_Click()
  For i = 0 To 1
    lstFTypes.Selected(i) = True
  Next
End Sub

Private Sub cmdSelNone_Click()
  For i = 0 To 1
    lstFTypes.Selected(i) = False
  Next
End Sub

Private Sub Form_Load()
  If GetSetting(App.Title, "Config", "SPLS", "True") = "True" Then
    chkSPLS.Value = 1
  End If
  If GetSetting(App.Title, "Config", "SRS", "False") = "True" Then
    chkSRS.Value = 1
  End If
End Sub
