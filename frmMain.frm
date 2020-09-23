VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Metronome 1.0"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCutTime 
      Caption         =   "Cut Time?"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox bpm 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "120"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame fraBeats 
      Caption         =   "Beats per Minute (BPM):"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer beat 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   1320
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private beats As Long, maxbeats As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub beat_Timer()
    If beats = maxbeats + 1 Then beats = 1
    If beats <= 1 Then
        sndPlaySound App.Path & "\beat.wav", 1
    Else
        sndPlaySound App.Path & "\crash.wav", 1
    End If
    beats = beats + 1
End Sub

Private Sub cmdStart_Click()
    If CLng(bpm) > 500 Or CLng(bpm) < 31 Then
        MsgBox "Please enter a BPM 31-500"
        Exit Sub
    End If
    beats = 1
    beat.Interval = 60000 / CLng(bpm)
    If chkCutTime.Value = vbChecked Then
        maxbeats = 2
        beat.Interval = beat.Interval / 2
    Else
        maxbeats = 4
    End If
    beat.Enabled = True
    chkCutTime.Enabled = False
    cmdStart.Enabled = False
    bpm.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    beat.Enabled = False
    chkCutTime.Enabled = True
    cmdStart.Enabled = True
    bpm.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Thank you for using the VB Metronome!" & vbCrLf & "Written by Steve McMaster"
End Sub
