VERSION 5.00
Begin VB.Form frmMetro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metronome"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmMetro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtBeats 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "4"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtNotes 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "4"
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer tmrMetro 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3840
      Top             =   2400
   End
   Begin VB.TextBox txtBPM 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "120"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4800
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Signature:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   2160
      X2              =   2520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   2040
      X2              =   2520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBottom 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Beats per minute:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4800
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmMetro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Beats As Integer
Dim Notes As Integer
Dim BPM As Integer
Dim Counter As Integer

Private Sub cmdChange_Click()
   cmdStart_Click
End Sub

Private Sub cmdStart_Click()
   Beats = Val(txtBeats.Text)
   Notes = Val(txtNotes.Text)
   BPM = Val(txtBPM.Text)
   tmrMetro.Interval = 1000 / (BPM / 60)
   Counter = 1
   lblBottom.Caption = Trim(Str(Beats))
   tmrMetro.Enabled = True
   cmdStop.Enabled = True
   cmdStart.Enabled = False
   cmdChange.Enabled = False
End Sub

Private Sub cmdStop_Click()
   tmrMetro.Enabled = False
   cmdStop.Enabled = False
   cmdStart.Enabled = True
   cmdChange.Enabled = False
   all_sounds_off
End Sub

Private Sub Form_Activate()
   DisableMenu
   If NumDevices > 0 Then
      Call program_change(10, 0, 117)
   End If
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
   all_sounds_off
End Sub

Private Sub tmrMetro_Timer()
   lblTop.Caption = Trim(Str(Counter))
   Counter = Counter Mod Notes
   Counter = Counter + 1
   all_sounds_off
   If NumDevices = 0 Then
      Beep
   Else
      If lblTop.Caption = "1" Then
         Call note_on(10, 38, 127)
      Else
         Call note_on(10, 45, 127)
      End If
   End If
End Sub

Private Sub txtBeats_Change()
   cmdChange.Enabled = True
End Sub

Private Sub txtBPM_Change()
   cmdChange.Enabled = True
End Sub

Private Sub txtNotes_Change()
   cmdChange.Enabled = True
End Sub
