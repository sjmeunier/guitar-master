VERSION 5.00
Begin VB.Form frmTuner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tuner"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "frmTuner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboTuning 
      Height          =   315
      ItemData        =   "frmTuner.frx":014A
      Left            =   960
      List            =   "frmTuner.frx":0178
      TabIndex        =   0
      Text            =   "Standard"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   960
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label s 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   5
         X1              =   0
         X2              =   1920
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   1920
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   1920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   1920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   1920
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   1920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Image Image1 
         Height          =   2085
         Left            =   0
         Picture         =   "frmTuner.frx":020A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   960
      X2              =   960
      Y1              =   1440
      Y2              =   3480
   End
   Begin VB.Label s1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label s2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label s3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "G"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label s4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "D"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label s5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label s6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3000
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tuning:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   3000
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmTuner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tune As String
Dim StrStart(0 To 5) As Integer

Private Sub cmdChange_Click()
   Tune = cboTuning.Text
   ChangeTune Tune, StrDelta(), StrText()
   s1.Caption = StrText(0)
   s2.Caption = StrText(1)
   s3.Caption = StrText(2)
   s4.Caption = StrText(3)
   s5.Caption = StrText(4)
   s6.Caption = StrText(5)
   StrStart(0) = 40 + StrDelta(0)
   StrStart(1) = 45 + StrDelta(1)
   StrStart(2) = 50 + StrDelta(2)
   StrStart(3) = 55 + StrDelta(3)
   StrStart(4) = 59 + StrDelta(4)
   StrStart(5) = 64 + StrDelta(5)
End Sub

Private Sub Form_Activate()
   DisableMenu
End Sub

Private Sub Form_Load()
   Dim i As Integer
   cboTuning.Text = "Standard"
   cmdChange_Click
   
   If FretboardCol = 1 Then
      Frame1.BackColor = vbWhite
      Image1.Picture = LoadPicture
      For i = 0 To 5
         Line3(i).BorderColor = vbBlack
      Next
   ElseIf FretboardCol = 2 Then
      Frame1.BackColor = vbBlack
      Image1.Picture = LoadPicture
      For i = 0 To 5
         Line3(i).BorderColor = vbWhite
      Next
   End If
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
   all_sounds_off
End Sub

Private Sub s_Click(Index As Integer)
   all_sounds_off
   Call note_on(5, StrStart(Index), 127)
End Sub


