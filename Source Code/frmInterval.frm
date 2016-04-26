VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInterval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interval Finder"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmInterval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   7470
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   5040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   6720
      TabIndex        =   25
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6120
      TabIndex        =   24
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   23
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   22
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   19
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "7"
      Height          =   255
      Index           =   10
      Left            =   6720
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "b7"
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "6/13"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   11
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "#5/b13"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "5"
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "b5/#11"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "4/11"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "b3/#9"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "2/9"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "b2/b9"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Root"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Counter As Integer

Private Sub cmdNext_Click()
   Dim i As Integer
   
   Counter = Counter + 1
   If Counter > 11 Then
      Counter = 0
   End If
   For i = 0 To 11
      lblNote(i).Caption = Trim(NoteName((i + Counter) Mod 12))
   Next
End Sub

Private Sub cmdPrev_Click()
   Dim i As Integer
   
   Counter = Counter - 1
   If Counter < 0 Then
      Counter = 11
   End If
   For i = 0 To 11
      lblNote(i).Caption = Trim(NoteName((i + Counter) Mod 12))
   Next
End Sub

Private Sub cmdPrint_Click()
   PrintData
End Sub

Private Sub Form_Activate()
   DisableMenu
   frmMain.mnuPrint.Visible = True
   frmMain.mnuPrint.Enabled = True
End Sub

Public Sub PrintData()
   'print
   Dim i As Integer
   Dim j As Integer
   Dim NumCopies As Integer
   Dim BeginPage As Integer
   Dim EndPage As Integer
   Dim temp As String
   diaCommon.CancelError = True
   On Error GoTo CancelErr
   diaCommon.ShowPrinter
   On Error GoTo PrintErr
   BeginPage = diaCommon.FromPage
   EndPage = diaCommon.ToPage
   NumCopies = diaCommon.Copies
   For i = 1 To NumCopies
      Printer.FontSize = 16
      Printer.FontBold = True
      Printer.Font = "Arial"
      Printer.Print "Intervals"
      Printer.Print
      Printer.FontSize = 12
      Printer.FontBold = False
      Printer.Print "Root"; Tab(8); "b2/b9"; Tab(16); "2/9"; Tab(24); "b3/#9"; Tab(32); "3"; Tab(40); "4/11"; Tab(48); "b5/#11"; Tab(56); "5"; Tab(64); "#5/b13"; Tab(72); "6"; Tab(80); "b7"; Tab(88); "7"
      Printer.Print lblNote(0).Caption; Tab(8); lblNote(1).Caption; Tab(16); lblNote(2).Caption; Tab(24); lblNote(3).Caption; Tab(32); lblNote(4).Caption; Tab(40); lblNote(5).Caption; Tab(48); lblNote(6).Caption; Tab(56); lblNote(7).Caption; Tab(64); lblNote(8).Caption; Tab(72); lblNote(9).Caption; Tab(80); lblNote(10).Caption; Tab(88); lblNote(11).Caption
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.FontSize = 8
      Printer.Print "Guitar Master Pro - Copyright 2000 Opus Software"
      Printer.EndDoc
   Next i
   Exit Sub
PrintErr:
   MsgBox "Chord was not printed", vbExclamation + vbOKOnly, "Print"
   Exit Sub
CancelErr:
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FillNames
   Counter = 0
   For i = 0 To 11
      lblNote(i).Caption = Trim(NoteName(i))
   Next
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub
