VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTranspose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transposer"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmTranspose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   2760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   960
   End
   Begin VB.Data dtaChords 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid grdScales 
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   3
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.Data dtaScale 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5760
      Top             =   960
   End
   Begin VB.CommandButton cmdDestPlay 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8400
      Picture         =   "frmTranspose.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2460
      Width           =   375
   End
   Begin VB.CommandButton cmdOrigPlay 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8400
      Picture         =   "frmTranspose.frx":02D4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2210
      Width           =   375
   End
   Begin VB.ComboBox cboScale 
      Height          =   315
      ItemData        =   "frmTranspose.frx":045E
      Left            =   1560
      List            =   "frmTranspose.frx":0460
      TabIndex        =   3
      Text            =   "Ionian (Major)"
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      ItemData        =   "frmTranspose.frx":0462
      Left            =   6240
      List            =   "frmTranspose.frx":048A
      TabIndex        =   2
      Text            =   "C"
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      ItemData        =   "frmTranspose.frx":04C6
      Left            =   6240
      List            =   "frmTranspose.frx":04EE
      TabIndex        =   1
      Text            =   "C"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdTranspose 
      Caption         =   "&Transpose"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdChords 
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   3
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8760
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Scale:"
      Height          =   255
      Left            =   -480
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Transpose From:"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Transpose To:"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmTranspose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim N(0 To 11) As Integer
Dim Nto(0 To 11) As Integer
Dim ChordInt(0 To 11) As Integer
Dim ChordIntNum As Integer

Private Sub cmdDestPlay_Click()
   Dim i As Integer
   Dim j As Integer
   Dim Counter As Integer
   Dim First As Integer
   Dim PNote As Integer
   Dim App As Integer
   Dim Note As String
   Dim tempnote As String
   App = 1
   
   tempnote = cboTo.Text
   If Len(tempnote) > 2 Then
      tempnote = Left(tempnote, 2)
   End If
   tempnote = UCase(tempnote)
   First = GetNoteNumber(tempnote) - 1
   Counter = First
   For i = 0 To NumNotes - 1
      Note = ToNote(i)
      If Len(Note) > 2 Then
         Note = Left(Note, 2)
      End If
      'play code
      PNote = GetNoteNumber(Note) - 1
      all_sounds_off
      If Counter < 12 Then
         Call note_on(4, 60 + PNote, 127)
      Else
      'playcode
         Call note_on(4, 72 + PNote, 127)
      End If
      If (Counter = 4) Or (Counter = 11) Then
         Counter = Counter + 1
      Else
         Counter = Counter + 2
      End If
      Timer1.Enabled = True
      While Timer1.Enabled
         DoEvents
      Wend
   Next
   all_sounds_off
   Call note_on(4, 72 + First, 127)
   Timer1.Enabled = True
   While Timer1.Enabled
      DoEvents
   Wend
   all_sounds_off
End Sub

Private Sub cmdOrigPlay_Click()
   Dim i As Integer
   Dim j As Integer
   Dim PNote As Integer
   Dim Counter As Integer
   Dim First As Integer
   Dim App As Integer
   Dim Note As String
   Dim tempnote As String
   App = 1
   
   tempnote = cboFrom.Text
   If Len(tempnote) > 2 Then
      tempnote = Left(tempnote, 2)
   End If
   tempnote = UCase(tempnote)
   First = GetNoteNumber(tempnote) - 1
   Counter = First
   For i = 0 To NumNotes - 1
      Note = OrigNote(i)
      If Len(Note) > 2 Then
         Note = Left(Note, 2)
      End If
       'play code
      PNote = GetNoteNumber(Note) - 1
      all_sounds_off
      If Counter < 12 Then
         Call note_on(4, 60 + PNote, 127)
      Else
      'playcode
         Call note_on(4, 72 + PNote, 127)
      End If
      If (Counter = 4) Or (Counter = 11) Then
         Counter = Counter + 1
      Else
         Counter = Counter + 2
      End If
      Timer1.Enabled = True
      While Timer1.Enabled
         DoEvents
      Wend
   Next
   all_sounds_off
   Call note_on(4, 72 + First, 127)
   Timer1.Enabled = True
   While Timer1.Enabled
      DoEvents
   Wend
   all_sounds_off
End Sub

Private Sub cmdPrint_Click()
   PrintData
End Sub

Public Sub PrintData()
     'print
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim NumCopies As Integer
   Dim BeginPage As Integer
   Dim EndPage As Integer
   Dim temp As String
   Dim ToScale As String
   Dim FromScale As String
   Dim ToChords As String
   Dim FromChords As String
   
   ToChords = ""
   FromChords = ""
   ToScale = ""
   FromScale = ""
   
   For j = 1 To grdScales.Cols - 1
      grdScales.Col = j
      grdScales.Row = 1
      temp = ""
      For k = 0 To (7 - Len(grdScales.Text))
         temp = temp + " "
      Next
      FromScale = FromScale + grdScales.Text + temp
      grdScales.Row = 2
      temp = ""
      For k = 0 To (7 - Len(grdScales.Text))
         temp = temp + " "
      Next
      ToScale = ToScale + grdScales.Text + temp
   Next
   If NumNotes = 7 Then
      For j = 1 To grdChords.Cols - 1
         grdChords.Col = j
         grdChords.Row = 1
         temp = ""
         For k = 0 To (7 - Len(grdChords.Text))
            temp = temp + " "
         Next
         FromChords = FromChords + grdChords.Text + temp
         grdChords.Row = 2
         temp = ""
         For k = 0 To (7 - Len(grdChords.Text))
            temp = temp + " "
         Next
         ToChords = ToChords + grdChords.Text + temp
      Next
   End If
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
      Printer.Print "Transposer"
      Printer.Print
      Printer.FontSize = 12
      Printer.Print "Scale:"
      Printer.FontBold = False
      Printer.Print Trim(cboScale.Text)
      Printer.Print
      Printer.FontBold = True
      Printer.Print "Transpose from " + Trim(cboFrom.Text) + " to " + Trim(cboTo.Text)
      Printer.Print
      Printer.Font = "Courier"
      Printer.Print "Scale"
      Printer.FontBold = False
      Printer.Print
      Printer.Print FromScale
      Printer.Print ToScale
      Printer.Print
      Printer.Print
      Printer.FontBold = True
      Printer.Print "Chords"
      Printer.Print
      Printer.FontBold = False
      Printer.Print FromChords
      Printer.Print ToChords
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.FontSize = 8
      Printer.Font = "Arial"
      Printer.Print "Guitar Master Pro - Copyright 2000 Opus Software"
      Printer.EndDoc
   Next i
   Exit Sub
PrintErr:
   MsgBox "Document was not printed", vbExclamation + vbOKOnly, "Print"
   Exit Sub
CancelErr:
End Sub

Private Sub cmdTranspose_Click()
   frmTranspose.MousePointer = 11
   
   cmdOrigPlay.Enabled = False
   cmdDestPlay.Enabled = False
   cmdPrint.Enabled = False
   frmMain.mnuPrint.Enabled = False
   
   If cboFrom.Text = "" Then
      MsgBox "No scale to transpose from", vbExclamation + vbOKOnly, "Error"
      Exit Sub
   End If
   If cboTo.Text = "" Then
      MsgBox "No scale to transpose to", vbExclamation + vbOKOnly, "Error"
      Exit Sub
   End If
   If grdChords.Cols > 0 Then
      grdChords.FixedCols = 0
   End If
   grdChords.FixedRows = 0
   grdChords.Cols = 0
   InputData
   DoConvert
   DoTranspose
   FindChords
   DoAnotherConvert
   ShowResult
   frmTranspose.MousePointer = 0

End Sub
Private Sub InputData()
   TScale = Trim(cboScale.Text)
   OrigKey = Trim(cboFrom.Text)
   ToKey = Trim(cboTo.Text)
End Sub

Private Sub DoConvert()
   Dim i As Integer
   
   'convert orig key to num
   OrigKeyNum = GetNoteNumber(OrigKey)
   'convert target key to num
   ToKeyNum = GetNoteNumber(ToKey)
     'set values for scales
   GetScales
End Sub

Private Sub DoTranspose()
   'find difference in keys
   Dim i As Integer
   Dim temp As Integer
   For i = 0 To NumNotes - 1
      OrigNoteNum(i) = (OrigNoteNum(i) + OrigKeyNum - 1)
   Next
   Diff = ToKeyNum - OrigKeyNum
   For i = 0 To NumNotes - 1
      ToNoteNum(i) = OrigNoteNum(i) + Diff
      If ToNoteNum(i) < 0 Then
         ToNoteNum(i) = 12 + ToNoteNum(i)
      End If
   Next
End Sub

Private Sub DoAnotherConvert()
   Dim i As Integer
   For i = 0 To NumNotes - 1
      OrigNote(i) = GetNoteText(OrigNoteNum(i) + 1)
      ToNote(i) = GetNoteText(ToNoteNum(i) + 1)
   Next
End Sub

Private Sub ShowResult()
   Dim i As Integer
   Dim wid As Integer
   
   grdScales.Cols = NumNotes + 1
   grdScales.ColWidth(0) = 2500
   grdScales.FixedCols = 1
   grdScales.FixedRows = 1
   wid = (grdScales.Width - 2580) / NumNotes
   For i = 1 To NumNotes
     grdScales.ColWidth(i) = wid
   Next
   
   grdScales.TextMatrix(1, 0) = TScale + " in " + OrigKey
   grdScales.TextMatrix(2, 0) = TScale + " in " + ToKey
   For i = 0 To NumNotes - 1
      grdScales.TextMatrix(1, i + 1) = OrigNote(i)
      grdScales.TextMatrix(2, i + 1) = ToNote(i)
   Next
   
   If NumNotes = 7 Then
      grdChords.Cols = NumNotes + 1
      grdChords.ColWidth(0) = 2500
      grdChords.FixedCols = 1
      grdChords.FixedRows = 1
      wid = (grdChords.Width - 2580) / NumNotes
      For i = 1 To NumNotes
         grdChords.ColWidth(i) = wid
      Next
   
      grdChords.TextMatrix(1, 0) = TScale + " in " + OrigKey
      grdChords.TextMatrix(2, 0) = TScale + " in " + ToKey
      For i = 0 To 6
         If Len(OrigNote(i)) > 2 Then
            grdChords.TextMatrix(1, i + 1) = Left(OrigNote(i), 2) + Trim(Chord(i))
         Else
            grdChords.TextMatrix(1, i + 1) = OrigNote(i) + Trim(Chord(i))
         End If
         If Len(ToNote(i)) > 2 Then
            grdChords.TextMatrix(2, i + 1) = Left(ToNote(i), 2) + Trim(Chord(i))
         Else
            grdChords.TextMatrix(2, i + 1) = ToNote(i) + Trim(Chord(i))
         End If
      Next
   End If
   If NumDevices > 0 Then
      cmdOrigPlay.Enabled = True
      cmdDestPlay.Enabled = True
   End If
   cmdPrint.Enabled = True
   frmMain.mnuPrint.Enabled = True
   
End Sub

Private Sub Form_Activate()
   DisableMenu
   frmMain.mnuPrint.Visible = True
   cmdPrint.Enabled = False
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   dtaScale.DatabaseName = "Chords.mdb"
   dtaScale.Refresh
   dtaScale.RecordSource = "SELECT * FROM Scales ORDER BY Name"
   dtaScale.Refresh
   dtaScale.Recordset.MoveFirst
   dtaScale.Recordset.MoveLast
   dtaScale.Recordset.MoveFirst
   cboScale.Clear
   For i = 1 To dtaScale.Recordset.RecordCount
      cboScale.AddItem dtaScale.Recordset.Fields(0).value
      dtaScale.Recordset.MoveNext
   Next
   cboScale.Text = "Ionian (Major)"
End Sub

Private Sub GetScales()
   Dim i As Integer
   
   dtaScale.RecordSource = "SELECT * FROM Scales WHERE Name = '" + cboScale.Text + "';"
   dtaScale.Refresh
   NumNotes = dtaScale.Recordset.Fields(1).value
   For i = 0 To NumNotes - 1
      OrigNoteNum(i) = dtaScale.Recordset.Fields(i + 2).value
   Next
   HaveChords = False
End Sub

Private Sub FindChords()
   Dim i As Integer
   Dim wid As Integer
   For i = 0 To NumNotes - 1
      N(i) = OrigNoteNum(i)
      Nto(i) = ToNoteNum(i)
      OrigNoteNum(i) = OrigNoteNum(i) Mod 12
      ToNoteNum(i) = ToNoteNum(i) Mod 12
   Next
   If NumNotes = 7 Then
      For i = 0 To 2
         If (N(i + 2) - N(i) = 4) And (N(i + 4) - N(i) = 7) Then
            Chord(i) = "maj"
         ElseIf (N(i + 2) - N(i) = 4) And (N(i + 4) - N(i) = 6) Then
            Chord(i) = "majb5"
        ElseIf (N(i + 2) - N(i) = 3) And (N(i + 4) - N(i) = 7) Then
            Chord(i) = "m"
         ElseIf (N(i + 2) - N(i) = 2) And (N(i + 4) - N(i) = 7) Then
            Chord(i) = "m"
         ElseIf (N(i + 2) - N(i) = 5) And (N(i + 4) - N(i) = 7) Then
            Chord(i) = "sus4"
         ElseIf (N(i + 2) - N(i) = 3) And (N(i + 4) - N(i) = 6) Then
            Chord(i) = "dim"
         ElseIf (N(i + 2) - N(i) = 4) And (N(i + 4) - N(i) = 8) Then
            Chord(i) = "aug"
         End If
      Next
      For i = 3 To 4
         If (N(i + 2) - N(i) = 4) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "maj"
         ElseIf (N(i + 2) - N(i) = 4) And (N(i) - N(i - 3) = 4) Then
            Chord(i) = "maj"
         ElseIf (N(i + 2) - N(i) = 3) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "m"
         ElseIf (N(i + 2) - N(i) = 2) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "sus2"
         ElseIf (N(i + 2) - N(i) = 3) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "sus4"
          ElseIf (N(i + 2) - N(i) = 3) And (N(i) - N(i - 3) = 6) Then
            Chord(i) = "dim"
         ElseIf (N(i + 2) - N(i) = 4) And (N(i) - N(i - 3) = 4) Then
            Chord(i) = "aug"
         Else
            Chord(i) = "xxxx"
        End If
      Next
      For i = 5 To 6
         If (N(i) - N(i - 5) = 8) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "maj"
         ElseIf (N(i) - N(i - 5) = 8) And (N(i) - N(i - 3) = 4) Then
            Chord(i) = "maj"
         ElseIf (N(i) - N(i - 5) = 9) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "m"
         ElseIf (N(i) - N(i - 5) = 10) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "sus2"
         ElseIf (N(i) - N(i - 5) = 7) And (N(i) - N(i - 3) = 5) Then
            Chord(i) = "sus4"
          ElseIf (N(i) - N(i - 5) = 9) And (N(i) - N(i - 3) = 6) Then
            Chord(i) = "dim"
         ElseIf (N(i) - N(i - 5) = 8) And (N(i) - N(i - 3) = 4) Then
            Chord(i) = "aug"
         Else
            Chord(i) = "xxxx"
        End If
      Next
         
   End If
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
   all_sounds_off
End Sub

Private Sub grdChords_Click()
   Dim i As Integer
   If grdChords.Col > 0 And grdChords.Row > 0 Then
      ShortSuffix = Chord(grdChords.Col - 1)
      If ShortSuffix = "xxxx" Then
         Exit Sub
      End If
      LongSuffix = LongChord(ShortSuffix)
      dtaChords.DatabaseName = "Chords.mdb"
      dtaChords.Refresh
      dtaChords.RecordSource = "SELECT * FROM ChordType WHERE NAME = '" + Trim(LongSuffix) + "'"
      dtaChords.Refresh
      ChordIntNum = dtaChords.Recordset.Fields(1).value
      For i = 0 To ChordIntNum - 1
         ChordInt(i) = dtaChords.Recordset.Fields(i + 2).value
      Next
      For i = 0 To ChordIntNum
         If grdChords.Row = 1 Then
            If OrigKeyNum - ToKeyNum > 5 Then
               Call note_on(3, 36 + N(grdChords.Col - 1) + ChordInt(i), 127)
            Else
               Call note_on(3, 48 + N(grdChords.Col - 1) + ChordInt(i), 127)
            End If
         ElseIf grdChords.Row = 2 Then
            If ToKeyNum - OrigKeyNum > 5 Then
               Call note_on(3, 36 + Nto(grdChords.Col - 1) + ChordInt(i), 127)
            Else
               Call note_on(3, 48 + Nto(grdChords.Col - 1) + ChordInt(i), 127)
            End If
         End If
      Next
      Timer2.Enabled = True
      While Timer2.Enabled
         DoEvents
      Wend
      'play chord
   End If
End Sub

Private Sub grdScales_Click()
   If grdScales.Col > 0 And grdScales.Row > 0 Then
      If grdScales.Row = 1 Then
          Call note_on(3, 60 + N(grdScales.Col - 1), 127)
      ElseIf grdScales.Row = 2 Then
          Call note_on(3, 60 + Nto(grdScales.Col - 1), 127)
      End If
      Timer2.Enabled = True
      While Timer2.Enabled
         DoEvents
      Wend
   End If
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
End Sub
