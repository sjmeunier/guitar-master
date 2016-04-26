VERSION 5.00
Begin VB.Form frmScale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scale Finder"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmScale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   5595
   Begin VB.TextBox txtMode 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   1695
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtScale 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Data dtaScales 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   11
      ItemData        =   "frmScale.frx":0442
      Left            =   4440
      List            =   "frmScale.frx":046A
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   10
      ItemData        =   "frmScale.frx":04A6
      Left            =   3600
      List            =   "frmScale.frx":04CE
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   9
      ItemData        =   "frmScale.frx":050A
      Left            =   2760
      List            =   "frmScale.frx":0532
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   8
      ItemData        =   "frmScale.frx":056E
      Left            =   1920
      List            =   "frmScale.frx":0596
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   7
      ItemData        =   "frmScale.frx":05D2
      Left            =   1080
      List            =   "frmScale.frx":05FA
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   6
      ItemData        =   "frmScale.frx":0636
      Left            =   240
      List            =   "frmScale.frx":065E
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   5
      ItemData        =   "frmScale.frx":069A
      Left            =   4440
      List            =   "frmScale.frx":06C2
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   4
      ItemData        =   "frmScale.frx":06FE
      Left            =   3600
      List            =   "frmScale.frx":0726
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   3
      ItemData        =   "frmScale.frx":0762
      Left            =   2760
      List            =   "frmScale.frx":078A
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   2
      ItemData        =   "frmScale.frx":07C6
      Left            =   1920
      List            =   "frmScale.frx":07EE
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   1
      ItemData        =   "frmScale.frx":082A
      Left            =   1080
      List            =   "frmScale.frx":0852
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox txtNote 
      Height          =   315
      Index           =   0
      ItemData        =   "frmScale.frx":088E
      Left            =   240
      List            =   "frmScale.frx":08B6
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Mode:"
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Scale:"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChordNote(0 To 11) As Integer
Dim ChordNs(0 To 11) As Integer
Dim RawNote(0 To 11) As Integer
Dim NoteCount As Integer
Dim Root As String
Dim RootNum As String
Dim NoteTxt As String

Private Sub cmdClear_Click()
   Dim i As Integer
   
   For i = 0 To 11
      txtNote(i).Text = ""
   Next
End Sub

Private Sub cmdFind_Click()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   Dim temp As String
   Dim tempnote As Integer
   Dim TotModeCount As Integer
   frmScale.MousePointer = 11
   txtScale.Text = ""
   txtMode.Text = ""
   NoteCount = 0
   For i = 0 To 11
      If txtNote(i).Text <> "" Then
         NoteTxt = Trim(UCase(txtNote(i).Text))
         ChordNs(i) = GetNoteNumber(NoteTxt)
         RawNote(i) = ChordNs(i)
         NoteCount = NoteCount + 1
      End If
   Next
   For j = 1 To NoteCount - 1
      For i = 1 To NoteCount - 1
         If ChordNs(i) < ChordNs(i - 1) Then
            ChordNs(i) = ChordNs(i) + 12
         End If
      Next
   Next
   For i = 1 To NoteCount - 1
      ChordNote(i) = ChordNs(i) - ChordNs(0)
   Next
   dtaScales.DatabaseName = "Chords.mdb"
   dtaScales.Refresh
   temp = "SELECT * FROM Scales WHERE [1] = " + Str(ChordNote(0))
   For i = 1 To NoteCount - 1
      temp = temp + " AND [" + Trim(Str(i + 1)) + "] = " + Str(ChordNote(i))
   Next
   temp = temp + " AND [NumNotes] = " + Str(NoteCount)
   dtaScales.RecordSource = temp
   dtaScales.Refresh
   
   On Error Resume Next
   dtaScales.Recordset.MoveFirst
   dtaScales.Recordset.MoveLast
   dtaScales.Recordset.MoveFirst
   If dtaScales.Recordset.RecordCount < 1 Then
      MsgBox "Scale was not found", vbOKOnly + vbInformation, "Message"
      frmScale.MousePointer = 0
      Exit Sub
   End If
   
   For i = 1 To dtaScales.Recordset.RecordCount
      txtScale.Text = txtScale.Text + Trim(UCase(txtNote(0).Text)) + " " + dtaScales.Recordset.Fields(0).value + Chr(13) + Chr(10)
      dtaScales.Recordset.MoveNext
   Next
   TotModeCount = NoteCount
   For j = 1 To TotModeCount - 1

      tempnote = RawNote(0)
      For i = 1 To NoteCount - 1
         RawNote(i - 1) = RawNote(i)
      Next
      RawNote(NoteCount - 1) = tempnote
      For i = 0 To 11
         ChordNs(i) = RawNote(i)
      Next
      For k = 1 To NoteCount - 1
         For i = 1 To NoteCount - 1
            If ChordNs(i) < ChordNs(i - 1) Then
               ChordNs(i) = ChordNs(i) + 12
            End If
         Next
      Next
      For i = 1 To NoteCount - 1
         ChordNote(i) = ChordNs(i) - ChordNs(0)
      Next
      temp = "SELECT * FROM Scales WHERE [1] = " + Str(ChordNote(0))
      For i = 1 To NoteCount - 1
         temp = temp + " AND [" + Trim(Str(i + 1)) + "] = " + Str(ChordNote(i))
      Next
      temp = temp + " AND [NumNotes] = " + Str(NoteCount)
      dtaScales.RecordSource = temp
      dtaScales.Refresh
   
      On Error Resume Next
      dtaScales.Recordset.MoveFirst
      dtaScales.Recordset.MoveLast
      dtaScales.Recordset.MoveFirst
      For i = 1 To dtaScales.Recordset.RecordCount
         txtMode.Text = txtMode.Text + Trim(UCase(txtNote(j).Text)) + " " + dtaScales.Recordset.Fields(0).value + Chr(13) + Chr(10)
         dtaScales.Recordset.MoveNext
      Next
   Next
   frmScale.MousePointer = 0

End Sub

Private Sub Form_Activate()
   DisableMenu
   txtScale.Text = ""
   txtMode.Text = ""
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub
