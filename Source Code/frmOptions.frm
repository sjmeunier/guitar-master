VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   6600
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGeneral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MIDI"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraGeneral 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4575
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6135
         Begin VB.Frame Frame5 
            Caption         =   "Chord Calculation"
            Height          =   1095
            Left            =   0
            TabIndex        =   23
            Top             =   2280
            Width           =   6135
            Begin VB.OptionButton optChordCalc 
               Caption         =   "Do not include any dead strings"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   3735
            End
            Begin VB.OptionButton optChordCalc 
               Caption         =   "Do not include dead strings between played strings"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   25
               Top             =   480
               Width           =   5055
            End
            Begin VB.OptionButton optChordCalc 
               Caption         =   "All chord combinations"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Fret Display"
            Height          =   1095
            Left            =   0
            TabIndex        =   19
            Top             =   3480
            Width           =   6135
            Begin VB.OptionButton optFretNote 
               Caption         =   "None"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   2655
            End
            Begin VB.OptionButton optFretNote 
               Caption         =   "Intervals"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   2895
            End
            Begin VB.OptionButton optFretNote 
               Caption         =   "Notes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Finger Streching for Chords"
            Height          =   855
            Left            =   0
            TabIndex        =   12
            Top             =   1320
            Width           =   6135
            Begin MSComctlLib.Slider sldStretch 
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Min             =   3
               Max             =   7
               SelStart        =   3
               Value           =   3
               TextPosition    =   1
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "7"
               Height          =   255
               Index           =   4
               Left            =   2490
               TabIndex        =   18
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "6"
               Height          =   255
               Index           =   3
               Left            =   1920
               TabIndex        =   17
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "5"
               Height          =   255
               Index           =   2
               Left            =   1330
               TabIndex        =   16
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "4"
               Height          =   255
               Index           =   1
               Left            =   770
               TabIndex        =   15
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "3"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   14
               Top             =   480
               Width           =   135
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Fretboard"
            Height          =   1095
            Left            =   0
            TabIndex        =   8
            Top             =   120
            Width           =   6135
            Begin VB.OptionButton optFretboard 
               Caption         =   "White on Black Fretboard"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   2415
            End
            Begin VB.OptionButton optFretboard 
               Caption         =   "Black on White Fretboard"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   480
               Width           =   3615
            End
            Begin VB.OptionButton optFretboard 
               Caption         =   "Colour Fretboard"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   2295
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   6015
         Begin VB.ListBox lstDevice 
            Height          =   840
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   4695
         End
         Begin VB.ListBox lstInstruments 
            Height          =   2010
            Left            =   1320
            TabIndex        =   3
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label Label2 
            Caption         =   "Instrument:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   6
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "MIDI Device:"
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
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   2175
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   5400
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Public Sub FillInstrument()
   Dim s As String

   Open "genmidi.dat" For Input As #1
   Do While Not EOF(1)
      Line Input #1, s
      lstInstruments.AddItem s
   Loop
   Close #1
End Sub

Public Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   DisableMenu
End Sub

Private Sub Form_Load()
   Dim FileNum
   FileNum = FreeFile

   'MIDI tab
   Call midi_listoutdevs(lstDevice)
   FillInstrument
   lstDevice.Text = lstDevice.List(DefMapper)
   lstInstruments.Text = lstInstruments.List(DefInstrument)

   optFretboard(FretboardCol).value = True
   optFretNote(FretNote).value = True
   optChordCalc(ChordCalc).value = True
   sldStretch.value = FingerStretch
   SSTab1.Tab = CurTab
End Sub

Private Sub lstDevice_Click()
   Dim FileNum
   Dim X  As Integer

   midi_out_close
   X = midi_out_open(lstDevice.ItemData(lstDevice.ListIndex))
   DefMapper = lstDevice.ListIndex
   DefDevice = lstDevice.ItemData(lstDevice.ListIndex)
   frmMain.sbrMain.Panels(2).Text = "Device: " + lstDevice.List(DefMapper)
   FileNum = FreeFile
   Open "Config.dat" For Binary As FileNum
   Put #FileNum, 3, DefMapper
   Put #FileNum, 5, DefDevice
   Close FileNum
End Sub

Private Sub lstInstruments_Click()
Dim FileNum
   Call program_change(0, 0, lstInstruments.ListIndex)
   Call program_change(1, 0, lstInstruments.ListIndex)
   Call program_change(2, 0, lstInstruments.ListIndex)
   Call program_change(3, 0, lstInstruments.ListIndex)
   Call program_change(4, 0, lstInstruments.ListIndex)
   Call program_change(5, 0, lstInstruments.ListIndex)
   Call program_change(6, 0, lstInstruments.ListIndex)
   Call program_change(7, 0, lstInstruments.ListIndex)
   Call program_change(8, 0, lstInstruments.ListIndex)
   Call program_change(9, 0, lstInstruments.ListIndex)
   Call program_change(11, 0, lstInstruments.ListIndex)
   DefInstrument = lstInstruments.ListIndex
   FileNum = FreeFile
   frmMain.sbrMain.Panels(1).Text = "Instrument: " + lstInstruments.List(DefInstrument)
   Open "Config.dat" For Binary As FileNum '
   Put #FileNum, 7, DefInstrument
   Close FileNum
End Sub

Private Sub optChordCalc_Click(Index As Integer)
   ChordCalc = Index
   WriteToConfig 27, ChordCalc
End Sub

Private Sub optFretBoard_Click(Index As Integer)
   FretboardCol = Index
   WriteToConfig 19, FretboardCol
End Sub

Private Sub optFretNote_Click(Index As Integer)
   FretNote = Index
   WriteToConfig 25, FretNote
End Sub

Private Sub sldStretch_Click()
   FingerStretch = sldStretch.value
   WriteToConfig 21, FingerStretch
End Sub

Private Sub WriteToConfig(Pos As Integer, Data As Integer)
   Dim FileNum
   On Error GoTo err:
   FileNum = FreeFile
   Open "config.dat" For Binary As FileNum
   Put #FileNum, Pos, Data
   Close FileNum
   Exit Sub
err:
   MsgBox "Unabled to write to the configuration file", vbOKOnly + vbExclamation, "Error"

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    CurTab = SSTab1.Tab
End Sub


