VERSION 5.00
Begin VB.Form frmStartUp 
   Caption         =   "StartUp"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ListBox lstInstruments 
         Height          =   2010
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ListBox lstDevice 
         Height          =   840
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   4695
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
         TabIndex        =   4
         Top             =   120
         Width           =   2175
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim Filenum
   Filenum = FreeFile

'MIDI tab
  Call midi_listoutdevs(lstDevice)
  FillInstrument
   lstDevice.Text = lstDevice.List(DefMapper)
   lstInstruments.Text = lstInstruments.List(DefInstrument)

End Sub

Public Sub cmdClose_Click()
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

Private Sub lstDevice_Click()
   Dim Filenum
   Dim X  As Integer

    midi_out_close
    X = midi_out_open(lstDevice.ItemData(lstDevice.ListIndex))
    DefMapper = lstDevice.ListIndex
    DefDevice = lstDevice.ItemData(lstDevice.ListIndex)
    frmMain.sbrMain.Panels(2).Text = "Device: " + lstDevice.List(DefMapper)
    Filenum = FreeFile
 Open "Config.dat" For Binary As Filenum
   Put #Filenum, 3, DefMapper
   Put #Filenum, 5, DefDevice
   Close Filenum


End Sub

Private Sub lstInstruments_Click()
Dim Filenum
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
   Filenum = FreeFile
   frmMain.sbrMain.Panels(1).Text = "Instrument: " + lstInstruments.List(DefInstrument)
   Open "Config.dat" For Binary As Filenum '
   Put #Filenum, 7, DefInstrument
   Close Filenum

End Sub


