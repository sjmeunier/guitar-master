VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Guitar Master Pro"
   ClientHeight    =   4440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6435
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4185
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":141E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1872
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":211A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":227A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":282E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":298A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Charts"
            Object.ToolTipText     =   "Chord Chart"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Finder"
            Object.ToolTipText     =   "Chord Finder"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NoteFinder"
            Object.ToolTipText     =   "Note Finder"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scale"
            Object.ToolTipText     =   "Scale Finder"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transposer"
            Object.ToolTipText     =   "Transposer"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Interval"
            Object.ToolTipText     =   "Interval Finder"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Metronome"
            Object.ToolTipText     =   "Metronome"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tuner"
            Object.ToolTipText     =   "Tuner"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Freq"
            Object.ToolTipText     =   "Frequency Finder"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Gloss"
            Object.ToolTipText     =   "Glossary"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuPrintTab 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuChordChart 
         Caption         =   "&Chord Chart"
      End
      Begin VB.Menu mnuChordFind 
         Caption         =   "C&hord Finder"
      End
      Begin VB.Menu mnuFreq 
         Caption         =   "&Frequency Finder"
      End
      Begin VB.Menu mnuInterval 
         Caption         =   "&Interval Finder"
      End
      Begin VB.Menu mnuMetro 
         Caption         =   "&Metronome"
      End
      Begin VB.Menu mnuNoteFind 
         Caption         =   "&Note Finder"
      End
      Begin VB.Menu mnuScale 
         Caption         =   "&Scale Finder"
      End
      Begin VB.Menu mnuTransposer 
         Caption         =   "&Transposer"
      End
      Begin VB.Menu mnuTuner 
         Caption         =   "&Tuner"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile &Horizontal"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuTileVer 
         Caption         =   "Tile &Vertical"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Icons"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGlossary 
         Caption         =   "&Glossary"
      End
      Begin VB.Menu mnuHelp2 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Sub MDIForm_Load()
   Dim FileNum
   Dim i As Integer
   On Error GoTo FileErr
   FileNum = FreeFile
   NumDevices = GetNumDevice
   
   If NumDevices = 0 Then
      MsgBox "No MIDI output device found on system. Unable to play MIDI sound", vbOKOnly + vbExclamation, "Warning"
      mnuOptions.Enabled = False
   End If
   
   Open "config.dat" For Binary As FileNum
   Get #FileNum, 1, PCSound
   Get #FileNum, 3, DefMapper
   Get #FileNum, 5, DefDevice
   Get #FileNum, 7, DefInstrument
   Get #FileNum, 9, IntervalSkill
   Get #FileNum, 11, ChordSkill
   Get #FileNum, 13, PitchSkill
   Get #FileNum, 15, ShowNeck
   Get #FileNum, 17, ModeSkill
   Get #FileNum, 19, FretboardCol
   Get #FileNum, 21, FingerStretch
   Get #FileNum, 23, Registered
   Get #FileNum, 25, FretNote
   Get #FileNum, 27, ChordCalc
   Close FileNum
   If Registered = 1 Then
      frmSplash.Show 1
   Else
      Registered = 0
      frmUnRegSplash.Show 1
   End If
   If NumDevices < DefDevice Then
      DefDevice = 0
   End If
   DisableMenu
   If NumDevices > 0 Then
      On Error GoTo err
      Call midi_listoutdevs(frmStartUp.lstDevice)
      frmStartUp.Hide
  
      frmStartUp.FillInstrument
      midi_out_close
      midi_out_open (DefDevice)
      Call program_change(0, 0, DefInstrument)
      sbrMain.Panels(1).Text = "Instrument: " + frmStartUp.lstInstruments.List(DefInstrument)
      sbrMain.Panels(2).Text = "Device: " + frmStartUp.lstDevice.List(DefMapper)
      frmStartUp.cmdClose_Click
   End If
 
   Exit Sub

FileErr:
   Unload frmSplash
   MsgBox "Unable to open config.dat", vbExclamation + vbOKOnly, "Error"
   Exit Sub
err:
   
End Sub

Private Sub MDIForm_Terminate()
   midi_out_close

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   all_sounds_off
   midi_out_close
End Sub

Private Sub mnuAbout_Click()
   If Registered = 1 Then
        frmAbout.Show 1
   Else
        frmUnRegAbout.Show 1
   End If
End Sub

Private Sub mnuArrange_Click()
   frmMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
   frmMain.Arrange vbCascade
End Sub

Private Sub mnuChordChart_Click()
   frmChords.Show
   frmChords.SetFocus
End Sub

Private Sub mnuChordFind_Click()
   frmChordFind.Show
   frmChordFind.SetFocus
End Sub

Private Sub mnuCopy_Click()
   CopyIt
End Sub

Private Sub mnuCut_Click()
   CutIt
End Sub

Private Sub mnuDelete_Click()
   DeleteIt
End Sub

Private Sub mnuExit_Click()
   midi_out_close
   End
End Sub

Private Sub mnuFreq_Click()
   frmFreq.Show
   frmFreq.SetFocus
End Sub

Private Sub mnuGlossary_Click()
   Dim nRet As Integer
   
   On Error GoTo err:

   nRet = OSWinHelp(Me.hWnd, "Glossary.hlp", 11, 0)
   Exit Sub
err:
    MsgBox "Cannot find glossary file", vbCritical + vbOKOnly, "Error"

End Sub

Private Sub mnuHelp2_Click()
   Dim nRet As Integer
   
   On Error GoTo err:

   nRet = OSWinHelp(Me.hWnd, App.HelpFile, 11, 0)
   Exit Sub
err:
    MsgBox "Cannot find help file", vbCritical + vbOKOnly, "Error"

End Sub

Private Sub mnuInterval_Click()
   frmInterval.Show
   frmInterval.SetFocus
End Sub

Private Sub mnuMetro_Click()
   frmMetro.Show
   frmMetro.SetFocus
End Sub


Private Sub OpenTab()
   frmMain.ActiveForm.OpenTab
End Sub

Private Sub mnuNoteFind_Click()
   frmNoteFind.Show
   frmNoteFind.SetFocus
End Sub

Private Sub mnuOptions_Click()
   frmOptions.Show
   frmOptions.SetFocus
End Sub

Private Sub mnuPaste_Click()
   PasteIt
End Sub

Private Sub mnuPrint_Click()
   frmMain.ActiveForm.PrintData
End Sub

Private Sub mnuPrintTab_Click()
  frmMain.ActiveForm.PrintTab
End Sub

Private Sub mnuSave_Click()
   frmMain.ActiveForm.Save
End Sub

Private Sub mnuSaveAs_Click()
   frmMain.ActiveForm.SaveAs
End Sub

Private Sub mnuScale_Click()
   frmScale.Show
   frmScale.SetFocus
End Sub

Private Sub mnuTileHor_Click()
   frmMain.Arrange vbTileHorizontal

End Sub

Private Sub mnuTileVer_Click()
      frmMain.Arrange vbVertical

End Sub

Private Sub mnuTransposer_Click()
  frmTranspose.Show
  frmTranspose.SetFocus
End Sub

Private Sub mnuTuner_Click()
   frmTuner.Show
   frmTuner.SetFocus
End Sub




Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Transposer" Then
       mnuTransposer_Click
    ElseIf Button.Key = "Charts" Then
       mnuChordChart_Click
    ElseIf Button.Key = "Finder" Then
       mnuChordFind_Click
    ElseIf Button.Key = "Metronome" Then
       mnuMetro_Click
    ElseIf Button.Key = "Tuner" Then
       mnuTuner_Click
     ElseIf Button.Key = "Freq" Then
       mnuFreq_Click
     ElseIf Button.Key = "Options" Then
       mnuOptions_Click
     ElseIf Button.Key = "Gloss" Then
       mnuGlossary_Click
     ElseIf Button.Key = "Help" Then
       mnuHelp2_Click
     ElseIf Button.Key = "Interval" Then
       mnuInterval_Click
     ElseIf Button.Key = "Scale" Then
       mnuScale_Click
     ElseIf Button.Key = "NoteFinder" Then
       mnuNoteFind_Click
       
   End If
End Sub
