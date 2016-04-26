VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Tab/Chord Viewer"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9285
   Begin RichTextLib.RichTextBox txtChords 
      Height          =   6650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11721
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmViewer.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog diaViewer 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   DisableMenu
   frmMain.mnuPrintTab.Visible = True
   frmMain.mnuSep5.Visible = True
   frmMain.mnuSave.Visible = True
   frmMain.mnuSaveAs.Visible = True
   frmMain.mnuSep4.Visible = True
   frmMain.mnuEdit.Visible = True
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Resize()
   Me.txtChords.Width = Me.Width - 120
   Me.txtChords.Height = Me.Height - 400
End Sub

Public Sub SaveAs()
   On Error GoTo err
   diaViewer.ShowSave
   FName = diaViewer.FileName
   txtChords.SaveFile FName
   Me.Caption = "Tab/Chord Editor - " + Trim(FName)
   Exit Sub
err:
End Sub

Public Sub PrintTab()
   On Error GoTo err
   diaViewer.Flags = cdlPDReturnDC + cdlPDNoPageNums
   If txtChords.SelLength = 0 Then
      diaViewer.Flags = diaViewer.Flags + cdlPDAllPages
   Else
      diaViewer.Flags = diaViewer.Flags + cdlPDSelection
   End If
   diaViewer.ShowPrinter
'   Printer.Print ""
   txtChords.SelPrint diaViewer.hDC
'   Printer.Print frmmain.activeform.txtchords.Text
   Printer.EndDoc
   Exit Sub
err:
   MsgBox "Cannot print document", vbOKOnly + vbCritical, "Error"
End Sub

Public Sub OpenTab()
   Dim Filt As String
   Filt = "Chord Files (*.crd)|*.crd|"
   Filt = Filt + "ChordPro Files (*.crdpro)|*.crdpro|"
   Filt = Filt + "Tab Files (*.tab)|*.tab|"
   Filt = Filt + "Bass Tab Files (*.btab)|*.btab|"
   Filt = Filt + "Text Files (*.txt)|*.txt"
   diaViewer.Filter = Filt
   On Error GoTo err
   diaViewer.ShowOpen
   FName = diaViewer.FileName
   Exit Sub
err:
   FName = "!!!!!$$$.%%%"
End Sub

Public Sub Save()
   If FName = "New" Then
      SaveAs
   Exit Sub
   End If
   On Error GoTo err:
   txtChords.SaveFile FName
   Exit Sub
err:
   MsgBox "File has not been saved", vbExclamation + vbOKOnly, "Error"
End Sub
