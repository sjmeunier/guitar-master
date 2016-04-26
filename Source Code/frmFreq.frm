VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFreq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frequency Finder"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   Icon            =   "frmFreq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   10680
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdFreq 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   13
      Cols            =   11
      ScrollBars      =   0
   End
   Begin MSComctlLib.StatusBar sbrFreq 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim i As Integer


Private Sub Form_Activate()
   DisableMenu
   frmMain.mnuPrint.Visible = True
   frmMain.mnuPrint.Enabled = True
End Sub

Public Sub PrintData()
'print
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim L As Integer
   
   Dim NumCopies As Integer
   Dim BeginPage As Integer
   Dim EndPage As Integer
   Dim temp As String
   Dim Freqs As String

   Freqs = ""
   diaCommon.CancelError = True
   On Error GoTo CancelErr
   diaCommon.ShowPrinter
   On Error GoTo PrintErr
   BeginPage = diaCommon.FromPage
   EndPage = diaCommon.ToPage
   NumCopies = diaCommon.Copies
   frmPrintWait.Show
   frmPrintWait.Refresh
   For i = 1 To NumCopies
      Printer.FontSize = 16
      Printer.FontBold = True
      Printer.Font = "Arial"

      Printer.Print "Frequency Table"
      Printer.Print
      Printer.FontSize = 8
      Printer.Font = "Courier"
      Printer.Print
      Printer.Print
      Printer.FontBold = False
      For j = 0 To grdFreq.Rows - 1
         Freqs = ""
         For L = 0 To grdFreq.Cols - 1
            grdFreq.Col = L
            grdFreq.Row = j
            temp = ""
            For k = 0 To (10 - Len(grdFreq.Text))
               temp = temp + " "
            Next
            Freqs = Freqs + grdFreq.Text + temp
         Next
         Printer.Print Freqs
      Next
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.FontSize = 8
      Printer.Font = "Arial"
      Printer.Print "Guitar Master Pro - Copyright 2000 Serge Meunier"
      Printer.EndDoc
   Next i
   frmPrintWait.Hide
   Exit Sub
PrintErr:
   MsgBox "Document was not printed", vbExclamation + vbOKOnly, "Print"
   frmPrintWait.Hide
   Exit Sub
CancelErr:
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim TmpHeight As Integer
   Dim TmpWidth As Integer
   FillNames
   grdFreq.ColWidth(0) = 800
   TmpHeight = grdFreq.RowHeight(0)
   TmpWidth = grdFreq.ColWidth(0)
   For i = 1 To 12
      grdFreq.TextMatrix(i, 0) = NoteName(i - 1)
      TmpHeight = TmpHeight + grdFreq.RowHeight(i)
   Next
   For i = 1 To 10
      grdFreq.TextMatrix(0, i) = "Octave " + Trim(Str(i))
      TmpWidth = TmpWidth + grdFreq.ColWidth(i)
   Next
   grdFreq.Width = TmpWidth + 90
   grdFreq.Height = TmpHeight + 90
   FillHz
End Sub

Private Sub FillHz()
   Dim i As Integer
   Dim j As Integer
   
   FillFreqs
   For i = 0 To 11
      For j = 0 To 9
         grdFreq.TextMatrix(i + 1, j + 1) = Trim(Str(Freqs(j, i))) + "Hz"
      Next
  Next
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
   all_sounds_off
End Sub

Private Sub grdFreq_Click()
   Dim N As Integer
   Dim Oct As Integer
   Timer1.Enabled = False
   If grdFreq.Col > 0 And grdFreq.Row > 0 Then
      N = grdFreq.Row - 1
      Oct = grdFreq.Col - 1
      If Oct = 9 And N > 7 Then
        Exit Sub
      End If
      
      sbrFreq.Panels(1).Text = "Note:" + NoteName(N)
      sbrFreq.Panels(2).Text = "Frequency:" + Str(Freqs(Oct, N)) + " Hz"
      sbrFreq.Panels(3).Text = "Wavelength:" + Str(330 / Freqs(Oct, N)) + " m"
      all_sounds_off
      Call note_on(0, (12 * (Oct + 1)) + N, 127)
      Timer1.Enabled = True
      While Timer1.Enabled = True
         DoEvents
      Wend
      all_sounds_off
   End If
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
End Sub
