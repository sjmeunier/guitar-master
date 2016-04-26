VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNoteFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note Finder"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   Icon            =   "frmNoteFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   11505
   Tag             =   "+4"
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   11505
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   11535
   End
   Begin VB.CommandButton cmdOrigPlay 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      Picture         =   "frmNoteFind.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   375
   End
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame frmNeck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1520
      Left            =   720
      TabIndex        =   14
      Top             =   2880
      Width           =   10695
      Begin VB.Shape Cur 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   200
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   21
         Left            =   10080
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   20
         Left            =   9825
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   19
         Left            =   9540
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   18
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   8925
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   16
         Left            =   8580
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   15
         Left            =   8220
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   14
         Left            =   7845
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   13
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   12
         Left            =   7020
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   11
         Left            =   6570
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   6090
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   9
         Left            =   5595
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   5055
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4485
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   3885
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2580
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1875
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1110
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   312
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   21
         Left            =   10080
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   20
         Left            =   9825
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   19
         Left            =   9540
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   18
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   8925
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   16
         Left            =   8580
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   15
         Left            =   8220
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   14
         Left            =   7845
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   13
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   12
         Left            =   7020
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   11
         Left            =   6570
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   6090
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   9
         Left            =   5595
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   5055
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4490
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   3885
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2580
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1875
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1110
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape A1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   312
         Shape           =   3  'Circle
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   21
         Left            =   10080
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   20
         Left            =   9825
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   19
         Left            =   9540
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   18
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   8925
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   16
         Left            =   8580
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   15
         Left            =   8220
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   14
         Left            =   7845
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   13
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   12
         Left            =   7020
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   11
         Left            =   6570
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   6090
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   9
         Left            =   5595
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   5055
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4485
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   3885
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2580
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1110
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape D1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   315
         Shape           =   3  'Circle
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   21
         Left            =   10086
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   20
         Left            =   9819
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   19
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   18
         Left            =   9235
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   8918
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   16
         Left            =   8581
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   15
         Left            =   8224
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   14
         Left            =   7846
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   13
         Left            =   7446
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   12
         Left            =   7022
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   11
         Left            =   6572
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   6096
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   9
         Left            =   5591
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   5057
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   3891
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2582
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1113
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape G1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   312
         Shape           =   3  'Circle
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   21
         Left            =   10086
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   20
         Left            =   9819
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   19
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   18
         Left            =   9235
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   8918
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   16
         Left            =   8581
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   15
         Left            =   8224
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   14
         Left            =   7846
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   13
         Left            =   7446
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   12
         Left            =   7022
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   11
         Left            =   6572
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   6096
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   9
         Left            =   5591
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   5057
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   3891
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2582
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1113
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape B1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   312
         Shape           =   3  'Circle
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   21
         Left            =   10086
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   20
         Left            =   9819
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   19
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   18
         Left            =   9235
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   17
         Left            =   8918
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   16
         Left            =   8581
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   15
         Left            =   8224
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   14
         Left            =   7846
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   13
         Left            =   7446
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   12
         Left            =   7022
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   11
         Left            =   6572
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   10
         Left            =   6096
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   9
         Left            =   5591
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   8
         Left            =   5057
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   7
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   6
         Left            =   3891
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   5
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   4
         Left            =   2582
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   3
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   2
         Left            =   1113
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   1
         Left            =   312
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   10
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   870
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   9
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   380
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   8
         Left            =   8918
         Shape           =   3  'Circle
         Top             =   630
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   7
         Left            =   8224
         Shape           =   3  'Circle
         Top             =   630
         Width           =   200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   5
         X1              =   0
         X2              =   10680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   6
         Left            =   7022
         Shape           =   3  'Circle
         Top             =   1130
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   5
         Left            =   7022
         Shape           =   3  'Circle
         Top             =   150
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   4
         Left            =   5591
         Shape           =   3  'Circle
         Top             =   630
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   870
         Width           =   195
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   390
         Width           =   195
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   1
         Left            =   3255
         Shape           =   3  'Circle
         Top             =   630
         Width           =   200
      End
      Begin VB.Shape Dot 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   630
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   10680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   10680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   10680
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   10680
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   0
         X2              =   10680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   15
         X1              =   8854
         X2              =   8854
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   20
         X1              =   10315
         X2              =   10315
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   19
         X1              =   10056
         X2              =   10056
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   9780
         X2              =   9780
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   9489
         X2              =   9489
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   9181
         X2              =   9181
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   14
         X1              =   8507
         X2              =   8507
         Y1              =   0
         Y2              =   1599
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   13
         X1              =   8140
         X2              =   8140
         Y1              =   0
         Y2              =   1599
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   7752
         X2              =   7752
         Y1              =   0
         Y2              =   1599
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   7340
         X2              =   7340
         Y1              =   0
         Y2              =   1599
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   6903
         X2              =   6903
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   6441
         X2              =   6441
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   8
         X1              =   5951
         X2              =   5951
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   5431
         X2              =   5431
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   4882
         X2              =   4882
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   4305
         X2              =   4305
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   3675
         X2              =   3675
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   3030
         X2              =   3030
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   2340
         X2              =   2340
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   1605
         X2              =   1605
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   21
         X1              =   10560
         X2              =   10560
         Y1              =   0
         Y2              =   1500
      End
      Begin VB.Image Image1 
         Height          =   1520
         Index           =   1
         Left            =   4320
         Picture         =   "frmNoteFind.frx":02D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   1515
         Index           =   0
         Left            =   0
         Picture         =   "frmNoteFind.frx":3797
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   1515
         Index           =   2
         Left            =   8520
         Picture         =   "frmNoteFind.frx":6C5A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.ComboBox cboTuning 
      Height          =   315
      ItemData        =   "frmNoteFind.frx":A11D
      Left            =   1920
      List            =   "frmNoteFind.frx":A14B
      TabIndex        =   12
      Text            =   "Standard"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cboKey 
      Height          =   315
      ItemData        =   "frmNoteFind.frx":A1DD
      Left            =   8040
      List            =   "frmNoteFind.frx":A205
      TabIndex        =   9
      Text            =   "C"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1800
      Top             =   360
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Data dtaScale 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cboScale 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblNote 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblFormula 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "Formula:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   720
      X2              =   720
      Y1              =   2880
      Y2              =   4380
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tuning:"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Key:"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label s1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2910
      Width           =   465
   End
   Begin VB.Label s2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3150
      Width           =   465
   End
   Begin VB.Label s3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3390
      Width           =   465
   End
   Begin VB.Label s4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3630
      Width           =   465
   End
   Begin VB.Label s5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3870
      Width           =   465
   End
   Begin VB.Label s6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4110
      Width           =   465
   End
   Begin VB.Shape E1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4110
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape A1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape D1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3630
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape G1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape B1 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3165
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape E2 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   2910
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   120
      X2              =   11400
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Scale:"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8520
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmNoteFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tune As String
Dim StartE1 As Integer
Dim StartA1 As Integer
Dim StartD1 As Integer
Dim StartG1 As Integer
Dim StartB1 As Integer
Dim StartE2 As Integer
Dim FretNum As Integer
Dim St As Integer
Dim FretNumB As Integer
Dim StB As Integer
Dim NScale As String
Dim NKey As String
Dim ScaleNote(0 To 11) As Integer
Dim ScaleForm(0 To 11) As String
Dim ScaleFormInt(0 To 11) As Integer
Dim ScaleNoteLetter(0 To 11) As String
Dim RootNote As Integer
Dim NumNotes As Integer

Private Sub cmdChange_Click()
   Dim i As Integer
   frmNoteFind.MousePointer = 11
   
   Tune = cboTuning.Text
   ChangeTune Tune, StrDelta(), StrText()
   s1.Caption = Left(StrText(0), 2)
   s2.Caption = Left(StrText(1), 2)
   s3.Caption = Left(StrText(2), 3)
   s4.Caption = Left(StrText(3), 4)
   s5.Caption = Left(StrText(4), 5)
   s6.Caption = Left(StrText(5), 6)
   StartE1 = StrDelta(0)
   StartA1 = StrDelta(1)
   StartD1 = StrDelta(2)
   StartG1 = StrDelta(3)
   StartB1 = StrDelta(4)
   StartE2 = StrDelta(5)
   For i = 0 To 22
      E1(i).Visible = False
      A1(i).Visible = False
      D1(i).Visible = False
      G1(i).Visible = False
      B1(i).Visible = False
      E2(i).Visible = False
   Next
   For i = 0 To NumNotes - 1
      DoE1 ScaleNote(i)
      DoA1 ScaleNote(i)
      DoD1 ScaleNote(i)
      DoG1 ScaleNote(i)
      DoB1 ScaleNote(i)
      DoE2 ScaleNote(i)
   Next
   frmNoteFind.MousePointer = 0
End Sub

Private Sub cmdFind_Click()
   Dim i As Integer
   frmNoteFind.MousePointer = 11
   
   cmdPrint.Enabled = False
   cmdOrigPlay.Enabled = False
   frmMain.mnuPrint.Enabled = False

   For i = 0 To 22
      E1(i).Visible = False
      A1(i).Visible = False
      D1(i).Visible = False
      G1(i).Visible = False
      B1(i).Visible = False
      E2(i).Visible = False
   Next
   NScale = Trim(cboScale.Text)
   NKey = Trim(cboKey.Text)
   
   If Len(NKey) > 2 Then
      NKey = Left(NKey, 2)
   End If
   RootNote = GetNoteNumber(NKey)
   RootNote = (RootNote + 7) Mod 12
   On Error GoTo err:
   dtaScale.DatabaseName = "Chords.mdb"
   dtaScale.Refresh
   dtaScale.RecordSource = "SELECT * FROM Scales WHERE Name = '" + NScale + "'"
   dtaScale.Refresh
   dtaScale.Recordset.MoveFirst
   dtaScale.Recordset.MoveLast
   dtaScale.Recordset.MoveFirst
   NumNotes = dtaScale.Recordset.Fields(1).value
   For i = 0 To NumNotes - 1
      ScaleNote(i) = dtaScale.Recordset.Fields(i + 2).value
      ScaleNoteLetter(i) = GetNoteText(((ScaleNote(i) + RootNote + 4) Mod 12) + 1)
      ScaleForm(i) = GetIntervalText(ScaleNote(i))
   Next
   lblNote.Caption = ""
   lblFormula.Caption = ""
   For i = 0 To NumNotes - 1
      lblNote.Caption = lblNote.Caption + ScaleNoteLetter(i) + " "
      lblFormula.Caption = lblFormula.Caption + ScaleForm(i) + " "
   Next
   For i = 0 To NumNotes - 1
      DoE1 ScaleNote(i)
      DoA1 ScaleNote(i)
      DoD1 ScaleNote(i)
      DoG1 ScaleNote(i)
      DoB1 ScaleNote(i)
      DoE2 ScaleNote(i)
   Next
   cmdChange.Enabled = True
   cmdPrint.Enabled = True
   frmMain.mnuPrint.Enabled = True
   frmNoteFind.MousePointer = 0
   cmdOrigPlay.Enabled = True
   Exit Sub
err:
   MsgBox "Cannot find scale", vbInformation + vbOKOnly, "Message"
   frmNoteFind.MousePointer = 0
End Sub

Private Sub cmdOrigPlay_Click()
   Dim i As Integer
   Dim temp As Integer
   Dim Flag As Boolean
   Flag = False
   temp = 0
   For i = 0 To NumNotes - 1
      all_sounds_off
      If temp > ((ScaleNote(i) + RootNote + 4) Mod 12) + 1 Then
         Flag = True
      End If
      If Not Flag Then
         Call note_on(3, 59 + ((ScaleNote(i) + RootNote + 4) Mod 12) + 1, 127)
      Else
         Call note_on(3, 71 + ((ScaleNote(i) + RootNote + 4) Mod 12) + 1, 127)
      End If
      temp = ((ScaleNote(i) + RootNote + 4) Mod 12) + 1
      Timer2.Enabled = True
      While Timer2.Enabled
         DoEvents
      Wend
   Next
   Call note_on(3, 59 + 12 + ((ScaleNote(0) + RootNote + 4) Mod 12) + 1, 127)
   Timer2.Enabled = True
   While Timer2.Enabled
      DoEvents
   Wend
   all_sounds_off
End Sub

Private Sub cmdPrint_Click()
   CaptureNeck
   PrintData
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
      Printer.Print cboKey.Text + " " + cboScale.Text
      Printer.Print
      Printer.FontSize = 12
      Printer.Print "Notes:"; Tab; Trim(lblNote.Caption)
      Printer.Print "Formula:"; Tab; Trim(lblFormula.Caption)
      Printer.PaintPicture Picture1.Picture, 200, 2000
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.FontSize = 8
      Printer.FontBold = False
      Printer.Print "Guitar Master Pro - Copyright 2000 Opus Software"
      Printer.EndDoc
   Next i
   Exit Sub
PrintErr:
   MsgBox "Chord was not printed", vbExclamation + vbOKOnly, "Print"
   Exit Sub
CancelErr:
End Sub
Private Sub CaptureNeck()
   Picture1.Picture = CaptureActiveWindow(((frmNoteFind.Left + 100) / Screen.TwipsPerPixelX), ((frmNoteFind.Top + frmNeck.Top + 1350) / Screen.TwipsPerPixelY), (frmNeck.Width + 750) / Screen.TwipsPerPixelX, frmNeck.Height / Screen.TwipsPerPixelY)
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
   Tune = "Standard"
   s1.Caption = "E"
   s2.Caption = "B"
   s3.Caption = "G"
   s4.Caption = "D"
   s5.Caption = "A"
   s6.Caption = "E"
   StartE1 = 0
   StartA1 = 0
   StartD1 = 0
   StartG1 = 0
   StartB1 = 0
   StartE2 = 0

   If FretboardCol = 1 Then
      frmNeck.BackColor = vbWhite
      For i = 0 To 2
         Image1(i).Picture = LoadPicture
      Next
      For i = 1 To 6
         Line3(i).BorderColor = vbBlack
      Next
      For i = 0 To 10
         Dot(i).FillColor = vbBlack
      Next
   ElseIf FretboardCol = 2 Then
      frmNeck.BackColor = vbBlack
      For i = 0 To 2
         Image1(i).Picture = LoadPicture
      Next
      For i = 1 To 6
         Line3(i).BorderColor = vbWhite
      Next
      For i = 0 To 10
         Dot(i).FillColor = vbWhite
      Next
   End If
   If Registered = 0 Then
      cboTuning.Enabled = False
      cmdChange.Enabled = False
   End If
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   all_sounds_off
   If Button = vbLeftButton Then
      If Index = 1 Then
         X = X + Image1(0).Width
      ElseIf Index = 2 Then
         X = X + Image1(0).Width + Image1(1).Width
      End If
      FretNum = GetFretNum(X)
      St = GetStringNum(Y)
      If FretNum > 22 Then
         Exit Sub
      End If
      PlayNote
   End If
End Sub

Private Sub PlayNote()
      If St = 1 Then
           Call note_on(3, 64 + StartE2 + FretNum, 127)
           Cur.Left = E2(FretNum).Left
           Cur.Top = E2(FretNum).Top
           Cur.Visible = True
      ElseIf St = 2 Then
           Call note_on(3, 59 + StartB1 + FretNum, 127)
           Cur.Left = B1(FretNum).Left
           Cur.Top = B1(FretNum).Top
           Cur.Visible = True
      ElseIf St = 3 Then
           Call note_on(3, 55 + StartG1 + FretNum, 127)
           Cur.Left = G1(FretNum).Left
           Cur.Top = G1(FretNum).Top
           Cur.Visible = True
      ElseIf St = 4 Then
           Call note_on(3, 50 + StartD1 + FretNum, 127)
           Cur.Left = D1(FretNum).Left
           Cur.Top = D1(FretNum).Top
           Cur.Visible = True
      ElseIf St = 5 Then
           Call note_on(3, 45 + StartA1 + FretNum, 127)
           Cur.Left = A1(FretNum).Left
           Cur.Top = A1(FretNum).Top
           Cur.Visible = True
      ElseIf St = 6 Then
           Call note_on(3, 40 + StartE1 + FretNum, 127)
           Cur.Left = E1(FretNum).Left
           Cur.Top = E1(FretNum).Top
           Cur.Visible = True
      End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      If Index = 1 Then
         X = X + Image1(0).Width
      ElseIf Index = 2 Then
         X = X + Image1(0).Width + Image1(1).Width
      End If
      FretNumB = GetFretNum(X)
      StB = GetStringNum(Y)
      If FretNumB > 22 Or (FretNumB = FretNum And StB = St) Then
         Exit Sub
      End If
      FretNum = FretNumB
      St = StB
      all_sounds_off
      PlayNote
   End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   all_sounds_off
   Cur.Visible = False
End Sub



Private Sub Timer1_Timer()
   Timer1.Enabled = False
End Sub

Private Sub DoE1(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  
  Counter = Note + RootNote - 24 + StartE1
  While Counter < 23
     If Counter > -1 Then
        E1(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        E1(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub DoA1(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  
  Counter = Note + RootNote - 24 - 5 + StartA1
  While Counter < 23
     If Counter > -1 Then
        A1(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        A1(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub DoD1(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  Counter = Note + RootNote - 24 - 10 + StartD1
  While Counter < 23
     If Counter > -1 Then
        D1(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        D1(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub DoG1(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  Counter = Note + RootNote - 24 - 15 + StartG1
  While Counter < 23
     If Counter > -1 Then
        G1(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        G1(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub DoB1(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  Counter = Note + RootNote - 24 - 19 + StartB1
  While Counter < 23
     If Counter > -1 Then
        B1(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        B1(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub DoE2(Note As Integer)
  Dim Counter As Integer
  Dim colCount
  Counter = Note + RootNote - 24 - 24 + StartE2
  While Counter < 23
     If Counter > -1 Then
        E2(Counter).FillColor = QBColor((colCount Mod 8) + 8)
        E2(Counter).Visible = True
     End If
     Counter = Counter + 12
     colCount = colCount + 1
  Wend
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
End Sub
