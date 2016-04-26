VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmChords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   11385
      TabIndex        =   50
      Top             =   5400
      Visible         =   0   'False
      Width           =   11415
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "&Goto"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   49
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox cboGoto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   48
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox cboEndFret 
      Height          =   315
      ItemData        =   "frmChords.frx":014A
      Left            =   10440
      List            =   "frmChords.frx":0190
      TabIndex        =   33
      Text            =   "22"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboStartFret 
      Height          =   315
      ItemData        =   "frmChords.frx":01E3
      Left            =   8520
      List            =   "frmChords.frx":0229
      TabIndex        =   32
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboTuning 
      Height          =   315
      ItemData        =   "frmChords.frx":027B
      Left            =   5520
      List            =   "frmChords.frx":02A9
      TabIndex        =   29
      Text            =   "Standard"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame frmNeck 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1520
      Left            =   600
      TabIndex        =   18
      Top             =   3120
      Width           =   10695
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   41
         Top             =   30
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   40
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   39
         Top             =   510
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   38
         Top             =   750
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   990
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblStrNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   36
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
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
      Begin VB.Shape E2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   195
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
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   30
         Visible         =   0   'False
         Width           =   200
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
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   270
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
         Index           =   22
         Left            =   10338
         Shape           =   3  'Circle
         Top             =   510
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
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   750
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
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   990
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
         Index           =   22
         Left            =   10335
         Shape           =   3  'Circle
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   0
         X2              =   10680
         Y1              =   120
         Y2              =   120
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
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   10680
         Y1              =   600
         Y2              =   600
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
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   10680
         Y1              =   1080
         Y2              =   1080
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
         Height          =   195
         Index           =   0
         Left            =   1869
         Shape           =   3  'Circle
         Top             =   630
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
         Index           =   2
         Left            =   4491
         Shape           =   3  'Circle
         Top             =   390
         Width           =   195
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
         Index           =   7
         Left            =   8224
         Shape           =   3  'Circle
         Top             =   630
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
         Index           =   10
         Left            =   9535
         Shape           =   3  'Circle
         Top             =   870
         Width           =   200
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
         Index           =   1
         X1              =   1605
         X2              =   1605
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
         Index           =   3
         X1              =   3030
         X2              =   3030
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
         Index           =   5
         X1              =   4305
         X2              =   4305
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
         Index           =   7
         X1              =   5431
         X2              =   5431
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
         Index           =   9
         X1              =   6441
         X2              =   6441
         Y1              =   0
         Y2              =   1500
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
         Index           =   11
         X1              =   7340
         X2              =   7340
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
         Index           =   13
         X1              =   8140
         X2              =   8140
         Y1              =   0
         Y2              =   1599
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
         Index           =   16
         X1              =   9181
         X2              =   9181
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
         Index           =   18
         X1              =   9780
         X2              =   9780
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
         Index           =   20
         X1              =   10315
         X2              =   10315
         Y1              =   0
         Y2              =   1500
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
      Begin VB.Image Image1 
         Height          =   1520
         Index           =   1
         Left            =   4320
         Picture         =   "frmChords.frx":033B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   1515
         Index           =   2
         Left            =   8520
         Picture         =   "frmChords.frx":37FE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   1515
         Index           =   0
         Left            =   0
         Picture         =   "frmChords.frx":6CC1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.Data dtaNotes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   2640
   End
   Begin VB.CommandButton cmdOrigPlay 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      Picture         =   "frmChords.frx":A184
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   2640
   End
   Begin VB.CommandButton cmdChordPlay 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      Picture         =   "frmChords.frx":A30E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.Data dtaChord 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cboSuffix 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.ComboBox cboRoot 
      Height          =   315
      ItemData        =   "frmChords.frx":A498
      Left            =   720
      List            =   "frmChords.frx":A49A
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   47
      Top             =   4350
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   46
      Top             =   4110
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   45
      Top             =   3870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   44
      Top             =   3630
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   43
      Top             =   3390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblZeroNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   42
      Top             =   3120
      Visible         =   0   'False
      Width           =   195
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
      Left            =   600
      TabIndex        =   35
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblFormula 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "End Fret:"
      Height          =   255
      Left            =   9600
      TabIndex        =   31
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Start Fret:"
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Tuning:"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCounter 
      Caption         =   "Chord 0 of 0"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   4350
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   4110
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   3870
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   21
      Top             =   3630
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   3390
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDead 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   19
      Top             =   3150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   600
      X2              =   600
      Y1              =   3120
      Y2              =   4640
   End
   Begin VB.Shape E2 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3150
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape B1 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape G1 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3630
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape D1 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape A1 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4110
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape E1 
      FillColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4350
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label s6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label s5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4110
      Width           =   135
   End
   Begin VB.Label s4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label s3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3630
      Width           =   135
   End
   Begin VB.Label s2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3390
      Width           =   135
   End
   Begin VB.Label s1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3135
      Width           =   135
   End
   Begin VB.Label lblNote 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label lblChord 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11280
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Chord:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11280
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmChords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChordType
    St(0 To 5) As Integer
End Type

Dim Notes(0 To 8) As Integer
Dim LetterNotes(0 To 8) As String
Dim Formula(0 To 8) As String
Dim FormNotes(0 To 8) As Integer
Dim NumInst(0 To 8) As Integer
Dim ChordList() As ChordType
Dim NumChords As Long
Dim StrStart(0 To 5) As Integer
Dim StrNotes(0 To 8) As String
Dim Root As Integer
Dim Suf As String
Dim NumNotes As Integer

Private Sub cmdChordPlay_Click()
   Dim i As Integer
   all_sounds_off
   For i = 0 To MaxFret
      If E1(i).Visible = True Then
         Call note_on(2, StrStart(0) + i, 127)
      End If
      If A1(i).Visible = True Then
         Call note_on(2, StrStart(1) + i, 127)
      End If
      If D1(i).Visible = True Then
         Call note_on(2, StrStart(2) + i, 127)
      End If
      If G1(i).Visible = True Then
         Call note_on(2, StrStart(3) + i, 127)
      End If
      If B1(i).Visible = True Then
         Call note_on(2, StrStart(4) + i, 127)
      End If
      If E2(i).Visible = True Then
         Call note_on(2, StrStart(5) + i, 127)
      End If
   Next
   Timer1.Enabled = True
   While Timer1.Enabled = True
      DoEvents
   Wend
   all_sounds_off
End Sub


Private Sub ChangeTuning()
   Dim i As Integer
   Dim Tune As String
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


Private Sub cmdFind_Click()
   Dim Root As String
   Dim i As Integer
   frmWaitChords.MousePointer = 11
   frmChords.MousePointer = 11
   frmWaitChords.Show
   frmWaitChords.Refresh
   cmdNext.Enabled = False
   cmdPrevious.Enabled = False
   cmdPrint.Enabled = False
   cmdFirst.Enabled = False
   cmdLast.Enabled = False
   frmMain.mnuPrint.Enabled = False
   ClearFrets
   
   ChangeTuning
   ShortSuffix = cboSuffix.Text
   LongSuffix = LongChord(ShortSuffix)
   lblChord.Caption = cboRoot.Text + " " + LongSuffix
   GetNotes


   Root = Trim(cboRoot.Text)
   If Len(Root) > 2 Then
      Root = Left(Root, 2)
   End If
   GetChordList
  
   CurPos = 1
   SetStringsTrue
   cboGoto.Clear
   For i = 1 To NumChords
      cboGoto.AddItem Trim(Str(i))
   Next
   cboGoto.Text = "1"
   FillInNotes
   If NumChords = 0 Then
      cmdChordPlay.Enabled = False
      cmdPrevious.Enabled = False
      cmdFirst.Enabled = False
      cmdLast.Enabled = False
      cmdNext.Enabled = False
      cmdGoto.Enabled = False
      cboGoto.Enabled = False
      lblCounter.Caption = ""
      MsgBox "No chord combinations found", vbInformation + vbOKOnly, "Cannot find Chords"
      Exit Sub
   End If
   
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   If NumDevices > 0 Then
      cmdChordPlay.Enabled = True
   End If
   cmdPrint.Enabled = True
   frmMain.mnuPrint.Enabled = True
   If NumChords > 1 Then
      cmdNext.Enabled = True
      cmdFirst.Enabled = True
      cmdLast.Enabled = True
      cmdGoto.Enabled = True
      cboGoto.Enabled = True
   End If
   
   Unload frmWaitChords
   frmChords.MousePointer = 0
End Sub

Private Sub FillInNotes()
   Dim i As Integer
   For i = 0 To 5
      lblStrNote(i).Caption = ""
      lblZeroNote(i).Caption = ""
      lblDead(i).Visible = False
   Next
   
   If NumChords > 0 Then
       For i = 0 To 5
         lblStrNote(i).Visible = True
         lblZeroNote(i).Visible = True
       Next
       If ChordList(CurPos).St(0) > -1 Then
          lblStrNote(0).Left = E1(ChordList(CurPos).St(0)).Left
       Else
          lblStrNote(0).Visible = False
          lblZeroNote(0).Visible = False
          lblDead(0).Visible = True
       End If
       If ChordList(CurPos).St(1) > -1 Then
          lblStrNote(1).Left = A1(ChordList(CurPos).St(1)).Left
       Else
          lblStrNote(1).Visible = False
          lblZeroNote(1).Visible = False
          lblDead(1).Visible = True
       End If
       If ChordList(CurPos).St(2) > -1 Then
          lblStrNote(2).Left = D1(ChordList(CurPos).St(2)).Left
       Else
          lblStrNote(2).Visible = False
          lblZeroNote(2).Visible = False
          lblDead(2).Visible = True
       End If
       If ChordList(CurPos).St(3) > -1 Then
          lblStrNote(3).Left = G1(ChordList(CurPos).St(3)).Left
       Else
          lblStrNote(3).Visible = False
          lblZeroNote(3).Visible = False
          lblDead(3).Visible = True
       End If
       If ChordList(CurPos).St(4) > -1 Then
          lblStrNote(4).Left = B1(ChordList(CurPos).St(4)).Left
       Else
          lblStrNote(4).Visible = False
          lblZeroNote(4).Visible = False
          lblDead(4).Visible = True
       End If
       If ChordList(CurPos).St(5) > -1 Then
          lblStrNote(5).Left = E2(ChordList(CurPos).St(5)).Left
       Else
          lblStrNote(5).Visible = False
          lblZeroNote(5).Visible = False
          lblDead(5).Visible = True
       End If
       If FretNote = 0 Then
          For i = 0 To 5
            If ChordList(CurPos).St(i) <> 0 Then
               lblStrNote(i).Caption = Left(GetNoteText(((StrStart(i) + ChordList(CurPos).St(i)) Mod 12) + 1), 2)
            Else
               lblZeroNote(i).Caption = Left(GetNoteText((StrStart(i) Mod 12) + 1), 2)
            End If
          Next
       ElseIf FretNote = 1 Then
          For i = 0 To 5
            If ChordList(CurPos).St(i) <> 0 Then
               lblStrNote(i).Caption = GetInterval(((StrStart(i) + ChordList(CurPos).St(i)) Mod 12) + 1)
            Else
               lblZeroNote(i).Caption = GetInterval((StrStart(i) Mod 12) + 1)
            End If
          Next
       Else
          For i = 0 To 5
             lblStrNote(i).Visible = False
             lblZeroNote(i).Visible = False
          Next
       End If
    End If
End Sub
Private Function GetInterval(N As Integer) As String
    Dim i As Integer
    For i = 0 To NumNotes - 1
       If Notes(i) = N Then
          GetInterval = Formula(i)
          Exit Function
       End If
    Next
    GetInterval = ""
End Function
Private Sub GetChordList()
   Dim i As Integer
   Dim h As Integer
   Dim j As Integer
   
   Dim MinPos As Integer
   Dim MaxPos As Integer
   Dim StartFret As Integer
   Dim EndFret As Integer
   Dim TempChord As ChordType
    
   ReDim ChordList(1 To 1) As ChordType
   NumChords = 0
   MinPos = 0
   MaxPos = 0
   StartFret = 0
   EndFret = 0
   If Val(cboStartFret) < 0 Or Val(cboEndFret) > MaxFret Then
      MsgBox "Invalid fret numbers", vbOKOnly + vbExclamation, "Error"
   End If
   If ChordCalc = 0 Or ChordCalc = 1 Then
      For h = 0 To 4
         For i = Val(cboStartFret.Text) To Val(cboEndFret.Text)
            For j = 1 To h
               TempChord.St(j - 1) = -1
            Next
            TempChord.St(h) = i
            If IsInChord(((TempChord.St(h) + StrStart(h)) Mod 12) + 1) Then
               MinPos = i
               MaxPos = i
               If MaxPos - FingerStretch < Val(cboStartFret.Text) Then
                  StartFret = Val(cboStartFret.Text)
               Else
                  StartFret = MaxPos - FingerStretch
               End If
               EndFret = MinPos + FingerStretch
               If EndFret > Val(cboEndFret.Text) Then
                  EndFret = Val(cboEndFret.Text)
               End If
               LoopStrings StartFret, EndFret, MinPos, MaxPos, TempChord, h + 1
            End If
         Next
      Next
   Else
      For i = Val(cboStartFret.Text) To Val(cboEndFret.Text)
         TempChord.St(0) = i
         If IsInChord(((TempChord.St(0) + StrStart(0)) Mod 12) + 1) Then
            MinPos = i
            MaxPos = i
            If MaxPos - FingerStretch < Val(cboStartFret.Text) Then
               StartFret = Val(cboStartFret.Text)
            Else
               StartFret = MaxPos - FingerStretch
            End If
            EndFret = MinPos + FingerStretch
            If EndFret > Val(cboEndFret.Text) Then
               EndFret = Val(cboEndFret.Text)
            End If
            LoopStrings StartFret, EndFret, MinPos, MaxPos, TempChord, h + 1
         End If
      Next
   End If
End Sub

Private Sub LoopStrings(oldStart As Integer, oldEnd As Integer, oldMin As Integer, oldMax As Integer, TChord As ChordType, Iteration As Integer)
   Dim i As Integer
   Dim j As Integer
   
   Dim MinPos As Integer
   Dim MaxPos As Integer
   Dim StartFret As Integer
   Dim EndFret As Integer
    
   If Iteration = 6 Then
      If ChordOK(TChord) Then
         NumChords = NumChords + 1
         ReDim Preserve ChordList(1 To NumChords)
         ChordList(NumChords) = TChord
      End If
   Else
      If oldStart > 0 Then
         TChord.St(Iteration) = 0
         If IsInChord(((TChord.St(Iteration) + StrStart(Iteration)) Mod 12) + 1) Then
            MinPos = oldMin
            MaxPos = oldMax
            If MaxPos - FingerStretch < Val(cboStartFret.Text) Then
               StartFret = Val(cboStartFret.Text)
            Else
               StartFret = MaxPos - FingerStretch
            End If
            EndFret = MinPos + FingerStretch
            If EndFret > Val(cboEndFret.Text) Then
               EndFret = Val(cboEndFret.Text)
            End If
            LoopStrings StartFret, EndFret, MinPos, MaxPos, TChord, Iteration + 1
         End If
      End If
      For i = oldStart To oldEnd
         TChord.St(Iteration) = i
         If IsInChord(((TChord.St(Iteration) + StrStart(Iteration)) Mod 12) + 1) Then
            If i < oldMin Then
               MinPos = i
            Else
               MinPos = oldMin
            End If
            If i > oldMax Then
               MaxPos = i
            Else
               MaxPos = oldMax
            End If
            If MaxPos - FingerStretch < Val(cboStartFret.Text) Then
               StartFret = Val(cboStartFret.Text)
            Else
               StartFret = MaxPos - FingerStretch
            End If
            EndFret = MinPos + FingerStretch
            If EndFret > Val(cboEndFret.Text) Then
               EndFret = Val(cboEndFret.Text)
            End If
            LoopStrings StartFret, EndFret, MinPos, MaxPos, TChord, Iteration + 1
         End If
      Next
      If ChordCalc = 0 Then
         TChord.St(Iteration) = -1
         LoopStrings oldStart, oldEnd, oldMin, oldMax, TChord, Iteration + 1
      ElseIf ChordCalc = 1 Then
         For j = Iteration To 5
            TChord.St(j) = -1
         Next
         LoopStrings oldStart, oldEnd, oldMin, oldMax, TChord, 6
      End If
   End If
End Sub
Private Function ChordOK(TChord As ChordType) As Boolean
   'Do checking if complete chord is made
   Dim i As Integer
   Dim j As Integer
   For i = 0 To 8
      NumInst(i) = 0
   Next
   For i = 0 To 5
      For j = 0 To NumNotes - 1
         If TChord.St(i) > -1 Then
            If (((TChord.St(i) + StrStart(i)) Mod 12) + 1) = Notes(j) Then
               NumInst(j) = NumInst(j) + 1
            End If
         End If
      Next
   Next

   'Special cases
   If ShortSuffix = "7b5" Then
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
   ElseIf ShortSuffix = "7#5" Then
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
   ElseIf ShortSuffix = "maj7#5" Then
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
   ElseIf ShortSuffix = "m7b5" Then
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
   ElseIf ShortSuffix = "add9" Then
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
   Else
    Select Case NumNotes
    Case 3
      For j = 0 To NumNotes - 1
        If NumInst(j) = 0 Then
          ChordOK = False
          Exit Function
        End If
      Next
      ChordOK = True
      Exit Function
    Case 4
      For j = 0 To NumNotes - 1
         If NumInst(j) = 0 Then
            If j <> 2 Then
               ChordOK = False
               Exit Function
            End If
         End If
      Next
      ChordOK = True
      Exit Function
    Case 5
      For j = 0 To NumNotes - 1
         If NumInst(j) = 0 Then
            If j <> 2 Then
               ChordOK = False
               Exit Function
            End If
         End If
      Next
      ChordOK = True
      Exit Function
    Case 6
      For j = 0 To NumNotes - 1
         If NumInst(j) = 0 Then
            If j <> 2 Then
               ChordOK = False
               Exit Function
            End If
         End If
      Next
      ChordOK = True
      Exit Function
    Case 7
      For j = 0 To NumNotes - 1
         If NumInst(j) = 0 Then
            If j <> 2 Then
               ChordOK = False
               Exit Function
            End If
         End If
      Next
      ChordOK = True
      Exit Function
    End Select
  End If
  '  ChordOK = False
End Function
Private Function IsInChord(CNote As Integer) As Boolean
   Dim i As Integer
    
   For i = 0 To NumNotes - 1
      If CNote = Notes(i) Then
         IsInChord = True
         Exit Function
      End If
   Next
   IsInChord = False
End Function

Private Sub cmdFirst_Click()
   CurPos = 1
   cmdPrevious.Enabled = False
   If NumChords > 1 Then
      cmdNext.Enabled = True
   End If
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   ClearFrets
   On Error Resume Next
   SetStringsTrue
   FillInNotes
   cboGoto.Text = Str(CurPos)
End Sub

Private Sub cmdGoto_Click()
   Dim GotoPos As Integer
   CurPos = Val(cboGoto.Text)
   cmdNext.Enabled = False
   cmdPrevious.Enabled = False
   If NumChords > 1 Then
      If CurPos < NumChords Then
         cmdNext.Enabled = True
      End If
      If CurPos > 1 Then
         cmdPrevious.Enabled = True
      End If
   End If
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   ClearFrets
   On Error Resume Next
   SetStringsTrue
   FillInNotes
   
End Sub

Private Sub cmdLast_Click()
   CurPos = NumChords
   If NumChords > 1 Then
      cmdPrevious.Enabled = True
   End If
   cmdNext.Enabled = False
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   ClearFrets
   On Error Resume Next
   SetStringsTrue
   FillInNotes
   cboGoto.Text = Str(CurPos)
End Sub

Private Sub cmdNext_Click()
   CurPos = CurPos + 1
   If CurPos = NumChords Then
      cmdNext.Enabled = False
   End If
   cmdPrevious.Enabled = True
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   ClearFrets
   On Error Resume Next
   SetStringsTrue
   FillInNotes
   cboGoto.Text = Str(CurPos)

End Sub

Private Sub cmdPrevious_Click()
   CurPos = CurPos - 1
   If CurPos = 1 Then
      cmdPrevious.Enabled = False
   End If
   cmdNext.Enabled = True
   lblCounter.Caption = "Chord " + Str(CurPos) + " of " + Str(NumChords)
   ClearFrets
   On Error Resume Next
   SetStringsTrue
   FillInNotes
   cboGoto.Text = Str(CurPos)
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
      Printer.Print lblChord.Caption
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
   Picture1.Picture = CaptureActiveWindow(((frmChords.Left + 100) / Screen.TwipsPerPixelX), ((frmChords.Top + frmNeck.Top + 1350) / Screen.TwipsPerPixelY), (frmNeck.Width + 550) / Screen.TwipsPerPixelX, frmNeck.Height / Screen.TwipsPerPixelY)
End Sub

Private Sub cmdPrint_Click()
    CaptureNeck
    PrintData
End Sub

Private Sub Form_Activate()
   DisableMenu
   frmMain.mnuPrint.Visible = True
   cmdPrint.Enabled = False
End Sub


Private Sub cmdOrigPlay_Click()
   Dim i As Integer
   Dim temp As Integer
   Dim Flag As Boolean
   Flag = False
   temp = 0
   For i = 0 To NumNotes - 1
      all_sounds_off
      If temp > Notes(i) Then
         Flag = True
      End If
      If Not Flag Then
         Call note_on(3, 59 + Notes(i), 127)
      Else
         Call note_on(3, 71 + Notes(i), 127)
      End If
      temp = Notes(i)
      Timer2.Enabled = True
      While Timer2.Enabled
         DoEvents
      Wend
   Next
   all_sounds_off
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   StrStart(0) = 40
   StrStart(1) = 45
   StrStart(2) = 50
   StrStart(3) = 55
   StrStart(4) = 59
   StrStart(5) = 64
   
   cboRoot.Clear
   cboRoot.AddItem "C"
   cboRoot.AddItem "C#/Db"
   cboRoot.AddItem "D"
   cboRoot.AddItem "D#/Eb"
   cboRoot.AddItem "E"
   cboRoot.AddItem "F"
   cboRoot.AddItem "F#/Gb"
   cboRoot.AddItem "G"
   cboRoot.AddItem "G#/Ab"
   cboRoot.AddItem "A"
   cboRoot.AddItem "A#/Bb"
   cboRoot.AddItem "B"
   cboRoot.Text = "C"
   If Registered = 1 Then
      cboSuffix.Clear
      cboSuffix.AddItem "aug"
      cboSuffix.AddItem "aug7"
      cboSuffix.AddItem "aug11"
      cboSuffix.AddItem "dim"
      cboSuffix.AddItem "dim7"
      cboSuffix.AddItem "7"
      cboSuffix.AddItem "7sus2"
      cboSuffix.AddItem "7sus4"
      cboSuffix.AddItem "7b5"
      cboSuffix.AddItem "7b9"
      cboSuffix.AddItem "7#5"
      cboSuffix.AddItem "7#5b9"
      cboSuffix.AddItem "7#9"
      cboSuffix.AddItem "9"
      cboSuffix.AddItem "9b5"
      cboSuffix.AddItem "9#5"
      cboSuffix.AddItem "11"
      cboSuffix.AddItem "13"
      cboSuffix.AddItem "maj"
      cboSuffix.AddItem "maj6/9"
      cboSuffix.AddItem "maj6/9 #11"
      cboSuffix.AddItem "6"
      cboSuffix.AddItem "6add9"
      cboSuffix.AddItem "maj7"
      cboSuffix.AddItem "maj7#5"
      cboSuffix.AddItem "maj7b3"
      cboSuffix.AddItem "maj7#11"
      cboSuffix.AddItem "maj9"
      cboSuffix.AddItem "add9"
      cboSuffix.AddItem "m"
      cboSuffix.AddItem "m6/9"
      cboSuffix.AddItem "mb6"
      cboSuffix.AddItem "m6"
      cboSuffix.AddItem "min/maj 7"
      cboSuffix.AddItem "m7"
      cboSuffix.AddItem "m7b3"
      cboSuffix.AddItem "m7b5"
      cboSuffix.AddItem "mb6"
      cboSuffix.AddItem "min/maj 9"
      cboSuffix.AddItem "m9"
      cboSuffix.AddItem "m11"
      cboSuffix.AddItem "m13"
      cboSuffix.AddItem "sus4"
      cboSuffix.AddItem "sus2"
   ElseIf Registered = 0 Then
      cboSuffix.Clear
      cboSuffix.AddItem "aug7"
      cboSuffix.AddItem "dim7"
      cboSuffix.AddItem "7"
      cboSuffix.AddItem "maj"
      cboSuffix.AddItem "6"
      cboSuffix.AddItem "maj7"
      cboSuffix.AddItem "m"
      cboSuffix.AddItem "m7"
      cboSuffix.AddItem "sus4"
      cboSuffix.AddItem "sus2"
   End If
   cboSuffix.Text = "maj"
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
   End If
   
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
End Sub


Private Sub FindNotes()
   Dim i As Integer
    
   On Error GoTo err:
   dtaNotes.DatabaseName = "Chords.mdb"
   dtaNotes.Refresh
   dtaNotes.RecordSource = "SELECT * FROM ChordTypeFind WHERE Name = '" + Trim(Suf) + "'"
   dtaNotes.Refresh
   NumNotes = dtaNotes.Recordset.Fields(1).value
   For i = 0 To NumNotes - 1
      Notes(i) = dtaNotes.Recordset.Fields(i + 2).value
      FormNotes(i) = Notes(i)
   Next
   Exit Sub
err:
    
   End Sub
Private Sub GetNotes()
   Dim i As Integer
   Dim j As Integer
   Dim temp As String
   Dim inttemp As Integer
   
   For i = 0 To 8
      LetterNotes(i) = ""
      Formula(i) = ""
   Next
   temp = Trim(UCase(cboRoot.Text))
   If Len(temp) > 2 Then
      temp = Left(temp, 2)
   End If
   If temp <> "" Then
      Root = GetNoteNumber(temp)
   End If
   Suf = LongSuffix
   FindNotes
   
   For i = 0 To NumNotes - 1
      Notes(i) = Notes(i) + Root
   Next
   For j = 1 To 3
      For i = 1 To NumNotes - 1
         If Notes(i) > 12 Then
            Notes(i) = Notes(i) - 12
         End If
      Next
   Next
   
   For i = 0 To NumNotes - 1
      LetterNotes(i) = GetNoteText(Notes(i))
      Formula(i) = GetIntervalText(FormNotes(i))
   Next
   lblNote.Caption = ""
   lblFormula.Caption = ""
   For i = 0 To NumNotes - 1
       lblNote.Caption = lblNote.Caption + LetterNotes(i) + " "
       lblFormula.Caption = lblFormula.Caption + Formula(i) + " "
   Next

   If NumDevices > 0 Then
      cmdOrigPlay.Enabled = True
   End If
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
End Sub

Private Sub SetStringsTrue()
   If ChordList(CurPos).St(0) > -1 Then
      E1(ChordList(CurPos).St(0)).Visible = True
   End If
   If ChordList(CurPos).St(1) > -1 Then
      A1(ChordList(CurPos).St(1)).Visible = True
   End If
   If ChordList(CurPos).St(2) > -1 Then
      D1(ChordList(CurPos).St(2)).Visible = True
   End If
   If ChordList(CurPos).St(3) > -1 Then
      G1(ChordList(CurPos).St(3)).Visible = True
   End If
   If ChordList(CurPos).St(4) > -1 Then
      B1(ChordList(CurPos).St(4)).Visible = True
   End If
   If ChordList(CurPos).St(5) > -1 Then
      E2(ChordList(CurPos).St(5)).Visible = True
   End If
End Sub

Private Sub ClearFrets()
   Dim i As Integer
   
   For i = 0 To MaxFret
      E1(i).Visible = False
      A1(i).Visible = False
      D1(i).Visible = False
      G1(i).Visible = False
      B1(i).Visible = False
      E2(i).Visible = False
   Next
End Sub
