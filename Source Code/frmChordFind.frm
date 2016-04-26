VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChordFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chord Finder"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmChordFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog diaCommon 
      Left            =   0
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstChordFind 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Fretboard"
      TabPicture(0)   =   "frmChordFind.frx":0742
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "frmChordFind.frx":075E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   11385
         TabIndex        =   48
         Top             =   4680
         Visible         =   0   'False
         Width           =   11415
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   30
         Top             =   960
         Width           =   11175
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   375
            Left            =   3600
            TabIndex        =   40
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   2520
            TabIndex        =   39
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdChordPlay 
            Enabled         =   0   'False
            Height          =   375
            Left            =   4920
            Picture         =   "frmChordFind.frx":077A
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1680
            Width           =   375
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   5640
            Top             =   1800
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   0
            ItemData        =   "frmChordFind.frx":0904
            Left            =   1080
            List            =   "frmChordFind.frx":092C
            TabIndex        =   37
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   1
            ItemData        =   "frmChordFind.frx":0968
            Left            =   1920
            List            =   "frmChordFind.frx":0990
            TabIndex        =   36
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   2
            ItemData        =   "frmChordFind.frx":09CC
            Left            =   2760
            List            =   "frmChordFind.frx":09F4
            TabIndex        =   35
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   3
            ItemData        =   "frmChordFind.frx":0A30
            Left            =   3600
            List            =   "frmChordFind.frx":0A58
            TabIndex        =   34
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   4
            ItemData        =   "frmChordFind.frx":0A94
            Left            =   4440
            List            =   "frmChordFind.frx":0ABC
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
         Begin VB.Data dtaChords 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   "Chords.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   600
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   5
            ItemData        =   "frmChordFind.frx":0AF8
            Left            =   5280
            List            =   "frmChordFind.frx":0B20
            TabIndex        =   32
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox txtNote 
            Height          =   315
            Index           =   6
            ItemData        =   "frmChordFind.frx":0B5C
            Left            =   6120
            List            =   "frmChordFind.frx":0B84
            TabIndex        =   31
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Chord:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   1680
            Width           =   975
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   11160
            Y1              =   1300
            Y2              =   1300
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Notes:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblChord 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1680
            TabIndex        =   41
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   0
            X2              =   11160
            Y1              =   1320
            Y2              =   1320
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11295
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   375
            Left            =   3600
            TabIndex        =   47
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton cmdChordPlayNeck 
            Height          =   375
            Left            =   5160
            Picture         =   "frmChordFind.frx":0BC0
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtChordList 
            BackColor       =   &H8000000F&
            Height          =   1575
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   2520
            Width           =   2415
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "&Change"
            Height          =   375
            Left            =   3600
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.ComboBox cboTuning 
            Height          =   315
            ItemData        =   "frmChordFind.frx":0D4A
            Left            =   1440
            List            =   "frmChordFind.frx":0D78
            TabIndex        =   9
            Text            =   "Standard"
            Top             =   120
            Width           =   2055
         End
         Begin VB.Frame frmNeck 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1520
            Left            =   600
            TabIndex        =   2
            Top             =   600
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
               Index           =   0
               Left            =   0
               TabIndex        =   8
               Top             =   1200
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
               Left            =   45
               TabIndex        =   7
               Top             =   960
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
               Left            =   45
               TabIndex        =   6
               Top             =   720
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
               Left            =   45
               TabIndex        =   5
               Top             =   480
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
               Left            =   45
               TabIndex        =   4
               Top             =   240
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
               Index           =   5
               Left            =   45
               TabIndex        =   3
               Top             =   0
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
               Height          =   195
               Index           =   2
               Left            =   1110
               Shape           =   3  'Circle
               Top             =   0
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
               Index           =   5
               X1              =   0
               X2              =   10680
               Y1              =   1320
               Y2              =   1320
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
               Index           =   21
               X1              =   10560
               X2              =   10560
               Y1              =   0
               Y2              =   1500
            End
            Begin VB.Image Image1 
               Height          =   1515
               Index           =   2
               Left            =   8520
               Picture         =   "frmChordFind.frx":0E0A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2175
            End
            Begin VB.Image Image1 
               Height          =   1515
               Index           =   1
               Left            =   4320
               Picture         =   "frmChordFind.frx":42CD
               Stretch         =   -1  'True
               Top             =   0
               Width           =   4215
            End
            Begin VB.Image Image1 
               Height          =   1515
               Index           =   0
               Left            =   0
               Picture         =   "frmChordFind.frx":7790
               Stretch         =   -1  'True
               Top             =   0
               Width           =   4335
            End
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   16
            Top             =   630
            Width           =   195
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   15
            Top             =   870
            Width           =   195
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   13
            Top             =   1350
            Width           =   195
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   14
            Top             =   1110
            Width           =   195
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   12
            Top             =   1590
            Width           =   195
         End
         Begin VB.Label lblZeroNote 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   11
            Top             =   1830
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "Chord List:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   855
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   11280
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tuning:"
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   120
            Width           =   735
         End
         Begin VB.Label s1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "E"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   615
            Width           =   360
         End
         Begin VB.Label s2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   870
            Width           =   360
         End
         Begin VB.Label s3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "G"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label s4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   1350
            Width           =   360
         End
         Begin VB.Label s5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   1590
            Width           =   360
         End
         Begin VB.Label s6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "E"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   1830
            Width           =   360
         End
         Begin VB.Shape E1 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   1830
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape A1 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   1590
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape D1 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   1350
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape G1 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   1110
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape B1 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   870
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape E2 
            FillColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   360
            Shape           =   3  'Circle
            Top             =   630
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            X1              =   600
            X2              =   600
            Y1              =   600
            Y2              =   2120
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   1830
            Width           =   135
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   1590
            Width           =   135
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   20
            Top             =   1350
            Width           =   135
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   1110
            Width           =   135
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   18
            Top             =   870
            Width           =   135
         End
         Begin VB.Label lblDead 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   17
            Top             =   630
            Width           =   135
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   11280
            Y1              =   2415
            Y2              =   2400
         End
      End
   End
End
Attribute VB_Name = "frmChordFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChordNote(0 To 8) As Integer
Dim ChordNs(0 To 8) As Integer
Dim NoteCount As Integer
Dim Root As String
Dim RootNum As String
Dim NoteTxt As String
Dim Suffix As String

Dim NumTemp As Integer
Dim NumList As Integer
Dim TempNotes(0 To 5) As Integer
Dim NoteList() As NotesType
Dim StrStart(0 To 5) As Integer
Dim StrFrets(0 To 5) As Integer
Dim St As Integer
Dim FretNum As Integer

Private Sub cmdChange_Click()
   Dim i As Integer
   Dim Tune As String
   Tune = cboTuning.Text
   ChangeTune Tune, StrDelta(), StrText()
   s1.Caption = Left(StrText(0), 2)
   s2.Caption = Left(StrText(1), 2)
   s3.Caption = Left(StrText(2), 2)
   s4.Caption = Left(StrText(3), 2)
   s5.Caption = Left(StrText(4), 2)
   s6.Caption = Left(StrText(5), 2)
   StrStart(0) = 40 + StrDelta(0)
   StrStart(1) = 45 + StrDelta(1)
   StrStart(2) = 50 + StrDelta(2)
   StrStart(3) = 55 + StrDelta(3)
   StrStart(4) = 59 + StrDelta(4)
   StrStart(5) = 64 + StrDelta(5)
   For i = 0 To 5
      If StrFrets(i) > -1 Then
         FillInNotes StrFrets(i), i
      End If
   Next
End Sub

Private Sub cmdChordPlay_Click()
   Dim i As Integer
   all_sounds_off
   For i = 0 To NoteCount - 1
      Call note_on(1, 60 + ChordNs(i), 127)
      Call note_on(11, 60 + ChordNs(i), 127)
   Next
   Timer1.Enabled = True
   While Timer1.Enabled = True
      DoEvents
   Wend
   all_sounds_off
End Sub

Private Sub cmdChordPlayNeck_Click()
   Dim i As Integer
   all_sounds_off
   For i = 0 To 5
      If StrFrets(i) > -1 Then
         Call note_on(1, StrStart(i) + StrFrets(i), 127)
      End If
   Next
   Timer1.Enabled = True
   While Timer1.Enabled = True
      DoEvents
   Wend
   all_sounds_off
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
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.PaintPicture Picture1.Picture, 200, 200
      Printer.FontSize = 12
      Printer.FontBold = False
      Printer.Print txtChordList.Text
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.FontSize = 8
      Printer.FontBold = False
      Printer.Print " Guitar Master Pro - Copyright 2000 Opus Software"
      Printer.EndDoc
   Next i
   Exit Sub
PrintErr:
   MsgBox "Chord was not printed", vbExclamation + vbOKOnly, "Print"
   Exit Sub
CancelErr:
   
End Sub
Private Sub CaptureNeck()
   Picture1.Picture = CaptureActiveWindow(((frmChordFind.Left + 400) / Screen.TwipsPerPixelX), ((frmChordFind.Top + frmNeck.Top + 1950) / Screen.TwipsPerPixelY), (frmNeck.Width + 600) / Screen.TwipsPerPixelX, frmNeck.Height / Screen.TwipsPerPixelY)
End Sub

Private Sub cmdPrint_Click()
    CaptureNeck
    PrintData
End Sub

Private Sub Form_Activate()
   DisableMenu
End Sub

Private Sub cmdClear_Click()
   Dim i As Integer
   
   For i = 0 To 6
      txtNote(i).Text = ""
   Next
End Sub

Private Sub cmdFind_Click()
   Dim i As Integer
   Dim j As Integer
   Dim temp As String
   
   frmChordFind.MousePointer = 11
   Suffix = ""
   Root = txtNote(0).Text
   NoteCount = 0
   For i = 0 To 6
      If txtNote(i).Text <> "" Then
         NoteTxt = Trim(UCase(txtNote(i).Text))
         ChordNs(i) = GetNoteNumber(NoteTxt)
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
      
   dtaChords.DatabaseName = "Chords.mdb"
   dtaChords.Refresh
   temp = "SELECT * FROM ChordType WHERE [NumNotes] = " + Str(NoteCount) + " AND [1] = " + Str(ChordNote(0))
   For i = 1 To NoteCount - 1
      temp = temp + " AND [" + Trim(Str(i + 1)) + "] = " + Str(ChordNote(i))
   Next
   On Error GoTo err:
   dtaChords.RecordSource = temp
   dtaChords.Refresh
   
   Suffix = dtaChords.Recordset.Fields(0).value
   lblChord.Caption = Trim(Root) + " " + Trim(Suffix)
   If NumDevices > 0 Then
      cmdChordPlay.Enabled = True
   End If
   frmChordFind.MousePointer = 0
   Exit Sub
err:
   MsgBox "Chord could not be found", vbOKOnly + vbExclamation, "Error"
   lblChord.Caption = ""
   frmChordFind.MousePointer = 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   cmdChordPlay.Enabled = False
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
   For i = 0 To 5
      StrFrets(i) = -1
   Next
   cboTuning.Text = "Standard"
   cmdChange_Click
   If Registered = 0 Then
      cboTuning.Enabled = False
      cmdChange.Enabled = False
   End If
   
End Sub

Private Sub Form_LostFocus()
   DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
   all_sounds_off
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmChordFind.MousePointer = 11
   all_sounds_off
   If Button = vbLeftButton Then
      If Index = 1 Then
         X = X + Image1(0).Width
      ElseIf Index = 2 Then
         X = X + Image1(0).Width + Image1(1).Width
      End If
      FretNum = GetFretNum(X)
      St = GetStringNumRev(Y)
      If FretNum > 22 Then
         Exit Sub
      End If
      MarkFret
      FillInNotes FretNum, St
      FindChordList
   End If
   frmChordFind.MousePointer = 0
End Sub

Private Sub FindChordList()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   Dim IsIn As Boolean
   Dim tempnote As Integer
   Dim temp As String
   Dim tempcount As Integer
   Dim Continued As Boolean
   
   IsIn = False
   NoteCount = 0
   For i = 0 To 8
      ChordNs(i) = 0
   Next
   For i = 0 To 5
      If StrFrets(i) > -1 Then
         IsIn = False
         tempnote = ((StrStart(i) + StrFrets(i)) Mod 12) + 1
         For j = 0 To NoteCount - 1
            If ChordNs(j) = tempnote Then
               IsIn = True
            End If
         Next
         If Not IsIn Then
            NoteCount = NoteCount + 1
            ChordNs(NoteCount - 1) = tempnote
         End If
      End If
   Next
   txtChordList.Text = ""
   If NoteCount > 1 Then
      'do permutation
      TotalChar = 0
      NextBlock = 0
      
      For i = 0 To NoteCount - 1
         char(i) = Chr(65 + ChordNs(i))
      Next
      NextBlock = NoteCount

      TotalChar = NextBlock - 1
      Initialize
      GenOutput (NoteCount)
         
      NumList = pcount
      ReDim Preserve NoteList(1 To NumList) As NotesType
      For i = 1 To NumList
         CnvToNum Results(i), NoteList(i)
      Next
      
'actual code
      On Error Resume Next
      dtaChords.RecordSource = "SELECT * FROM ChordType WHERE [NumNotes] = " + Str(NoteCount)
      dtaChords.Refresh
      dtaChords.Recordset.MoveFirst
      dtaChords.Recordset.MoveLast
      dtaChords.Recordset.MoveFirst

      For k = 1 To NumList
         Suffix = ""
         Root = GetNoteText(NoteList(k).Notes(0))
         For j = 1 To NoteCount - 1
            For i = 1 To NoteCount - 1
               If NoteList(k).Notes(i) < NoteList(k).Notes(i - 1) Then
                  NoteList(k).Notes(i) = NoteList(k).Notes(i) + 12
               End If
            Next
         Next
         For i = 1 To NoteCount - 1
           ChordNote(i) = NoteList(k).Notes(i) - NoteList(k).Notes(0)
         Next
         dtaChords.Recordset.MoveFirst
         For i = 1 To dtaChords.Recordset.RecordCount
            Select Case NoteCount
            Case 3
               If ((dtaChords.Recordset.Fields(2).value = ChordNote(0)) And (dtaChords.Recordset.Fields(3).value = ChordNote(1)) And (dtaChords.Recordset.Fields(4).value = ChordNote(2))) Then
                  Suffix = dtaChords.Recordset.Fields(0).value
                  If Suffix <> "" Then
                     txtChordList.Text = txtChordList.Text + Trim(Root) + " " + Trim(Suffix) + Chr(13) + Chr(10)
                  End If
               End If
            Case 4
               If ((dtaChords.Recordset.Fields(2).value = ChordNote(0)) And (dtaChords.Recordset.Fields(3).value = ChordNote(1)) And (dtaChords.Recordset.Fields(4).value = ChordNote(2)) And (dtaChords.Recordset.Fields(5).value = ChordNote(3))) Then
                  Suffix = dtaChords.Recordset.Fields(0).value
                  If Suffix <> "" Then
                     txtChordList.Text = txtChordList.Text + Trim(Root) + " " + Trim(Suffix) + Chr(13) + Chr(10)
                  End If
               End If
            Case 5
               If ((dtaChords.Recordset.Fields(2).value = ChordNote(0)) And (dtaChords.Recordset.Fields(3).value = ChordNote(1)) And (dtaChords.Recordset.Fields(4).value = ChordNote(2)) And (dtaChords.Recordset.Fields(5).value = ChordNote(3)) And (dtaChords.Recordset.Fields(6).value = ChordNote(4))) Then
                  Suffix = dtaChords.Recordset.Fields(0).value
                  If Suffix <> "" Then
                     txtChordList.Text = txtChordList.Text + Trim(Root) + " " + Trim(Suffix) + Chr(13) + Chr(10)
                  End If
               End If
            Case 6
               If ((dtaChords.Recordset.Fields(2).value = ChordNote(0)) And (dtaChords.Recordset.Fields(3).value = ChordNote(1)) And (dtaChords.Recordset.Fields(4).value = ChordNote(2)) And (dtaChords.Recordset.Fields(5).value = ChordNote(3)) And (dtaChords.Recordset.Fields(6).value = ChordNote(4)) And (dtaChords.Recordset.Fields(7).value = ChordNote(5))) Then
                  Suffix = dtaChords.Recordset.Fields(0).value
                  If Suffix <> "" Then
                     txtChordList.Text = txtChordList.Text + Trim(Root) + " " + Trim(Suffix) + Chr(13) + Chr(10)
                  End If
               End If
            End Select
            dtaChords.Recordset.MoveNext
         Next
      Next

 
   End If
End Sub

Private Sub CnvToNum(ListStr As String, ListNum As NotesType)
   Dim i As Integer
   Dim Ln As Integer
   Ln = Len(ListStr)
   For i = 1 To Ln
      ListNum.Notes(i - 1) = Asc(Mid(ListStr, i, 1)) - 65
   Next
End Sub
Private Sub MarkFret()
   Dim i As Integer
   
   Select Case St
   Case 5
      If E2(FretNum).Visible = True Then
         E2(FretNum).Visible = False
         lblDead(5).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            E2(i).Visible = False
         Next
         E2(FretNum).Visible = True
         lblDead(5).Visible = False
         StrFrets(St) = FretNum
      End If
   Case 4
      If B1(FretNum).Visible = True Then
         B1(FretNum).Visible = False
         lblDead(4).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            B1(i).Visible = False
         Next
         B1(FretNum).Visible = True
         lblDead(4).Visible = False
         StrFrets(St) = FretNum
      End If
   Case 3
      If G1(FretNum).Visible = True Then
         G1(FretNum).Visible = False
         lblDead(3).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            G1(i).Visible = False
         Next
         G1(FretNum).Visible = True
         lblDead(3).Visible = False
         StrFrets(St) = FretNum
      End If
   Case 2
      If D1(FretNum).Visible = True Then
         D1(FretNum).Visible = False
         lblDead(2).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            D1(i).Visible = False
         Next
         D1(FretNum).Visible = True
         lblDead(2).Visible = False
         StrFrets(St) = FretNum
      End If
   Case 1
      If A1(FretNum).Visible = True Then
         A1(FretNum).Visible = False
         lblDead(1).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            A1(i).Visible = False
         Next
         A1(FretNum).Visible = True
         lblDead(1).Visible = False
         StrFrets(St) = FretNum
      End If
   Case 0
      If E1(FretNum).Visible = True Then
         E1(FretNum).Visible = False
         lblDead(0).Visible = True
         lblStrNote(St).Caption = ""
         lblZeroNote(St).Caption = ""
         StrFrets(St) = -1
      Else
         For i = 0 To 22
            E1(i).Visible = False
         Next
         E1(FretNum).Visible = True
         lblDead(0).Visible = False
         StrFrets(St) = FretNum
      End If
   End Select
End Sub

Private Sub lblStrNote_Click(Index As Integer)
   frmChordFind.MousePointer = 11
   St = Index
   FretNum = GetFretNum(lblStrNote(Index).Left)
   If lblStrNote(Index).Visible = True Then
      lblStrNote(Index).Visible = False
      lblZeroNote(Index).Caption = ""
      MarkFret
      FindChordList
   Else
      FretNum = GetFretNum(lblStrNote(Index).Left)
      lblStrNote(Index).Visible = True
      lblZeroNote(Index).Caption = ""
      MarkFret
      FindChordList
      FillInNotes FretNum, St
   End If
   frmChordFind.MousePointer = 0
End Sub

Private Sub lblZeroNote_Click(Index As Integer)
   frmChordFind.MousePointer = 11
   Dim i As Integer
   St = Index
   FretNum = 0
   For i = 1 To FretNum
      If Index = 0 Then
         E1(i).Visible = False
      ElseIf Index = 1 Then
         A1(i).Visible = False
      ElseIf Index = 2 Then
         D1(i).Visible = False
      ElseIf Index = 3 Then
         G1(i).Visible = False
      ElseIf Index = 4 Then
         B1(i).Visible = False
      ElseIf Index = 5 Then
         E2(i).Visible = False
      End If
   Next
   lblStrNote(Index).Visible = False
   MarkFret
   If lblZeroNote(Index).Caption = "" Then
      lblZeroNote(Index).Visible = True
      FillInNotes FretNum, Index
   Else
      lblZeroNote(Index).Caption = ""
   End If
   FindChordList
   frmChordFind.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
End Sub

Private Sub FillInNotes(FretPos As Integer, s As Integer)
   Dim i As Integer
   lblStrNote(s).Caption = ""
   lblZeroNote(s).Caption = ""
   
   If FretPos > -1 Then
      lblStrNote(s).Visible = True
      lblZeroNote(s).Visible = True
      Select Case s
      Case 5
         lblStrNote(s).Left = E2(FretPos).Left
      Case 4
         lblStrNote(s).Left = B1(FretPos).Left
      Case 3
         lblStrNote(s).Left = G1(FretPos).Left
      Case 2
         lblStrNote(s).Left = D1(FretPos).Left
      Case 1
         lblStrNote(s).Left = A1(FretPos).Left
      Case 0
         lblStrNote(s).Left = E1(FretPos).Left
      End Select
      If FretNote <> 2 Then
         If FretPos <> 0 Then
            lblStrNote(s).Caption = Left(GetNoteText(((StrStart(s) + FretPos) Mod 12) + 1), 2)
         Else
            lblZeroNote(s).Caption = Left(GetNoteText((StrStart(s) Mod 12) + 1), 2)
         End If
      Else
         lblStrNote(s).Visible = False
         lblZeroNote(s).Visible = False
      End If
   End If
End Sub
