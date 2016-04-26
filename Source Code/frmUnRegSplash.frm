VERSION 5.00
Begin VB.Form frmUnRegSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      Begin VB.CommandButton cmdEnterCode 
         Caption         =   "&Enter Registration Code"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblDaysLeft 
         BackStyle       =   0  'Transparent
         Caption         =   "Visit www.opussoftware.co.za to register"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   5
         Top             =   705
         Width           =   105
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 95/98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   4
         Top             =   2400
         Width           =   2280
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Opus Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   3030
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   2820
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   1080
         Picture         =   "frmUnRegSplash.frx":0000
         Top             =   1200
         Width           =   2400
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   1560
         Picture         =   "frmUnRegSplash.frx":23C9
         Top             =   360
         Width           =   4140
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   3120
   End
End
Attribute VB_Name = "frmUnRegSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Counter As Integer

Private Sub cmdEnterCode_Click()
   Timer1.Enabled = False
   frmRegCode.Show 1
   Timer1.Enabled = True
End Sub

Private Sub Form_Load()
   Counter = 10
  ' ExpireCheck
End Sub


Private Sub Timer1_Timer()
   If Counter = 0 Then
     Unload Me
   End If
   Counter = Counter - 1
End Sub

