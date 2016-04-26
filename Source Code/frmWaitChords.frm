VERSION 5.00
Begin VB.Form frmWaitChords 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please wait..."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Calculating Chords"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmWaitChords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
