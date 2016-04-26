VERSION 5.00
Begin VB.Form frmPrintWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please wait"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Printing..."
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrintWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
