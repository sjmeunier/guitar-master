VERSION 5.00
Begin VB.Form frmRegCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Registration Code"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "frmRegCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the 10 digit registration code which was sent to you"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmRegCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
   Unload Me
End Sub

Private Sub cmdEnter_Click()
    Dim FileNum
    If txtRegCode.Text = "E63FG739AC" Or txtRegCode.Text = "CC9346TR2Y" Or txtRegCode.Text = "4X94TMN518" Or txtRegCode.Text = "3RY573PDQ6" Or txtRegCode.Text = "G3T61R573H" Then
        Registered = 1
        On Error GoTo err:
        FileNum = FreeFile
        Open "config.dat" For Binary As FileNum
        Put #FileNum, 23, Registered
        Close FileNum
        MsgBox "Registration code has been accepted", vbOKOnly + vbExclamation, "Invalid"
        frmUnRegSplash.Counter = 1
        Unload Me
    Else
        MsgBox "Invalid registration code entered", vbOKOnly + vbExclamation, "Invalid"
    End If
   Exit Sub
err:
   MsgBox "Unabled to write to the configuration file", vbOKOnly + vbExclamation, "Error"

End Sub

