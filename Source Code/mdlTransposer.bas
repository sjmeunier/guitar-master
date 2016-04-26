Attribute VB_Name = "mdlTransposer"
Option Explicit

Public TScale As String
Public OrigKey As String
Public ToKey As String
Public Chord(0 To 11) As String
Public OrigNote(0 To 11) As String
Public ToNote(0 To 11) As String
Public NumNotes As Integer

Public OrigNoteNum(0 To 11) As Integer
Public ToNoteNum(0 To 11) As Integer
Public OrigKeyNum As Integer
Public ToKeyNum As Integer
Public Diff As Integer
Public HaveChords As Boolean

