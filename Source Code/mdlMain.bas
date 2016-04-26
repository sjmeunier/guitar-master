Attribute VB_Name = "mdlMain"
Option Explicit
'MIDI Channels
'0 - Frequency Finer
'1 - Chord FInder
'2- Chord Charts
'3- Note Finer
'4- Transposer
'5- Tuner
'10-Metronome

Public Cancelled As Boolean
Public PCSound As Boolean

Public CurPos
Public MaxPos

Public FName As String

Public DefMapper As Integer
Public DefDevice As Integer
Public DefInstrument As Integer
Public NumDevices As Integer
Public IntervalSkill As Integer
Public ChordSkill As Integer
Public PitchSkill As Integer
Public ModeSkill As Integer
Public ShowNeck As Integer
Public Volume As Long
Public FretboardCol As Integer
Public FingerStretch As Integer
Public ChordCalc As Integer
Public Registered As Integer
Public FretNote As Integer
Public Const MaxFret = 22
Public DaysLeft As Integer
Public Tuning As String
Public LongSuffix As String
Public ShortSuffix As String
Public CurTab As Integer
Public StrDelta(0 To 5) As Integer
Public StrText(0 To 5) As String

Type NotesType
   Notes(0 To 8) As Integer
   NumNotes As Integer
End Type
