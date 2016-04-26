Attribute VB_Name = "mdlMainSub"
Option Explicit

Public Sub DisableMenu()
   frmMain.mnuNew.Visible = False
   frmMain.mnuSave.Visible = False
   frmMain.mnuSaveAs.Visible = False
   frmMain.mnuSep4.Visible = False
   frmMain.mnuSep5.Visible = False
   frmMain.mnuOpen.Visible = False
   frmMain.mnuPrint.Visible = False
   frmMain.mnuPrintTab.Visible = False
   frmMain.mnuEdit.Visible = False
End Sub

Public Function FindChordSuffix(CNote() As Integer, Notes As Integer) As String
   '3-note chords
   If Notes = 3 Then
      'major
      If CNote(1) = 4 And CNote(2) = 7 Then
         FindChordSuffix = "maj"
         Exit Function
      End If
      'minor
      If CNote(1) = 3 And CNote(2) = 7 Then
         FindChordSuffix = "m"
         Exit Function
      End If
      'sus4
      If CNote(1) = 5 And CNote(2) = 7 Then
         FindChordSuffix = "sus4"
         Exit Function
      End If
      'sus2
      If CNote(1) = 2 And CNote(2) = 7 Then
         FindChordSuffix = "sus2"
         Exit Function
      End If
      'maj7
      If CNote(1) = 7 And CNote(2) = 11 Then
         FindChordSuffix = "maj7"
         Exit Function
      End If
      If CNote(1) = 4 And CNote(2) = 11 Then
         FindChordSuffix = "maj7"
         Exit Function
      End If
      'maj6
      If CNote(1) = 7 And CNote(2) = 9 Then
         FindChordSuffix = "maj7"
         Exit Function
      End If
      If CNote(1) = 4 And CNote(2) = 9 Then
         FindChordSuffix = "maj7"
         Exit Function
      End If
      'augmented
      If CNote(1) = 4 And CNote(2) = 8 Then
         FindChordSuffix = "aug"
         Exit Function
      End If
      'diminished
      If CNote(1) = 4 And CNote(2) = 6 Then
         FindChordSuffix = "dim"
         Exit Function
      End If
   End If
   
   '4 note chords
   If Notes = 4 Then
      'major 7th
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 11 Then
         FindChordSuffix = "maj7"
         Exit Function
      End If
      'dominant 7th
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 10 Then
         FindChordSuffix = "dom7"
         Exit Function
      End If
      'minor 7th
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 10 Then
         FindChordSuffix = "min7"
         Exit Function
      End If
      'minor 7th flat 5th
      If CNote(1) = 3 And CNote(2) = 6 And CNote(3) = 10 Then
         FindChordSuffix = "min7b5"
         Exit Function
      End If
      'diminished 7th
      If CNote(1) = 3 And CNote(2) = 6 And CNote(3) = 9 Then
         FindChordSuffix = "dim7"
         Exit Function
      End If
      'dominant 7th flat 5th
      If CNote(1) = 4 And CNote(2) = 6 And CNote(3) = 10 Then
         FindChordSuffix = "dom7b5"
         Exit Function
      End If
      'dominat 7th sharp 5th
      If CNote(1) = 4 And CNote(2) = 8 And CNote(3) = 10 Then
         FindChordSuffix = "dom7#5"
         Exit Function
      End If
      'minor/major 7th
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 11 Then
         FindChordSuffix = "min/maj7"
         Exit Function
      End If
      'major 6th
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 9 Then
         FindChordSuffix = "maj6"
         Exit Function
      End If
      'minor 6th
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 9 Then
         FindChordSuffix = "min6"
         Exit Function
      End If
      'major 9th
      If CNote(1) = 4 And CNote(2) = 11 And CNote(3) = 14 Then
         FindChordSuffix = "maj9"
         Exit Function
      End If
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 14 Then
         FindChordSuffix = "maj9"
         Exit Function
      End If
      'dominant 9th
      If CNote(1) = 4 And CNote(2) = 10 And CNote(3) = 14 Then
         FindChordSuffix = "dom9"
         Exit Function
      End If
      'minor 9th
      If CNote(1) = 3 And CNote(2) = 10 And CNote(3) = 14 Then
         FindChordSuffix = "min9"
         Exit Function
      End If
      'dominant 7th flat 9
      If CNote(1) = 4 And CNote(2) = 10 And CNote(3) = 13 Then
         FindChordSuffix = "dom7b9"
         Exit Function
      End If
      'dominant 7th sharp 9
      If CNote(1) = 4 And CNote(2) = 10 And CNote(3) = 15 Then
         FindChordSuffix = "dom7#9"
         Exit Function
      End If
      'minor flat 6th
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 8 Then
         FindChordSuffix = "minb6"
         Exit Function
      End If
      'add 9
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 14 Then
         FindChordSuffix = "add9"
         Exit Function
      End If
      'min/maj 9th
      If CNote(1) = 3 And CNote(2) = 11 And CNote(3) = 14 Then
         FindChordSuffix = "min/maj9"
         Exit Function
      End If
      'minor 13th
      If CNote(1) = 3 And CNote(2) = 10 And CNote(3) = 21 Then
         FindChordSuffix = "min13"
         Exit Function
      End If
   End If
   
   '5-note chords
   If Notes = 5 Then
      'dominant 11th
      If CNote(1) = 4 And CNote(2) = 5 And CNote(3) = 7 And CNote(4) = 10 Then
         FindChordSuffix = "dom11"
         Exit Function
      End If
      If CNote(1) = 4 And CNote(2) = 7 And CNote(3) = 10 And CNote(4) = 17 Then
         FindChordSuffix = "dom11"
         Exit Function
      End If
      'minor 11th
      If CNote(1) = 3 And CNote(2) = 5 And CNote(3) = 7 And CNote(4) = 10 Then
         FindChordSuffix = "min11"
         Exit Function
      End If
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 10 And CNote(4) = 17 Then
         FindChordSuffix = "min11"
         Exit Function
      End If
      'dominant 13th
      If CNote(1) = 4 And CNote(2) = 10 And CNote(3) = 14 And CNote(4) = 21 Then
         FindChordSuffix = "dom13"
         Exit Function
      End If
      'dominant 9 flat 5
      If CNote(1) = 4 And CNote(2) = 6 And CNote(3) = 10 And CNote(4) = 14 Then
         FindChordSuffix = "9b5"
         Exit Function
      End If
      'min/maj 9th
      If CNote(1) = 3 And CNote(2) = 7 And CNote(3) = 11 And CNote(4) = 14 Then
         FindChordSuffix = "min/maj9"
         Exit Function
      End If
   End If
   FindChordSuffix = ""
End Function

Public Function GetNoteNumber(N As String) As Integer
   N = Trim(UCase(N))
   If Len(N) > 2 Then
      N = Left(N, 2)
   End If
   If N = "C" Then
      GetNoteNumber = 1
      Exit Function
   ElseIf N = "C#" Or N = "DB" Then
      GetNoteNumber = 2
      Exit Function
   ElseIf N = "D" Then
      GetNoteNumber = 3
      Exit Function
   ElseIf N = "D#" Or N = "EB" Then
      GetNoteNumber = 4
      Exit Function
   ElseIf N = "E" Then
      GetNoteNumber = 5
      Exit Function
   ElseIf N = "F" Then
      GetNoteNumber = 6
      Exit Function
   ElseIf N = "F#" Or N = "GB" Then
      GetNoteNumber = 7
      Exit Function
   ElseIf N = "G" Then
      GetNoteNumber = 8
      Exit Function
   ElseIf N = "G#" Or N = "AB" Then
      GetNoteNumber = 9
      Exit Function
   ElseIf N = "A" Then
      GetNoteNumber = 10
      Exit Function
   ElseIf N = "A#" Or N = "BB" Then
      GetNoteNumber = 11
      Exit Function
   ElseIf N = "B" Then
      GetNoteNumber = 12
      Exit Function
   End If
   GetNoteNumber = 0
End Function
Public Function GetNoteText(N As Integer) As String
   Select Case N
   Case 1
      GetNoteText = "C"
      Exit Function
   Case 2
      GetNoteText = "C#/Db"
      Exit Function
   Case 3
      GetNoteText = "D"
      Exit Function
   Case 4
      GetNoteText = "D#/Eb"
      Exit Function
   Case 5
      GetNoteText = "E"
      Exit Function
   Case 6
      GetNoteText = "F"
      Exit Function
   Case 7
      GetNoteText = "F#/Gb"
      Exit Function
   Case 8
      GetNoteText = "G"
      Exit Function
   Case 9
      GetNoteText = "G#/Ab"
      Exit Function
   Case 10
      GetNoteText = "A"
      Exit Function
   Case 11
      GetNoteText = "A#/Bb"
      Exit Function
   Case 12
      GetNoteText = "B"
      Exit Function
   End Select
End Function
Public Function GetIntervalText(N As Integer) As String
   Select Case N
   Case 0
      GetIntervalText = "1"
      Exit Function
   Case 1
      GetIntervalText = "b2"
      Exit Function
   Case 2
      GetIntervalText = "2"
      Exit Function
   Case 3
      GetIntervalText = "b3"
      Exit Function
   Case 4
      GetIntervalText = "3"
      Exit Function
   Case 5
      GetIntervalText = "4"
      Exit Function
   Case 6
      GetIntervalText = "4+"
      Exit Function
   Case 7
      GetIntervalText = "5"
      Exit Function
   Case 8
      GetIntervalText = "b6"
      Exit Function
   Case 9
      GetIntervalText = "6"
      Exit Function
   Case 10
      GetIntervalText = "b7"
      Exit Function
   Case 11
      GetIntervalText = "7"
      Exit Function
   Case 12
      GetIntervalText = "8"
      Exit Function
   Case 13
      GetIntervalText = "b9"
      Exit Function
   Case 14
      GetIntervalText = "9"
      Exit Function
   Case 15
      GetIntervalText = "b10"
      Exit Function
   Case 16
      GetIntervalText = "10"
      Exit Function
   Case 17
      GetIntervalText = "b11"
      Exit Function
   Case 18
      GetIntervalText = "11"
      Exit Function
   Case 19
      GetIntervalText = "b12"
      Exit Function
   Case 20
      GetIntervalText = "12"
      Exit Function
   Case 21
      GetIntervalText = "b13"
      Exit Function
   Case 22
      GetIntervalText = "13"
      Exit Function
   End Select
   GetIntervalText = ""
End Function
Public Function LongChord(SSuffix As String) As String
   SSuffix = Trim(SSuffix)
   If SSuffix = "maj" Then
      LongChord = "Major"
      Exit Function
   ElseIf SSuffix = "m" Then
      LongChord = "Minor"
      Exit Function
   ElseIf SSuffix = "aug" Then
      LongChord = "Augmented"
      Exit Function
   ElseIf SSuffix = "aug7" Then
      LongChord = "Augmented 7th"
      Exit Function
   ElseIf SSuffix = "aug11" Then
      LongChord = "Augmented 11th"
      Exit Function
   ElseIf SSuffix = "dim" Then
      LongChord = "Diminished"
      Exit Function
   ElseIf SSuffix = "dim7" Then
      LongChord = "Diminished 7th"
      Exit Function
   ElseIf SSuffix = "7" Then
      LongChord = "Dominant 7th"
      Exit Function
   ElseIf SSuffix = "7sus2" Then
      LongChord = "7th Suspended 2nd"
      Exit Function
   ElseIf SSuffix = "7sus4" Then
      LongChord = "7th Suspended 4th"
      Exit Function
   ElseIf SSuffix = "7b5" Then
      LongChord = "Dominant 7th b5th"
      Exit Function
   ElseIf SSuffix = "7b9" Then
      LongChord = "Dominant 7th b9th"
      Exit Function
   ElseIf SSuffix = "7#5" Then
      LongChord = "Dominant 7th #5th"
      Exit Function
   ElseIf SSuffix = "7#5b9" Then
      LongChord = "Dominant 7th #5th b9th"
      Exit Function
   ElseIf SSuffix = "7#9" Then
      LongChord = "Dominant 7th #9th"
      Exit Function
   ElseIf SSuffix = "9" Then
      LongChord = "Dominant 9th"
      Exit Function
   ElseIf SSuffix = "9b5" Then
      LongChord = "Dominant 9th b5th"
      Exit Function
   ElseIf SSuffix = "9#5" Then
      LongChord = "Dominant 9th #5th"
      Exit Function
   ElseIf SSuffix = "11" Then
      LongChord = "Dominant 11th"
      Exit Function
   ElseIf SSuffix = "13" Then
      LongChord = "Dominant 13th"
      Exit Function
   ElseIf SSuffix = "maj 6/9" Then
      LongChord = "Major 6/9"
      Exit Function
   ElseIf SSuffix = "maj 6/9 #11" Then
      LongChord = "Major 6/9 #11th"
      Exit Function
   ElseIf SSuffix = "6" Then
      LongChord = "Major 6th"
      Exit Function
   ElseIf SSuffix = "6add9" Then
      LongChord = "Major 6th Added 9th"
      Exit Function
   ElseIf SSuffix = "maj7" Then
      LongChord = "Major 7th"
      Exit Function
   ElseIf SSuffix = "maj7#5" Then
      LongChord = "Major 7th #5th"
      Exit Function
   ElseIf SSuffix = "maj7#11" Then
      LongChord = "Major 7th #11th"
      Exit Function
   ElseIf SSuffix = "maj9" Then
      LongChord = "Major 9th"
      Exit Function
   ElseIf SSuffix = "add9" Then
      LongChord = "Major Added 9th"
      Exit Function
   ElseIf SSuffix = "m6/9" Then
      LongChord = "Minor 6/9"
      Exit Function
   ElseIf SSuffix = "mb6" Then
      LongChord = "Minor b6th"
      Exit Function
   ElseIf SSuffix = "m6" Then
      LongChord = "Minor 6th"
      Exit Function
   ElseIf SSuffix = "min/maj 7" Then
      LongChord = "Minor/Major 7th"
      Exit Function
   ElseIf SSuffix = "m7" Then
      LongChord = "Minor 7th"
      Exit Function
   ElseIf SSuffix = "m7b5" Then
      LongChord = "Minor 7th b5th"
      Exit Function
   ElseIf SSuffix = "m7b3" Then
      LongChord = "Minor 7th b3rd"
      Exit Function
   ElseIf SSuffix = "min/maj 9" Then
      LongChord = "Minor/Major 9th"
      Exit Function
   ElseIf SSuffix = "m9" Then
      LongChord = "Minor 9th"
      Exit Function
   ElseIf SSuffix = "m11" Then
      LongChord = "Minor 11th"
      Exit Function
   ElseIf SSuffix = "m13" Then
      LongChord = "Minor 13th"
      Exit Function
   ElseIf SSuffix = "sus4" Then
      LongChord = "Suspended 4th"
      Exit Function
   ElseIf SSuffix = "sus2" Then
      LongChord = "Suspended 2nd"
      Exit Function
   End If
   LongChord = ""
End Function

Public Function GetFretNum(X As Single) As Single
    If X < BFret1 Then
        GetFretNum = 1
    ElseIf X < BFret2 Then
        GetFretNum = 2
    ElseIf X < BFret3 Then
        GetFretNum = 3
    ElseIf X < BFret4 Then
        GetFretNum = 4
    ElseIf X < BFret5 Then
        GetFretNum = 5
    ElseIf X < BFret6 Then
        GetFretNum = 6
    ElseIf X < BFret7 Then
        GetFretNum = 7
    ElseIf X < BFret8 Then
        GetFretNum = 8
    ElseIf X < BFret9 Then
        GetFretNum = 9
    ElseIf X < BFret10 Then
        GetFretNum = 10
    ElseIf X < BFret11 Then
        GetFretNum = 11
    ElseIf X < BFret12 Then
        GetFretNum = 12
    ElseIf X < BFret13 Then
        GetFretNum = 13
    ElseIf X < BFret14 Then
        GetFretNum = 14
    ElseIf X < BFret15 Then
        GetFretNum = 15
    ElseIf X < BFret16 Then
        GetFretNum = 16
    ElseIf X < BFret17 Then
        GetFretNum = 17
    ElseIf X < BFret18 Then
        GetFretNum = 18
    ElseIf X < BFret19 Then
        GetFretNum = 19
    ElseIf X < BFret20 Then
        GetFretNum = 20
    ElseIf X < BFret21 Then
        GetFretNum = 21
    ElseIf X < BFret22 Then
        GetFretNum = 22
    Else
        GetFretNum = 100
    End If
End Function

Public Function GetStringNum(Y As Single) As Single
    If Y < 250 Then
        GetStringNum = 1
    ElseIf Y < 495 Then
        GetStringNum = 2
    ElseIf Y < 725 Then
        GetStringNum = 3
    ElseIf Y < 970 Then
        GetStringNum = 4
    ElseIf Y < 1210 Then
        GetStringNum = 5
    ElseIf Y < 1450 Then
        GetStringNum = 6
    Else
        GetStringNum = 6
    End If
        
End Function

Public Function GetStringNumRev(Y As Single) As Single
    If Y < 250 Then
        GetStringNumRev = 5
    ElseIf Y < 495 Then
        GetStringNumRev = 4
    ElseIf Y < 725 Then
        GetStringNumRev = 3
    ElseIf Y < 970 Then
        GetStringNumRev = 2
    ElseIf Y < 1210 Then
        GetStringNumRev = 1
    ElseIf Y < 1450 Then
        GetStringNumRev = 0
    Else
        GetStringNumRev = 0
    End If
        
End Function

Public Sub ChangeTune(Tuning As String, SDelta() As Integer, SText() As String)
   Tuning = Trim(Tuning)
   If Tuning = "Standard" Then
       SText(0) = "E"
       SText(1) = "B"
       SText(2) = "G"
       SText(3) = "D"
       SText(4) = "A"
       SText(5) = "E"
       SDelta(0) = 0
       SDelta(1) = 0
       SDelta(2) = 0
       SDelta(3) = 0
       SDelta(4) = 0
       SDelta(5) = 0
   ElseIf Tuning = "Dropped D" Then
       SText(0) = "E"
       SText(1) = "B"
       SText(2) = "G"
       SText(3) = "D"
       SText(4) = "A"
       SText(5) = "D"
       SDelta(0) = -2
       SDelta(1) = 0
       SDelta(2) = 0
       SDelta(3) = 0
       SDelta(4) = 0
       SDelta(5) = 0
   ElseIf Tuning = "Dropped A" Then
       SText(0) = "E"
       SText(1) = "C#/Db"
       SText(2) = "A"
       SText(3) = "E"
       SText(4) = "A"
       SText(5) = "E"
       SDelta(0) = 0
       SDelta(1) = 0
       SDelta(2) = 2
       SDelta(3) = 2
       SDelta(4) = 2
       SDelta(5) = 0
   ElseIf Tuning = "Open C" Then
       SText(0) = "E"
       SText(1) = "C"
       SText(2) = "G"
       SText(3) = "C"
       SText(4) = "G"
       SText(5) = "C"
       SDelta(0) = -4
       SDelta(1) = -2
       SDelta(2) = -2
       SDelta(3) = 0
       SDelta(4) = 1
       SDelta(5) = 0
   ElseIf Tuning = "Open Cm" Then
       SText(0) = "D#/Eb"
       SText(1) = "C"
       SText(2) = "G"
       SText(3) = "C"
       SText(4) = "G"
       SText(5) = "C"
       SDelta(0) = -4
       SDelta(1) = -2
       SDelta(2) = -2
       SDelta(3) = 0
       SDelta(4) = 1
       SDelta(5) = -1
   ElseIf Tuning = "Open D" Then
       SText(0) = "D"
       SText(1) = "A"
       SText(2) = "F#/Gb"
       SText(3) = "D"
       SText(4) = "A"
       SText(5) = "D"
       SDelta(0) = -2
       SDelta(1) = 0
       SDelta(2) = 0
       SDelta(3) = -1
       SDelta(4) = -2
       SDelta(5) = -2
   ElseIf Tuning = "Open Dm" Then
       SText(0) = "D"
       SText(1) = "A"
       SText(2) = "F"
       SText(3) = "D"
       SText(4) = "A"
       SText(5) = "D"
       SDelta(0) = -2
       SDelta(1) = 0
       SDelta(2) = 0
       SDelta(3) = -2
       SDelta(4) = -2
       SDelta(5) = -2
   ElseIf Tuning = "Open Dsus4" Then
       SText(0) = "D"
       SText(1) = "A"
       SText(2) = "G"
       SText(3) = "D"
       SText(4) = "A"
       SText(5) = "D"
       SDelta(0) = -2
       SDelta(1) = 0
       SDelta(2) = 0
       SDelta(3) = 0
       SDelta(4) = -2
       SDelta(5) = -2
   ElseIf Tuning = "Open E" Then
       SText(0) = "E"
       SText(1) = "B"
       SText(2) = "G#/Ab"
       SText(3) = "E"
       SText(4) = "B"
       SText(5) = "E"
       SDelta(0) = 0
       SDelta(1) = 2
       SDelta(2) = 2
       SDelta(3) = 1
       SDelta(4) = 0
       SDelta(5) = 0
   ElseIf Tuning = "Open Em" Then
       SText(0) = "E"
       SText(1) = "B"
       SText(2) = "G"
       SText(3) = "E"
       SText(4) = "B"
       SText(5) = "E"
       SDelta(0) = 0
       SDelta(1) = 2
       SDelta(2) = 2
       SDelta(3) = 0
       SDelta(4) = 0
       SDelta(5) = 0
   ElseIf Tuning = "Open F" Then
       SText(0) = "C"
       SText(1) = "A"
       SText(2) = "F"
       SText(3) = "C"
       SText(4) = "F"
       SText(5) = "C"
       SDelta(0) = -4
       SDelta(1) = 1
       SDelta(2) = 3
       SDelta(3) = -2
       SDelta(4) = -2
       SDelta(5) = -4
   ElseIf Tuning = "Open G" Then
       SText(0) = "D"
       SText(1) = "B"
       SText(2) = "G"
       SText(3) = "D"
       SText(4) = "G"
       SText(5) = "D"
       SDelta(0) = -2
       SDelta(1) = -2
       SDelta(2) = 0
       SDelta(3) = 0
       SDelta(4) = 0
       SDelta(5) = -2
   ElseIf Tuning = "Standard Bass" Then
       SText(0) = "C"
       SText(1) = "G"
       SText(2) = "D"
       SText(3) = "A"
       SText(4) = "E"
       SText(5) = "B"
       SDelta(0) = -19
       SDelta(1) = -19
       SDelta(2) = -19
       SDelta(3) = -19
       SDelta(4) = -20
       SDelta(5) = -20
   ElseIf Tuning = "Dropped D Bass" Then
       SText(0) = "C"
       SText(1) = "G"
       SText(2) = "D"
       SText(3) = "A"
       SText(4) = "D"
       SText(5) = "B"
       SDelta(0) = -19
       SDelta(1) = -21
       SDelta(2) = -19
       SDelta(3) = -19
       SDelta(4) = -20
       SDelta(5) = -20
   End If
End Sub
