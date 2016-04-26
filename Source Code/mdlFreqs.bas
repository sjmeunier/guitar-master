Attribute VB_Name = "mdlFreqs"
Option Explicit

Public Freqs(0 To 9, 0 To 11) As Single
Public NoteName(0 To 11) As String

Public Sub FillFreqs()
   Dim i As Integer
   
   For i = 0 To 9
      Freqs(i, 0) = 16.35 * (2 ^ i)
      Freqs(i, 1) = 17.3 * (2 ^ i)
      Freqs(i, 2) = 18.35 * (2 ^ i)
      Freqs(i, 3) = 19.5 * (2 ^ i)
      Freqs(i, 4) = 20.6 * (2 ^ i)
      Freqs(i, 5) = 21.8 * (2 ^ i)
      Freqs(i, 6) = 23.1 * (2 ^ i)
      Freqs(i, 7) = 24.5 * (2 ^ i)
      Freqs(i, 8) = 26 * (2 ^ i)
      Freqs(i, 9) = 27.5 * (2 ^ i)
      Freqs(i, 10) = 29.1 * (2 ^ i)
      Freqs(i, 11) = 30.9 * (2 ^ i)
   Next
End Sub
Public Sub FillNames()

   NoteName(0) = "C"
   NoteName(1) = "C#/Db"
   NoteName(2) = "D"
   NoteName(3) = "D#/Eb"
   NoteName(4) = "E"
   NoteName(5) = "F"
   NoteName(6) = "F#/Gb"
   NoteName(7) = "G"
   NoteName(8) = "G#/Ab"
   NoteName(9) = "A"
   NoteName(10) = "A#/Bb"
   NoteName(11) = "B"
   
End Sub
