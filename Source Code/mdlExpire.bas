Attribute VB_Name = "mdlExpire"
Option Explicit

Public Function ExpireCheck() As Boolean
   Dim FileNum As Integer
   Dim Loaded As Integer
   Dim ThisDay As Integer
   Dim FirstDay As Integer
   Dim ThisMonth As Integer
   Dim FirstMonth As Integer
   
   On Error Resume Next
   FileNum = FreeFile
   Open "chords.dat" For Input As FileNum
   
   Input #FileNum, Loaded
   
   If Loaded = 1 Then
      Input #FileNum, FirstDay
      Input #FileNum, FirstMonth
      ThisDay = Day(Now)
      ThisMonth = Month(Now)
      
      Select Case FirstMonth
         Case 1, 3, 5, 7, 8, 10, 12
            If ThisMonth = (FirstMonth + 1) Then
               If (ThisDay + (31 - FirstDay)) > 30 Then
                  
                  GoTo TooLate
               End If
            ElseIf ThisMonth = FirstMonth Then
               If (ThisDay - FirstDay) > 30 Then
                  GoTo TooLate
               End If
            Else
               GoTo TooLate
            End If
         Case 4, 6, 9, 11
            If ThisMonth = (FirstMonth + 1) Then
               If (ThisDay + (31 - FirstDay)) > 30 Then
                  GoTo TooLate
               End If
            ElseIf ThisMonth = FirstMonth Then
               If (ThisDay - FirstDay) > 30 Then
                  GoTo TooLate
               End If
            Else
               GoTo TooLate
            End If
         Case 2
            If ThisMonth = (FirstMonth + 1) Then
               If (ThisDay + (31 - FirstDay)) > 30 Then
                  GoTo TooLate
               End If
            ElseIf ThisMonth = FirstMonth Then
               If (ThisDay - FirstDay) > 30 Then
                  GoTo TooLate
               End If
            Else
               GoTo TooLate
            End If
      End Select
      DaysLeft = 29
   Else
      Close FileNum
      Open "chords.dat" For Output As FileNum
      Loaded = 1
      Print #FileNum, Loaded
      FirstDay = Day(Now)
      FirstMonth = Month(Now)
      Print #FileNum, FirstDay
      Print #FileNum, FirstMonth
      DaysLeft = 30
   End If
   
   Close FileNum
   ExpireCheck = False
   Exit Function
TooLate:
   ExpireCheck = True
   Exit Function
err:
   ExpireCheck = True
   Exit Function

End Function
