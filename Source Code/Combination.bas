Attribute VB_Name = "combination"
'Source code developed by Joseph Ninan
'2nd Year BTech Computer Science and Engineering
'Sree Chitra Thirunal College of Engineering
'Papanamcode, Trivandrum-18
'Affliated to University of Kerala
'Residential Address
'Liju Bhavan, Muttampuram Lane, Sreekariyam PO
'Trivandrum
'Kerala state
'India
'PIN 695017
'Tel No 0091-471-449977
'email josephninan@crosswinds.net   liju_trv@yahoo.com
'Web Site http://www.jofu.8m.com
Public char(260) As String
Public TotalChar, NextBlock, ASCIIStart, Max, Counter As Integer
Public Ca, Cb, Cc, Cd, Ce, Cf, Cg, Ch, Cbi As Integer
Public NoOfComb As Long
Public Results() As String
Public Const TOTALUP = 26
Public Const TOTALLOW = 26
Public Const TOTALDIGITS = 10
Public fact As Double

Public Sub Initialize()
    pcount = 0
    NoOfComb = 0
End Sub

'Now what should i do
'I have got all the input chars in char(Totalchar)
'I have got a fixed length in n=frmmain.txtlength
'First i have to get all the n letter combinations
'Then i have to permute them
'How to do?
'I got the code for doing nC3
'Will try to do the rest later

Public Sub GenOutput(cho As Integer)
Max = TotalChar + 1 - c
StartTime = Time
Select Case cho
Case 1:
    For Ca = 0 To Max - 1
            OpenForms = DoEvents
            NoOfComb = NoOfComb + 1
            comb = char(Ca)
            permute (comb)
            'NoOfComb = NoOfComb + 1
            'Debug.Print comb, NoOfComb
    Next Ca

Case 2:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cbi = 0 To Max - Ca - 2
            NoOfComb = NoOfComb + 1
            comb = char(Cbi) & char(Cbi + Ca + 1)
            permute (comb)
        Next Cbi
    Next Ca
Case 3:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cbi = 0 To Max - Ca - Cb - 3
                OpenForms = DoEvents
                NoOfComb = NoOfComb + 1
                comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2)
                permute (comb)
                'NoOfComb = NoOfComb + 1
                'Debug.Print comb, NoOfComb
            Next Cbi
        Next Cb
    Next Ca
Case 4:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cbi = 0 To Max - Ca - Cb - Cc - 4
                    NoOfComb = NoOfComb + 1
                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3)
                    permute (comb)
                    'NoOfComb = NoOfComb + 1
                    'Debug.Print comb, NoOfComb
                Next Cbi
            Next Cc
        Next Cb
    Next Ca
Case 5:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Cbi = 0 To Max - Ca - Cb - Cc - Cd - 5
                    OpenForms = DoEvents
                    NoOfComb = NoOfComb + 1
                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4)
                    permute (comb)
                        'NoOfComb = NoOfComb + 1
                        'Debug.Print comb, NoOfComb
                    Next Cbi
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 6:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - 6
                            NoOfComb = NoOfComb + 1
                            comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5)
                            permute (comb)
                        Next Cbi
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 7:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - 7
                                OpenForms = DoEvents
                                NoOfComb = NoOfComb + 1
                                comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6)
                                permute (comb)
                            Next Cbi
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 8:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cg = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf
                                For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg - 8
                                    OpenForms = DoEvents
                                    NoOfComb = NoOfComb + 1
                                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + 7)
                                    permute (comb)
                                Next Cbi
                            Next Cg
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 9:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cg = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf
                                For Ch = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg
                                    For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg - Ch - 9
                                        OpenForms = DoEvents
                                        NoOfComb = NoOfComb + 1
                                        comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + 7) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + Ch + 8)
                                        permute (comb)
                                    Next Cbi
                                Next Ch
                            Next Cg
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca

End Select
'frmMain.txtStatus.Text = result
'Debug.Print Len(result) / (cho + 2)
'Debug.Print "Total No of combinations processed"; NoOfComb
'Debug.Print StartTime, "Started"
'Debug.Print Time, "Ended"
'Debug.Print "Over"


End Sub

