Attribute VB_Name = "Permutation"
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
' or planet@jofu.8m.com
'Web Site http://www.jofu.8m.com

Public lett(50) As String
Public perm, word As String
Public L, Pa, Pb, Pc, Pd, Pe, Pf, Pg, Ph, Pi, Pj, Pmi As Integer
Public pcount As Double

Public Sub permute(word As String)
L = Len(word)
For Pmi = 1 To L: lett(Pmi) = Mid(word, Pmi, 1): Next Pmi
Select Case L
Case 1:
    pcount = pcount + 1
    perm = word
    ReDim Preserve Results(1 To pcount) As String
    Results(pcount) = perm
Case 2:
    For Pa = 1 To L
                Pb = 3 - Pa
                pcount = pcount + 1
                perm = lett(Pa) & lett(Pb)
                ReDim Preserve Results(1 To pcount) As String
                Results(pcount) = perm
    Next Pa

Case 3:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                Pc = 6 - (Pa + Pb)
                pcount = pcount + 1
                perm = lett(Pa) & lett(Pb) & lett(Pc)
                ReDim Preserve Results(1 To pcount) As String
                Results(pcount) = perm
            End If
        Next Pb
    Next Pa
    
Case 4:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        Pd = 10 - (Pa + Pb + Pc)
                        pcount = pcount + 1
                        perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd)
                        ReDim Preserve Results(1 To pcount) As String
                        Results(pcount) = perm
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa
Case 5:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        For Pd = 1 To L
                            If Pd <> Pc And Pd <> Pb And Pd <> Pa Then
                                Pe = 15 - (Pa + Pb + Pc + Pd)
                                pcount = pcount + 1
                                perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd) & lett(Pe)
                                ReDim Preserve Results(1 To pcount) As String
                                Results(pcount) = perm
                            End If
                        Next Pd
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa
Case 6:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        For Pd = 1 To L
                            If Pd <> Pc And Pd <> Pb And Pd <> Pa Then
                                For Pe = 1 To L
                                    If Pe <> Pd And Pe <> Pc And Pe <> Pb And Pe <> Pa Then
                                        Pf = 21 - (Pa + Pb + Pc + Pd + Pe)
                                        pcount = pcount + 1
                                        perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd) & lett(Pe) & lett(Pf)
                                        ReDim Preserve Results(1 To pcount) As String
                                        Results(pcount) = perm
                                    End If
                                Next Pe
                            End If
                        Next Pd
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa
Case 7:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        For Pd = 1 To L
                            If Pd <> Pc And Pd <> Pb And Pd <> Pa Then
                                For Pe = 1 To L
                                    If Pe <> Pd And Pe <> Pc And Pe <> Pb And Pe <> Pa Then
                                        For Pf = 1 To L
                                            If Pf <> Pe And Pf <> Pd And Pf <> Pc And Pf <> Pb And Pf <> Pa Then
                                                Pg = 28 - (Pa + Pb + Pc + Pd + Pe + Pf)
                                                pcount = pcount + 1
                                                perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd) & lett(Pe) & lett(Pf) & lett(Pg)
                                                ReDim Preserve Results(1 To pcount) As String
                                                Results(pcount) = perm
                                            End If
                                        Next Pf
                                    End If
                                Next Pe
                            End If
                        Next Pd
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa
Case 8:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        For Pd = 1 To L
                            If Pd <> Pc And Pd <> Pb And Pd <> Pa Then
                                For Pe = 1 To L
                                    If Pe <> Pd And Pe <> Pc And Pe <> Pb And Pe <> Pa Then
                                        For Pf = 1 To L
                                            If Pf <> Pe And Pf <> Pd And Pf <> Pc And Pf <> Pb And Pf <> Pa Then
                                                For Pg = 1 To L
                                                    If Pg <> Pf And Pg <> Pe And Pg <> Pd And Pg <> Pc And Pg <> Pb And Pg <> Pa Then
                                                        Ph = 36 - (Pa + Pb + Pc + Pd + Pe + Pf + Pg)
                                                        pcount = pcount + 1
                                                        perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd) & lett(Pe) & lett(Pf) & lett(Pg) & lett(Ph)
                                                        ReDim Preserve Results(1 To pcount) As String
                                                        Results(pcount) = perm
                                                    End If
                                                Next Pg
                                            End If
                                        Next Pf
                                    End If
                                Next Pe
                            End If
                        Next Pd
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa
Case 9:
    For Pa = 1 To L
        For Pb = 1 To L
            If Pa <> Pb Then
                For Pc = 1 To L
                    If Pc <> Pa And Pc <> Pb Then
                        For Pd = 1 To L
                            If Pd <> Pc And Pd <> Pb And Pd <> Pa Then
                                For Pe = 1 To L
                                    If Pe <> Pd And Pe <> Pc And Pe <> Pb And Pe <> Pa Then
                                        For Pf = 1 To L
                                            If Pf <> Pe And Pf <> Pd And Pf <> Pc And Pf <> Pb And Pf <> Pa Then
                                                For Pg = 1 To L
                                                    If Pg <> Pf And Pg <> Pe And Pg <> Pd And Pg <> Pc And Pg <> Pb And Pg <> Pa Then
                                                        For Ph = 1 To L
                                                            If Ph <> Pg And Ph <> Pf And Ph <> Pe And Ph <> Pd And Ph <> Pc And Ph <> Pb And Ph <> Pa Then
                                                                Pi = 45 - (Pa + Pb + Pc + Pd + Pe + Pf + Pg + Ph)
                                                                pcount = pcount + 1
                                                                perm = lett(Pa) & lett(Pb) & lett(Pc) & lett(Pd) & lett(Pe) & lett(Pf) & lett(Pg) & lett(Ph) & lett(Pi)
                                                                ReDim Preserve Results(1 To pcount) As String
                                                                Results(pcount) = perm
                                                           End If
                                                        Next Ph
                                                    End If
                                                Next Pg
                                            End If
                                        Next Pf
                                    End If
                                Next Pe
                            End If
                        Next Pd
                    End If
                Next Pc
            End If
        Next Pb
    Next Pa

            
End Select
'Debug.Print "Finished processing "; pcount; " No of permutations"
'frmMain.txtStatus.Text = result

End Sub
