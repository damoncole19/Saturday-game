Dim FrontSidePressesWon
Dim FrontSidePressesLost
Dim BackSidePressesWon
Dim BackSidePressesLost
' ... add any other variable declarations here

Sub Nassau()
'
'Load Hole Handicap Rankings
'
ClearComments
Audit.ComboBox1.Clear
Range("c7").Select
For Z = 1 To 18
    HoleHdcp(Z) = ActiveCell.Offset(0, Z).Value
Next
Range("Ae9", "Ae32").Value = 0
Range("B8").Select
For entry = 0 To 23
    For Item = 0 To 27
        scores(entry, Item) = ActiveCell.Offset(entry + 1, Item).Value
    Next
Next

For loopcounter = 0 To 23
    If scores(loopcounter, 24) = "Y" Then
        Player1 = scores(loopcounter, 0)
        For LoopCounter1 = 0 To 23
            Player2 = ""
            If scores(LoopCounter1, 0) <> "" And scores(LoopCounter1, 24) = "Y" Then
                TotalNassau = 0
                FrontNineNassau = 0
                BackNineNassau = 0
                
                Player2 = scores(LoopCounter1, 0)
                '
                'Insert Calc Code Here for Nassau
                '
                p1Cap = scores(loopcounter, 1): p2Cap = scores(LoopCounter1, 1)
                If p1Cap - p2Cap < 0 Then
                    P1AdjCap = 0: P2AdjCap = p2Cap - p1Cap
                Else
                    P1AdjCap = p1Cap - p2Cap: P2AdjCap = 0:
                End If
                For LoopCounter3 = 2 To 10
                    p1HoleScore = scores(loopcounter, LoopCounter3)
                    P2HoleScore = scores(LoopCounter1, LoopCounter3)
                    P1AdjHoleScore = p1HoleScore: P2AdjHoleScore = P2HoleScore
                    If P1AdjCap >= HoleHdcp(LoopCounter3 - 1) + 18 Then
                        P1AdjHoleScore = p1HoleScore - 2
                        ElseIf P1AdjCap >= HoleHdcp(LoopCounter3 - 1) Then
                            P1AdjHoleScore = p1HoleScore - 1
                        ElseIf P1AdjCap = HoleHdcp(LoopCounter3 - 1) Then
                            P1AdjHoleScore = p1HoleScore
                    End If
               
                    If P2AdjCap >= HoleHdcp(LoopCounter3 - 1) + 18 Then
                        P2AdjHoleScore = P2HoleScore - 2
                        ElseIf P2AdjCap >= HoleHdcp(LoopCounter3 - 1) Then
                            P2AdjHoleScore = P2HoleScore - 1
                        ElseIf P2AdjCap = HoleHdcp(LoopCounter3 - 1) Then
                            P2AdjHoleScore = P2HoleScore
                    End If
                    
                    PressArray(1, LoopCounter3 - 1) = P1AdjHoleScore
                    PressArray(2, LoopCounter3 - 1) = P2AdjHoleScore
                    If P1AdjHoleScore < P2AdjHoleScore Then
                        FrontNineNassau = FrontNineNassau + 1
                        TotalNassau = TotalNassau + 1
                        ElseIf P1AdjHoleScore > P2AdjHoleScore Then
                            FrontNineNassau = FrontNineNassau - 1
                            TotalNassau = TotalNassau - 1
                    End If
                    '
                    'Add Audit code 1Up 1Dn
                    '
                Next
                For LoopCounter3 = 11 To 19
                    p1HoleScore = scores(loopcounter, LoopCounter3)
                    P2HoleScore = scores(LoopCounter1, LoopCounter3)
                    P1AdjHoleScore = p1HoleScore: P2AdjHoleScore = P2HoleScore
                    If P1AdjCap >= HoleHdcp(LoopCounter3 - 1) + 18 Then
                        P1AdjHoleScore = p1HoleScore - 2
                        ElseIf P1AdjCap >= HoleHdcp(LoopCounter3 - 1) Then
                        P1AdjHoleScore = p1HoleScore - 1
                        ElseIf P1AdjCap = HoleHdcp(LoopCounter3 - 1) Then
                            P1AdjHoleScore = p1HoleScore
                    End If
                
                    If P2AdjCap >= HoleHdcp(LoopCounter3 - 1) + 18 Then
                        P2AdjHoleScore = P2HoleScore - 2
                        ElseIf P2AdjCap >= HoleHdcp(LoopCounter3 - 1) Then
                            P2AdjHoleScore = P2HoleScore - 1
                        ElseIf P2AdjCap = HoleHdcp(LoopCounter3 - 1) Then
                            P2AdjHoleScore = P2HoleScore
                    End If
                    
                    PressArray(1, LoopCounter3 - 1) = P1AdjHoleScore
                    PressArray(2, LoopCounter3 - 1) = P2AdjHoleScore
                    If P1AdjHoleScore < P2AdjHoleScore Then
                        BackNineNassau = BackNineNassau + 1
                        TotalNassau = TotalNassau + 1
                    ElseIf P1AdjHoleScore > P2AdjHoleScore Then
                            BackNineNassau = BackNineNassau - 1
                            TotalNassau = TotalNassau - 1
                    End If
                    
                Next
                '
                'Update the Nassau base numbers
                '
                
            End If
            'ActiveCell.Offset(loopcounter + 1, 29).Select
            
            If FrontNineNassau < 0 Then
               P1NassauDollars = P1NassauDollars - Range("f6").Value
            ElseIf FrontNineNassau > 0 Then
               P1NassauDollars = P1NassauDollars + Range("f6").Value
            End If
            If BackNineNassau < 0 Then
               P1NassauDollars = P1NassauDollars - Range("f6").Value
            ElseIf BackNineNassau > 0 Then
               P1NassauDollars = P1NassauDollars + Range("f6").Value
            End If
            If TotalNassau < 0 Then
               P1NassauDollars = P1NassauDollars - Range("f6").Value
            ElseIf TotalNassau > 0 Then
               P1NassauDollars = P1NassauDollars + Range("f6").Value
            End If
            If Player2 > "" And Player1 > "" And scores(loopcounter, 25) = "Y" And scores(LoopCounter1, 25) = "Y" Then
                Calc_Presses
                ActiveCell.Offset(loopcounter + 1, 29).Value = ActiveCell.Offset(loopcounter + 1, 29).Value + P1NassauDollars + PressDollars
                If Player1 <> Player2 Then
                        Comment = Comment + Player1 & " vs. " & Player2 + vbCrLf & _
                        "Dollars: " + Str(P1NassauDollars + PressDollars) + Chr(10) & _
                        "Frt:" & Str(FrontNineNassau) & _
                        ", Bck:" & Str(BackNineNassau) & _
                        ", Tot:" & Str(TotalNassau) & _
                        " FPW:" & Str(FrontSidePressesWon) & " FPL:" & Str(FrontSidePressesLost) & " BPW:" & Str(BackSidePressesWon) & " BPL:" & Str(BackSidePressesLost) & Chr(10) & Chr(10)
                        'ActiveCell.Offset(loopcounter + 1, 0).ClearComments
                        'ActiveCell.Offset(loopcounter + 1, 0).AddComment Comment
                        AuditArray(loopcounter, 1) = Player1
                        AuditArray(loopcounter, 2) = Comment
                        
                        
                End If
                P1NassauDollars = 0: FrontNineNassau = 0: BackNineNassau = 0: TotalNassau = 0: PressDollars = 0
                FrontSidePressesWon = 0: FrontSidePressesLost = 0: BackSidePressesWon = 0: BackSidePressesLost = 0
            End If
             
            'ActiveCell.Offset(loopCounter + 1, 29).Value = ActiveCell.Offset(loopCounter + 1, 29).Value + P1NassauDollars + PressDollars
            P1NassauDollars = 0: FrontNineNassau = 0: BackNineNassau = 0: TotalNassau = 0: PressDollars = 0
            FrontSidePressesWon = 0: FrontSidePressesLost = 0: BackSidePressesWon = 0: BackSidePressesLost = 0
            
        Next
        If Player1 <> "" Then
            Audit.ComboBox1.AddItem Player1
        End If
        Comment = ""
    End If
Next
If Audit.ComboBox1.ListCount > 0 Then
    Audit.ComboBox1.ListIndex = 0
End If
CalcSkins
'Audit.Show

'Comments_AutoSize
End Sub
