Attribute VB_Name = "MainCode"
Dim Comment As Variant
Dim FrontSidePressesWon
Dim FrontSidePressesLost
Dim BackSidePressesWon
Dim BackSidePressesLost
Dim PressDollars
Dim PressDetailsFront(10, 4)
Public scores(0 To 23, 0 To 27)
Dim PressArray(2, 18)
Dim AuditPressArray(24, 10)
Public AuditArray(24, 2)
Public HoleHdcp(18)
Public Player1
Public Player2

Sub NewRound()
    Range("Ae9", "Ae32").Value = 0
    Range("Af9", "Af32").Value = 0
    Range("Ag9", "Ag32").Value = 0
    Range("Ah9", "Ah32").Value = 0
    Range("B9", "u32").Value = ""
    Range("Z9", "AA32").Value = "Y"
    Range("AB9", "AC32").Value = "N"
    Range("B9:u32").Interior.ColorIndex = 0
    Range("b10:ai10").Interior.ColorIndex = 20
    Range("b12:ai12").Interior.ColorIndex = 20
    Range("b14:ai14").Interior.ColorIndex = 20
    Range("b16:ai16").Interior.ColorIndex = 20
    Range("b18:ai18").Interior.ColorIndex = 20
    Range("b20:ai20").Interior.ColorIndex = 20
    Range("b22:ai22").Interior.ColorIndex = 20
    Range("b24:ai24").Interior.ColorIndex = 20
    Range("b26:ai26").Interior.ColorIndex = 20
    Range("b28:ai28").Interior.ColorIndex = 20
    Range("b30:ai30").Interior.ColorIndex = 20
    Range("b32:ai32").Interior.ColorIndex = 20
    
End Sub

Sub ShowScores()
'
' ShowPayouts Macro
'
'
    Columns("b:cr").Select
    Selection.EntireColumn.Hidden = True
    Columns("b:x").Select
    Selection.EntireColumn.Hidden = False
    Range("a1").Select
    
End Sub
Sub ShowSettings()
'
' ShowPayouts Macro
'
'
    Columns("b:cr").Select
    Selection.EntireColumn.Hidden = True
    Columns("az:cr").Select
    Selection.EntireColumn.Hidden = False
    Range("a1").Select
    
End Sub
Sub ShowInstructions()
'
' ShowPayouts Macro
'
'
    Columns("b:cr").Select
    Selection.EntireColumn.Hidden = False
    'Columns("z:ax").Select
    'Selection.EntireColumn.Hidden = False
    Range("a1").Select
    
End Sub
Dim FrontSidePressesWon
Dim FrontSidePressesLost
Dim BackSidePressesWon
Dim BackSidePressesLost
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

'
' Calculate Auto Presses Macro
'
'
Sub Calc_Presses()
    Dim frontgames(10) As Integer
    Dim backgames(10) As Integer
    For l = 1 To 10
        frontgames(l) = 0
        backgames(l) = 0
    Next l
    
    FrontSidePressesWon = 0
    FrontSidePressesLost = 0
    BackSidePressesWon = 0
    BackSidePressesLost = 0
    
    '
    'Front Nine
    '
    CntPresses = 0
    Delta = 0

    For b = 1 To 9
        If PressArray(1, b) < PressArray(2, b) Then
            Delta = 1
        End If
        If PressArray(1, b) > PressArray(2, b) Then
            Delta = -1
        End If
        If PressArray(1, b) = PressArray(2, b) Then
            Delta = 0
        End If
        For c = 0 To CntPresses
            frontgames(c) = frontgames(c) + Delta
        Next c
        If Abs(frontgames(CntPresses)) = Range("L6").Value And b < 9 Then
            CntPresses = CntPresses + 1
            'At this point var B holds the hole number where press starts
            'PressDetailsFront(CntPresses, 0) = frontgames(c)
            
        End If
        AuditPressArray(0, 10) = Player1 + " " + Player2
        
        AuditPressArray(0, b) = AuditPressArray(0, b) + Delta
    Next b
    For g = 1 To CntPresses
        If frontgames(g) > 0 Then
            FrontSidePressesWon = FrontSidePressesWon + 1
        End If
        If frontgames(g) < 0 Then
            FrontSidePressesLost = FrontSidePressesLost + 1
        End If
    Next g
    '
    'Back Nine
    '
    CntPresses = 0
    Delta = 0
    For b = 10 To 18
        If PressArray(1, b) < PressArray(2, b) Then
            Delta = 1
        End If
        If PressArray(1, b) > PressArray(2, b) Then
            Delta = -1
        End If
        If PressArray(1, b) = PressArray(2, b) Then
            Delta = 0
        End If
        For c = 0 To CntPresses
            backgames(c) = backgames(c) + Delta
        Next c
        If Abs(backgames(CntPresses)) = Range("L6").Value And b < 18 Then
            CntPresses = CntPresses + 1
        End If
    Next b
    For g = 1 To CntPresses
        If backgames(g) > 0 Then
            BackSidePressesWon = BackSidePressesWon + 1
        End If
        If backgames(g) < 0 Then
            BackSidePressesLost = BackSidePressesLost + 1
        End If
    Next g
    PressDollars = (FrontSidePressesWon + BackSidePressesWon) * Range("F6").Value
    PressDollars = PressDollars - (FrontSidePressesLost + BackSidePressesLost) * Range("F6").Value
    
    
    
End Sub
Sub SortPlayers()
'
' SortPlayers Macro
'
'
    Range("b9:ac32").Select
    Selection.Sort Key1:=Range("b9"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
End Sub
Sub CalcSkins()
'
' Fill the array with players and scores
'
Range("Ac8").Select
Range("ag9", "ag32").Value = 0
Range("AG8").Value = "Each Gross Skin" + vbCrLf + "$0.00"
For loopcounter = 1 To 32
    If ActiveCell.Offset(loopcounter, 0).Value = "Y" Then
        GrossSkinPlayerCount = GrossSkinPlayerCount + 1
        ActiveCell.Offset(loopcounter, 4) = Range("Y6").Value * -1
    End If
Next
Range("B8").Select
For entry = 0 To 23
    For Item = 0 To 27
        scores(entry, Item) = ActiveCell.Offset(entry + 1, Item).Value
    Next
Next
'
' Reset colorindex on range of scores
'
Range("B9:u32").Interior.ColorIndex = 0
Range("b10:ai10").Interior.ColorIndex = 20
Range("b12:ai12").Interior.ColorIndex = 20
Range("b14:ai14").Interior.ColorIndex = 20
Range("b16:ai16").Interior.ColorIndex = 20
Range("b18:ai18").Interior.ColorIndex = 20
Range("b20:ai20").Interior.ColorIndex = 20
Range("b22:ai22").Interior.ColorIndex = 20
Range("b24:ai24").Interior.ColorIndex = 20
Range("b26:ai26").Interior.ColorIndex = 20
Range("b28:ai28").Interior.ColorIndex = 20
Range("b30:ai30").Interior.ColorIndex = 20
Range("b32:ai32").Interior.ColorIndex = 20
'
' Find the Gross Skins
'
Range("AC8").Select
For loopcounter = 1 To 32
    If ActiveCell.Offset(loopcounter, 0).Value = "Y" Then
        GrossSkinsPLayerCount = GrossSkinsPLayerCount + 1
    End If
Next
For x = 2 To 19
    Prevlow = 10
    LowScoreCount = 1
    For i = 0 To 23
        If scores(i, 27) = "Y" Then
            'LowScoreCount = 1
            HoleScore = scores(i, x)
            If HoleScore = Prevlow Then
                LowScoreCount = LowScoreCount + 1
            End If
            If HoleScore < Prevlow Then
                Prevlow = HoleScore
                LowScoreCount = 1
                LowScoreX = x
                LowScoreI = i
            End If
        End If
    Next
    If LowScoreCount = 1 Then
        Range("b8").Select
        ActiveCell.Offset(LowScoreI + 1, LowScoreX).Interior.ColorIndex = 36
    End If
    Prevlow = 10
Next
'
'Do Totals based on colorindex
'

Range("a8").Select
a = 0
GrossSkincount = 0
For entry = 0 To 23
    If ActiveCell.Offset(entry + 1, 28).Value = "Y" Then
        For Item = 2 To 20
            a = ActiveCell.Offset(entry + 1, Item).Interior.ColorIndex
            If a = 36 Then
                TotalGrossSkinCount = TotalGrossSkinCount + 1
            End If
        Next
    End If
Next

For entry = 0 To 23
    For Item = 2 To 19
        a = ActiveCell.Offset(entry + 1, Item).Interior.ColorIndex
        If a = 36 Then
            GrossSkincount = GrossSkincount + 1
        End If
    Next
    If GrossSkincount > 0 Then
        ActiveCell.Offset((entry + 1), 32).Value = (((Range("Y6").Value * GrossSkinsPLayerCount) / TotalGrossSkinCount) * GrossSkincount) - Range("Y6").Value
        Range("AG8").Value = "Each Gross Skin" + vbCrLf + "$" + Str(Round((Range("Y6").Value * GrossSkinsPLayerCount / TotalGrossSkinCount), 2))
    GrossSkincount = 0
    End If
Next
NetSkins
End Sub

'
'   Calculate Net Skins
'
Sub NetSkins()
Dim HoleHdcp(18)
'
'   Load Hole Handicaps
'
Range("AB8").Select
Range("AF9", "AF32").Value = 0
Range("AF8").Value = "Each Net Skin" + vbCrLf + "$0.00"
For loopcounter = 1 To 32
    If ActiveCell.Offset(loopcounter, 0).Value = "Y" Then
        NetSkinsPLayerCount = NetSkinsPLayerCount + 1
        ActiveCell.Offset(loopcounter, 4) = Range("R6").Value * -1
    End If
Next
Range("c7").Select
For Z = 0 To 17
    HoleHdcp(Z) = ActiveCell.Offset(0, Z + 1).Value
Next
'
'   Find the Net Skins
'
For x = 2 To 19
    Prevlow = 10
    LowScoreCount = 1
    For i = 0 To 23
        If scores(i, 26) = "Y" Then
            'LowScoreCount = 1
            '
            '   Adjust hole score for handicap
            '
            HoleScore = scores(i, x)
                If scores(i, 1) >= HoleHdcp(x - 2) Then
                    HoleScore = HoleScore - 1
                End If
                If HoleScore = Prevlow Then
                    LowScoreCount = LowScoreCount + 1
                End If
                If HoleScore < Prevlow Then
                    Prevlow = HoleScore
                    LowScoreCount = 1
                    LowScoreX = x
                    LowScoreI = i
                End If
       End If
    Next
    If LowScoreCount = 1 Then
        Range("b8").Select
        a = ActiveCell.Offset(LowScoreI + 1, LowScoreX).Interior.ColorIndex
            If a = 36 Then
                ActiveCell.Offset(LowScoreI + 1, LowScoreX).Interior.ColorIndex = 38
            Else
                ActiveCell.Offset(LowScoreI + 1, LowScoreX).Interior.ColorIndex = 37 'here is the culprit
            End If
    End If
    Prevlow = 10
Next
'
'Do Totals based on colorindex
'
Range("A8").Select
a = 0
NetSkincount = 0

For entry = 0 To 23
    For Item = 2 To 20
        a = ActiveCell.Offset(entry + 1, Item).Interior.ColorIndex
        If a = 37 Or a = 38 Then
            TotalNetSkinCount = TotalNetSkinCount + 1
        End If
    Next
Next

For entry = 0 To 23
    For Item = 2 To 20
        a = ActiveCell.Offset(entry + 1, Item).Interior.ColorIndex
        If a = 37 Or a = 38 Then
            NetSkincount = NetSkincount + 1
        End If
    Next
    If NetSkincount > 0 Then
        ActiveCell.Offset((entry + 1), 31).Value = (((Range("r6").Value * NetSkinsPLayerCount) / TotalNetSkinCount) * NetSkincount) - Range("R6").Value
        'NetSkinDisplay = Str(Round((Range("R6").Value * NetSkinsPLayerCount / TotalNetSkinCount), 2))
        Range("AF8").Value = "Each Net Skin" + vbCrLf + "$" + Str(Round((Range("R6").Value * NetSkinsPLayerCount / TotalNetSkinCount), 2))
        NetSkincount = 0
    End If
Next


End Sub
'
' Autosize Comment Box
' posted by Dana DeLouis  2000-09-16
'
Sub Comments_AutoSize()
  Dim MyComments As Comment
  Dim lArea As Long
  For Each MyComments In ActiveSheet.Comments
    With MyComments
      .Shape.TextFrame.AutoSize = True
      If .Shape.Width > 300 Then
        lArea = .Shape.Width * .Shape.Height
        .Shape.Width = 200
        ' An adjustment factor of 1.1 seems to work ok.
        .Shape.Height = (lArea / 200) * 1.1
      End If
    End With
  Next
End Sub
Sub ClearComments()
'
' ClearComments Macro
'
'
    Range("B9:B32").Select
    Selection.ClearComments
End Sub
Sub AuditDialog()
    Audit.Show
    
End Sub
