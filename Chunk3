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
