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
