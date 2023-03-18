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
    With Range("Ae9:Ah32")
        .Value = 0
    End With
    
    With Range("B9:U32")
        .Value = ""
        .Interior.ColorIndex = 0
    End With
    
    With Range("Z9:AA32")
        .Value = "Y"
    End With
    
    With Range("AB9:AC32")
        .Value = "N"
    End With
    
    With Range("B10:AI32").Interior
        .ColorIndex = 0
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    
    With Range("B10:B32", "D10:D32", "F10:F32", "H10:H32", "J10:J32", "L10:L32", "N10:N32", "P10:P32", "R10:R32", "T10:T32", "V10:V32", "X10:X32")
        .Interior.ColorIndex = 20
    End With
End Sub

Sub ShowScores()
    ' This sub-routine shows only columns B to X and hides all other columns in the active worksheet.
    
    ' Select all columns from B to CR in the active worksheet.
    Columns("B:CR").Select
    
    ' Hide all selected columns.
    Selection.EntireColumn.Hidden = True
    
    ' Select all columns from B to X in the active worksheet.
    Columns("B:X").Select
    
    ' Unhide all selected columns.
    Selection.EntireColumn.Hidden = False
    
    ' Move the active cell to cell A1 in the worksheet.
    Range("A1").Select
End Sub

Sub ShowSettings()
    ' This sub-routine shows only columns AZ to CR and hides all other columns in the active worksheet.
    
    ' Select all columns from B to CR in the active worksheet.
    Columns("B:CR").Select
    
    ' Hide all selected columns.
    Selection.EntireColumn.Hidden = True
    
    ' Select all columns from AZ to CR in the active worksheet.
    Columns("AZ:CR").Select
    
    ' Unhide all selected columns.
    Selection.EntireColumn.Hidden = False
    
    ' Move the active cell to cell A1 in the worksheet.
    Range("A1").Select
End Sub

Sub ShowInstructions()
    ' This sub-routine shows all columns from B to CR in the active worksheet, and then moves the active cell to cell A1.
    
    ' Select all columns from B to CR in the active worksheet.
    Columns("B:CR").Select
    
    ' Unhide all selected columns.
    Selection.EntireColumn.Hidden = False
    
    ' Move the active cell to cell A1 in the worksheet.
    Range("A1").Select
End Sub
