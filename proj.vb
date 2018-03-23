Sub Schaltfläche2_Klicken()
        
End Sub
Sub Schaltfläche1_Klicken()

Dim score As Integer, result As String

    If D3.Value = "ja" Then
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Ttt"
    Range("D26").Select
    zz++
    
    Else

End Sub
Sub Schaltflächettt4()

End Sub
Sub comparison()
For i = 2 To 1000
    For j = 2 To 1000
        If Worksheets(Worksheet).Range("A" & i).Value = Worksheets(Worksheet).Range("L" & j).Value Then
            Worksheets(worksheet).Range("N" & j).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

        End If
    Next j
Next i
End Sub