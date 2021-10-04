Sub Macro()

    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
    Range("B12").Select
    Selection.AutoFill Destination:=Range("B12:F12"), Type:=xlFillDefault
    Range("B12:F12").Select

End Sub


Sub Macro_2()

    Range("J12").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-11]C:R[-2]C)"
    Range("J12").Select
    Selection.AutoFill Destination:=Range("J12:M12"), Type:=xlFillDefault
    Range("J12:M12").Select
  
End Sub

