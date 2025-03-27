Sub Macro1()
    Set ws = ThisWorkbook.Worksheets("Planilha1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Define the dynamic range
    Set DataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Store the range in a variant array
    cadastro_instrumentos = DataRange.Value
        
    For i = 1 To lastRow
        Rows("16:17").Select
        Range("C16").Activate
        Selection.Copy
        Rows("18:18").Select
        Selection.Insert Shift:=xlDown
        
        Range("A18:B19").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 2) & "-" & cadastro_instrumentos(i, 1) & "-" & cadastro_instrumentos(i, 3)
        
        Range("D18").Select
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 6)
        
        Range("E18").Select
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 7)
        
        Range("F18").Select
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 9)

    Next i
    
End Sub
