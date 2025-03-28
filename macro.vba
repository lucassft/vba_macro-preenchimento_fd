Public Sub Macro1()
    Set ws = ThisWorkbook.Worksheets("Planilha1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Define the dynamic range
    Set DataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Store the range in a variant array
    cadastro_instrumentos = DataRange.Value
        
    For i = lastRow To 1 Step -1
        Rows("16:17").Select
        Range("C16").Activate
        Selection.Copy
        Rows("18:18").Select
        Selection.Insert Shift:=xlDown
        
        Range("A18:B19").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 2) & "-" & cadastro_instrumentos(i, 1) & "-" & cadastro_instrumentos(i, 3)
        
        Range("E18").Select
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 6)
        
        Range("G18").Select
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 9)
        
        Range("I18").Select
        If cadastro_instrumentos(i, 7) <> "" And cadastro_instrumentos(i, 8) <> "" Then
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 7) & " / " & vbNewLine & cadastro_instrumentos(i, 8)
        
        ElseIf cadastro_instrumentos(i, 7) <> "" Then
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 7)
        
        Else
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 8)
            
        End If
        
    Next i
    
End Sub
