'--------------------------------------------------------------------------------
'    Para usar esse código, basicamente vamos copiar + colar a LI em uma aba
'    separada, neste caso, chamada "Planilha1" (vide linha 7). Então, adequamos 
'    os índices do array desta para preencher a FD nos campos corretos, atentando 
'    à lógica de programação.
'--------------------------------------------------------------------------------

Public Sub Macro1()
    Set ws = ThisWorkbook.Worksheets("Planilha1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Define the dynamic range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Store the range in a variant array
    cadastro_instrumentos = dataRange.Value
        
    For i = lastRow To 3 Step -1
        Rows("17:18").Select
        Range("C17").Activate
        Selection.Copy
        Rows("19:19").Select
        Selection.Insert Shift:=xlDown
        
        Range("A19:B20").Select 'identificação (tag)
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 1)
        
        Range("G19").Select 'função
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 2)
        
        Range("I19").Select 'serviço
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 3)
        
        Range("K19").Select 'linha / equipamento
        If (cadastro_instrumentos(i, 4) <> "" And cadastro_instrumentos(i, 4) <> "-") And (cadastro_instrumentos(i, 5) <> "" And cadastro_instrumentos(i, 5) <> "-") Then
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 4) & " / " & vbNewLine & cadastro_instrumentos(i, 5)
        ElseIf cadastro_instrumentos(i, 4) <> "" And cadastro_instrumentos(i, 4) <> "-" Then
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 4)
        Else
            ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 5)
        End If
                
        Range("M19").Select 'fluxograma
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 6)
                
        Range("O19").Select 'fd
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 7)
                        
        Range("Q19").Select 'tipo de sinal
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 8)
                        
        Range("S19").Select 'rede
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 9)
                        
        Range("U19").Select 'diag. de interl.
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 10)
                        
        Range("W19").Select 'detalhe típico
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 11)
                        
        Range("Y19").Select 'planta de locação
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 12)
                        
        Range("AA19").Select 'observação
        ActiveCell.FormulaR1C1 = cadastro_instrumentos(i, 13)
        
    Next i
    
End Sub
