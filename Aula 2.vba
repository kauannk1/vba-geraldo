Sub getValue()
    'Dá o valor da célula
    v_x = Worksheets("Planilha1").Range("A1").Value
    MsgBox v_x
End Sub
    'Setando valor
Sub setValue()
    Worksheets(1).Range("A1").Value = 123.45
End Sub

Sub countBooks()
    'Conta quantas Workbooks estão abertas
    'WorkBooks = Projetos
    MsgBox Workbooks.Count
End Sub

Sub clearCell()
    'Limpa a célula
    Range("A1").ClearContents
End Sub

Sub copyCell()
    'Vai ativar a planilha um (ja está ativa)
    Worksheets("Planilha1").Activate
    'Copiar o Range A1 e colar no Range B1
    Range("A1").Copy Range("B1")
End Sub

    'Adicionando + Projetos
Sub addWorkBook()
    Workbooks.Add
End Sub
    
    'Macro para seleção
Sub Select_Range_A1_D6()
    Range("A1:D6").Select
End Sub
    
    'Pulando Linha e Coluna Específica
Sub TesteOffset()
    Range("A1").Offset(RowOffSet:=1, ColumnOffset:=1).Select
    
End Sub
            
Sub TesteOffset_abreviado()
    Range("A1").Offset(1, 1).Select
End Sub

Sub TesteOffset_voltando()
    Range("B2").Offset(-1, -1).Select
End Sub

Sub TesteOffset_em_grupo()
    Range("A1:C3").Offset(2, 2).Select
End Sub
    
    'Aumenta a seleção
Sub TesteResize()
    Range("A5").Resize(RowSize:=3, ColumnSize:=3).Select
End Sub

Sub TesteResize_abreviado()
    Range("A5").Resize(3, 3).Select
End Sub

Sub TesteResize_emitindo_columnsize()
    Range("A5").Resize(3).Select
End Sub

Sub TesteResize_emitindo_Row()
    Range("A5").Resize(, 3).Select
End Sub

    'Mesclando Resize e Offset
Sub Resize_Offset()
    Range("B1").Offset(, -1).Resize(, 3).Select
End Sub

Sub Exercicio()
    Range("A2").Offset(3, 2).Select
End Sub

Sub Exercício2()
    Range("E8").Offset(-5, -3).Select
End Sub

Sub Exercício3()
    Range("A2").Resize(, 3).Select
End Sub

