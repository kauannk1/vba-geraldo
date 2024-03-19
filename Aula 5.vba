Sub definindoRange()
    Worksheets("Planilha1").Range("A1:C3").Value = 123

End Sub

Sub propriedadeText()
    MsgBox Range("A1").Text 'Lê com a formatação definida na planilha
    
End Sub

Sub countTest()
    MsgBox Range("A1:C3").Count
    'Conta a quantidade de células no intervalo
End Sub

Sub columnTest()
    MsgBox Range("F3").Column
    'Dá o número da coluna da célula
End Sub

Sub rowsTest()
    MsgBox Range("F3").Row
    'Conta o número de linhas da célula
End Sub

Sub addressTest()
    MsgBox Range(Cells(1, 1), Cells(5, 5)).Address
    'Dá o endereço absoluto da célula
End Sub

Sub hasformulaTest()
    Dim TemFormula As Variant
    TemFormula = Range("A1:A2").HasFormula
    
    If TypeName(TemFormula) = "Null" Then
    MsgBox "Misturado!"
    End If

End Sub

Sub fontTest()
    Range("A1").Font.Bold = True
    'Dá a característica Negrito para a célula
End Sub

Sub fontcolorTest()
    Range("A1").Font.Color = RGB(0, 0, 0)
    Range("A1").Interior.Color = RGB(255, 0, 0)
    'Definindo cor RGB para a célula
End Sub

Sub copypasteTest()

    Range("A1:A12").Select
    Selection.Copy
    Range("C1").Select
    Sheets("Planilha1").Paste
End Sub

