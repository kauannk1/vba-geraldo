Sub showmax()
Dim max As Double

max = WorksheetFunction.max(Range("A:A")) 'Dá o maior valor do intervalo
MsgBox max

End Sub


Sub showmax_2()
Dim max As Double

max = WorksheetFunction.Large(Range("A:A"), 2) 'Vai me dar o segundo maior valor
MsgBox max

End Sub

Sub getpreco()

Dim nr_produto As Variant
Dim preco As Double

nr_produto = InputBox("Informe o número do produto")
Sheets("Planilha1").Activate
preco = WorksheetFunction. _
    VLookup(nr_produto, Range("tab"), 2, False) 'False para não ser sensitive case
MsgBox nr_produto & " preço " & preco

End Sub

Function VerificaCelula(p_valorCelula) As Boolean

    If IsNumeric(p_valorCelula) Then
        VerificaCelula = True
        
    Else
        VerificaCelula = False
    End If
End Function

Sub testeVerificaCelula()

    Dim valor_retornado As Boolean
    
    valor_retornado = VerificaCelula(ActiveCell.Value)
    
    If (valor_retornado = False) Then
        MsgBox ("Não é um número")
        Exit Sub
    End If
End Sub
