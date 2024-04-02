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


Function Tabuada(numero As Integer, inicio As Integer, fim As Integer) As String
    Dim resultado As String
    Dim multiplyer As Integer
    
    ' Verifica se o início é menor ou igual ao final da tabuada
    If inicio <= fim Then
        For multiplyer = inicio To fim
            resultado = resultado & numero & " x " & multiplyer & " = " & numero * multiplyer & Chr(13)
        Next multiplyer
    Else
        resultado = "O início da tabuada deve ser menor ou igual ao final da tabuada."
    End If
    
    ' Retorna a tabuada calculada
    Tabuada = resultado
End Function

Sub MostrarTabuada()
    Dim numero As Integer
    Dim inicio As Integer
    Dim fim As Integer
    Dim tabuada_resultado As String
    
    numero = InputBox("Digite o número para o qual deseja calcular a tabuada:")
    
    inicio = InputBox("Digite o início da tabuada:")
    
    fim = InputBox("Digite o final da tabuada:")
    
    tabuada_resultado = Tabuada(numero, inicio, fim)
    
    MsgBox tabuada_resultado
End Sub
