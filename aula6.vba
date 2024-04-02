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

Function fn_maior_menor(num1 As Integer, num2 As Integer, num3 As Integer) As String
    Dim maior As Integer
    Dim menor As Integer
    
    If (num1 > num2) Then
        maior = num1
        menor = num2
    Else
        maior = num2
        menor = num1
    End If
    If (num3 > maior) Then
        maior = num3
    End If
        
    If (num3 < menor) Then
        menor = num3
    End If
    
    fn_maior_menor = "O maior valor é o " & maior & " e o menor valor é o " & menor
    
End Function

Sub maior_menor()
    Dim num1 As Integer
    Dim num2 As Integer
    Dim num3 As Integer
    Dim result As String
    
    num1 = InputBox("Digite o primeiro número")
    num2 = InputBox("Digite o segundo número")
    num3 = InputBox("Digite o terceiro número")

    result = fn_maior_menor(num1, num2, num3)
    
    MsgBox result
End Sub
