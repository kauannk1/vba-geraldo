Sub exercicio_1()
    Dim dolar As Double
    Dim euro As Double
    Dim valor As Double
    
    
    dolar = InputBox("Digite a cotação do dólar")
    euro = InputBox("Digite a cotação do euro")
    valor = InputBox("Digite o valor a ser convertido em R$")
    valordolar = valor * dolar
    valoreuro = valor * euro
    MsgBox "R$" & valor & " convertido em dólar é: " & "$" & valordolar & Chr(13) & "R$" & valor & " convertido em euro é: " & "€" & valoreuro
    
End Sub

Sub exercicio_2()
    Dim n1 As Integer
    Dim n2 As Integer
    
    n1 = InputBox("Digite um número inteiro e real")
    n2 = InputBox("Digite outro número inteiro e real")
    soma = n1 + n2
    
    
    If n2 = 0 Then
    MsgBox "O segundo número deve ser maior que 0"
    End If
    
    MsgBox "A soma " & n1 & " + " & n2 & " equivale a: " & n1 + n2 & Chr(13) & _
    "A subtração " & n1 & " - " & n2 & " equivale a: " & n1 - n2 & Chr(13) & _
    "A multiplicação " & n1 & " * " & n2 & " equivale a: " & n1 * n2 & Chr(13) & _
    "A divisão " & n1 & " / " & n2 & " equivale a: " & n1 / n2
    
End Sub
