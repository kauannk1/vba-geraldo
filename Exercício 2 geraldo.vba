'Exercício 1
Sub Exercício_1()
    Dim numero As Integer
    numero = InputBox("Digite um número inteiro: ")
    If numero > 20 Then
    MsgBox ("O número digitado: " & numero & " é maior que 20")
    End If
    
End Sub

'Exercício 2
Sub Exercício_2()
    Dim numero1 As Integer
    Dim numero2 As Integer
    
    numero1 = InputBox("Digite um número inteiro: ")
    numero2 = InputBox("Digite outro número inteiro: ")
    soma = numero1 + numero2
    If soma > 10 Then
    MsgBox ("O somatório de: " & numero1 & numero2 & " = " & soma)
    End If
End Sub
