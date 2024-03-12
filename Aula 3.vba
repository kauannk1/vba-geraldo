Sub TesteComentario()
    v_x = 0 'x é uma variável
    'apresentar o resultado
    MsgBox v_x
End Sub


Sub TesteComentário_2()
    'Declaração das variáveis
    Dim v_x, v_y, v_z As Integer 'Integer = número inteiro
    
    'Inicia a rotina definindo valores para as variáveis
    v_x = 100
    v_y = 200
    
    'Somando as duas variáveis em uma só:
    v_z = v_x + v_y
    'Criando caixa de mensagem:
    MsgBox ("O resultado da soma é: " & v_z)
    MsgBox "O resultado da soma de " & v_x & " e " & v_y & " é " & v_z 'Utilize o & para juntar as variáveis no msgbox

End Sub

Sub testInsert() 'Inserindo número na célula
    Dim v_numero As Integer '**** Caso seja omitido atuará como Variant
    v_numero = 100
    Range("A1").Value = v_numero
End Sub

Sub counTest()
    Static contador As Integer
    contador = contador + 1
    msg = "Números de execuções: " & contador
    MsgBox msg
End Sub

Sub Calculadora()
    Static num1 As Integer
    Static num2 As Integer
    num1 = InputBox("Digite um número")
    num2 = InputBox("Digite outro número")
    result = num1 + num2
    MsgBox ("O resultado da soma dos dois números é: " & result)
    
End Sub


Sub Exercise1()
    Dim v_1 As Integer
    Dim v_2 As Integer
    v_1 = 10
    v_2 = 20
    Range("A1:A10").Value = v_1
    Range("B1:B10").Value = v_2
            
End Sub
