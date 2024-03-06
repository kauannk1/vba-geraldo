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
