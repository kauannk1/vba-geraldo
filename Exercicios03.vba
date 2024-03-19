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

Sub exercicio_3()
    Dim pi As Double
    Dim raio As Double
    Dim area As Double
    
    
    pi = 3.14159
    raio = InputBox("Digite o raio da circunferência em centímetros")
    area = pi * (raio * raio)
    MsgBox "A área de uma circunferência de raio " & raio & " é: " & area & "cm²"
    
End Sub

Sub exercicio_4()
    Dim n1 As Integer
    Dim n2 As Integer
    
    n1 = InputBox("Digite um número inteiro")
    n2 = InputBox("Digite outro número inteiro")
    MsgBox "A fórmula: " & "(" & n1 & "+" & n2 & ")²" & " = " & (n1 * n1) + (n2 * n2) + (n1 * n2)
End Sub

Sub exercicio_5()
    Dim juros As Double
    Dim iniValue As Double
    Dim tax As Double
    Dim period As Integer
    
    iniValue = InputBox("Digite o valor inicial ")
    tax = InputBox("Digite a taxa unitária em porcentagem por período")
    tax = tax / 100
    period = InputBox("Digite o período")
    
    juros = iniValue * tax * period
    result = juros + iniValue
    MsgBox "O juros simples equivale a: R$ " & Format(juros, "0.00") & Chr(13) & _
    "Acrescido do valor inicial fica: R$ " & Format(result, "0.00")
    
End Sub

Sub exercicio_6()
    Dim juroscomp As Double
    Dim iniValue As Double
    Dim tax As Double
    Dim period As Integer
    Dim result As Double
    
    iniValue = InputBox("Digite o valor inicial ")
    tax = InputBox("Digite a taxa unitária em porcentagem por período")
    tax = tax / 100
    period = InputBox("Digite o período")
    
    juroscomp = iniValue * (1 + tax) ^ period
    result = juroscomp + iniValue
    MsgBox "O juros composto equivale a: R$ " & Format(juroscomp, "0.00") & Chr(13) & _
    "Acrescido do valor inicial fica: R$ " & Format(result, "0.00")
End Sub

Sub exercicio_7()

    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    
    dia = InputBox("Digite o dia")
    mes = InputBox("Digite o mês")
    ano = InputBox("Digite o ano em quatro caracteres")
    
    
    MsgBox Format(dia, "00") & "/" & Format(mes, "00") & "/" & ano
End Sub

Sub exercicio_8()
    Dim cand1 As String
    Dim cand2 As String
    Dim cand3 As String
    Dim voto1 As Integer
    Dim voto2 As Integer
    Dim voto3 As Integer
    
    cand1 = InputBox("Digite o nome do primeiro candidato: ")
    voto1 = InputBox("Digite o número de votos do primeiro candidato: ")
    
    cand2 = InputBox("Digite o nome do segundo candidato: ")
    voto2 = InputBox("Digite o número de votos do segundo candidato: ")
    
    cand3 = InputBox("Digite o nome do terceiro candidato: ")
    voto3 = InputBox("Digite o número de votos do terceiro candidato: ")
    
    If voto1 > voto2 And voto1 > voto3 Then
        MsgBox cand1 & " é o ganhador, com " & voto1 & " votos"
    
    ElseIf voto2 > voto1 And voto2 > voto3 Then
        MsgBox cand2 & " é o ganhador, com " & voto2 & " votos"

    Else
        MsgBox cand3 & " é o ganhador, com " & voto3 & " votos"
    End If
    
End Sub

Sub exercicio_9()
    Dim escolha As Integer
    
    Dim num1 As Integer
    Dim num2 As Integer
    
    num1 = InputBox("Digite um número inteiro: ")
    num2 = InputBox("Digite outro número inteiro: ")
    escolha = InputBox("Escolha uma das opções a seguir: " & Chr(13) & Chr(13) & _
    "(1) --------> Soma dos números" & Chr(13) & _
    "(2) --------> Subtração dos números" & Chr(13) & _
    "(3) --------> Multiplicação dos números" & Chr(13) & _
    "(4) --------> Divisão dos números")
    
    If escolha = 1 Then
        MsgBox num1 & "+" & num2 & " = " & num1 + num2
        ElseIf escolha = 2 Then MsgBox num1 & "-" & num2 & " = " & num1 - num2
        ElseIf escolha = 3 Then MsgBox num1 & "*" & num2 & " = " & num1 * num2
        Else: MsgBox num1 & "/" & num2 & " = " & num1 / num2
    End If
End Sub

Sub exercicio_10()
    Dim idade As Integer
    Dim sexo As String
    Dim salario As Double
    Dim contador As Integer
    Dim somaSalario As Double
    Dim maiorIdade As Integer
    Dim menorIdade As Integer
    Dim qntmulhersalAlto As Integer
    Dim mediaSalario As Double
    
    ' iniciando os contadores
    contador = 0
    somaSalario = 0
    maiorIdade = 0
    menorIdade = 999 'Garante que qualquer idade seja considerada menor inicia
    
    Do
        idade = InputBox("Digite a idade (ou uma idade negativa para sair):")
        
        If idade < 0 Then
            MsgBox "Você inseriu uma idade negativa. O loop será encerrado."
            Exit Do
        End If
        
        If idade > maiorIdade Then
            maiorIdade = idade
        End If
        
        If idade < menorIdade Then
            menorIdade = idade
        End If
        
        sexo = InputBox("Digite o sexo (M/F):")
        
        salario = InputBox("Digite o salário:")
        
        If salario < 0 Then
            MsgBox "O salário não pode ser negativo. Por favor, insira um valor válido."
            salario = InputBox("Digite o salário:")
        End If
        
        somaSalario = somaSalario + salario
        
        If sexo = "F" And salario > 600 Then
            qntmulhersalAlto = qntmulhersalAlto + 1
        End If
        
        contador = contador + 1
        
    Loop
    
    mediaSalario = somaSalario / contador
    
    MsgBox "Média do salário dos habitantes: R$" & Format(mediaSalario, "0.00") & Chr(13) & _
           "Maior idade do grupo: " & maiorIdade & Chr(13) & _
           "Menor idade do grupo: " & menorIdade & Chr(13) & _
           "Quantidade de mulheres com salários superiores a R$ 600,00: " & qntmulhersalAlto
End Sub

Sub exercicio_11()
    Dim escolha As Integer

    Dim quantcachquente As Integer
    Dim quanthamb As Integer
    Dim quantxtudo As Integer
    Dim quantrefri As Integer
    Dim quantsuco As Integer
    
    Dim Somacachquente As Integer
    Dim Somahamb As Integer
    Dim Somaxtudo As Integer
    Dim Somarefri As Integer
    Dim Somasuco As Integer
    
    Dim valorhamb As Double
    Dim valorcachquente As Double
    Dim valorxtudo As Double
    Dim valorrefri As Double
    Dim valorsuco As Double
    
    
    
    Do
        escolha = InputBox("Qual o código do produto?" & Chr(13) & Chr(13) & _
            "(100) - Cachorro Quente ----- R$ 3,50" & Chr(13) & _
            "(101) - Hambúrguer --------- R$ 3,00" & Chr(13) & _
            "(102) - X-Tudo --------------- R$ 5,00" & Chr(13) & _
            "(103) - Refrigerante ---------- R$ 2,50" & Chr(13) & _
            "(104) - Suco ------------------ R$ 1,50" & Chr(13) & Chr(13) & _
            "(  0  ) - Encerrar o pedido")
            
        If escolha = 0 Then
            MsgBox "Você encerrou o pedido."
            Exit Do
        End If
            
            
        If escolha = 100 Then
            quantcachquente = InputBox("Qual a quantidade de Cachorro Quente desejada?")
            Somacachquente = Somacachquente + quantcachquente
            valorcachquente = Somacachquente * 3.5
            
        ElseIf escolha = 101 Then
            quanthamb = InputBox("Qual a quantidade de Hambúrguer desejada?")
            Somahamb = Somahamb + quanthamb
            valorhamb = Somahamb * 4
            
        ElseIf escolha = 102 Then
            quantxtudo = InputBox("Qual a quantidade de X-Tudo desejada?")
            Somaxtudo = Somaxtudo + quantxtudo
            valorxtudo = Somaxtudo * 5
            
        ElseIf escolha = 103 Then
            quantrefri = InputBox("Qual a quantidade de Refrigerante Quente desejada?")
            Somarefri = Somarefri + quantrefri
            valorrefri = Somarefri * 2.5
            
        ElseIf escolha = 104 Then
            quantsuco = InputBox("Qual a quantidade de Suco Quente desejada?")
            Somasuco = Somasuco + quantsuco
            valorsuco = Somasuco * 1.5
        Else
            MsgBox "Por favor, digite algum código válido!"
        End If
        
    Loop
            Total = valorcachquente + valorhamb + valorxtudo + valorrefri + valorsuco
            MsgBox " O valor de " & Somacachquente & " Cachorros Quentes é: R$ " & Format(valorcachquente, "0.00") & Chr(13) & _
                    "----------------------------------------------------------" & Chr(13) & _
                    " O valor de " & Somahamb & " Hambúrgueres é: R$ " & Format(valorhamb, "0.00") & Chr(13) & _
                    "----------------------------------------------------------" & Chr(13) & _
                    " O valor de " & Somaxtudo & " X-Tudos é: R$ " & Format(valorxtudo, "0.00") & Chr(13) & _
                    "----------------------------------------------------------" & Chr(13) & _
                    " O valor de " & Somarefri & " Refrigerantes é: R$ " & Format(valorrefri, "0.00") & Chr(13) & _
                    "----------------------------------------------------------" & Chr(13) & _
                    " O valor de " & Somasuco & " Sucos é: R$ " & Format(valorsuco, "0.00") & Chr(13) & _
                    "----------------------------------------------------------" & Chr(13) & _
                    " O valor total do pedido ficou R$ " & Format(Total, "0.00")

End Sub
