Sub exercicio_1()

    Dim pesoideal As Double
    Dim pesoatual As Double
    Dim sexo As String
    Dim altura As Double
    

    pesoatual = InputBox("Digite o peso atual: ")
    If pesoatual < 0 Then
        MsgBox "O peso não pode ser negativo."
        Exit Sub
    End If
    altura = InputBox("Digite a altura (em metros): ")
    sexo = InputBox("Digite o sexo (em letra maiúscula): " & Chr(13) & _
                    "( M ) ------> Masculino " & Chr(13) & _
                    "( F ) ------> Feminino ")
    If sexo = "M" Then
        pesoideal = (72.7 * altura) - 58
    ElseIf sexo = "F" Then
        pesoideal = (62.1 * altura) - 44.7
    Else
        MsgBox "Digite uma das opções válidas."
        Exit Sub
    End If

    If pesoatual > pesoideal Then
        MsgBox "O paciente está ACIMA do seu peso ideal"
    End If
End Sub

Sub exercicio_2()
    
    Dim inisalario As Double
    Dim finalsalario As Double
    Dim aumento As Double
    Dim tempo As Integer
    Dim cargo As String
    
    cargo = InputBox("Digite o cargo do funcionário (com letra inicial maiúscula): R$ " & Chr(13) & _
                    " ( Gerente ) " & Chr(13) & _
                    " ( Engenheiro ) " & Chr(13) & _
                    " ( Técnico ) ")
                    
    If cargo = "Gerente" Then
        cargo = gerente
    ElseIf cargo = "Engenheiro" Then
        cargo = engenheiro
    ElseIf cargo = "Técnico" Then
        cargo = tecnico
    Else
        MsgBox "Digite um cargo válido."
        Exit Sub
    End If
    
    
    inisalario = InputBox("Digite o salário do funcionario: ")
    
    tempo = InputBox("Digite o tempo de serviço (em anos) do funcionário: ")
    
    If cargo = gerente And tempo >= 5 Then
        aumento = inisalario * (10 / 100)
    ElseIf cargo = gerente And tempo >= 3 And tempo < 5 Then aumento = inisalario * (9 / 100)
    ElseIf cargo = gerente And tempo < 3 Then aumento = inisalario * (8 / 100)
    End If
    
    If cargo = engenheiro And tempo >= 5 Then
        aumento = inisalario * (11 / 100)
    ElseIf cargo = engenheiro And tempo >= 3 And tempo < 5 Then aumento = inisalario * (10 / 100)
    ElseIf cargo = engenheiro And tempo < 3 Then aumento = inisalario * (9 / 100)
    End If
    
    If cargo = tecnico And tempo >= 5 Then
        aumento = inisalario * (12 / 100)
    ElseIf cargo = tecnico And tempo >= 3 And tempo < 5 Then aumento = inisalario * (11 / 100)
    ElseIf cargo = tecnico And tempo < 3 Then aumento = inisalario * (10 / 100)
    End If
    
    finalsalario = inisalario + aumento
    
        MsgBox "Salário antigo: R$ " & inisalario & Chr(13) & _
                " Aumento de: R$ " & aumento & Chr(13) & _
                " Novo salario: R$ " & finalsalario
    
    
    
    
End Sub
