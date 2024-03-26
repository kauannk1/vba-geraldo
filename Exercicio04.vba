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

Sub exercicio_3()

    Dim salario As Double
    Dim imposto As Double
    Dim nome As String

    nome = InputBox("Digite o nome do funcionário: ")
    salario = InputBox("Digite o salário do funcionário: R$ ")
    
    
    If salario < 1903.98 Then
        MsgBox "O funcionário não paga IRPF"
        
    ElseIf salario >= 1903.99 And salario < 2826.65 Then
        imposto = salario * 0.075
        MsgBox "A parcela a deduzir do IRPF é R$ " & imposto
        
    ElseIf salario >= 2826.66 And salario < 3751.05 Then
        imposto = salario * 0.15
        MsgBox "A parcela a deduzir do IRPF é R$ " & imposto
        
    ElseIf salario >= 3751.06 And salario <= 4664.68 Then
        imposto = salario * 0.225
        MsgBox "A parcela a deduzir do IRPF é R$ " & imposto
        
    ElseIf salario > 4664.68 Then
        imposto = salario * 0.275
        MsgBox "A parcela a deduzir do IRPF é R$ " & imposto
    End If

End Sub


Sub exercicio_4()

    Dim n1 As Double
    Dim n2 As Double
    Dim n3 As Double
    Dim n4 As Double
    Dim PR As Double
    Dim ED As Double
    
    Dim media_n As Double
    Dim finalMedia As Double
    
    Dim disciplina As String
    Dim nome As String
    Dim situation As String
    
    
    nome = InputBox("Digite o nome do aluno: ")
    disciplina = InputBox("Digite o nome da disciplina: ")
    
    n1 = InputBox("Digite a primeira nota")
    n2 = InputBox("Digite a segunda nota")
    n3 = InputBox("Digite a terceira nota")
    n4 = InputBox("Digite a quarta nota")
    
    media_n = (n1 + n2 + n3 + n4) / 4
    
    PR = InputBox("Digite a nota do Provão")
    ED = InputBox("Digite a nota do Estudo Dirigido")
    
    finalMedia = (media_n * 0.2) + (ED * 0.2) + (PR * 0.6) / (100 / 100)
    
    If finalMedia >= 6 Then
        situation = "APROVADO"
    Else
        situation = "REPROVADO"
    End If
    
    MsgBox "Nome: " & nome & Chr(13) & _
            "Disciplina: " & disciplina & Chr(13) & Chr(13) & _
            "Nota 1: " & n1 & Chr(13) & _
            "Nota 2: " & n2 & Chr(13) & _
            "Nota 3: " & n3 & Chr(13) & _
            "Nota 4: " & n4 & Chr(13) & _
            "Média das 4 notas: " & media_n & Chr(13) & Chr(13) & _
            "Nota Provão: " & PR & Chr(13) & _
            "Nota Estudo Dirigido: " & ED & Chr(13) & Chr(13) & _
            "Média final do aluno: " & finalMedia & Chr(13) & Chr(13) & _
            "Situação do aluno: " & situation
            
End Sub

