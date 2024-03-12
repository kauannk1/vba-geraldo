' TESTANDO O GO TO
Sub testGoTo()
    Nome = InputBox("Digite seu nome")
    If Nome <> "Kauan" Then
        GoTo NomeErrado
    End If
    MsgBox "Seja Bem-vindo Kauan"
    '...Código vba)
    Exit Sub
    
NomeErrado:
    MsgBox "Desculpe, Somente Kauan pode executar essa tarefa"
End Sub
' TESTANTO IF E ELSE
Sub getHour()
    If Time < 0.5 Then
        MsgBox "Bom Dia!"
    Else
        MsgBox "Boa Tarde!"
        'End if vem depois do Else
    End If
End Sub

Sub manhã_tarde_noite()
    Dim msg As String
    If Time < 0.5 Then msg = "Manhã"
    If Time >= 0.5 Then msg = "Tarde"
    If Time >= 0.75 Then msg = "Noite"
    
    MsgBox "Boa " & msg & "!"
End Sub

    'Mesma sub passada porém com End If
Sub manhã_tarde_noite_2()
    Dim msg As String
    If Time < 0.5 Then
        msg = "Manhã"
    End If
    If Time >= 0.5 And Time < 0.75 Then
        msg = "Tarde"
    End If
    If Time >= 0.75 Then
        msg = "Noite"
    MsgBox "Boa " & msg & "!"
End Sub
    'Mesma sub passada porém com ElseIf
Sub manhã_tarde_noite_3()
    Dim msg As String
    If Time < 0.5 Then
        msg = "Manhã"
        
    'Se não, se:
    ElseIf Time >= 0.5 And Time < 0.75 Then
        msg = "Tarde"
        
    Else
        msg = "Noite"
    End If
    
    MsgBox "Boa " & msg & "!"
End Sub

Sub showSale()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um Valor Inteiro: ")
    If quantidade > 0 Then desconto = 10
    If quantidade >= 25 Then desconto = 15
    If quantidade >= 50 Then desconto = 20
    If quantidade >= 75 Then desconto = 25
    
    MsgBox "Desconto : " & desconto & "%"
    
End Sub

    'Mesma sub anterior utilizando Elseif
Sub showSale_2()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um Valor Inteiro: ")
    If quantidade > 0 And quantidade < 25 Then
    desconto = 10
    
    ElseIf quantidade >= 25 And quantidade < 50 Then
    desconto = 15
    
    ElseIf quantidade >= 50 And quantidade < 75 Then
    desconto = 20
    
    ElseIf quantidade >= 75 Then
    desconto = 25
    End If
    
    MsgBox "Desconto: " & desconto & "%"
End Sub

Sub showSale_3()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um Valor Inteiro: ")
    
    Select Case quantidade
        Case 0 To 24
            desconto = 10
        Case 25 To 49
            desconto = 15
        Case 50 To 74
            desconto = 20
        Case Is >= 75
            desconto = 25
    End Select
    
    MsgBox "Desconto: " & desconto & "%"
End Sub

' Aplicando a declaração na mesma linha do CASE

Sub showsale_4()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um Valor Inteiro: ")
    
    Select Case quantidade
        Case 0 To 24: desconto = 10
        Case 24 To 49: desconto = 15
        Case 50 To 74: desconto = 20
        Case Is >= 75: desconto = 25
    End Select
    
    MsgBox "Desconto: " & desconto & "%"
End Sub

'Cascateamento de Select Cases

Sub VerificaCelula()
    Dim msg As String
    'Verificando se há conteúdo
    Select Case IsEmpty(ActiveCell)
        Case Truef
        msg = "está vazia"
        Case Else
            'Verificado se a célula selecionada tem fórmula
            Select Case ActiveCell.HasFormula
            Case True
            msg = "tem uma fórmula"
            Case False
                'Verificando se a célula selecionada tem Número
                Select Case IsNumeric(ActiveCell)
                    Case True
                    msg = "tem número"
                    Case Else
                        msg = "tem texto"
                End Select
            End Select
        End Select
        
    MsgBox "Célula: " & ActiveCell.Address & " " & msg
End Sub


Sub preencherRange()
    
    Dim contador As Long 'Usado para armazenar valores longos

    For contador = 0 To 19
        ActiveCell.Offset(contador, 0) = Rnd 'Random'
    Next contador
    
End Sub

Sub preencherRange_2()
    Dim contador As Long 'Usado para armazenar valores longos
    
    For contador = 0 To 19 Step 2 'Step = pular de 2 em 2 células
        ActiveCell.Offset(contador, 0) = Rnd
    Next contador
End Sub

'Utilizando For e Next

Sub exitforDemo()
    Dim valor_max As Double
    Dim linha As Long
    
    valor_max = WorksheetFunction.Max(Range("A:A")) 'função do excel para procurar valor máximo
    For linha = 1 To Rows.Count 'contar as linhas
        If (Range("A1").Offset(linha - 1, 0).Value = valor_max) Then 'linha - 1 para que ele não ignore a primeira linha
            Range("A1").Offset(linha - 1, 0).Activate 'deixar a linha mais alta ativada (clicada)
        
            MsgBox "Valor máximo é a linha: " & linha
            Exit For
        End If
    Next linha
End Sub


Sub fillCells()
    Dim linha As Long
    Dim coluna As Long
    
    For coluna = 1 To 5
        For linha = 1 To 12 'referece diretamente a numeragem da linha
        Cells(linha, coluna) = Rnd
        Next linha
    Next coluna
End Sub

Sub do_While() 'Faca enquanto a condição for verdadeira
    Do While ActiveCell.Value <> Empty 'Faça enquanto célula ativa for diferente de vazio:
        ActiveCell.Value = ActiveCell.Value * 2
        ActiveCell.Offset(1, 0).Select 'Selecionar a próxima
    Loop
        
End Sub

Sub do_While_2()
    Do
        ActiveCell.Value = ActiveCell.Value * 2
        ActiveCell.Offset(1, 0).Select
    Loop While ActiveCell.Value <> Empty
End Sub



'Mesma função porém com condição diferente

Sub do_Until() 'Faça até que a condição for verdadeira
    Do Until IsEmpty(ActiveCell.Value)
        ActiveCell.Value = ActiveCell.Value * 2
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

Sub do_Until_2()
    Do
        ActiveCell.Value = ActiveCell.Value * 2
        ActiveCell.Offset(1, 0).Select
    Loop Until IsEmpty(ActiveCell.Value)
End Sub
