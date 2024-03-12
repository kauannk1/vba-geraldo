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
    If quantidade > 0 Then desconto = 1
    If quantidade >= 25 Then desconto = 15
    If quantidade >= 50 Then desconto = 2
    If quantidade >= 75 Then desconto = 25
    
    MsgBox "Desconto : " & desconto & "%"
    
End Sub
