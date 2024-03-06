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
