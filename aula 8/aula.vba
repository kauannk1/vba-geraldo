Sub button()

    Dim resp As Integer
    
    resp = MsgBox("Você almoçou hoje?", vbYesNo)
    
    Select Case resp

        Case vbYes
            MsgBox ("A resposta foi positiva")
        Case vbNo
            MsgBox ("A resposta foi negativa")
    End Select
End Sub

Sub button2()
    Dim resp As Integer
    Dim config As Integer
    
    config = vbYesNo + vbQuestion + vbDefaultButton2
    
    resp = MsgBox("Emitir o relatório mensal?", config)
    
    If (resp = vbYes) Then
        MsgBox ("Invocando a função RunReport")
    End If
    
End Sub

Sub commandBars()
    Application.commandBars.executeMso ("FormatCellsFontDialog")
End Sub

Sub commandBars2()
    Application.commandBars.executeMso ("NumberFormatsDialog")
End Sub


Sub FileName2()
    Dim FileNames As Variant
    Dim msg As String
    Dim i As Integer
    
    FileNames = Application.GetOpenFilename(MultiSelect:=True)
    
    If (IsArray(FileNames)) Then
        'apresenta o caminho completo e o nome dos arquivos
        msg = "Você selecionou: " & Chr(13)
        
        For i = LBound(FileNames) To UBound(FileNames)
        msg = msg & FileNames(i) & Chr(13)
        
        Next i
        MsgBox (msg)
    Else
    'Botao de cancelar clicado
        MsgBox ("Nenhum arquivo foi selecionado")
    End If
End Sub


Sub getPaste()
    With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = Application.DefaultFilePath & "\"
    .title = "Por favor, selecione uma pasta para o backup"
    .Show
    
        If (.SelectedItems.Count = 0) Then
            MsgBox ("Cancelado")
        Else
            MsgBox .SelectedItems(1)
        End If
    End With
End Sub

Sub title()
    Dim msg As String, titulo As String
    Dim config As Integer, resp As Integer
    
    msg = "Você deseja emitir o relatório mensal?"
    msg = msg & Chr(13) & Chr(13)
    msg = msg & "O relatório mensal será emitido "
    msg = msg & "em aproximadamente 15 minutos. Serão "
    msg = msg & "geradas 30 páginas para "
    msg = msg & "todos os escritórios de vendas do "
    msg = msg & "mês atual"
    
    titulo = "Fatec RP - Programação de MicroInformática"
    config = vbYesNo + vbQuestion + vbDefaultButton1
    resp = MsgBox(msg, config, titulo)
    
    If resp = vbYes Then
        MsgBox ("A impressão será iniciada em segundos.")
    End If
End Sub

Sub nameTest()
    Dim padrao As String
    Dim nome As String
    
    nome_padrao = Application.UserName

    nome = InputBox("Qual é o seu nome? ", "Saudações", nome_padrao)
End Sub

Sub inputTest2()
    Dim v_prompt As String, v_caption As String
    Dim valor_default As Integer
    Dim NumSheeets As String
    
    v_prompt = "Quantas planilhas você deseja adicionar?"
    v_caption = "Diga-me..."
    valor_default = 1
    NumSheets = InputBox(v_prompt, v_caption, valor_default)
    
    If (NumSheets = "   ") Then
        Exit Sub 'Fluxo cancelado
    End If
    
    If (IsNumeric(NumSheets)) Then
        If (NumSheets > 0) Then
            Sheets.Add Count:=NumSheets
        Else
        MsgBox ("Número inválido")
        End If
    End If
End Sub

Sub FileName()
    Dim Finfo As String
    Dim FilterIndex As Integer
    Dim titulo As String
    Dim FileName As Variant
    
    Finfo = "Text Files (*.txt), *.txt," & _
    "Lotus Files (*.prn), *.prn," & _
    "Comma Separated Files (*.csv), *.csv," & _
    "ASCII Files (*.asc), *.asc," & _
    "All Files (*.*), *.*"
    
    FilterIndex = 5
    
    titulo = "Selecione um arquivo para importar"
    
    FileName = Application.GetOpenFilename(Finfo, FilterIndex, titulo)
    
    If (FileName = False) Then
        MsgBox ("Nenhum arquivo foi selecionado")
    Else
        MsgBox ("Você selecionou" & FileName)
    End If
End Sub
