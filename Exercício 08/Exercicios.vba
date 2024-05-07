Sub Exercicio_1()
    Dim numeros(1 To 5) As Integer
    Dim i As Integer, j As Integer, temp As Integer

    For i = 1 To 5
        numeros(i) = InputBox("Digite o " & i & "º número.")
    Next i
    
    For i = 1 To 4
        For j = i + 1 To 5
            If numeros(i) > numeros(j) Then
                temp = numeros(i)
                numeros(i) = numeros(j)
                numeros(j) = temp
            End If
        Next j
    Next i
    
   Dim numerosOrdenados As String
    For i = 1 To 5
        numerosOrdenados = numerosOrdenados & CStr(numeros(i)) & ", " 'Conveter o vetor para string
    Next i
    
    numerosOrdenados = Left(numerosOrdenados, Len(numerosOrdenados) - 2) ' Remove a última vírgula e espaço
    'Left = retorna um numero específico de caracteres do inicio
    'Len = retorna o numero de caracteres de uma string
    
    MsgBox "Os números do vetor ordenados são: " & Chr(13) & numerosOrdenados
End Sub

Sub Exercicio_2()
    
    Dim num(1 To 10) As Integer
    Dim i As Integer
    
    For i = 1 To 10
        num(i) = InputBox("Digite o " & i & "º número.")
    Next i
    
    msg = ""
    
    For i = 1 To 10
        If num(i) Mod 2 = 0 Then
            msg = msg & i & "º Elemento: " & num(i) & " é par" & Chr(13)
        Else
            msg = msg & i & "º Elemento: " & num(i) & " é impar" & Chr(13)
        End If
    Next i
    
    MsgBox msg
    
End Sub

Sub Exercicio_3()
    
    Dim num(1 To 8) As Integer
    Dim i As Integer, j As Integer
    
    j = 0
    
    For i = 1 To 8
        num(i) = InputBox("Digite o " & i & "º número.")
    Next i
    
    msg = ""
    
    For i = 1 To 8
       msg = msg & i & "º Elemento: " & num(i) & Chr(13)
       
       If num(i) Mod 6 = 0 Then
            j = j + 1
        End If
    Next i
    
    
    msg = msg & Chr(13) & "Total de múltiplos de 6: " & j
    

    MsgBox msg

End Sub

