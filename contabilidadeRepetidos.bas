Option Explicit

Sub VerificaValoresImpares()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    ' Ajuste o nome da planilha de origem
    Set wsOrigem = ThisWorkbook.Sheets("NomeDaSuaPlanilha")
    
    ' Verifica/Cria a planilha de destino “ValoresNaoRepetidos”
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("ValoresNaoRepetidos")
    On Error GoTo 0
    
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Worksheets.Add
        wsDestino.Name = "ValoresNaoRepetidos"
    Else
        wsDestino.Cells.Clear ' Limpar conteúdo antigo toda vez que rodar
    End If
    
    ' Achar última linha usada (coluna A como base); ajuste conforme necessário
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    
    ' Criar dicionários para contagem das ocorrências em cada coluna (J e K)
    Dim dictJ As Object, dictK As Object
    Set dictJ = CreateObject("Scripting.Dictionary")
    Set dictK = CreateObject("Scripting.Dictionary")
    
    Dim valorJ As Variant, valorK As Variant
    
    ' --- 1º loop: contar quantas vezes cada valor aparece em J e K ---
    For i = 2 To ultimaLinha ' Pressupondo que a linha 1 é cabeçalho
        valorJ = wsOrigem.Cells(i, "J").Value
        valorK = wsOrigem.Cells(i, "K").Value
        
        ' Se não estiver vazio em J, conta no dictJ
        If Not IsEmpty(valorJ) And Trim(valorJ & "") <> "" Then
            If Not dictJ.Exists(valorJ) Then
                dictJ.Add valorJ, 1
            Else
                dictJ(valorJ) = dictJ(valorJ) + 1
            End If
        End If
        
        ' Se não estiver vazio em K, conta no dictK
        If Not IsEmpty(valorK) And Trim(valorK & "") <> "" Then
            If Not dictK.Exists(valorK) Then
                dictK.Add valorK, 1
            Else
                dictK(valorK) = dictK(valorK) + 1
            End If
        End If
    Next i
    
    ' Copiar o cabeçalho para a planilha de destino (opcional)
    wsOrigem.Rows(1).Copy Destination:=wsDestino.Rows(1)
    
    Dim linhaDestino As Long
    linhaDestino = 2
    
    ' --- 2º loop: copiar linhas com ocorrência ímpar em J OU K ---
    For i = 2 To ultimaLinha
        valorJ = wsOrigem.Cells(i, "J").Value
        valorK = wsOrigem.Cells(i, "K").Value
        
        Dim jImpar As Boolean, kImpar As Boolean
        jImpar = False
        kImpar = False
        
        ' Verifica se J não é vazio e se está ímpar
        If Not IsEmpty(valorJ) And Trim(valorJ & "") <> "" Then
            If dictJ(valorJ) Mod 2 <> 0 Then
                jImpar = True
            End If
        End If
        
        ' Verifica se K não é vazio e se está ímpar
        If Not IsEmpty(valorK) And Trim(valorK & "") <> "" Then
            If dictK(valorK) Mod 2 <> 0 Then
                kImpar = True
            End If
        End If
        
        ' Se J for ímpar OU K for ímpar, copia a linha
        If jImpar Or kImpar Then
            wsOrigem.Rows(i).Copy wsDestino.Rows(linhaDestino)
            linhaDestino = linhaDestino + 1
        End If
        
    Next i
    
    ' Ajuste de colunas (opcional)
    wsDestino.Columns.AutoFit
    
    MsgBox "Análise concluída! Linhas com valores de ocorrência ímpar (J ou K) foram copiadas para 'ValoresNaoRepetidos'.", vbInformation
End Sub
