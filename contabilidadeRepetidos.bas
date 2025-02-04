Option Explicit

Sub VerificaValoresImpares()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    ' Ajustar o nome da planilha de origem:
    Set wsOrigem = ThisWorkbook.Sheets("NomeDaSuaPlanilha")
    
    ' Verifica se já existe a planilha de destino “ValoresNaoRepetidos”.
    ' Se existir, vamos limpá-la. Se não existir, criamos.
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("ValoresNaoRepetidos")
    On Error GoTo 0
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Worksheets.Add
        wsDestino.Name = "ValoresNaoRepetidos"
    Else
        ' Limpar conteúdo antigo, se for o caso
        wsDestino.Cells.Clear
    End If
    
    ' Encontrar a última linha preenchida na planilha de origem (coluna A como referência, por exemplo).
    ' Ajuste conforme a sua estrutura.
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    
    ' Criar um dicionário para contar ocorrências
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Primeiro loop: armazenar contagem das ocorrências de cada valor em J
    For i = 2 To ultimaLinha ' assumindo cabeçalho na linha 1
        Dim valor As Variant
        valor = wsOrigem.Cells(i, "J").Value
        
        ' Se quiser ignorar células vazias, pode colocar If valor <> "" Then ...
        If Not dict.Exists(valor) Then
            dict.Add valor, 1
        Else
            dict(valor) = dict(valor) + 1
        End If
    Next i
    
    ' Agora vamos copiar o cabeçalho para a planilha de destino
    wsOrigem.Rows(1).Copy Destination:=wsDestino.Rows(1)
    
    Dim linhaDestino As Long
    linhaDestino = 2
    
    ' Segundo loop: para cada linha, verificar se a contagem do valor em J é ímpar
    For i = 2 To ultimaLinha
        Dim valAtual As Variant
        valAtual = wsOrigem.Cells(i, "J").Value
        
        If dict(valAtual) Mod 2 <> 0 Then
            ' Se for ímpar, copiamos a linha inteira para a aba de destino
            wsOrigem.Rows(i).Copy wsDestino.Rows(linhaDestino)
            linhaDestino = linhaDestino + 1
        End If
    Next i
    
    ' Opcional: ajustar colunas na planilha de destino
    wsDestino.Columns.AutoFit
    
    MsgBox "Análise concluída. Linhas com valores ímpares de ocorrência foram copiadas para 'ValoresNaoRepetidos'."
End Sub
