Option Explicit

Sub ApenasCroqui()

    Dim wsGeral As Worksheet, wsCroqui As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim proximaLinhaCroqui As Long
    
    ' Ajuste para o nome da aba de origem
    Set wsGeral = ThisWorkbook.Worksheets("geral")
    
    ' Verifica se a planilha "Croqui" já existe
    On Error Resume Next
    Set wsCroqui = ThisWorkbook.Worksheets("Croqui")
    On Error GoTo 0
    
    ' Se não existir, cria
    If wsCroqui Is Nothing Then
        Set wsCroqui = ThisWorkbook.Worksheets.Add(After:=wsGeral)
        wsCroqui.Name = "Croqui"
    End If
    
    ' Limpa o conteúdo da planilha Croqui
    wsCroqui.Cells.Clear
    
    ' Descobre última linha preenchida na aba "geral" (p. ex. coluna A)
    ultimaLinha = wsGeral.Cells(wsGeral.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Não há dados suficientes em 'geral' para copiar.", vbExclamation
        Exit Sub
    End If
    
    ' Cria cabeçalhos na Croqui
    wsCroqui.Range("A1").Value = "DATA"
    wsCroqui.Range("B1").Value = "HISTORICO"
    wsCroqui.Range("C1").Value = "PREFIXO"
    wsCroqui.Range("D1").Value = "NO. TITULO"
    wsCroqui.Range("E1").Value = "PARCELA"
    wsCroqui.Range("F1").Value = "TIPO"
    wsCroqui.Range("G1").Value = "EMISSAO"
    wsCroqui.Range("H1").Value = "VENCTO"
    ' As 2 ou 3 colunas de BAIXA, etc., se quiser, podem ser copiadas também;
    ' ou ajuste conforme a sua necessidade exata.
    
    wsCroqui.Range("I1").Value = "Valor Unificado"
    wsCroqui.Range("J1").Value = "Origem (Débito/Crédito)"
    wsCroqui.Range("K1").Value = "Qtd. Vezes Repetido"
    wsCroqui.Range("L1").Value = "Módulo (Original)"
    
    ' Começamos a preencher a partir da linha 2 em Croqui
    proximaLinhaCroqui = 2
    
    ' Loop por todas as linhas de dados da planilha Geral
    For i = 2 To ultimaLinha
        
        Dim valDeb As Variant, valCred As Variant
        Dim modulo As String
        
        valDeb = wsGeral.Cells(i, "J").Value    ' Coluna J = Débito
        valCred = wsGeral.Cells(i, "K").Value   ' Coluna K = Crédito
        modulo = wsGeral.Cells(i, "L").Value    ' Coluna L = Módulo (Financeiro/Contabilidade)
        
        ' Se houver Débito (não vazio e diferente de 0), cria uma linha em Croqui
        If Not IsEmpty(valDeb) And valDeb <> 0 Then
            ' Copiamos as colunas A–H para Croqui
            wsCroqui.Range("A" & proximaLinhaCroqui & ":H" & proximaLinhaCroqui).Value = _
                wsGeral.Range("A" & i & ":H" & i).Value
            
            ' Valor unificado (coluna I)
            wsCroqui.Cells(proximaLinhaCroqui, "I").Value = valDeb
            ' Origem (coluna J)
            wsCroqui.Cells(proximaLinhaCroqui, "J").Value = "Debito"
            ' Módulo (coluna L)
            wsCroqui.Cells(proximaLinhaCroqui, "L").Value = modulo
            
            ' Fórmula para contar frequência (coluna K)
            ' Faremos a fórmula depois de preencher tudo, ou já deixamos com COUNTIF
            wsCroqui.Cells(proximaLinhaCroqui, "K").Formula = _
                "=COUNTIF($I:$I," & wsCroqui.Cells(proximaLinhaCroqui, "I").Address(False, False) & ")"
            
            proximaLinhaCroqui = proximaLinhaCroqui + 1
        End If
        
        ' Se houver Crédito (não vazio e diferente de 0), cria outra linha em Croqui
        If Not IsEmpty(valCred) And valCred <> 0 Then
            wsCroqui.Range("A" & proximaLinhaCroqui & ":H" & proximaLinhaCroqui).Value = _
                wsGeral.Range("A" & i & ":H" & i).Value
            
            wsCroqui.Cells(proximaLinhaCroqui, "I").Value = valCred
            wsCroqui.Cells(proximaLinhaCroqui, "J").Value = "Credito"
            wsCroqui.Cells(proximaLinhaCroqui, "L").Value = modulo
            wsCroqui.Cells(proximaLinhaCroqui, "K").Formula = _
                "=COUNTIF($I:$I," & wsCroqui.Cells(proximaLinhaCroqui, "I").Address(False, False) & ")"
            
            proximaLinhaCroqui = proximaLinhaCroqui + 1
        End If
        
        ' Obs: Caso uma linha não tenha nem débito nem crédito, não vai pra Croqui;
        ' se você quiser levar "linhas zeradas" também, basta adaptar.
        
    Next i
    
    ' Ajusta a largura das colunas
    wsCroqui.Columns.AutoFit
    
    MsgBox "A planilha 'Croqui' foi gerada/atualizada sem divergências!", vbInformation

End Sub


