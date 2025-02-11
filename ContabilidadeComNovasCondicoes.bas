Option Explicit

Sub TratarNovasRegrasCorrigido()

    '-----------------------------------------------------------
    ' 1) Definir e localizar a aba "Croqui"
    '-----------------------------------------------------------
    Dim wsCroqui As Worksheet
    Const NOME_ABA_CROQUI As String = "Croqui"
    
    On Error Resume Next
    Set wsCroqui = ThisWorkbook.Worksheets(NOME_ABA_CROQUI)
    On Error GoTo 0
    
    If wsCroqui Is Nothing Then
        MsgBox "A planilha '" & NOME_ABA_CROQUI & "' não existe!", vbExclamation
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' 2) REMOVER LINHAS DA PLANILHA "Croqui" CUJA COLUNA I ESTEJA VAZIA
    '    (de baixo para cima)
    '-----------------------------------------------------------
    Dim lastRowCroqui As Long
    lastRowCroqui = wsCroqui.Cells(wsCroqui.Rows.Count, "I").End(xlUp).Row
    
    Dim r As Long
    For r = lastRowCroqui To 2 Step -1
        If IsEmpty(wsCroqui.Cells(r, "I")) Or Trim(wsCroqui.Cells(r, "I").Value) = "" Then
            wsCroqui.Rows(r).Delete
        End If
    Next r
    
    '-----------------------------------------------------------
    ' 3) CRIAR (OU LIMPAR) AS ABAS "Averiguar" E "Erros Encontrados"
    '-----------------------------------------------------------
    Dim wsAveriguar As Worksheet, wsErros As Worksheet
    Dim ultimaLinhaCroqui As Long
    ultimaLinhaCroqui = wsCroqui.Cells(wsCroqui.Rows.Count, 1).End(xlUp).Row
    
    Dim NOME_ABA_AVERIGUAR As String: NOME_ABA_AVERIGUAR = "Averiguar"
    Dim NOME_ABA_ERROS As String: NOME_ABA_ERROS = "Erros Encontrados"
    
    ' Criar/Limpar aba "Averiguar"
    On Error Resume Next
    Set wsAveriguar = ThisWorkbook.Worksheets(NOME_ABA_AVERIGUAR)
    On Error GoTo 0
    If wsAveriguar Is Nothing Then
        Set wsAveriguar = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAveriguar.Name = NOME_ABA_AVERIGUAR
    Else
        wsAveriguar.Cells.Clear
    End If
    
    ' Criar/Limpar aba "Erros Encontrados"
    On Error Resume Next
    Set wsErros = ThisWorkbook.Worksheets(NOME_ABA_ERROS)
    On Error GoTo 0
    If wsErros Is Nothing Then
        Set wsErros = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsErros.Name = NOME_ABA_ERROS
    Else
        wsErros.Cells.Clear
    End If
    
    '-----------------------------------------------------------
    ' 4) Se não houver dados na aba Croqui além do cabeçalho, sair
    '-----------------------------------------------------------
    If ultimaLinhaCroqui <= 1 Then
        MsgBox "Não há dados em Croqui após remover linhas com I vazio.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' 5) COPIAR CABEÇALHOS DA ABA CROQUI PARA AMBAS AS ABAS
    '    E CRIAR A COLUNA "Observações" EM CADA UMA
    '-----------------------------------------------------------
    Dim ultimaColCroqui As Long
    ultimaColCroqui = wsCroqui.Cells(1, wsCroqui.Columns.Count).End(xlToLeft).Column
    
    Dim rngCabecalho As Range
    Set rngCabecalho = wsCroqui.Range(wsCroqui.Cells(1, 1), wsCroqui.Cells(1, ultimaColCroqui))
    
    rngCabecalho.Copy wsAveriguar.Range("A1")
    rngCabecalho.Copy wsErros.Range("A1")
    
    ' Criar coluna "Observações" depois do último cabeçalho em cada aba
    wsAveriguar.Cells(1, ultimaColCroqui + 1).Value = "Observações"
    wsErros.Cells(1, ultimaColCroqui + 1).Value = "Observações"
    
    ' Próximas linhas livres nessas abas
    Dim proxLinAveriguar As Long: proxLinAveriguar = 2
    Dim proxLinErros As Long: proxLinErros = 2
    
    '-----------------------------------------------------------
    ' 6) COLETAR VALORES ÚNICOS DA COLUNA K (LINHAS 2 ATÉ O FIM)
    '-----------------------------------------------------------
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim rngColK As Range
    Set rngColK = wsCroqui.Range("K2:K" & ultimaLinhaCroqui)  ' Ajuste se K não for a col. 11
    
    Dim cel As Range
    For Each cel In rngColK
        If Not IsEmpty(cel.Value) Then
            If Not dic.Exists(cel.Value) Then
                dic.Add cel.Value, cel.Value
            End If
        End If
    Next cel
    
    If dic.Count = 0 Then
        MsgBox "Não foram encontrados valores na coluna K!", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' 7) ORDENAR ESSES VALORES (BubbleSort simples)
    '-----------------------------------------------------------
    Dim arrValoresK As Variant
    arrValoresK = dic.Items
    
    Dim i As Long, pass As Long
    Dim swapped As Boolean, temp As Variant
    Dim firstIndex As Long, lastIndex As Long
    firstIndex = LBound(arrValoresK)
    lastIndex = UBound(arrValoresK)
    
    Dim upper As Long: upper = lastIndex
    
    For pass = firstIndex To lastIndex - 1
        swapped = False
        For i = firstIndex To lastIndex - 1
            If arrValoresK(i) > arrValoresK(i + 1) Then
                temp = arrValoresK(i)
                arrValoresK(i) = arrValoresK(i + 1)
                arrValoresK(i + 1) = temp
                swapped = True
            End If
        Next i
        If Not swapped Then Exit For
    Next pass
    
    '-----------------------------------------------------------
    ' 8) PARA CADA valorK ORDENADO, FILTRAR NA COLUNA K
    '-----------------------------------------------------------
    
    Dim valorK As Variant
    
    For i = LBound(arrValoresK) To UBound(arrValoresK)
        
        valorK = arrValoresK(i)
        
        ' Limpar qualquer filtro existente antes
        If wsCroqui.AutoFilterMode Then
            wsCroqui.AutoFilter.ShowAllData
        End If
        
        ' Filtrar a coluna K
        wsCroqui.Range("A1").AutoFilter Field:=11, Criteria1:=valorK
        
        ' Pegar as linhas visíveis (excluindo o cabeçalho)
        Dim rngDados As Range, rngVisivelK As Range
        Set rngDados = wsCroqui.Range("A1").CurrentRegion
        
        On Error Resume Next
        Set rngVisivelK = rngDados.Offset(1, 0).Resize(rngDados.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not rngVisivelK Is Nothing Then
            
            Select Case True
                
                '-------------------------------------------------------
                ' CASO K=1  (SEM SEGUNDO FILTRO EM I)
                '-------------------------------------------------------
                Case (valorK = 1)
                    
                    ' Copiar todas as linhas para "Averiguar"
                    ' com a mensagem "valor não repetido na base. Valor único"
                    CopiarLinhas rngVisivelK, wsAveriguar, ultimaColCroqui, _
                                 proxLinAveriguar, "valor não repetido na base. Valor único"
                
                '-------------------------------------------------------
                ' CASO K=2..6 => FILTRAR A COLUNA I DE UM EM UM
                '-------------------------------------------------------
                Case (valorK >= 2 And valorK <= 6)
                    
                    ' Recolher valores únicos da coluna I dentro desse subset (K=valorK)
                    Dim dicI As Object
                    Set dicI = CreateObject("Scripting.Dictionary")
                    
                    Dim cI As Range
                    For Each cI In rngVisivelK.Columns(9).Cells  ' I = col 9
                        If Not IsEmpty(cI.Value) Then
                            If Not dicI.Exists(cI.Value) Then
                                dicI.Add cI.Value, cI.Value
                            End If
                        End If
                    Next cI
                    
                    If dicI.Count > 0 Then
                        Dim arrI As Variant
                        arrI = dicI.Items
                        
                        ' (Opcional) Ordenar arrI se quiser filtrar em ordem
                        ' BubbleSort simples (pode ser omitido se não for obrigatório)
                        Dim ii As Long, passI As Long
                        Dim swappedI As Boolean, tempI As Variant
                        For passI = LBound(arrI) To UBound(arrI) - 1
                            swappedI = False
                            For ii = LBound(arrI) To UBound(arrI) - 1
                                If arrI(ii) > arrI(ii + 1) Then
                                    tempI = arrI(ii)
                                    arrI(ii) = arrI(ii + 1)
                                    arrI(ii + 1) = tempI
                                    swappedI = True
                                End If
                            Next ii
                            If Not swappedI Then Exit For
                        Next passI
                        
                        Dim j As Long
                        For j = LBound(arrI) To UBound(arrI)
                            Dim valorI As Variant
                            valorI = arrI(j)
                            
                            ' Aplicar segundo filtro na coluna I
                            If wsCroqui.AutoFilterMode Then
                                ' manter o filtro de K
                            End If
                            wsCroqui.Range("A1").AutoFilter Field:=9, Criteria1:=valorI
                            
                            ' Agora, rngVisivelK e rngVisivelI => subset com K=valorK e I=valorI
                            Dim rngVisivelI As Range
                            On Error Resume Next
                            Set rngVisivelI = rngDados.Offset(1, 0).Resize(rngDados.Rows.Count - 1) _
                                             .SpecialCells(xlCellTypeVisible)
                            On Error GoTo 0
                            
                            If Not rngVisivelI Is Nothing Then
                                
                                ' APLICAR AS REGRAS ESPECÍFICAS DE CADA K
                                Select Case valorK
                                    
                                    Case 2
                                        Call TrataK2(rngVisivelI, wsAveriguar, ultimaColCroqui, proxLinAveriguar)
                                    
                                    Case 3
                                        Call TrataK3(rngVisivelI, wsAveriguar, wsErros, ultimaColCroqui, _
                                                     proxLinAveriguar, proxLinErros)
                                    
                                    Case 4
                                        ' Lógica de 4 => Checa 2C2D ou 4C ou 4D, e 2Fin/2Contab
                                        ' Se falhar, copiar para Erros
                                        Call ChecarK4(rngVisivelI, wsErros, ultimaColCroqui, proxLinErros)
                                    
                                    Case 5
                                        ' Sempre copiar as 5 linhas para Erros
                                        CopiarLinhas rngVisivelI, wsErros, ultimaColCroqui, proxLinErros, _
                                                     "Valor se encontra 5x vezes na tabela"
                                    
                                    Case 6
                                        ' Checa distribuição par, 3Fin/3Contab, etc.
                                        Call ChecarK6(rngVisivelI, wsErros, ultimaColCroqui, proxLinErros)
                                End Select
                                
                            End If
                            
                            ' Remover o filtro da Coluna I, mas manter K
                            wsCroqui.Range("A1").AutoFilter Field:=9
                            
                        Next j
                        
                    End If
                
                '-------------------------------------------------------
                ' CASO (valorK > 6 E IMPAR)
                '-------------------------------------------------------
                Case (valorK > 6 And (valorK Mod 2 <> 0))
                    ' Copiar todas as linhas para ErrosEncontrados
                    Dim msgOdd As String
                    msgOdd = "valor " & valorK & "x na tabela. Favor averiguar."
                    CopiarLinhas rngVisivelK, wsErros, ultimaColCroqui, proxLinErros, msgOdd
                
                '-------------------------------------------------------
                ' CASO (valorK > 6 E PAR)
                '-------------------------------------------------------
                Case (valorK > 6 And (valorK Mod 2 = 0))
                    ' Lógica dos pares > 6 => Checar "metade" etc.
                    ' Se falhar, manda p/ Erros.
                    Call ChecarKParAcima6(valorK, rngVisivelK, wsErros, ultimaColCroqui, proxLinErros)
                    
            End Select
            
        End If
        
        ' Limpar todos os filtros antes de passar pro próximo valorK
        If wsCroqui.AutoFilterMode Then
            wsCroqui.AutoFilter.ShowAllData
        End If
        
    Next i
    
    MsgBox "Processo concluído de acordo com o pseudocódigo corrigido!", vbInformation

End Sub

'==========================================================
'                 SUBS AUXILIARES
'==========================================================

Private Sub CopiarLinhas(ByVal rngLinhas As Range, _
                         ByVal wsDestino As Worksheet, _
                         ByVal ultimaColCroqui As Long, _
                         ByRef proxLinha As Long, _
                         ByVal msgObserv As String)
    
    Dim qtde As Long
    qtde = rngLinhas.Rows.Count
    
    rngLinhas.Copy wsDestino.Range("A" & proxLinha)
    
    Dim rngObs As Range
    Set rngObs = wsDestino.Range(wsDestino.Cells(proxLinha, ultimaColCroqui + 1), _
                                 wsDestino.Cells(proxLinha + qtde - 1, ultimaColCroqui + 1))
    rngObs.Value = msgObserv
    
    proxLinha = proxLinha + qtde

End Sub

'==========================================================
' Trata K=2
' Se houver erro (não 2 créditos ou 2 débitos, ou não 1 Fin / 1 Contab),
' copia o bloco para "Averiguar"
'==========================================================
Private Sub TrataK2(ByVal rngVisivel As Range, _
                    ByVal wsAveriguar As Worksheet, _
                    ByVal ultimaCol As Long, _
                    ByRef proxLinhaAveriguar As Long)

    Dim countCred As Long, countDeb As Long
    Dim countFin As Long, countContab As Long
    
    Dim c As Range
    
    ' Coluna J=10 (crédito/débito)
    For Each c In rngVisivel.Columns(10).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "crédito", "credito": countCred = countCred + 1
            Case "débito", "debito":   countDeb = countDeb + 1
        End Select
    Next c
    
    ' Coluna L=12 (financeiro/contabilidade)
    For Each c In rngVisivel.Columns(12).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "financeiro":      countFin = countFin + 1
            Case "contabilidade":   countContab = countContab + 1
        End Select
    Next c
    
    Dim erroK2 As Boolean
    erroK2 = False
    
    ' Precisa 2 créditos ou 2 débitos
    If Not ((countCred = 2) Or (countDeb = 2)) Then
        erroK2 = True
    End If
    
    ' Precisa 1 financeiro e 1 contabilidade
    If Not (countFin = 1 And countContab = 1) Then
        erroK2 = True
    End If
    
    If erroK2 Then
        CopiarLinhas rngVisivel, wsAveriguar, ultimaCol, proxLinhaAveriguar, _
                     "K=2 inconsistente: crédito/débito ou contab/financeiro."
    End If
    
End Sub

'==========================================================
' Trata K=3
' - se as 3 forem iguais (coluna J), copia TODAS para ErrosEncontrados
' - se 2 iguais e 1 diferente, copia só a divergente para Averiguar
'==========================================================
Private Sub TrataK3(ByVal rngVisivel As Range, _
                    ByVal wsAveriguar As Worksheet, _
                    ByVal wsErros As Worksheet, _
                    ByVal ultimaCol As Long, _
                    ByRef proxLinhaAveriguar As Long, _
                    ByRef proxLinhaErros As Long)

    Dim dictJ As Object
    Set dictJ = CreateObject("Scripting.Dictionary")
    
    Dim c As Range, valJ As String
    For Each c In rngVisivel.Columns(10).Cells
        valJ = LCase(Trim(c.Value & ""))
        If Not dictJ.Exists(valJ) Then
            dictJ.Add valJ, 1
        Else
            dictJ(valJ) = dictJ(valJ) + 1
        End If
    Next c
    
    If dictJ.Count = 1 Then
        ' São 3 iguais
        CopiarLinhas rngVisivel, wsErros, ultimaCol, proxLinhaErros, "valor repetido 3x na base"
    Else
        ' Deve ter 2 de um tipo e 1 de outro
        Dim countCred3 As Long, countDeb3 As Long
        For Each c In rngVisivel.Rows
            Select Case LCase(Trim(c.Cells(1, 10).Value & ""))
                Case "crédito", "credito": countCred3 = countCred3 + 1
                Case "débito", "debito":   countDeb3 = countDeb3 + 1
            End Select
        Next c
        
        Dim tipoErrado As String
        If countCred3 = 1 And countDeb3 = 2 Then
            tipoErrado = "credito"
        ElseIf countCred3 = 2 And countDeb3 = 1 Then
            tipoErrado = "debito"
        Else
            ' Caso estranho
            Exit Sub
        End If
        
        Dim rngErrada As Range
        Dim rw As Range
        For Each rw In rngVisivel.Rows
            valJ = LCase(Trim(rw.Cells(1, 10).Value & ""))
            ' Checar com ou sem acento => ("credito","crédito")
            If valJ = tipoErrado Or valJ = tipoErrado & "s" Or valJ = "crédito" And tipoErrado = "credito" _
               Or valJ = "débitos" And tipoErrado = "debito" Then
                If rngErrada Is Nothing Then
                    Set rngErrada = rw
                Else
                    Set rngErrada = Union(rngErrada, rw)
                End If
            End If
        Next rw
        
        If Not rngErrada Is Nothing Then
            CopiarLinhas rngErrada, wsAveriguar, ultimaCol, proxLinhaAveriguar, "K=3 => linha divergente"
        End If
    End If
    
End Sub

'==========================================================
' Checar K=4 => Se não atender (2C2D ou 4C ou 4D) + (2 Fin / 2 Contab),
' copiar tudo para ErrosEncontrados
'==========================================================
Private Sub ChecarK4(ByVal rngVisivel As Range, _
                     ByVal wsErros As Worksheet, _
                     ByVal ultimaCol As Long, _
                     ByRef proxLinhaErros As Long)

    Dim countCred As Long, countDeb As Long
    Dim countFin As Long, countContab As Long
    Dim c As Range
    
    For Each c In rngVisivel.Columns(10).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "credito", "crédito": countCred = countCred + 1
            Case "debito", "débito":   countDeb = countDeb + 1
        End Select
    Next c
    
    For Each c In rngVisivel.Columns(12).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "financeiro":      countFin = countFin + 1
            Case "contabilidade":   countContab = countContab + 1
        End Select
    Next c
    
    Dim erro As Boolean: erro = False
    ' (2C2D) ou (4C) ou (4D)
    If Not ((countCred = 2 And countDeb = 2) Or (countCred = 4) Or (countDeb = 4)) Then
        erro = True
    End If
    ' 2 Fin + 2 Contab
    If Not (countFin = 2 And countContab = 2) Then
        erro = True
    End If
    
    If erro Then
        CopiarLinhas rngVisivel, wsErros, ultimaCol, proxLinhaErros, "valor repetido 4x na base com erro"
    End If
End Sub

'==========================================================
' Checar K=6 => se não atender às regras (par, 3Fin/3Contab, etc.),
' copiar tudo p/ ErrosEncontrados
'==========================================================
Private Sub ChecarK6(ByVal rngVisivel As Range, _
                     ByVal wsErros As Worksheet, _
                     ByVal ultimaCol As Long, _
                     ByRef proxLinhaErros As Long)

    Dim countCred As Long, countDeb As Long
    Dim countFin As Long, countContab As Long
    Dim c As Range
    
    For Each c In rngVisivel.Columns(10).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "credito", "crédito": countCred = countCred + 1
            Case "debito", "débito":   countDeb = countDeb + 1
        End Select
    Next c
    
    For Each c In rngVisivel.Columns(12).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "financeiro":      countFin = countFin + 1
            Case "contabilidade":   countContab = countContab + 1
        End Select
    Next c
    
    Dim erro As Boolean: erro = False
    
    ' Deve ter 6 no total
    If (countCred + countDeb) <> 6 Then
        erro = True
    End If
    
    ' 3 fin, 3 contab
    If Not (countFin = 3 And countContab = 3) Then
        erro = True
    End If
    
    ' Se 4C2D ou 4D2C => checar se aqueles 2 têm 1 fin e 1 contab
    If (countCred = 4 And countDeb = 2) Or (countCred = 2 And countDeb = 4) Then
        Dim tipoMenor As String
        If countCred = 2 Then
            tipoMenor = "credito"
        Else
            tipoMenor = "debito"
        End If
        
        Dim countMenorFin As Long, countMenorContab As Long
        Dim rw As Range
        
        For Each rw In rngVisivel.Rows
            Dim valJ As String
            valJ = LCase(Trim(rw.Cells(1, 10).Value & ""))
            
            If valJ = tipoMenor Or valJ = tipoMenor & "s" Or _
               (tipoMenor = "credito" And (valJ = "crédito" Or valJ = "créditos")) Or _
               (tipoMenor = "debito" And (valJ = "débito" Or valJ = "débitos")) Then
               
                Dim valL As String
                valL = LCase(Trim(rw.Cells(1, 12).Value & ""))
                If valL = "financeiro" Then
                    countMenorFin = countMenorFin + 1
                ElseIf valL = "contabilidade" Then
                    countMenorContab = countMenorContab + 1
                End If
            End If
        Next rw
        
        If Not (countMenorFin = 1 And countMenorContab = 1) Then
            erro = True
        End If
    End If
    
    If erro Then
        CopiarLinhas rngVisivel, wsErros, ultimaCol, proxLinhaErros, "valor repetido 6x na base com erro"
    End If
End Sub

'==========================================================
' Checar valores pares > 6 (8,10,12,...).
' Exemplo: exigir que countCred=metade, countDeb=metade,
'          countFin=metade, countContab=metade
'          Se não atender => Erros
'==========================================================
Private Sub ChecarKParAcima6(ByVal valorK As Long, _
                             ByVal rngVisivel As Range, _
                             ByVal wsErros As Worksheet, _
                             ByVal ultimaCol As Long, _
                             ByRef proxLinhaErros As Long)

    Dim countCred As Long, countDeb As Long
    Dim countFin As Long, countContab As Long
    
    Dim c As Range
    
    For Each c In rngVisivel.Columns(10).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "credito", "crédito": countCred = countCred + 1
            Case "debito", "débito":   countDeb = countDeb + 1
        End Select
    Next c
    
    For Each c In rngVisivel.Columns(12).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "financeiro":      countFin = countFin + 1
            Case "contabilidade":   countContab = countContab + 1
        End Select
    Next c
    
    Dim erro As Boolean: erro = False
    Dim metade As Long
    metade = valorK / 2
    
    ' Exemplo: half credit, half debit, half fin, half contab
    If Not (countCred = metade And countDeb = metade) Then
        erro = True
    End If
    
    If Not (countFin = metade And countContab = metade) Then
        erro = True
    End If
    
    If erro Then
        CopiarLinhas rngVisivel, wsErros, ultimaCol, proxLinhaErros, _
                     "Valor repetido " & valorK & "x (par) na base com erro."
    End If
End Sub


