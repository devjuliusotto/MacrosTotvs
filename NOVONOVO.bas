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
    ' 2) REMOVER LINHAS CUJA COLUNA I ESTEJA VAZIA
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
    
    On Error Resume Next
    Set wsAveriguar = ThisWorkbook.Worksheets(NOME_ABA_AVERIGUAR)
    On Error GoTo 0
    If wsAveriguar Is Nothing Then
        Set wsAveriguar = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAveriguar.Name = NOME_ABA_AVERIGUAR
    Else
        wsAveriguar.Cells.Clear
    End If
    
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
    ' 4) Se não houver dados além do cabeçalho, sair
    '-----------------------------------------------------------
    If ultimaLinhaCroqui <= 1 Then
        MsgBox "Não há dados em Croqui após remover linhas vazias na Coluna I.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' 5) Copiar cabeçalhos para "Averiguar" e "Erros Encontrados"
    '    e criar a coluna "Observações" em cada uma
    '-----------------------------------------------------------
    Dim ultimaColCroqui As Long
    ultimaColCroqui = wsCroqui.Cells(1, wsCroqui.Columns.Count).End(xlToLeft).Column
    
    Dim rngCabecalho As Range
    Set rngCabecalho = wsCroqui.Range(wsCroqui.Cells(1, 1), wsCroqui.Cells(1, ultimaColCroqui))
    
    rngCabecalho.Copy wsAveriguar.Range("A1")
    rngCabecalho.Copy wsErros.Range("A1")
    
    wsAveriguar.Cells(1, ultimaColCroqui + 1).Value = "Observações"
    wsErros.Cells(1, ultimaColCroqui + 1).Value = "Observações"
    
    Dim proxLinAveriguar As Long: proxLinAveriguar = 2
    Dim proxLinErros As Long: proxLinErros = 2
    
    '-----------------------------------------------------------
    ' 6) COLETAR VALORES ÚNICOS DA COLUNA K
    '-----------------------------------------------------------
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim rngColK As Range
    Set rngColK = wsCroqui.Range("K2:K" & ultimaLinhaCroqui)
    
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
        
        Else
        Dim valoresConcatenados As String
        valoresConcatenados = Join(dic.Items, ", ")
        
      '  MsgBox "Os seguintes valores foram encontrados na coluna K: " & valoresConcatenados, vbExclamation, "Valores Coletados" ' ----------------------------------------------------------XXXXXXXXXXXXXXXXXXXXXX

    End If
        

    
    ' Ordenar array de valores K
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
    
    '-----------------------------------------------------------********************************************************************************************************************************************
    ' 7) PARA CADA valorK, FILTRAR K E DEPOIS I
    '-------------------------------------------------------------------------------------------------------------------*****************************************************************************************
    Dim valorK As Variant
    
    For i = LBound(arrValoresK) To UBound(arrValoresK)
    

        valorK = arrValoresK(i)
        
        
        ' Remove filtros anteriores somente para iniciar esse K
        If wsCroqui.AutoFilterMode Then
            wsCroqui.AutoFilter.ShowAllData
        End If
        
        ' ---> Filtrar K=valorK <---
        wsCroqui.Range("A1").AutoFilter Field:=11, Criteria1:="=" & valorK
        
        Dim rngFiltradoK As Range
        On Error Resume Next
        Set rngFiltradoK = wsCroqui.AutoFilter.Range  ' Range filtrado (com K=valorK)
        On Error GoTo 0
        
        If Not rngFiltradoK Is Nothing Then
            Dim rngVisivelK As Range
            On Error Resume Next
            Set rngVisivelK = rngFiltradoK.Offset(1, 0).Resize(rngFiltradoK.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If Not rngVisivelK Is Nothing Then
                
                ' De acordo com o valorK, temos regras diferentes:
                Select Case True
                    
                    Case (valorK = 1)
                        ' Se K=1 => manda tudo para Averiguar
                        CopiarLinhas rngVisivelK, wsAveriguar, ultimaColCroqui, proxLinAveriguar, _
                                     "K=1 => valor não repetido na base"
                    
                    Case (valorK >= 2 And valorK <= 6)
                        
                        ' Precisamos analisar TODOS os valores distintos de I, um por um,
                        ' mantendo o filtro de K (não limpamos).
                                              ' Precisamos analisar TODOS os valores distintos de I, um por um,
                                            ' mantendo o filtro de K (não limpamos).
                                            Dim dicI As Object
                                            Set dicI = CreateObject("Scripting.Dictionary")
                                            
                                            Dim cI As Range
                                            For Each cI In rngVisivelK.Columns(9).Cells ' col. 9 = I
                                                If Not IsEmpty(cI.Value) Then
                                                    ' Pegar o texto formatado (R$ 1.300,00 etc.)
                                                    Dim txtI As String
                                                    txtI = Trim(cI.Text)
                                                    If Not dicI.Exists(txtI) Then
                                                        dicI.Add txtI, txtI
                                                    End If
                                                End If
                                            Next cI
                                            
                                            ' Exibir os valores distintos encontrados e a quantidade total
                                            If dicI.Count = 0 Then
                                                MsgBox "Não foram encontrados valores distintos na coluna I!", vbInformation, "Valores Distintos"
                                            Else
                                                Dim valoresIConcatenados As String
                                                valoresIConcatenados = Join(dicI.Items, ", ")
                                                
                                                MsgBox "Foram encontrados " & dicI.Count & " valores distintos na coluna I: " & vbCrLf & valoresIConcatenados, _
                                                       vbInformation, "Valores Distintos"
                                            End If
                                                                    
                        If dicI.Count > 0 Then
                            
                            Dim arrI As Variant
                            arrI = dicI.Items
                            
                            ' Ordenar arrI (opcional)
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
                                Dim valorI As String
                                valorI = arrI(j)
                                
                                ' ----> FILTRAR I=valorI, mas mantendo K <----
                                wsCroqui.Range("A1").AutoFilter Field:=9, Criteria1:="=" & valorI
                                
                                Dim rngFiltradoI As Range
                                Set rngFiltradoI = wsCroqui.AutoFilter.Range  ' ainda com K ativo
                                
                                Dim rngVisivelI As Range
                                On Error Resume Next
                                Set rngVisivelI = rngFiltradoI.Offset(1, 0).Resize(rngFiltradoI.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
                                On Error GoTo 0
                                
                                If Not rngVisivelI Is Nothing Then
                                    
                                    ' Chama a sub apropriada
                                    Select Case valorK
                                        Case 2
                                            TrataK2 rngVisivelI, wsAveriguar, ultimaColCroqui, proxLinAveriguar
                                        Case 3
                                            TrataK3 rngVisivelI, wsAveriguar, wsErros, ultimaColCroqui, _
                                                     proxLinAveriguar, proxLinErros
                                        Case 4
                                            ChecarK4 rngVisivelI, wsErros, ultimaColCroqui, proxLinErros
                                        Case 5
                                            ' Sempre erro/s
                                            CopiarLinhas rngVisivelI, wsErros, ultimaColCroqui, proxLinErros, _
                                                         "Valor se encontra 5x vezes na tabela"
                                        Case 6
                                            ChecarK6 rngVisivelI, wsErros, ultimaColCroqui, proxLinErros
                                    End Select
                                    
                                End If
                                
                                ' *Remover SOMENTE o filtro de I*, mantendo K
                                wsCroqui.Range("A1").AutoFilter Field:=9
                                
                            Next j ' próximo valor de I
                        End If
                        
                    Case (valorK > 6 And (valorK Mod 2 <> 0))
                        ' Impar > 6 => Erros
                        Dim msgOdd As String
                        msgOdd = "valor " & valorK & "x na base (ímpar). Favor averiguar."
                        CopiarLinhas rngVisivelK, wsErros, ultimaColCroqui, proxLinErros, msgOdd
                    
                    Case (valorK > 6 And (valorK Mod 2 = 0))
                        ' Par > 6 => ChecarKParAcima6
                        ChecarKParAcima6 valorK, rngVisivelK, wsErros, ultimaColCroqui, proxLinErros
                        
                End Select
                
            End If
        End If
        
        If valorK = 1 Then
            If wsCroqui.AutoFilterMode Then
            wsCroqui.AutoFilter.ShowAllData
            End If
        End If
        
        
        ' <-- Após esgotar todos os valores de I, agora sim limpamos TUDO,
        '     para passar ao próximo K. -->

        
    Next i
    
    MsgBox "Processo concluído de acordo com o pseudocódigo corrigido!", vbInformation
    
            If wsCroqui.AutoFilterMode Then
            wsCroqui.AutoFilter.ShowAllData
            End If

End Sub

'==========================================================
' CopiarLinhas
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
' TrataK2
'==========================================================
Private Sub TrataK2(ByVal rngVisivel As Range, _
                    ByVal wsAveriguar As Worksheet, _
                    ByVal ultimaCol As Long, _
                    ByRef proxLinhaAveriguar As Long)
    
    ' (Exemplo de lógica para K=2)
    Dim countCred As Long, countDeb As Long
    Dim countFin As Long, countContab As Long
    Dim c As Range
    
    ' Contar créditos/débitos (coluna J=10)
    For Each c In rngVisivel.Columns(10).Cells
        Select Case LCase(Trim(c.Value & ""))
            Case "credito", "crédito", "creditos", "créditos"
                countCred = countCred + 1
            Case "debito", "débito", "debitos", "débitos"
                countDeb = countDeb + 1
        End Select
    Next c
    
    Dim erroK2 As Boolean: erroK2 = False
    
    ' Precisamos 2 créditos OU 2 débitos
    If Not (countCred = 2 Or countDeb = 2) Then
        erroK2 = True
    End If
    
    ' Se não errou, verificar finance/contab (col L=12)
    If Not erroK2 Then
        For Each c In rngVisivel.Columns(12).Cells
            Select Case LCase(Trim(c.Value & ""))
                Case "financeiro", "financeiros"
                    countFin = countFin + 1
                Case "contabilidade", "contabilidades"
                    countContab = countContab + 1
            End Select
        Next c
        
        ' Precisamos 1 Fin + 1 Contab
        If Not (countFin = 1 And countContab = 1) Then
            erroK2 = True
        End If
    End If
    
    If erroK2 Then
        CopiarLinhas rngVisivel, wsAveriguar, ultimaCol, proxLinhaAveriguar, _
                     "K=2 inconsistente: crédito/débito ou contab/financeiro."
    End If
End Sub

'==========================================================
' TrataK3
'==========================================================
Private Sub TrataK3(ByVal rngVisivel As Range, _
                    ByVal wsAveriguar As Worksheet, _
                    ByVal wsErros As Worksheet, _
                    ByVal ultimaCol As Long, _
                    ByRef proxLinhaAveriguar As Long, _
                    ByRef proxLinhaErros As Long)
    
    ' (Exemplo de lógica para K=3)
    Dim dictJ As Object
    Set dictJ = CreateObject("Scripting.Dictionary")
    
    Dim c As Range, valJ As String
    For Each c In rngVisivel.Columns(10).Cells
        valJ = LCase(Trim(c.Value & ""))
        Select Case valJ
            Case "credito", "crédito", "creditos", "créditos"
                valJ = "credito"
            Case "debito", "débito", "debitos", "débitos"
                valJ = "debito"
        End Select
        
        If Not dictJ.Exists(valJ) Then
            dictJ.Add valJ, 1
        Else
            dictJ(valJ) = dictJ(valJ) + 1
        End If
    Next c
    
    ' Se só tem 1 tipo => 3 iguais
    If dictJ.Count = 1 Then
        CopiarLinhas rngVisivel, wsErros, ultimaCol, proxLinhaErros, "K=3 => 3 iguais na coluna J"
    Else
        ' 2 de um tipo e 1 de outro
        Dim countCred3 As Long, countDeb3 As Long
        Dim rw As Range
        
        For Each rw In rngVisivel.Rows
            valJ = LCase(Trim(rw.Cells(1, 10).Value & ""))
            Select Case valJ
                Case "credito", "crédito", "creditos", "créditos"
                    countCred3 = countCred3 + 1
                Case "debito", "débito", "debitos", "débitos"
                    countDeb3 = countDeb3 + 1
            End Select
        Next rw
        
        Dim tipoErrado As String
        If countCred3 = 1 And countDeb3 = 2 Then
            tipoErrado = "credito"
        ElseIf countCred3 = 2 And countDeb3 = 1 Then
            tipoErrado = "debito"
        Else
            Exit Sub
        End If
        
        ' Copiar só a divergente para Averiguar
        Dim rngErrada As Range
        For Each rw In rngVisivel.Rows
            valJ = LCase(Trim(rw.Cells(1, 10).Value & ""))
            If valJ = "credito" Or valJ = "crédito" Or valJ = "creditos" Or valJ = "créditos" Then
                valJ = "credito"
            ElseIf valJ = "debito" Or valJ = "débito" Or valJ = "debitos" Or valJ = "débitos" Then
                valJ = "debito"
            End If
            
            If valJ = tipoErrado Then
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
' ChecarK4, ChecarK6, ChecarKParAcima6
' (Cada um com sua lógica)
'==========================================================
Private Sub ChecarK4(ByVal rngVisivel As Range, _
                     ByVal wsErros As Worksheet, _
                     ByVal ultimaCol As Long, _
                     ByRef proxLinhaErros As Long)
    ' ...
End Sub

Private Sub ChecarK6(ByVal rngVisivel As Range, _
                     ByVal wsErros As Worksheet, _
                     ByVal ultimaCol As Long, _
                     ByRef proxLinhaErros As Long)
    ' ...
End Sub

Private Sub ChecarKParAcima6(ByVal valorK As Long, _
                             ByVal rngVisivel As Range, _
                             ByVal wsErros As Worksheet, _
                             ByVal ultimaCol As Long, _
                             ByRef proxLinhaErros As Long)
    ' ...
End Sub


