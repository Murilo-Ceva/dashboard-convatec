Attribute VB_Name = "modBaseWV"
Option Explicit
' v12i CONVATEC (10/11/2026): DATA DO AGENDAMENTO e AGENDADO EM 24H ficam em AF:AG, logo depois de AE, para o analista.
'   A atualização move as duas para o fim antes de começar (o modelo A:AT continua igual) e devolve para AF:AG no fim,
'   mesmo se der erro. A lista SIM/NÃO vem do nome listaSimNao (Agendamentos!R2:R3): com "SIM,NÃO" o Excel pt-BR
'   mostrava um item só.
' v12h CONVATEC (10/11/2026): na Base por NF Oficial, se o cliente agendou (AE = SIM) para depois da previsão,
'   o prazo de ATRASO, DIAS DE ATRASO e OTD passa a ser a NOVA DATA PREV ENTREGA (data agendada).
'   Só a coluna DATA DO AGENDAMENTO (gravada pelo evento da aba) acompanha a nota como coluna adicional.
' v12g CONVATEC (10/11/2026): colunas adicionais digitadas à mão (sem fórmula, com cabeçalho, depois de AT),
'   como DATA AGENDADA, acompanham a nota pela chave CT-e + NF. Antes ficavam presas à posição da linha.
' v12f CONVATEC (10/11/2026): na Base por NF Oficial, AC/AD/AE deixam de ser Ano/Mês/Qtde. NFs e viram
'   CLIENTE AGENDADO (fórmula na aba Agendamentos: CNPJ, CNPJ truncado pelo BI ou nome), CONTATO (e-mail/portal)
'   e AGENDAMENTO REALIZADO (manual, lista SIM/NÃO, preservado). A aba Base continua com Ano/Mês/Qtde.
' v12e CONVATEC (10/11/2026): modo ACUMULADO pela chave CT-e. CT-e que volta no 238 tem as linhas
'   substituídas pelos dados novos (campos manuais preservados); CT-e novo entra; CT-e que não veio
'   no 238 do dia continua na base como estava. Não gera duplicidade de CT-e nem de CT-e + NF.
' v12d CONVATEC (10/11/2026): Prazo BID por modal. RODOAEREO (DESC_OPERACAO do 238) usa a aba TT Aereo
'   (UF + capital/interior, prazo convencional); rodoviário usa TT LTL e, sem a cidade, a aba Transit time.
'   Modalidade (J) = DESC_OPERACAO do 238. Exceção SEM PRAZO BID sai do resultado da fórmula.
' v12c CONVATEC (10/11/2026): Prazo BID vira fórmula PROCV(Cidade&UF; 'Transit time'!B:C). Muda a aba e a base acompanha.
' v12b CONVATEC (10/11/2026): sem a coluna N DN (M) na Base e na Base por NF Oficial. A macro
'   monta a cópia de trabalho no modelo antigo (com M) e apaga M antes de gravar; na planilha o OTIF fica em AT.
' v12 CONVATEC: o cliente filtrado vem da célula cfgCliente (aba Config Follow); padrão CONVATEC.
' v11: só entram linhas do cliente HPE (238: RAZAO_SOCIAL_CLIENTE; BI: NOME_REMETENTE); aviso claro de Somente Leitura.
' v10: Transit time com chave CIDADE+UF; colunas sem cabeçalho são mantidas, salvo fórmula com #REF!.
' v9 (Base HPE): usa a aba Base por NF Oficial; a macro grava A:AU e preserva as colunas
' adicionais com cabeçalho (AV em diante), estendendo as fórmulas até a última linha.
' Tabelas dinâmicas são reapontadas e atualizadas.
' Evidenciado (AB) sempre vem do relatório; Cod e Justificativa (Z/AA) vêm do relatório e
' só mantêm o valor anterior quando o relatório vier em branco.
Private Const DETAIL_SHEET As String = "Base por NF Oficial"
Private Const DETAIL_COLS As Long = 47   'colunas da cópia de trabalho (com N DN)
Private Const WB_COLS As Long = 46       'colunas na planilha (sem N DN)
Private Const COL_NDN As Long = 13
'Linha do cliente na aba Agendamentos: CNPJ exato; senão CNPJ que o BI truncou em 6 dígitos (6,08849E+13); senão nome.
'Prazo efetivo na cópia de trabalho: AD = CLIENTE AGENDADO, AF = AGENDAMENTO REALIZADO, AK = NOVA DATA PREV ENTREGA.
Private Const PRAZO_AG As String = "IF(AND(AD2=""SIM"",AF2=""SIM"",ISNUMBER(AK2)),MAX(INT(U2),INT(AK2)),INT(U2))"
Private Const AGENDA_POS As String = "IFERROR(MATCH(--L2,Agendamentos!$K:$K,0),IFERROR(MATCH(IF(--L2=ROUND(--L2,6-LEN(--L2)),ROUND(--L2,6-LEN(--L2)),""#""),Agendamentos!$L:$L,0),MATCH(IF(TRIM(K2)="""",""#"",TRIM(UPPER(K2))),Agendamentos!$M:$M,0)))"
Private Const BASE_COLS As Long = 32    'colunas da Base na cópia de trabalho (com N DN)
'Texto procurado no nome do cliente. O 238 pode vir com todos os clientes da CEVA.
Private Const CLIENT_FILTER_PADRAO As String = "CONVATEC"

' v8: aceita SharePoint/OneDrive e salva o backup em caminho local quando necessário.
' Regra do modelo atual: previsão = emissão CT-e + Prazo BID em dias úteis (sem sábado, domingo e datas da aba Feriados).
' CTE 24 HRS usa a diferença exata B-D; NF sem horário equivale a meia-noite.
' AGING emissão = hoje - emissão. AGING previsão = previsão - hoje.
' Semana ISO 8601. OTIF aguarda confirmação de entrega completa.
' Nenhuma ausência de ocorrência é tratada como prova de entrega completa.
Private mExceptionRow As Long
Private mRelData As Variant
Private mBIData As Variant
Private mLastMessage As String

Public Sub AtualizarBaseWV()
    Dim p238 As Variant, p2749 As Variant
    p238 = Application.GetOpenFilename("Relatórios Excel (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , "Selecione o relatório 238")
    If VarType(p238) = vbBoolean Then Exit Sub
    p2749 = Application.GetOpenFilename("Relatórios Excel (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , "Selecione o relatório BI (2749 ou 2969)")
    If VarType(p2749) = vbBoolean Then Exit Sub
    AtualizarBaseWVArquivos CStr(p238), CStr(p2749), False
End Sub

Public Function ResultadoUltimaAtualizacaoWV() As String
    ResultadoUltimaAtualizacaoWV = mLastMessage
End Function

' Entrada automatizada para teste ou integração. Não ignora nenhuma validação.
Public Sub AtualizarBaseWVArquivos(ByVal arquivo238 As String, ByVal arquivo2749 As String, Optional ByVal silencioso As Boolean = False)
    Dim stage As Workbook, src238 As Workbook, srcBI As Workbook, rollback As Workbook
    Dim detail As Worksheet, bas As Worksheet, rel As Worksheet, bi As Worksheet, ex As Worksheet, target As Worksheet
    Dim hRel As Object, hBI As Object, dRel As Object, dNF As Object, dBI As Object, dTotal As Object
    Dim dIndex As Object, transit As Object, holidays As Object, missing As Object, manual As Object
    Dim nfKeys As Object, record As Variant, total As Variant, nr As Variant, key As Variant, cte As Variant
    Dim dataBase() As Variant, dataDetail() As Variant, displayNF As String, hasBI As Boolean
    Dim nBase As Long, nDetail As Long, rb As Long, rd As Long, nfCount As Long, c As Long, restored As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldSecurity As Long, oldStatus As Variant, backupPath As String, phase As String, errMsg As String, errNum As Long
    Dim commitStarted As Boolean, committed As Boolean, newSheets As Collection, names As Variant, name As Variant
    Dim nfMissing As Long, sumNF As Double, sumFreight As Double, tolerance As Double
    Dim pivotMsg As String, filterMsg As String, removed238 As Long, removedBI As Long
    Dim oldDet As Variant, oldBas As Variant, keepDet As Collection, keepBas As Collection, oldD As Object
    Dim manualExtra As Object, restoredExtra As Long
    oldCalc = Application.Calculation: oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    oldSecurity = Application.AutomationSecurity: oldStatus = Application.StatusBar
    Set newSheets = New Collection
    On Error GoTo Failed
    phase = "validar o arquivo de trabalho"
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 701, , "Salve a pasta de trabalho como .xlsm antes de atualizar."
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 702, , "A planilha está aberta como Somente Leitura (veja a barra de título). " & _
        "Feche e abra de novo sem Somente Leitura. Se outra pessoa estiver com ela aberta, peça para fechar."
    If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 502, , "A estrutura da pasta de trabalho está protegida."
    Set detail = SheetIfExists(ThisWorkbook, DETAIL_SHEET)
    If detail Is Nothing Then Err.Raise vbObjectError + 503, , "Não encontrei a aba Base por NF Oficial."
    phase = "mover DATA DO AGENDAMENTO e AGENDADO EM 24H para o fim"
    MoverAgendaParaFim detail
    phase = "validar o arquivo de trabalho"
    If SheetIfExists(ThisWorkbook, "Feriados") Is Nothing Then Err.Raise vbObjectError + 511, , "Não encontrei a aba Feriados, usada no cálculo da previsão."
    ValidateModel detail
    If LCase$(arquivo238) = LCase$(arquivo2749) Then Err.Raise vbObjectError + 504, , "Selecione dois arquivos diferentes: relatório 238 e relatório 2749."
    ValidateFile arquivo238: ValidateFile arquivo2749
    Set manual = CaptureManual(detail)
    Set manualExtra = CaptureManualExtras(detail)
    Set transit = TransitForWorkbook(ThisWorkbook, detail)
    Set holidays = CreateObject("Scripting.Dictionary")
    Set missing = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3
    phase = "abrir e validar os relatórios"
    Application.StatusBar = phase
    Set src238 = Workbooks.Open(Filename:=arquivo238, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Local:=True)
    Set srcBI = Workbooks.Open(Filename:=arquivo2749, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Local:=True)
    Set stage = Workbooks.Add(xlWBATWorksheet)
    Set rel = stage.Worksheets(1): rel.Name = "REL_238"
    Set bi = stage.Worksheets.Add(After:=rel): bi.Name = "BI_2749"
    ImportReport src238, rel, Array("NR_CTRC", "DT_EMISSAO", "NUMERO_DA_NF", "UF_DEST", "CIDADE_DEST", "RAZA_SOCIAL")
    ImportReport srcBI, bi, Array("NR_CTRC", "NR_NF", "DT_EMISSAO_NF", "VLR_MERC_NF", "QTD_VOL", "PESO_BRUTO_CTE", "PESO_CUBADO_CTE", "VLR_PRESTACAO_CTE", "VLR_MERCADORIA_CTE")
    phase = "filtrar o cliente " & CLIENT_FILTER
    removed238 = FilterClientRows(rel, "RAZAO_SOCIAL_CLIENTE")
    If removed238 = -2 Then Err.Raise vbObjectError + 711, , "O relatório 238 não tem nenhuma linha do cliente " & CLIENT_FILTER & "."
    removedBI = FilterClientRows(bi, "NOME_REMETENTE")
    If removedBI = -2 Then Err.Raise vbObjectError + 712, , "O relatório BI não tem nenhuma linha do remetente " & CLIENT_FILTER & "."
    filterMsg = "Filtro " & CLIENT_FILTER & ": " & FilterText(removed238, "238", "RAZAO_SOCIAL_CLIENTE") & "; " & _
        FilterText(removedBI, "BI", "NOME_REMETENTE") & "."
    CloseWorkbookSafe src238
    CloseWorkbookSafe srcBI
    Set hRel = BuildHeaderMap(rel): Set hBI = BuildHeaderMap(bi)
    CacheSheet rel: CacheSheet bi
    ValidateRows rel, hRel, False: ValidateRows bi, hBI, True
    phase = "consolidar CT-es e NFs"
    Application.StatusBar = phase
    Set dRel = Build238Dictionary(rel, hRel)
    Set dNF = Build238NFDictionary(rel, hRel)
    Set dBI = BuildBIDictionary(bi, hBI)
    Set dTotal = BuildBICTEDictionary(bi, hBI, dBI)
    Set dIndex = IndexNFUnionByCte(dNF, dBI)
    If dRel.Count = 0 Or dBI.Count = 0 Then Err.Raise vbObjectError + 505, , "Os relatórios não geraram CT-es e NFs válidos."
    nBase = dRel.Count: nDetail = CountMatchedNFUnion(dIndex, dRel)
    If nDetail = 0 Then Err.Raise vbObjectError + 506, , "Os relatórios 238 e 2749 não têm correspondências. Confira os arquivos selecionados."
    phase = "manter os CT-es anteriores que não vieram no 238"
    Set keepDet = OldRowsToKeep(detail, WB_COLS, dRel, dIndex, oldDet)
    Set target = SheetIfExists(ThisWorkbook, "Base")
    Set keepBas = New Collection
    If Not target Is Nothing Then Set keepBas = OldRowsToKeep(target, BASE_COLS - 1, dRel, Nothing, oldBas)
    nDetail = nDetail + keepDet.Count: nBase = nBase + keepBas.Count
    If nDetail > detail.Rows.Count - 1 Then Err.Raise vbObjectError + 507, , "A quantidade de NFs excede o limite de linhas do Excel."
    Set detail = CopyDetailToStage(detail, stage, nDetail)
    EnsureDetailModel detail
    Set target = SheetIfExists(ThisWorkbook, "Base")
    If target Is Nothing Then
        Set bas = stage.Worksheets.Add(After:=detail): bas.Name = "Base"
        detail.Range("A1:AF2").Copy Destination:=bas.Range("A1")
        bas.Cells(1, 3).Value2 = "Notas fiscais do CTe": bas.Cells(1, 5).Value2 = "Val.NFs (Total CTe)"
        bas.Cells(1, 16).Value2 = "Qtde. Vol.": bas.Cells(1, 19).Value2 = "Valor Frete"
    Else
        If target.ProtectContents Then Err.Raise vbObjectError + 508, , "A aba Base está protegida."
        target.Copy After:=stage.Worksheets(stage.Worksheets.Count)
        Set bas = stage.Worksheets(stage.Worksheets.Count)
        If EhNDN(bas.Cells(1, COL_NDN).Value2) Then Err.Raise vbObjectError + 513, , _
            "A coluna M da aba Base ainda é N DN. Esta versão trabalha sem ela: apague a coluna M da aba Base."
        bas.Columns(COL_NDN).Insert Shift:=xlToRight
        bas.Cells(1, COL_NDN).Value2 = "N  DN"
    End If
    ThisWorkbook.Worksheets("Feriados").Copy After:=stage.Worksheets(stage.Worksheets.Count)
    ThisWorkbook.Worksheets("Transit time").Copy After:=stage.Worksheets(stage.Worksheets.Count)
    ThisWorkbook.Worksheets("TT LTL").Copy After:=stage.Worksheets(stage.Worksheets.Count)
    ThisWorkbook.Worksheets("TT Aereo").Copy After:=stage.Worksheets(stage.Worksheets.Count)
    If SheetIfExists(ThisWorkbook, "Agendamentos") Is Nothing Then Err.Raise vbObjectError + 514, , "Não encontrei a aba Agendamentos (lista de clientes com agendamento)."
    ThisWorkbook.Worksheets("Agendamentos").Copy After:=stage.Worksheets(stage.Worksheets.Count)
    Set ex = stage.Worksheets.Add(After:=bas): ex.Name = "Exceções"
    ex.Range("A1:F1").Value = Array("Tipo", "CT-e", "NF", "Detalhe", "Origem", "Data")
    mExceptionRow = 2
    ClearData detail: ClearData bas
    ReDim dataBase(1 To nBase, 1 To 32): ReDim dataDetail(1 To nDetail, 1 To 47)
    For Each cte In dRel.Keys
        record = dRel(cte): rb = rb + 1: total = Empty
        hasBI = dTotal.Exists(CStr(cte))
        If hasBI Then total = dTotal(cte)
        If dIndex.Exists(CStr(cte)) Then
            Set nfKeys = dIndex(cte): nfCount = nfKeys.Count
            displayNF = JoinNFList(nfKeys, dNF, dBI)
        Else
            nfCount = 0: displayNF = ""
        End If
        FillOutputRow dataBase, rb, record, total, hasBI, displayNF, nfCount, transit, holidays, missing, ex
        sumNF = 0: sumFreight = 0: nfMissing = 0
        If nfCount > 0 Then
            For Each key In nfKeys.Keys
                rd = rd + 1: nr = Empty: hasBI = dBI.Exists(CStr(key))
                If hasBI Then nr = dBI(key)
                If dNF.Exists(CStr(key)) Then
                    displayNF = CStr(dNF(key))
                Else
                    displayNF = CStr(nr(1))
                End If
                FillOutputRow dataDetail, rd, record, nr, hasBI, displayNF, nfCount, transit, holidays, missing, ex
                dataDetail(rd, 32) = Empty   'AGENDAMENTO REALIZADO: campo do analista
                If hasBI Then
                    dataDetail(rd, 5) = nr(20): dataDetail(rd, 16) = nr(9)
                    sumNF = sumNF + CDbl(nr(20))
                Else
                    nfMissing = nfMissing + 1
                    LogException ex, "NF SEM BI", CStr(cte), displayNF, "Valor e volume da NF não disponíveis.", "REL_238 x BI_2749"
                End If
                If dTotal.Exists(CStr(cte)) Then
                    dataDetail(rd, 17) = total(10): dataDetail(rd, 18) = total(11)
                    dataDetail(rd, 19) = CDbl(total(12)) / nfCount
                    sumFreight = sumFreight + CDbl(dataDetail(rd, 19))
                End If
            Next key
        Else
            LogException ex, "CTE SEM NF", CStr(cte), "", "CT-e sem NF nos relatórios.", "REL_238"
        End If
        If dTotal.Exists(CStr(cte)) Then
            If nfCount > 0 And Abs(sumFreight - CDbl(total(12))) > 0.000001 Then Err.Raise vbObjectError + 509, , "O rateio do frete não fecha para CT-e " & CStr(cte)
            If nfMissing = 0 And nfCount > 0 Then
                dataBase(rb, 5) = sumNF
                If Abs(sumNF - CDbl(total(3))) > 0.01 Then LogException ex, "VALOR NF DIFERENTE DO TOTAL CTE", CStr(cte), "", "Soma NF: " & sumNF & "; total BI: " & total(3), "BI_2749"
            Else
                dataBase(rb, 5) = total(3)
            End If
        Else
            LogException ex, "CTE SEM BI", CStr(cte), "", "Totais financeiros indisponíveis.", "REL_238 x BI_2749"
        End If
    Next cte
    For Each key In dBI.Keys
        nr = dBI(key)
        If Not dRel.Exists(CStr(nr(0))) Then LogException ex, "CTE SOMENTE NO BI", CStr(nr(0)), CStr(nr(1)), "Não constar no 238 do dia: se já estava na Base, ficou como estava; se é novo, não entrou.", "BI_2749"
    Next key
    AppendOldRows dataDetail, rd, oldDet, keepDet, WB_COLS
    Set oldD = AppendOldRows(dataBase, rb, oldBas, keepBas, BASE_COLS - 1)
    phase = "preencher fórmulas e restaurar campos manuais"
    Application.StatusBar = phase
    WriteData bas, dataBase, nBase, 32: WriteData detail, dataDetail, nDetail, 47
    ApplyCoreFormulas bas, nBase + 1: ApplyCoreFormulas detail, nDetail + 1
    'Data NF da Base é PROCV no BI do dia: CT-e antigo fora do BI mantém a data que já tinha.
    For Each key In oldD.Keys
        bas.Cells(CLng(key), 4).Value2 = oldD(key)
    Next key
    ApplyDetailFormulas detail, nDetail + 1
    restored = RestoreManual(detail, manual, ex)
    detail.Calculate: bas.Calculate
    LogSemPrazo detail, ex
    ValidateFormulaResults detail, nDetail + 1
    ValidateFormulaResults bas, nBase + 1
    detail.Columns(COL_NDN).Delete Shift:=xlToLeft
    bas.Columns(COL_NDN).Delete Shift:=xlToLeft
    names = Array("REL_238", "BI_2749", "Base", DETAIL_SHEET, "Exceções")
    For Each name In names
        Set target = SheetIfExists(ThisWorkbook, CStr(name))
        If Not target Is Nothing Then
            If target.ProtectContents Then Err.Raise vbObjectError + 510, , "A aba " & name & " está protegida."
        End If
    Next name
    phase = "criar cópia de segurança"
    backupPath = BackupFilePath(ThisWorkbook)
    ThisWorkbook.SaveCopyAs backupPath
    phase = "gravar as bases validadas"
    commitStarted = True
    For Each name In names
        Set target = SheetIfExists(ThisWorkbook, CStr(name))
        If target Is Nothing Then
            stage.Worksheets(CStr(name)).Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            newSheets.Add CStr(name)
        Else
            CommitSheet stage.Worksheets(CStr(name)), target
        End If
    Next name
    ThisWorkbook.Worksheets(DETAIL_SHEET).Calculate
    ThisWorkbook.Worksheets("Base").Calculate
    committed = True
    ListaAgendamento ThisWorkbook.Worksheets(DETAIL_SHEET), nDetail + 1
    restoredExtra = RestoreManualExtras(ThisWorkbook.Worksheets(DETAIL_SHEET), manualExtra, nDetail + 1)
    pivotMsg = RefreshDetailPivots(nDetail + 1)
    mLastMessage = nBase & " CT-es na Base; " & nDetail & " NFs na Base por NF Oficial." & vbCrLf & _
        "Modo acumulado: " & keepBas.Count & " CT-es anteriores (" & keepDet.Count & " NFs) mantidos sem alteração; " & _
        dRel.Count & " CT-es do 238 do dia novos ou atualizados." & vbCrLf & _
        (restored + restoredExtra) & " campos manuais preservados. " & (mExceptionRow - 2) & " exceções registradas." & vbCrLf & _
        filterMsg & vbCrLf & _
        "Frete rateado somente na Base por NF Oficial. OTIF pendente de validação." & vbCrLf & _
        pivotMsg & vbCrLf & "Backup: " & backupPath & vbCrLf & "Salve a pasta de trabalho para manter a atualização."
    GoTo Finally
Failed:
    errNum = Err.Number: errMsg = Err.Description
    On Error Resume Next
    If commitStarted And Not committed Then
        Set rollback = Workbooks.Open(Filename:=backupPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        If Not rollback Is Nothing Then
            For Each name In names
                Set target = Nothing: Set target = SheetIfExists(rollback, CStr(name))
                If Not target Is Nothing Then CommitSheet target, ThisWorkbook.Worksheets(CStr(name))
            Next name
            For Each name In newSheets
                ThisWorkbook.Worksheets(CStr(name)).Delete
            Next name
            CloseWorkbookSafe rollback
        End If
    End If
    On Error GoTo 0
    mLastMessage = "Atualização interrompida ao " & phase & "." & vbCrLf & "Erro " & errNum & ": " & errMsg
    If Not commitStarted Then mLastMessage = mLastMessage & vbCrLf & "Nenhuma base foi alterada."
    If commitStarted Then mLastMessage = mLastMessage & vbCrLf & "Cópia anterior disponível em: " & backupPath
Finally:
    On Error Resume Next
    MoverAgendaParaAE ThisWorkbook.Worksheets(DETAIL_SHEET)
    CloseWorkbookSafe src238
    CloseWorkbookSafe srcBI
    CloseWorkbookSafe stage
    mRelData = Empty: mBIData = Empty
    Application.AutomationSecurity = oldSecurity: Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.StatusBar = oldStatus
    On Error GoTo 0
    If Not silencioso Then
        If committed Then MsgBox mLastMessage, vbInformation, "Base WV atualizada" Else MsgBox mLastMessage, vbCritical, "Base WV"
    End If
End Sub

Private Sub CloseWorkbookSafe(ByRef wb As Workbook)
    'Some Excel/OneDrive combinations raise 1004 while closing a read-only
    'source workbook. Closing is cleanup and must never roll back the update.
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    wb.Saved = True
    wb.Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Err.Clear
    Set wb = Nothing
    On Error GoTo 0
End Sub

Private Function BackupFilePath(ByVal wb As Workbook) As String
    Dim folder As String, extension As String
    folder = wb.Path
    If LCase$(Left$(folder, 7)) = "http://" Or LCase$(Left$(folder, 8)) = "https://" Or Len(folder) = 0 Then
        folder = Environ$("TEMP")
    End If
    extension = Mid$(wb.Name, InStrRev(wb.Name, "."))
    If Len(extension) = 0 Then extension = ".xlsm"
    BackupFilePath = folder & Application.PathSeparator & "Backup_WV_" & _
        Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 100), "0") & extension
End Function

Private Function SheetIfExists(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetIfExists = wb.Worksheets(name)
    On Error GoTo 0
End Function

Private Sub ValidateFile(ByVal path As String)
    Dim book As Workbook
    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 601, , "Arquivo não encontrado: " & path
    For Each book In Application.Workbooks
        If LCase$(book.FullName) = LCase$(path) Then Err.Raise vbObjectError + 602, , "O arquivo selecionado já está aberto. Feche-o antes da importação: " & path
    Next book
End Sub

Private Sub ValidateModel(ByVal ws As Worksheet)
    Dim lastCol As Long
    If ws.ProtectContents Then Err.Raise vbObjectError + 603, , "A aba Base por NF Oficial está protegida."
    lastCol = LastUsedColumn(ws)
    If NormalizeHeader(ws.Cells(1, 1).Value2) <> "CTE" Or NormalizeHeader(ws.Cells(1, 3).Value2) <> "NOTA FISCAL" Then _
        Err.Raise vbObjectError + 604, , "Modelo incompatível: A deve ser Cte e C deve ser Nota fiscal."
    If EhNDN(ws.Cells(1, COL_NDN).Value2) Then Err.Raise vbObjectError + 605, , _
        "A coluna M ainda é N DN. Esta versão trabalha sem ela: apague a coluna M da aba " & ws.Name & "."
    If lastCol < WB_COLS Then Err.Raise vbObjectError + 607, , "A aba Base por NF Oficial deve ter pelo menos " & WB_COLS & " colunas. Foram encontradas " & lastCol & "."
    If NormalizeHeader(ws.Cells(1, WB_COLS).Value2) <> "OTIF" Then Err.Raise vbObjectError + 606, , "Modelo inválido: AT deve ser OTIF."
End Sub

Private Function EhNDN(ByVal v As Variant) As Boolean
    EhNDN = (Replace(NormalizeHeader(v), " ", "") = "NDN")
End Function

Private Sub EnsureDetailModel(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, tbl As ListObject
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)
    If lastCol = 46 Then
        'A planilha legada possui Transit Transportador em AL. A nova estrutura
        'insere Região em AL e desloca Transit e os indicadores até AU.
        ws.Columns("AL").Insert Shift:=xlToRight
        ws.Range("AM1:AM" & lastRow).Copy
        ws.Range("AL1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ws.Range("AL2:AL" & lastRow).ClearContents
    End If
    If LastUsedColumn(ws) <> 47 Then Err.Raise vbObjectError + 608, , "Não foi possível ampliar a Base por NF Oficial para 47 colunas."

    ws.Cells(1, 30).Value2 = "CLIENTE AGENDADO"
    ws.Cells(1, 31).Value2 = "CONTATO"
    ws.Cells(1, 32).Value2 = "AGENDAMENTO REALIZADO"
    ws.Cells(1, 33).Value2 = "Status"
    ws.Cells(1, 34).Value2 = "Data coleta transportador"
    ws.Cells(1, 35).Value2 = "POSSUE OCORRENCIA?"
    ws.Cells(1, 36).Value2 = "OBSERVAÇÃO"
    ws.Cells(1, 37).Value2 = "NOVA DATA PREV ENTREGA"
    ws.Cells(1, 38).Value2 = "Região"
    ws.Cells(1, 39).Value2 = "Transit Transportador"
    ws.Cells(1, 40).Value2 = "AGING EMISSÃO NF"
    ws.Cells(1, 41).Value2 = "AGING EMISSÃO CTE"
    ws.Cells(1, 42).Value2 = "CTE 24 HRS"
    ws.Cells(1, 43).Value2 = "OTD DO CLIENTE (previsão X DATA DE ENTREGA)"
    ws.Cells(1, 44).Value2 = "AGING DA PREVISÃO"
    ws.Cells(1, 45).Value2 = "WEEK DA PREVISÃO (SEMANA DO ANO)"
    ws.Cells(1, 46).Value2 = "WEEK DA ENTREGA (SEMANA DO ANO)"
    ws.Cells(1, 47).Value2 = "OTIF"
    For Each tbl In ws.ListObjects
        If tbl.Range.Row = 1 And tbl.Range.Column = 1 Then tbl.Resize ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 47))
    Next tbl
End Sub

Private Sub ImportReport(ByVal source As Workbook, ByVal dest As Worksheet, ByVal required As Variant)
    Dim ws As Worksheet, chosen As Worksheet, r As Long, c As Long, headerRow As Long, hits As Long, valid As Boolean
    Dim map As Object, item As Variant, maxR As Long, key As String
    For Each ws In source.Worksheets
        maxR = Application.Min(30, LastUsedRow(ws))
        For r = 1 To maxR
            Set map = CreateObject("Scripting.Dictionary")
            For c = 1 To LastUsedColumn(ws)
                key = NormalizeHeader(ws.Cells(r, c).Value2)
                If Len(key) > 0 Then map(key) = c
            Next c
            valid = True
            For Each item In required
                If Not map.Exists(NormalizeHeader(item)) Then valid = False
            Next item
            If valid Then
                hits = hits + 1: Set chosen = ws: headerRow = r
                Exit For
            End If
        Next r
    Next ws
    If hits <> 1 Then Err.Raise vbObjectError + 606, , "Não foi possível identificar uma única aba de " & dest.Name & " no arquivo " & source.Name & ". Cabeçalhos obrigatórios: " & Join(required, ", ")
    If LastUsedRow(chosen) <= headerRow Then Err.Raise vbObjectError + 607, , "O relatório " & dest.Name & " está vazio."
    dest.Range("A1").Resize(LastUsedRow(chosen) - headerRow + 1, LastUsedColumn(chosen)).Value2 = chosen.Range(chosen.Cells(headerRow, 1), chosen.Cells(LastUsedRow(chosen), LastUsedColumn(chosen))).Value2
End Sub

'Mantém só as linhas cujo cliente contém CLIENT_FILTER. Linha com cliente em branco também sai.
'Retorna quantas linhas saíram; -1 se a coluna não existe (nada filtrado); -2 se não sobrou nenhuma.
Private Function FilterClientRows(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim data As Variant, result() As Variant, col As Long, c As Long, r As Long, kept As Long
    Dim lastR As Long, lastC As Long
    lastR = LastUsedRow(ws): lastC = LastUsedColumn(ws)
    If lastR < 2 Then FilterClientRows = -2: Exit Function
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value2
    For c = 1 To lastC
        If NormalizeHeader(data(1, c)) = NormalizeHeader(header) Then col = c: Exit For
    Next c
    If col = 0 Then FilterClientRows = -1: Exit Function
    ReDim result(1 To lastR, 1 To lastC)
    For c = 1 To lastC: result(1, c) = data(1, c): Next c
    kept = 1
    For r = 2 To lastR
        If InStr(1, TextValue(data(r, col)), CLIENT_FILTER, vbTextCompare) > 0 Then
            kept = kept + 1
            For c = 1 To lastC: result(kept, c) = data(r, c): Next c
        End If
    Next r
    If kept = 1 Then FilterClientRows = -2: Exit Function
    FilterClientRows = lastR - kept
    If FilterClientRows > 0 Then
        ws.Cells.ClearContents
        ws.Range("A1").Resize(lastR, lastC).Value2 = result
    End If
End Function

Private Function FilterText(ByVal removed As Long, ByVal report As String, ByVal header As String) As String
    If removed = -1 Then
        FilterText = report & " sem a coluna " & header & ", sem filtro"
    ElseIf removed = 0 Then
        FilterText = report & " já vinha só com o cliente"
    Else
        FilterText = removed & " linhas de outros clientes descartadas do " & report
    End If
End Function

Private Sub CacheSheet(ByVal ws As Worksheet)
    If ws.Name = "REL_238" Then
        mRelData = ws.Range(ws.Cells(1, 1), ws.Cells(LastUsedRow(ws), LastUsedColumn(ws))).Value2
    Else
        mBIData = ws.Range(ws.Cells(1, 1), ws.Cells(LastUsedRow(ws), LastUsedColumn(ws))).Value2
    End If
End Sub

Private Sub ValidateRows(ByVal ws As Worksheet, ByVal headers As Object, ByVal isBI As Boolean)
    Dim r As Long, item As Variant, v As Variant, fields As Variant, id As String, dates As Variant
    If isBI Then
        fields = Array("VLR_MERC_NF", "QTD_VOL", "PESO_BRUTO_CTE", "PESO_CUBADO_CTE", "VLR_PRESTACAO_CTE", "VLR_MERCADORIA_CTE")
        dates = Array("DT_EMISSAO_NF", "DT_INICIO_TRANSP", "DT_ENTREGA", "DT_BAIXA", "DT_UPLOAD")
    Else
        fields = Array(): dates = Array("DT_EMISSAO", "DT_RECEBIMENTO", "DT_ENTREGA_CTRC", "DT_BAIXA", "DATA_DE_UPLOAD")
    End If
    For r = 2 To LastUsedRow(ws)
        id = CleanId(CellByHeader(ws, r, headers, "NR_CTRC"))
        If Len(id) = 0 Then
            If Application.CountA(ws.Rows(r)) > 0 Then Err.Raise vbObjectError + 608, , ws.Name & " linha " & r & ": CT-e vazio."
        Else
            If isBI Then
                If Len(CleanId(CellByHeader(ws, r, headers, "NR_NF"))) = 0 Then Err.Raise vbObjectError + 609, , ws.Name & " linha " & r & ": NF vazia."
            End If
            For Each item In fields
                v = CellByHeader(ws, r, headers, CStr(item))
                If IsError(v) Then Err.Raise vbObjectError + 610, , ws.Name & " linha " & r & ": erro em " & item
                If Not IsNumeric(v) Or Len(TextValue(v)) = 0 Then Err.Raise vbObjectError + 611, , ws.Name & " linha " & r & ": valor ausente ou inválido em " & item
                If CDbl(v) < 0 Then Err.Raise vbObjectError + 612, , ws.Name & " linha " & r & ": valor negativo em " & item
            Next item
            For Each item In dates
                v = CellByHeader(ws, r, headers, CStr(item))
                If Len(TextValue(v)) > 0 Then
                    If Not IsDate(ParseDateValue(v)) Then Err.Raise vbObjectError + 613, , ws.Name & " linha " & r & ": data inválida em " & item
                End If
            Next item
        End If
    Next r
End Sub

Private Function TransitForWorkbook(ByVal wb As Workbook, ByVal detail As Worksheet) As Object
    Dim ws As Worksheet, result As Object, r As Long, key As String, v As Variant
    Set ws = SheetIfExists(wb, "Transit time")
    If Not ws Is Nothing Then
        Set TransitForWorkbook = BuildTransitDictionary(ws): Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    For r = 2 To LastUsedRow(detail)
        If Len(TextValue(detail.Cells(r, 1).Value2)) > 0 Then
            key = NormalizeText(detail.Cells(r, 14).Value2 & "_" & detail.Cells(r, 13).Value2)
            v = detail.Cells(r, 19).Value2
            If Len(TextValue(v)) > 0 And IsNumeric(v) Then
                If result.Exists(key) Then
                    If CDbl(result(key)) <> CDbl(v) Then Err.Raise vbObjectError + 614, , "Prazo BID conflitante para " & key
                End If
                result(key) = CDbl(v)
            End If
        End If
    Next r
    Set TransitForWorkbook = result
End Function

Private Function CaptureManual(ByVal ws As Worksheet) As Object
    Dim result As Object, rowData As Object, r As Long, col As Variant, key As String, wbCol As Long
    Set result = CreateObject("Scripting.Dictionary")
    'Pela coluna A: colunas adicionais podem ter fórmula preenchida até o fim da planilha.
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Len(TextValue(ws.Cells(r, 1).Value2)) > 0 Then
            key = CleanId(ws.Cells(r, 1).Value2) & "|" & CleanNFDisplay(CStr(ws.Cells(r, 3).Value2))
            If result.Exists(key) Then Err.Raise vbObjectError + 615, , "Chave CT-e + NF duplicada na Base por NF Oficial: " & key
            Set rowData = CreateObject("Scripting.Dictionary")
            For Each col In Array(25, 26, 27, 32, 33, 34, 35, 36, 37, 47)
                wbCol = CLng(col) - 1   'na planilha, sem a coluna M
                If Not ws.Cells(r, wbCol).HasFormula Then
                    If Len(TextValue(ws.Cells(r, wbCol).Value2)) > 0 Then rowData.Add CStr(col), ws.Cells(r, wbCol).Value2
                End If
            Next col
            result.Add key, rowData
        End If
    Next r
    Set CaptureManual = result
End Function

Private Function RestoreManual(ByVal ws As Worksheet, ByVal manual As Object, ByVal ex As Worksheet) As Long
    Dim r As Long, key As String, col As Variant, rowData As Object, used As Object, k As Variant
    Set used = CreateObject("Scripting.Dictionary")
    For r = 2 To LastUsedRow(ws)
        key = CleanId(ws.Cells(r, 1).Value2) & "|" & CleanNFDisplay(CStr(ws.Cells(r, 3).Value2))
        If manual.Exists(key) Then
            Set rowData = manual(key): used(key) = True
            For Each col In rowData.Keys
                'Cod e Justificativa: o relatório vence; o valor anterior só volta se ele vier em branco.
                If (CLng(col) <> 26 And CLng(col) <> 27) Or Len(TextValue(ws.Cells(r, CLng(col)).Value2)) = 0 Then
                    ws.Cells(r, CLng(col)).Value2 = rowData(col)
                    RestoreManual = RestoreManual + 1
                End If
            Next col
        End If
    Next r
    For Each k In manual.Keys
        Set rowData = manual(k)
        If rowData.Count > 0 And Not used.Exists(k) Then LogException ex, "MANUAL SEM CORRESPONDÊNCIA", Split(CStr(k), "|")(0), Split(CStr(k), "|")(1), "Campos manuais disponíveis na cópia de segurança. Registro ausente da nova base.", "Base por NF Oficial anterior"
    Next k
End Function

'Linhas da planilha (sem N DN) cujo CT-e não veio no 238 do dia. Na Base por NF Oficial, um CT-e que
'voltou sem nenhuma NF também mantém as linhas antigas. Devolve os números das linhas e os valores em data.
' Colunas depois de AT com cabeçalho e sem fórmula na linha 2 são digitadas pelo analista (ex.: DATA AGENDADA).
' Guarda o valor de cada nota pela chave CT-e|NF, para regravar depois que a base é remontada.
Private Function CaptureManualExtras(ByVal ws As Worksheet) As Object
    Dim result As Object, colData As Object, c As Long, r As Long, lastRow As Long, lastCol As Long, k As String
    Set result = CreateObject("Scripting.Dictionary")
    Set CaptureManualExtras = result
    lastCol = LastUsedColumn(ws)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For c = WB_COLS + 1 To lastCol
        If NormalizeHeader(ws.Cells(1, c).Value2) = "DATA DO AGENDAMENTO" And Not ws.Cells(2, c).HasFormula Then
            Set colData = CreateObject("Scripting.Dictionary")
            For r = 2 To lastRow
                If Len(TextValue(ws.Cells(r, 1).Value2)) > 0 And Not ws.Cells(r, c).HasFormula Then
                    If Len(TextValue(ws.Cells(r, c).Value2)) > 0 Then
                        k = CleanId(ws.Cells(r, 1).Value2) & "|" & CleanNFDisplay(CStr(ws.Cells(r, 3).Value2))
                        colData(k) = ws.Cells(r, c).Value2
                    End If
                End If
            Next r
            result.Add c, colData
        End If
    Next c
End Function

Private Function RestoreManualExtras(ByVal ws As Worksheet, ByVal extras As Object, ByVal lastRow As Long) As Long
    Dim c As Variant, colData As Object, r As Long, k As String, maxRow As Long
    maxRow = Application.Max(lastRow, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    For Each c In extras.Keys
        Set colData = extras(c)
        ws.Range(ws.Cells(2, CLng(c)), ws.Cells(maxRow, CLng(c))).ClearContents
        For r = 2 To lastRow
            k = CleanId(ws.Cells(r, 1).Value2) & "|" & CleanNFDisplay(CStr(ws.Cells(r, 3).Value2))
            If colData.Exists(k) Then
                ws.Cells(r, CLng(c)).Value2 = colData(k)
                RestoreManualExtras = RestoreManualExtras + 1
            End If
        Next r
    Next c
End Function

Private Function OldRowsToKeep(ByVal ws As Worksheet, ByVal nCols As Long, ByVal dRel As Object, _
    ByVal dIndex As Object, ByRef data As Variant) As Collection
    Dim result As Collection, lastRow As Long, r As Long, cte As String, seen As Object, k As String
    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set OldRowsToKeep = result
    If lastRow < 2 Then Exit Function
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, nCols)).Value2
    For r = 2 To lastRow
        cte = CleanId(data(r, 1))
        If Len(cte) > 0 Then
            If Not dRel.Exists(cte) Then
                result.Add r
            ElseIf Not dIndex Is Nothing Then
                If Not dIndex.Exists(cte) Then result.Add r
            End If
            If dIndex Is Nothing Then k = cte Else k = cte & "|" & CleanNFDisplay(CStr(data(r, 3)))
            If seen.Exists(k) Then Err.Raise vbObjectError + 617, , "Registro duplicado em " & ws.Name & " (linha " & r & "): " & k & ". Apague a linha repetida e rode de novo."
            seen.Add k, True
        End If
    Next r
End Function

'Copia as linhas mantidas para o fim do array de saída (inserindo a coluna N DN vazia).
'Devolve, por linha da planilha de trabalho, a Data NF (coluna D) antiga para regravar depois das fórmulas.
Private Function AppendOldRows(ByRef outputData As Variant, ByRef outRow As Long, ByRef data As Variant, _
    ByVal rowsToKeep As Collection, ByVal nCols As Long) As Object
    Dim d As Object, r As Variant, c As Long, sc As Long, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each r In rowsToKeep
        outRow = outRow + 1
        For c = 1 To nCols
            v = data(CLng(r), c)
            If IsError(v) Then v = Empty
            If c >= COL_NDN Then sc = c + 1 Else sc = c
            outputData(outRow, sc) = v
        Next c
        If Not IsEmpty(outputData(outRow, 4)) Then d(outRow + 1) = outputData(outRow, 4)
    Next r
    Set AppendOldRows = d
End Function

Private Function ColunaCab(ByVal ws As Worksheet, ByVal nome As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If NormalizeHeader(ws.Cells(1, c).Value2) = nome Then ColunaCab = c: Exit Function
    Next c
End Function

'As duas colunas do agendamento só se movem juntas e quando estão lado a lado.
Private Sub MoverAgendaParaFim(ByVal ws As Worksheet)
    Dim c1 As Long, c2 As Long, ultima As Long, ev As Boolean
    c1 = ColunaCab(ws, "DATA DO AGENDAMENTO"): c2 = ColunaCab(ws, "AGENDADO EM 24H")
    If c1 = 0 Or c2 <> c1 + 1 Or c1 > WB_COLS Then Exit Sub
    'Com filtro aplicado o Excel não recorta colunas inteiras (erro 1004 de vários intervalos).
    If ws.FilterMode Then ws.ShowAllData
    ultima = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ev = Application.EnableEvents: Application.EnableEvents = False
    ws.Range(ws.Columns(c1), ws.Columns(c2)).Cut
    ws.Columns(ultima + 1).Insert Shift:=xlToRight
    Application.CutCopyMode = False
    Application.EnableEvents = ev
End Sub

Private Sub MoverAgendaParaAE(ByVal ws As Worksheet)
    Dim c1 As Long, c2 As Long, ev As Boolean
    c1 = ColunaCab(ws, "DATA DO AGENDAMENTO"): c2 = ColunaCab(ws, "AGENDADO EM 24H")
    If c1 = 0 Or c2 <> c1 + 1 Or c1 <= WB_COLS Then Exit Sub
    If NormalizeHeader(ws.Cells(1, 31).Value2) <> "AGENDAMENTO REALIZADO" Then Exit Sub
    ev = Application.EnableEvents: Application.EnableEvents = False
    ws.Range(ws.Columns(c1), ws.Columns(c2)).Cut
    ws.Columns(32).Insert Shift:=xlToRight
    Application.CutCopyMode = False
    Application.EnableEvents = ev
End Sub

Private Sub ListaAgendamento(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range("AE2:AE" & Application.Max(lastRow, 2))
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=listaSimNao"
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

Private Sub ClearData(ByVal ws As Worksheet)
    If LastUsedRow(ws) > 1 Then ws.Range(ws.Cells(2, 1), ws.Cells(LastUsedRow(ws), LastUsedColumn(ws))).ClearContents
End Sub

Private Sub WriteData(ByVal ws As Worksheet, ByRef data As Variant, ByVal n As Long, ByVal cols As Long)
    Dim tbl As ListObject
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, cols)).NumberFormat = "General"
    ws.Range("A2:A" & n + 1).NumberFormat = "@": ws.Range("C2:C" & n + 1).NumberFormat = "@"
    ws.Range("I2:I" & n + 1).NumberFormat = "@": ws.Range("L2:M" & n + 1).NumberFormat = "@"
    ws.Range("A2").Resize(n, cols).Value = data
    ws.Range("B2:B" & n + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("D2:D" & n + 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("F2:F" & n + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("U2:V" & n + 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("E2:E" & n + 1).NumberFormat = "#,##0.00"
    ws.Range("Q2:S" & n + 1).NumberFormat = "#,##0.00"
    If cols = 47 Then
        ws.Range("AH2:AH" & n + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        ws.Range("AK2:AK" & n + 1).NumberFormat = "dd/mm/yyyy"
        ws.Range("AN2:AO" & n + 1).NumberFormat = "0"
        ws.Range("AR2:AT" & n + 1).NumberFormat = "0"
        ws.Range("AP2:AQ" & n + 1).WrapText = True
        ws.Range("AU2:AU" & n + 1).WrapText = True
    End If
    For Each tbl In ws.ListObjects
        If tbl.Range.Row = 1 And tbl.Range.Column = 1 Then tbl.Resize ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, cols))
    Next tbl
End Sub

Private Sub FormulaDown(ByVal ws As Worksheet, ByVal col As String, ByVal lastRow As Long, ByVal formula As String)
    If lastRow < 2 Then Exit Sub
    ws.Range(col & "2").Formula = formula
    If lastRow > 2 Then ws.Range(col & "2:" & col & lastRow).FillDown
End Sub

Private Sub ApplyCoreFormulas(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws.Name = "Base" Then
        'CT-e identifiers are imported as text; tolerate the numeric type used by BI_2749.
        'A missing lookup stays blank so an empty/raw 2749 sheet does not abort the update.
        Dim biMap As Object, cteRef As String, dateRef As String
        Set biMap = BuildHeaderMap(ws.Parent.Worksheets("BI_2749"))
        cteRef = BIColumnRef(ws.Parent.Worksheets("BI_2749"), biMap("NR_CTRC"))
        dateRef = BIColumnRef(ws.Parent.Worksheets("BI_2749"), biMap("DT_EMISSAO_NF"))
        FormulaDown ws, "D", lastRow, "=IFERROR(INDEX(" & dateRef & ",MATCH(A2," & cteRef & ",0)),IFERROR(INDEX(" & dateRef & ",MATCH(VALUE(A2)," & cteRef & ",0)),""""))"
    End If
    FormulaDown ws, "T", lastRow, "=IF(A2="""","""",IF(ISNUMBER(SEARCH(""AERE"",J2)),IFERROR(1*VLOOKUP(O2&IF(COUNTIFS('TT Aereo'!$K:$K,O2,'TT Aereo'!$L:$L,N2)>0,""C"",""I""),'TT Aereo'!$A:$E,5,0),""""),IFERROR(1*VLOOKUP(N2&O2,'TT LTL'!$C:$D,2,0),IFERROR(1*VLOOKUP(N2&O2,'Transit time'!$B:$C,2,0),""""))))"
    FormulaDown ws, "U", lastRow, "=IF(A2="""","""",IF(OR(B2="""",T2=""""),"""",WORKDAY.INTL(B2,T2,1," & HolidayRef(ws.Parent.Worksheets("Feriados")) & ")))"
    FormulaDown ws, "W", lastRow, "=IF(A2="""","""",IF(U2="""",""SEM PREVISÃO"",IF(V2="""",IF(TODAY()<=INT(U2),""EM TRÂNSITO NO PRAZO"",""EM TRÂNSITO FORA DO PRAZO""),IF(INT(V2)<=INT(U2),""NÃO"",""SIM""))))"
    FormulaDown ws, "X", lastRow, "=IF(A2="""","""",IF(U2="""","""",MAX(0,IF(V2="""",TODAY(),INT(V2))-INT(U2))))"
    If ws.Name = DETAIL_SHEET Then
        'Cliente que agendou para depois da previsão: o prazo é a data agendada.
        FormulaDown ws, "W", lastRow, "=IF(A2="""","""",IF(U2="""",""SEM PREVISÃO"",IF(V2="""",IF(TODAY()<=" & PRAZO_AG & ",""EM TRÂNSITO NO PRAZO"",""EM TRÂNSITO FORA DO PRAZO""),IF(INT(V2)<=" & PRAZO_AG & ",""NÃO"",""SIM""))))"
        FormulaDown ws, "X", lastRow, "=IF(A2="""","""",IF(U2="""","""",MAX(0,IF(V2="""",TODAY(),INT(V2))-" & PRAZO_AG & ")))"
        'Letras da cópia de trabalho (com N DN): AD = CLIENTE AGENDADO, AE = CONTATO. K = Nome do DN, L = CGC DEST.
        FormulaDown ws, "AD", lastRow, "=IF(A2="""","""",IF(ISNUMBER(" & AGENDA_POS & "),""SIM"",""NÃO""))"
        FormulaDown ws, "AE", lastRow, "=IF(AD2<>""SIM"","""",IFERROR(INDEX(Agendamentos!$N:$N," & AGENDA_POS & ")&"""",""""))"
    Else
        FormulaDown ws, "AD", lastRow, "=IF(A2="""","""",IF(B2="""","""",YEAR(B2)))"
        FormulaDown ws, "AE", lastRow, "=IF(A2="""","""",IF(B2="""","""",MONTH(B2)))"
    End If
End Sub

'Datas contíguas a partir de A2. A observação no fim da aba não entra no intervalo.
Private Function HolidayRef(ByVal ws As Worksheet) As String
    Dim r As Long
    r = 2
    Do While IsDate(ws.Cells(r + 1, 1).Value) And Not IsEmpty(ws.Cells(r + 1, 1).Value)
        r = r + 1
    Loop
    HolidayRef = "Feriados!$A$2:$A$" & r
End Function

Private Function BIColumnRef(ByVal bi As Worksheet, ByVal col As Long) As String
    Dim letter As String
    letter = Split(bi.Cells(1, col).Address(True, False), "$")(0)
    BIColumnRef = "BI_2749!$" & letter & ":$" & letter
End Function

Private Sub ApplyDetailFormulas(ByVal ws As Worksheet, ByVal lastRow As Long)
    FormulaDown ws, "AL", lastRow, "=IF(A2="""","""",IF(O2="""","""",IF(OR(O2=""AC"",O2=""AP"",O2=""AM"",O2=""PA"",O2=""RO"",O2=""RR"",O2=""TO""),""NORTE"",IF(OR(O2=""AL"",O2=""BA"",O2=""CE"",O2=""MA"",O2=""PB"",O2=""PE"",O2=""PI"",O2=""RN"",O2=""SE""),""NORDESTE"",IF(OR(O2=""DF"",O2=""GO"",O2=""MT"",O2=""MS""),""CENTRO-OESTE"",IF(OR(O2=""ES"",O2=""MG"",O2=""RJ"",O2=""SP""),""SUDESTE"",IF(OR(O2=""PR"",O2=""RS"",O2=""SC""),""SUL"",""UF INVÁLIDA"")))))))"
    FormulaDown ws, "AM", lastRow, "=IF(A2="""","""",IF(T2="""","""",T2))"
    FormulaDown ws, "AN", lastRow, "=IF(A2="""","""",IF(D2="""","""",IF(ISNUMBER(D2),TODAY()-INT(D2),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AO", lastRow, "=IF(A2="""","""",IF(B2="""","""",IF(ISNUMBER(B2),TODAY()-INT(B2),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AP", lastRow, "=IF(A2="""","""",IF(OR(B2="""",D2=""""),""SEM DATA"",IF(AND(ISNUMBER(B2),ISNUMBER(D2)),IF(B2<D2,""DATAS INCONSISTENTES"",IF(B2-D2<=1,""ATÉ 24 HRS"",""ACIMA DE 24 HRS"")),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AQ", lastRow, "=IF(A2="""","""",IF(U2="""",""SEM PREVISÃO"",IF(ISNUMBER(U2),IF(V2="""",IF(TODAY()<=" & PRAZO_AG & ",""EM TRÂNSITO NO PRAZO"",""EM TRÂNSITO FORA DO PRAZO""),IF(ISNUMBER(V2),IF(INT(V2)<=" & PRAZO_AG & ",""OTD"",""FORA OTD""),""DATA INVÁLIDA"")),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AR", lastRow, "=IF(A2="""","""",IF(U2="""","""",IF(ISNUMBER(U2),INT(U2)-TODAY(),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AS", lastRow, "=IF(A2="""","""",IF(U2="""","""",IF(ISNUMBER(U2),WEEKNUM(U2,21),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AT", lastRow, "=IF(A2="""","""",IF(V2="""","""",IF(ISNUMBER(V2),WEEKNUM(V2,21),""DATA INVÁLIDA"")))"
    FormulaDown ws, "AU", lastRow, "=IF(A2="""","""",""PENDENTE DE VALIDAÇÃO"")"
End Sub

Private Sub ValidateFormulaResults(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim errors As Range, cell As Range, firstBad As String
    On Error Resume Next
    Set errors = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, LastUsedColumn(ws))).SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo 0
    If Not errors Is Nothing Then
        For Each cell In errors.Cells
            'Data.nf is a lookup against the selected 2749. A missing key is
            'allowed here; all calculated Base por NF Oficial indicators remain strict.
            If Not (ws.Name = "Base" And cell.Column = 4) Then
                firstBad = cell.Address(False, False)
                Exit For
            End If
        Next cell
        If Len(firstBad) > 0 Then Err.Raise vbObjectError + 616, , "Erro de fórmula em " & ws.Name & "!" & firstBad
    End If
End Sub

Private Sub CommitSheet(ByVal source As Worksheet, ByVal dest As Worksheet)
    Dim rows As Long, cols As Long, oldRows As Long, keepExtras As Boolean, tbl As ListObject
    rows = LastUsedRow(source): cols = LastUsedColumn(source)
    If dest.Name = DETAIL_SHEET Then AlignDetailWidth dest, cols
    oldRows = LastUsedRow(dest)
    'A cópia de trabalho tem só A:AU. Na restauração do backup vêm todas as colunas.
    keepExtras = (dest.Name = DETAIL_SHEET And cols = WB_COLS And LastUsedColumn(dest) > WB_COLS)
    If keepExtras Then
        dest.Range(dest.Cells(1, 1), dest.Cells(oldRows, WB_COLS)).ClearContents
    Else
        dest.UsedRange.ClearContents
    End If
    For Each tbl In dest.ListObjects
        If tbl.Range.Row = 1 And tbl.Range.Column = 1 Then tbl.Resize dest.Range(dest.Cells(1, 1), dest.Cells(rows, cols))
    Next tbl
    source.Range(source.Cells(1, 1), source.Cells(rows, cols)).Copy
    dest.Range("A1").PasteSpecial xlPasteFormats
    dest.Range("A1").Resize(rows, cols).Formula = source.Range(source.Cells(1, 1), source.Cells(rows, cols)).Formula
    Application.CutCopyMode = False
    If keepExtras Then ExtendExtraColumns dest, rows, oldRows
End Sub

Private Sub AlignDetailWidth(ByVal ws As Worksheet, ByVal expectedColumns As Long)
    If LastUsedColumn(ws) >= WB_COLS And expectedColumns >= WB_COLS Then Exit Sub
    Err.Raise vbObjectError + 609, , "Não é possível alinhar a Base por NF Oficial de " & LastUsedColumn(ws) & " para " & expectedColumns & " colunas."
End Sub

'Monta a cópia de trabalho só com cabeçalho e formatos de A:AU. Copiar a aba inteira trava o
'Excel quando há colunas adicionais com fórmula preenchida até o fim da planilha.
Private Function CopyDetailToStage(ByVal source As Worksheet, ByVal stage As Workbook, ByVal nDetail As Long) As Worksheet
    Dim dest As Worksheet, lastCol As Long, lastRow As Long
    lastCol = Application.Min(LastUsedColumn(source), WB_COLS)
    lastRow = Application.Max(2, source.Cells(source.Rows.Count, 1).End(xlUp).Row)
    Set dest = stage.Worksheets.Add(After:=stage.Worksheets(stage.Worksheets.Count))
    dest.Name = DETAIL_SHEET
    source.Range(source.Cells(1, 1), source.Cells(lastRow, lastCol)).Copy
    dest.Range("A1").PasteSpecial xlPasteFormats
    dest.Range("A1").Resize(1, lastCol).Value2 = source.Range(source.Cells(1, 1), source.Cells(1, lastCol)).Value2
    dest.Range("A2").Resize(1, lastCol).Copy
    dest.Range("A2").Resize(nDetail, lastCol).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    dest.Columns(COL_NDN).Insert Shift:=xlToRight
    dest.Cells(1, COL_NDN).Value2 = "N  DN"
    Set CopyDetailToStage = dest
End Function

'Usa a fórmula da linha 2 como modelo de cada coluna adicional.
Private Sub ExtendExtraColumns(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal oldRows As Long)
    Dim c As Long, lastCol As Long, maxRow As Long
    lastCol = LastUsedColumn(ws)
    maxRow = Application.Max(lastRow, oldRows)
    For c = WB_COLS + 1 To lastCol
        If Len(TextValue(ws.Cells(1, c).Value2)) = 0 And InStr(1, CStr(ws.Cells(2, c).Formula), "#REF!", vbTextCompare) > 0 Then
            ws.Range(ws.Cells(2, c), ws.Cells(maxRow, c)).ClearContents
        Else
            'Sem isto, entrega vazia vira 0 e o ATRASO2 acusa ENTREGUE NO PRAZO.
            If NormalizeHeader(ws.Cells(1, c).Value2) = "DATA DA ENTREGA2" Then ws.Cells(2, c).Formula = "=IF(V2="""","""",V2)"
            If oldRows > lastRow Then ws.Range(ws.Cells(lastRow + 1, c), ws.Cells(oldRows, c)).ClearContents
            If ws.Cells(2, c).HasFormula And lastRow > 2 Then ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).FillDown
        End If
    Next c
End Sub

'Falha aqui não desfaz a base já gravada: vira aviso na mensagem final.
Private Function RefreshDetailPivots(ByVal lastRow As Long) As String
    Dim ws As Worksheet, detail As Worksheet, pt As PivotTable, cache As PivotCache
    Dim lastCol As Long, ok As Long, failed As String
    On Error Resume Next
    Application.Calculate
    Set detail = ThisWorkbook.Worksheets(DETAIL_SHEET)
    lastCol = detail.Cells(1, detail.Columns.Count).End(xlToLeft).Column
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            If InStr(1, pt.SourceData, DETAIL_SHEET, vbTextCompare) > 0 Then
                Err.Clear
                'Um cache para cada tabela: compartilhar o mesmo falhou no teste.
                Set cache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=detail.Range(detail.Cells(1, 1), detail.Cells(lastRow, lastCol)))
                pt.ChangePivotCache cache
                pt.RefreshTable
                If Err.Number = 0 Then
                    ok = ok + 1
                Else
                    failed = failed & " " & ws.Name & "!" & pt.Name
                End If
            End If
        Next pt
    Next ws
    On Error GoTo 0
    RefreshDetailPivots = ok & " tabelas dinâmicas atualizadas."
    If Len(failed) > 0 Then RefreshDetailPivots = RefreshDetailPivots & " ATENÇÃO, não atualizadas:" & failed & ". Use Dados > Atualizar Tudo."
End Function

Private Function BuildBICTEDictionary(ByVal ws As Worksheet, ByVal headers As Object, ByVal dBI As Object) As Object
    Dim result As Object, key As Variant, record As Variant, aggregate As Variant, cte As String, col As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In dBI.Keys
        record = dBI(key): cte = CStr(record(0))
        If result.Exists(cte) Then
            aggregate = result(cte)
            For Each col In Array(3, 10, 11, 12)
                If Abs(CDbl(aggregate(col)) - CDbl(record(col))) > 0.000001 Then Err.Raise vbObjectError + 624, , "Total de CT-e divergente entre NFs: CT-e " & cte & ". Confira mercadoria, peso e frete no BI."
            Next col
            aggregate(9) = CDbl(aggregate(9)) + CDbl(record(9))
            MergeDateMin aggregate, 2, record(2)
            MergeDateMin aggregate, 4, record(4)
            MergeDateMax aggregate, 19, record(19)
            result(cte) = aggregate
        Else
            result.Add cte, record
        End If
    Next key
    Set BuildBICTEDictionary = result
End Function


Private Function Build238Dictionary(ByVal ws As Worksheet, ByVal headers As Object) As Object
    Dim result As Object
    Dim lastRow As Long, rowNumber As Long
    Dim cte As String, record As Variant
    Dim currentDate As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    lastRow = LastUsedRow(ws)

    For rowNumber = 2 To lastRow
        cte = CleanId(CellByHeader(ws, rowNumber, headers, "NR_CTRC"))
        If Len(cte) > 0 Then
            If result.Exists(cte) Then
                record = result(cte)
            Else
                ReDim record(0 To 19)
                record(0) = cte
            End If

            MergeDateMin record, 1, CellByHeader(ws, rowNumber, headers, "DT_EMISSAO")
            MergeDateMax record, 2, CellByHeader(ws, rowNumber, headers, "DT_ENTREGA_CTRC")
            MergeDateMax record, 3, CellByHeader(ws, rowNumber, headers, "DT_BAIXA")
            MergeText record, 4, CellByHeader(ws, rowNumber, headers, "CD_OCORR_ENTREGA")
            MergeText record, 5, CellByHeader(ws, rowNumber, headers, "DESC_OCORR_ENTREGA")
            MergeText record, 6, CellByHeader(ws, rowNumber, headers, "RAZA_SOCIAL")
            MergeDateMin record, 7, CellByHeader(ws, rowNumber, headers, "DT_RECEBIMENTO")
            MergeText record, 8, CellByHeader(ws, rowNumber, headers, "DESC_OPERACAO")
            If Len(TextValue(record(8))) = 0 Then MergeText record, 8, CellByHeader(ws, rowNumber, headers, "DESC_TP_VIAGEM")
            MergeText record, 9, CellByHeader(ws, rowNumber, headers, "TRANSIT_TIME")
            MergeDateMax record, 10, CellByHeader(ws, rowNumber, headers, "PREV_ENTREGA")
            MergeText record, 11, CellByHeader(ws, rowNumber, headers, "TIPO_UPLOAD")
            MergeDateMax record, 12, CellByHeader(ws, rowNumber, headers, "DATA_DE_UPLOAD")
            MergeText record, 13, CellByHeader(ws, rowNumber, headers, "NOME_DEST")
            MergeText record, 14, CellByHeader(ws, rowNumber, headers, "UF_DEST")
            MergeText record, 15, CellByHeader(ws, rowNumber, headers, "CIDADE_DEST")
            MergeText record, 16, CellByHeader(ws, rowNumber, headers, "NUMERO_DA_NF")
            MergeText record, 17, CellByHeader(ws, rowNumber, headers, "OBS_BAIXA")
            MergeText record, 18, CellByHeader(ws, rowNumber, headers, "RAZAO_SOCIAL_CLIENTE")

            currentDate = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_OCOREENCIA"))
            If IsDate(currentDate) Then
                If Not IsDate(record(19)) Then
                    record(19) = currentDate
                ElseIf CDbl(CDate(currentDate)) > CDbl(CDate(record(19))) Then
                    record(19) = currentDate
                End If
            End If

            If result.Exists(cte) Then
                result(cte) = record
            Else
                result.Add cte, record
            End If
        End If
    Next rowNumber

    Set Build238Dictionary = result
End Function

Private Function Build238NFDictionary(ByVal ws As Worksheet, ByVal headers As Object) As Object
    Dim result As Object
    Dim lastRow As Long, rowNumber As Long
    Dim cte As String, rawList As String, nf As Variant, key As String
    Dim nfs As Collection

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    lastRow = LastUsedRow(ws)

    For rowNumber = 2 To lastRow
        cte = CleanId(CellByHeader(ws, rowNumber, headers, "NR_CTRC"))
        rawList = TextValue(CellByHeader(ws, rowNumber, headers, "NUMERO_DA_NF"))
        If Len(cte) > 0 And Len(rawList) > 0 Then
            Set nfs = ExtractNFList(rawList)
            For Each nf In nfs
                key = cte & "|" & NormalizeNFKey(CStr(nf))
                If Len(NormalizeNFKey(CStr(nf))) > 0 Then
                    If result.Exists(key) Then
                        If CleanNFDisplay(CStr(result(key))) <> CleanNFDisplay(CStr(nf)) Then Err.Raise vbObjectError + 621, , "NF com séries diferentes no mesmo CT-e: " & key
                    Else
                        result.Add key, CleanNFDisplay(CStr(nf))
                    End If
                End If
            Next nf
        End If
    Next rowNumber

    Set Build238NFDictionary = result
End Function

Private Function BuildBIDictionary(ByVal ws As Worksheet, ByVal headers As Object) As Object
    Dim result As Object
    Dim lastRow As Long, rowNumber As Long
    Dim cte As String, nf As String, key As String
    Dim record As Variant, priorRecord As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    lastRow = LastUsedRow(ws)

    For rowNumber = 2 To lastRow
        cte = CleanId(CellByHeader(ws, rowNumber, headers, "NR_CTRC"))
        nf = CleanId(CellByHeader(ws, rowNumber, headers, "NR_NF"))
        If Len(cte) > 0 And Len(nf) > 0 Then
            key = cte & "|" & NormalizeNFKey(nf)
            record = ReadBIRecord(ws, rowNumber, headers, cte, nf)
            If result.Exists(key) Then
                priorRecord = result(key)
                If DateScore(record(13)) >= DateScore(priorRecord(13)) Then result(key) = record
            Else
                result.Add key, record
            End If
        End If
    Next rowNumber

    Set BuildBIDictionary = result
End Function

Private Function ReadBIRecord(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headers As Object, _
                              ByVal cte As String, ByVal nf As String) As Variant
    Dim record(0 To 24) As Variant
    record(0) = cte
    record(1) = nf
    record(2) = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_EMISSAO_NF"))
    record(3) = CellByHeader(ws, rowNumber, headers, "VLR_MERCADORIA_CTE")
    record(4) = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_INICIO_TRANSP"))
    record(5) = TextValue(CellByHeader(ws, rowNumber, headers, "NOME_REMETENTE"))
    record(6) = TextValue(CellByHeader(ws, rowNumber, headers, "CD_UF_REMETENTE"))
    record(7) = DocumentId(CellByHeader(ws, rowNumber, headers, "CPF_CNPJ_REMETENTE"))
    record(8) = DocumentId(CellByHeader(ws, rowNumber, headers, "CPF_CNPJ_DESTINATARIO"))
    record(9) = CellByHeader(ws, rowNumber, headers, "QTD_VOL")
    record(10) = CellByHeader(ws, rowNumber, headers, "PESO_BRUTO_CTE")
    record(11) = CellByHeader(ws, rowNumber, headers, "PESO_CUBADO_CTE")
    record(12) = CellByHeader(ws, rowNumber, headers, "VLR_PRESTACAO_CTE")
    record(13) = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_BAIXA"))
    record(14) = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_UPLOAD"))
    record(15) = TextValue(CellByHeader(ws, rowNumber, headers, "NOME_DEST"))
    record(16) = TextValue(CellByHeader(ws, rowNumber, headers, "CIDADE_DEST"))
    record(17) = TextValue(CellByHeader(ws, rowNumber, headers, "CD_UF"))
    record(18) = TextValue(CellByHeader(ws, rowNumber, headers, "NOME_AGENTE"))
    record(19) = ParseDateValue(CellByHeader(ws, rowNumber, headers, "DT_ENTREGA"))
    record(20) = CellByHeader(ws, rowNumber, headers, "VLR_MERC_NF")
    record(21) = TextValue(CellByHeader(ws, rowNumber, headers, "STATUS"))
    record(22) = TextValue(CellByHeader(ws, rowNumber, headers, "OCORR_TRANSPORTE"))
    record(23) = TextValue(CellByHeader(ws, rowNumber, headers, "OBS_OCORR_TRANSPORTE"))
    record(24) = TextValue(CellByHeader(ws, rowNumber, headers, "OBS_ENTREGA"))
    ReadBIRecord = record
End Function

Private Function IndexNFUnionByCte(ByVal d238NFs As Object, ByVal dBI As Object) As Object
    Dim result As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For Each key In d238NFs.Keys
        AddNFKeyToIndex result, CStr(key)
    Next key
    For Each key In dBI.Keys
        AddNFKeyToIndex result, CStr(key)
    Next key

    Set IndexNFUnionByCte = result
End Function

Private Sub AddNFKeyToIndex(ByVal index As Object, ByVal compositeKey As String)
    Dim separatorPosition As Long, cte As String, items As Object
    separatorPosition = InStr(1, compositeKey, "|", vbBinaryCompare)
    If separatorPosition <= 1 Then Exit Sub
    cte = Left$(compositeKey, separatorPosition - 1)
    If index.Exists(cte) Then
        Set items = index(cte)
    Else
        Set items = CreateObject("Scripting.Dictionary")
        items.CompareMode = vbTextCompare
        index.Add cte, items
    End If
    If Not items.Exists(compositeKey) Then items.Add compositeKey, True
End Sub

Private Function JoinNFList(ByVal items As Object, ByVal d238NFs As Object, ByVal dBI As Object) As String
    Dim key As Variant, record As Variant
    Dim nf As String, result As String

    For Each key In items.Keys
        If d238NFs.Exists(CStr(key)) Then
            nf = CStr(d238NFs(CStr(key)))
        ElseIf dBI.Exists(CStr(key)) Then
            record = dBI(CStr(key))
            nf = CStr(record(1))
        Else
            nf = ""
        End If
        If Len(nf) > 0 Then
            If Len(result) > 0 Then result = result & " -"
            result = result & nf
        End If
    Next key

    JoinNFList = result
End Function

Private Sub FillOutputRow(ByRef outputData As Variant, ByVal outputRow As Long, _
                          ByVal cteRecord As Variant, ByVal biRecord As Variant, ByVal hasBI As Boolean, _
                          ByVal displayNF As String, ByVal nfCount As Long, _
                          ByVal dTransit As Object, ByVal dHolidays As Object, _
                          ByVal dMissingTransit As Object, ByVal wsEx As Worksheet)
    Dim coleta As Variant, nomeRemetente As String, nomeDestino As String
    Dim cidade As String, ufDestino As String, transportadora As String
    Dim prazo As Variant, justificativa As String, previsaoBID As Variant
    Dim nfForException As String, hasEvidence As Boolean

    If IsDate(cteRecord(7)) Then
        coleta = cteRecord(7)
    Else
        If hasBI Then
            If IsDate(biRecord(4)) Then coleta = biRecord(4) Else coleta = cteRecord(1)
        Else
            coleta = cteRecord(1)
        End If
    End If

    nomeRemetente = TextValue(cteRecord(18))
    If Len(nomeRemetente) = 0 And hasBI Then nomeRemetente = TextValue(biRecord(5))
    nomeDestino = TextValue(cteRecord(13))
    If Len(nomeDestino) = 0 And hasBI Then nomeDestino = TextValue(biRecord(15))
    cidade = TextValue(cteRecord(15))
    If Len(cidade) = 0 And hasBI Then cidade = TextValue(biRecord(16))
    ufDestino = TextValue(cteRecord(14))
    If Len(ufDestino) = 0 And hasBI Then ufDestino = TextValue(biRecord(17))
    transportadora = TextValue(cteRecord(6))

    prazo = Empty   'Prazo BID (T) é fórmula por modal: ver ApplyCoreFormulas e LogSemPrazo.
    'A coluna U é calculada por fórmula em ApplyCoreFormulas (dias úteis).

    justificativa = CombineText(TextValue(cteRecord(5)), TextValue(cteRecord(17)))

    outputData(outputRow, 1) = cteRecord(0)
    outputData(outputRow, 2) = cteRecord(1)
    outputData(outputRow, 3) = displayNF
    If hasBI Then outputData(outputRow, 4) = biRecord(2)
    If hasBI Then outputData(outputRow, 5) = biRecord(3)
    outputData(outputRow, 6) = coleta
    outputData(outputRow, 7) = nomeRemetente
    If hasBI Then outputData(outputRow, 8) = biRecord(6)
    If hasBI Then outputData(outputRow, 9) = biRecord(7)
    outputData(outputRow, 10) = cteRecord(8)
    outputData(outputRow, 11) = nomeDestino
    If hasBI Then outputData(outputRow, 12) = biRecord(8)
    outputData(outputRow, 13) = ""
    outputData(outputRow, 14) = cidade
    outputData(outputRow, 15) = ufDestino
    If hasBI Then outputData(outputRow, 16) = biRecord(9)
    If hasBI Then outputData(outputRow, 17) = biRecord(10)
    If hasBI Then outputData(outputRow, 18) = biRecord(11)
    If hasBI Then outputData(outputRow, 19) = biRecord(12)
    outputData(outputRow, 20) = prazo
    outputData(outputRow, 21) = previsaoBID
    If IsDate(cteRecord(2)) Then
        outputData(outputRow, 22) = cteRecord(2)
    ElseIf hasBI Then
        If IsDate(biRecord(19)) Then outputData(outputRow, 22) = biRecord(19)
    End If
    outputData(outputRow, 23) = ""
    outputData(outputRow, 24) = ""
    outputData(outputRow, 25) = ""
    outputData(outputRow, 26) = cteRecord(4)
    outputData(outputRow, 27) = justificativa
    hasEvidence = (Len(TextValue(cteRecord(11))) > 0 Or IsDate(cteRecord(12)))
    If hasBI Then
        If IsDate(biRecord(14)) Then hasEvidence = True
    End If
    If hasEvidence Then
        outputData(outputRow, 28) = "SIM"
    Else
        outputData(outputRow, 28) = "NAO"
    End If
    outputData(outputRow, 29) = transportadora
    outputData(outputRow, 30) = ""
    outputData(outputRow, 31) = ""
    outputData(outputRow, 32) = nfCount
End Sub

Private Function BuildHeaderMap(ByVal ws As Worksheet) As Object
    Dim result As Object
    Dim lastCol As Long, col As Long, header As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    lastCol = LastUsedColumn(ws)
    For col = 1 To lastCol
        header = NormalizeHeader(ws.Cells(1, col).Value2)
        If Len(header) > 0 Then
            If Not result.Exists(header) Then
                result.Add header, col
            ElseIf header = "VLR_MERCADORIA_CTE" Then
                result(header) = col
            End If
        End If
    Next col
    Set BuildHeaderMap = result
End Function

Private Function CellByHeader(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                              ByVal headers As Object, ByVal header As String) As Variant
    Dim key As String
    key = NormalizeHeader(header)
    If headers.Exists(key) Then
        If ws.Name = "REL_238" Then
            CellByHeader = mRelData(rowNumber, CLng(headers(key)))
        Else
            CellByHeader = mBIData(rowNumber, CLng(headers(key)))
        End If
    Else
        CellByHeader = Empty
    End If
End Function

Private Function BuildTransitDictionary(ByVal ws As Worksheet) As Object
    Dim result As Object
    Dim lastRow As Long, rowNumber As Long
    Dim uf As String, cityKey As String, key As String, prazo As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    'UsedRange, e não LastUsedRow: a aba costuma ficar filtrada e o filtro dela deve ser mantido.
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    For rowNumber = 2 To lastRow
        uf = NormalizeText(ws.Cells(rowNumber, 1).Value2)
        cityKey = NormalizeText(ws.Cells(rowNumber, 2).Value2)
        prazo = ws.Cells(rowNumber, 3).Value2
        If Len(cityKey) > 0 And IsNumeric(prazo) Then
            If Len(uf) = 0 And InStr(1, cityKey, "_", vbBinaryCompare) = 0 And Len(cityKey) <= 2 Then
                key = ""   'Linha de prazo por UF (coluna B só com a sigla): não é usada pela macro.
            ElseIf Len(uf) = 0 Or Left$(cityKey, Len(uf) + 1) = uf & "_" Then
                key = cityKey
            ElseIf Right$(cityKey, Len(uf)) = uf And Len(cityKey) > Len(uf) Then
                key = cityKey   'Formato CIDADE+UF, ex.: PALMASTO.
            Else
                key = uf & "_" & cityKey
                'A coluna A pode estar trocada (ex.: RN em BrasIliaDF): a chave CIDADE+UF da coluna B também vale.
                If Not result.Exists(cityKey) Then result(cityKey) = CDbl(prazo)
            End If
            If Len(key) > 0 Then
                If CDbl(prazo) < 0 Then Err.Raise vbObjectError + 622, , "Prazo BID negativo: " & key
                If result.Exists(key) Then
                    If CDbl(result(key)) <> CDbl(prazo) Then Err.Raise vbObjectError + 623, , "Prazo BID conflitante: " & key
                End If
                result(key) = CDbl(prazo)
            End If
        End If
    Next rowNumber

    Set BuildTransitDictionary = result
End Function

Private Function TransitDays(ByVal dTransit As Object, ByVal uf As String, ByVal city As String) As Variant
    Dim key As String
    key = NormalizeText(uf & "_" & city)
    If dTransit.Exists(key) Then
        TransitDays = dTransit(key)
        Exit Function
    End If
    key = NormalizeText(city & uf)
    If dTransit.Exists(key) Then
        TransitDays = dTransit(key)
    Else
        TransitDays = Empty
    End If
End Function

Private Function CountMatchedNFUnion(ByVal dByCte As Object, ByVal d238 As Object) As Long
    Dim cte As Variant, items As Object
    For Each cte In d238.Keys
        If dByCte.Exists(CStr(cte)) Then
            Set items = dByCte(CStr(cte))
            CountMatchedNFUnion = CountMatchedNFUnion + items.Count
        End If
    Next cte
End Function

Private Sub LogException(ByVal ws As Worksheet, ByVal exceptionType As String, _
                         ByVal cte As String, ByVal nf As String, ByVal detail As String, ByVal sourceName As String)
    ws.Cells(mExceptionRow, 1).Value = exceptionType
    ws.Cells(mExceptionRow, 2).Value = cte
    ws.Cells(mExceptionRow, 3).Value = nf
    ws.Cells(mExceptionRow, 4).Value = detail
    ws.Cells(mExceptionRow, 5).Value = sourceName
    ws.Cells(mExceptionRow, 6).Value = Now
    mExceptionRow = mExceptionRow + 1
End Sub

Private Sub MergeText(ByRef record As Variant, ByVal index As Long, ByVal newValue As Variant)
    Dim text As String
    text = TextValue(newValue)
    If Len(TextValue(record(index))) = 0 And Len(text) > 0 Then record(index) = text
End Sub

Private Sub MergeDateMin(ByRef record As Variant, ByVal index As Long, ByVal newValue As Variant)
    Dim parsed As Variant
    parsed = ParseDateValue(newValue)
    If IsDate(parsed) Then
        If Not IsDate(record(index)) Then
            record(index) = parsed
        ElseIf CDbl(CDate(parsed)) < CDbl(CDate(record(index))) Then
            record(index) = parsed
        End If
    End If
End Sub

Private Sub MergeDateMax(ByRef record As Variant, ByVal index As Long, ByVal newValue As Variant)
    Dim parsed As Variant
    parsed = ParseDateValue(newValue)
    If IsDate(parsed) Then
        If Not IsDate(record(index)) Then
            record(index) = parsed
        ElseIf CDbl(CDate(parsed)) > CDbl(CDate(record(index))) Then
            record(index) = parsed
        End If
    End If
End Sub

Private Function ParseDateValue(ByVal value As Variant) As Variant
    Dim text As String, pieces() As String, dateParts() As String, timeParts() As String
    Dim result As Date
    On Error GoTo InvalidDate

    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then GoTo InvalidDate
    If IsNumeric(value) Then
        If CDbl(value) > 0 Then
            ParseDateValue = CDate(CDbl(value))
            Exit Function
        End If
    End If
    If VarType(value) = vbDate Then
        ParseDateValue = CDate(value)
        Exit Function
    End If

    text = Trim$(CStr(value))
    pieces = Split(text, " ")
    If InStr(pieces(0), "-") > 0 Then
        dateParts = Split(pieces(0), "-")
        If UBound(dateParts) = 2 Then
            text = dateParts(2) & "/" & dateParts(1) & "/" & dateParts(0)
            pieces(0) = text
        End If
    End If
    dateParts = Split(pieces(0), "/")
    If UBound(dateParts) <> 2 Then GoTo InvalidDate
    result = DateSerial(CInt(dateParts(2)), CInt(dateParts(1)), CInt(dateParts(0)))
    If Day(result) <> CInt(dateParts(0)) Or Month(result) <> CInt(dateParts(1)) Or Year(result) <> CInt(dateParts(2)) Then GoTo InvalidDate
    If UBound(pieces) >= 1 Then
        timeParts = Split(pieces(1), ":")
        If UBound(timeParts) >= 1 Then
            If UBound(timeParts) >= 2 Then
                result = result + TimeSerial(CInt(timeParts(0)), CInt(timeParts(1)), CInt(timeParts(2)))
            Else
                result = result + TimeSerial(CInt(timeParts(0)), CInt(timeParts(1)), 0)
            End If
        End If
    End If
    ParseDateValue = result
    Exit Function

InvalidDate:
    ParseDateValue = Empty
    Err.Clear
End Function

Private Function DateScore(ByVal value As Variant) As Double
    If IsDate(value) Then DateScore = CDbl(CDate(value)) Else DateScore = 0
End Function

Private Function ExtractNFList(ByVal rawValue As String) As Collection
    Dim result As New Collection
    Dim seen As Object, regex As Object, matches As Object, match As Object
    Dim prepared As String, parts() As String, part As Variant, nf As String, key As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[0-9]+\s*/\s*[0-9]+"

    If regex.Test(rawValue) Then
        Set matches = regex.Execute(rawValue)
        For Each match In matches
            nf = CleanNFDisplay(CStr(match.Value))
            key = NormalizeNFKey(nf)
            If Len(key) > 0 Then
                If Not seen.Exists(key) Then
                    seen.Add key, True
                    result.Add nf
                End If
            End If
        Next match
    Else
        prepared = Replace(rawValue, vbCrLf, "|")
        prepared = Replace(prepared, vbCr, "|")
        prepared = Replace(prepared, vbLf, "|")
        prepared = Replace(prepared, ";", "|")
        prepared = Replace(prepared, ",", "|")
        prepared = Replace(prepared, " -", "|")
        parts = Split(prepared, "|")
        For Each part In parts
            nf = CleanNFDisplay(CStr(part))
            key = NormalizeNFKey(nf)
            If Len(key) > 0 Then
                If Not seen.Exists(key) Then
                    seen.Add key, True
                    result.Add nf
                End If
            End If
        Next part
    End If

    Set ExtractNFList = result
End Function

Private Function CleanNFDisplay(ByVal value As String) As String
    Dim text As String
    text = Trim$(Replace(value, ChrW(160), " "))
    text = Replace(text, " ", "")
    Do While Len(text) > 0 And Left$(text, 1) = "-"
        text = Mid$(text, 2)
    Loop
    Do While Len(text) > 0 And Right$(text, 1) = "-"
        text = Left$(text, Len(text) - 1)
    Loop
    CleanNFDisplay = text
End Function

Private Function NormalizeNFKey(ByVal value As String) As String
    Dim text As String, digits As String, position As Long, character As String
    text = CleanNFDisplay(value)
    position = InStr(1, text, "/", vbBinaryCompare)
    If position > 0 Then text = Left$(text, position - 1)
    For position = 1 To Len(text)
        character = Mid$(text, position, 1)
        If character >= "0" And character <= "9" Then digits = digits & character
    Next position
    Do While Len(digits) > 1 And Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop
    NormalizeNFKey = digits
End Function

Private Function CleanId(ByVal value As Variant) As String
    Dim text As String
    If IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then
        CleanId = Format$(CDbl(value), "0")
    Else
        text = Trim$(Replace(CStr(value), ChrW(160), " "))
        If Right$(text, 2) = ".0" And IsNumeric(text) Then text = Left$(text, Len(text) - 2)
        CleanId = text
    End If
End Function

Private Function DocumentId(ByVal value As Variant) As String
    If IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then
        DocumentId = Format$(CDbl(value), "0")
    Else
        DocumentId = Trim$(Replace(CStr(value), ChrW(160), " "))
    End If
End Function

Private Function TextValue(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    TextValue = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Private Function CombineText(ByVal firstText As String, ByVal secondText As String) As String
    If Len(firstText) = 0 Then
        CombineText = secondText
    ElseIf Len(secondText) = 0 Then
        CombineText = firstText
    ElseIf InStr(1, firstText, secondText, vbTextCompare) > 0 Then
        CombineText = firstText
    Else
        CombineText = firstText & " | " & secondText
    End If
End Function

Private Function NormalizeHeader(ByVal value As Variant) As String
    NormalizeHeader = UCase$(Trim$(Replace(TextValue(value), ChrW(160), " ")))
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    Dim text As String
    text = UCase$(Trim$(Replace(TextValue(value), ChrW(160), " ")))
    text = Replace(text, ChrW(193), "A")
    text = Replace(text, ChrW(192), "A")
    text = Replace(text, ChrW(194), "A")
    text = Replace(text, ChrW(195), "A")
    text = Replace(text, ChrW(196), "A")
    text = Replace(text, ChrW(201), "E")
    text = Replace(text, ChrW(200), "E")
    text = Replace(text, ChrW(202), "E")
    text = Replace(text, ChrW(203), "E")
    text = Replace(text, ChrW(205), "I")
    text = Replace(text, ChrW(204), "I")
    text = Replace(text, ChrW(206), "I")
    text = Replace(text, ChrW(207), "I")
    text = Replace(text, ChrW(211), "O")
    text = Replace(text, ChrW(210), "O")
    text = Replace(text, ChrW(212), "O")
    text = Replace(text, ChrW(213), "O")
    text = Replace(text, ChrW(214), "O")
    text = Replace(text, ChrW(218), "U")
    text = Replace(text, ChrW(217), "U")
    text = Replace(text, ChrW(219), "U")
    text = Replace(text, ChrW(220), "U")
    text = Replace(text, ChrW(199), "C")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    NormalizeText = text
End Function

Private Sub ShowAllRows(ByVal ws As Worksheet)
    'O Find ignora linhas ocultas por filtro e cortaria dados sem avisar.
    If ws.FilterMode Then ws.ShowAllData
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    ShowAllRows ws
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                              LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then LastUsedRow = 1 Else LastUsedRow = found.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim found As Range
    ShowAllRows ws
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                              LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then LastUsedColumn = 1 Else LastUsedColumn = found.Column
End Function

'Lê o texto do cliente na célula nomeada cfgCliente. Em branco ou sem a célula, usa o padrão.
Private Function CLIENT_FILTER() As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(ThisWorkbook.Names("cfgCliente").RefersToRange.Value2))
    On Error GoTo 0
    If Len(v) = 0 Then v = CLIENT_FILTER_PADRAO
    CLIENT_FILTER = v
End Function

'Registra em Exceções as NFs cujo Prazo BID (T, na cópia de trabalho) ficou vazio.
Private Sub LogSemPrazo(ByVal ws As Worksheet, ByVal ex As Worksheet)
    Dim r As Long
    For r = 2 To LastUsedRow(ws)
        If Len(TextValue(ws.Cells(r, 1).Value2)) > 0 And Len(TextValue(ws.Cells(r, 20).Value2)) = 0 Then
            LogException ex, "SEM PRAZO BID", CStr(ws.Cells(r, 1).Value2), CStr(ws.Cells(r, 3).Value2), _
                "Sem prazo para " & TextValue(ws.Cells(r, 14).Value2) & "-" & TextValue(ws.Cells(r, 15).Value2) & _
                " (" & TextValue(ws.Cells(r, 10).Value2) & ").", "TT LTL / TT Aereo / Transit time"
        End If
    Next r
End Sub
