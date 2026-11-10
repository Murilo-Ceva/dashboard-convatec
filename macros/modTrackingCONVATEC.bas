Attribute VB_Name = "modTrackingCONVATEC"
Option Explicit
' Gera, por responsável da coluna "resp cobrança", um anexo com as notas ainda não entregues
' e um RASCUNHO no Outlook. Esta macro não dispara e-mail: o envio é sempre manual.
' As colunas são localizadas pelo cabeçalho, então podem mudar de posição.
' Os e-mails ficam em campos para preencher:
'   aba Contatos: um e-mail por transportadora (CIDADE e UF em branco valem para todas as rotas dela);
'                 uma linha com CIDADE e UF preenchidas vale só para aquela rota e tem prioridade.
'   aba Config Follow: CC fixo (cfgCC) e nome do cliente no e-mail (cfgNomeEmail).
'   aba Contatos, coluna E_MAIL CÓPIA (CC): cópia só daquele responsável; soma-se ao CC fixo.
Private Const ABA_BASE As String = "Base por NF Oficial"
Private Const ABA_CONTATOS As String = "Contatos"
Private Const ABA_CONFIG As String = "Config Follow"
Private Const TITULO As String = "Follow ConvaTec"
Private mResultado As String

Public Sub GerarTrackingTransportadores()
    ExecutarTracking False
End Sub

Public Function ResultadoUltimoTracking() As String
    ResultadoUltimoTracking = mResultado
End Function

' Coloca os botões na aba Config Follow. Pode ser executada de novo sem duplicar.
Public Sub CriarBotoesFollow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ABA_CONFIG)
    PoeBotao ws, "btnAtualizarBase", "1. Atualizar base (238 + BI)", "AtualizarBaseWV", 0
    PoeBotao ws, "btnPrepararContatos", "2. Listar transportadoras em Contatos", "PrepararContatos", 1
    PoeBotao ws, "btnGerarFollow", "3. Gerar follow (rascunhos Outlook)", "GerarTrackingTransportadores", 2
End Sub

Private Sub PoeBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal acao As String, ByVal i As Long)
    Dim b As Object, ancora As Range
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
    Set ancora = ws.Range("D3")
    Set b = ws.Buttons.Add(ancora.Left + 10, ancora.Top + i * 44, 250, 36)
    b.Name = nome
    b.Caption = texto
    b.OnAction = acao
End Sub

' Acrescenta na aba Contatos cada transportadora da base que ainda não tem linha,
' com o e-mail em branco para preencher. Não altera nem apaga linhas existentes.
Public Sub PrepararContatos()
    Dim n As Long, lista As String
    On Error GoTo Falhou
    n = IncluirTransportadoras(lista)
    If n = 0 Then
        MsgBox "Todas as transportadoras da base já estão na aba " & ABA_CONTATOS & ".", vbInformation, TITULO
    Else
        MsgBox n & " transportadoras incluídas na aba " & ABA_CONTATOS & "." & vbCrLf & _
               "Preencha RESPONSÁVEL e E_MAIL (campos amarelos):" & lista, vbInformation, TITULO
        ThisWorkbook.Worksheets(ABA_CONTATOS).Activate
    End If
    Exit Sub
Falhou:
    MsgBox Err.Description, vbCritical, TITULO
End Sub

Public Function IncluirTransportadoras(ByRef lista As String) As Long
    Dim wb As Worksheet, wc As Worksheet, cT As Long, r As Long, lastRow As Long, lastCol As Long
    Dim existentes As Object, novos As Object, t As String, k As Variant, linha As Long, ultimaC As Long
    Set wb = ThisWorkbook.Worksheets(ABA_BASE)
    Set wc = ThisWorkbook.Worksheets(ABA_CONTATOS)
    lastCol = wb.Cells(1, wb.Columns.Count).End(xlToLeft).Column
    For r = 1 To lastCol
        If Norm(wb.Cells(1, r).Value2) = "TRANSP." Then cT = r: Exit For
    Next r
    If cT = 0 Then Err.Raise vbObjectError + 705, , "Não encontrei a coluna Transp. na aba " & ABA_BASE & "."
    Set existentes = CreateObject("Scripting.Dictionary")
    ultimaC = wc.Cells(wc.Rows.Count, 1).End(xlUp).Row
    For r = 2 To ultimaC
        t = UCase$(Texto(wc.Cells(r, 1).Value2))
        If Len(t) > 0 Then existentes(t) = True
    Next r
    Set novos = CreateObject("Scripting.Dictionary")
    lastRow = wb.Cells(wb.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        t = Texto(wb.Cells(r, cT).Value2)
        If Len(t) > 0 And Len(Texto(wb.Cells(r, 1).Value2)) > 0 Then
            If Not existentes.Exists(UCase$(t)) And Not novos.Exists(UCase$(t)) Then novos(UCase$(t)) = t
        End If
    Next r
    linha = Application.Max(2, ultimaC + 1)
    For Each k In novos.Keys
        wc.Cells(linha, 1).Value2 = novos(k)
        wc.Range(wc.Cells(linha, 4), wc.Cells(linha, 5)).Interior.Color = RGB(255, 242, 204)
        If Not wc.Cells(linha, 6).HasFormula Then wc.Cells(linha, 6).Formula = "=IF(A" & linha & "="""","""",TRIM(UPPER(A" & linha & "))&""_""&TRIM(UPPER(B" & linha & "))&""_""&TRIM(UPPER(C" & linha & ")))"
        lista = lista & vbCrLf & "  " & novos(k)
        linha = linha + 1
    Next k
    IncluirTransportadoras = novos.Count
End Function

Public Sub ExecutarTracking(ByVal silencioso As Boolean)
    Dim ws As Worksheet, data As Variant, h As Object, grupos As Object, nomes As Object
    Dim lastRow As Long, lastCol As Long, r As Long, resp As String, key As Variant
    Dim cCte As Long, cData As Long, cNF As Long, cRem As Long, cDN As Long, cCid As Long
    Dim ccs As Object, ccLinha As String, cUF As Long, cPrev As Long, cEnt As Long, cResp As Long, cTransp As Long, incluidas As Long, listaInc As String
    Dim entregues As Long, abertas As Long, semResp As Long, listaSem As String
    Dim pasta As String, resumo As String, msg As String, ol As Object, erroOutlook As String
    Dim arquivo As String, nome As String, usados As Object, feitos As Long, falhas As String
    Dim qtdNF As Long, qtdCte As Long, qtdVenc As Long
    Dim oldScreen As Boolean, oldAlerts As Boolean

    oldScreen = Application.ScreenUpdating: oldAlerts = Application.DisplayAlerts
    On Error GoTo Falhou
    Set ws = ThisWorkbook.Worksheets(ABA_BASE)
    Application.Calculate

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set h = CreateObject("Scripting.Dictionary")
    For r = 1 To lastCol
        If Not h.Exists(Norm(ws.Cells(1, r).Value2)) Then h.Add Norm(ws.Cells(1, r).Value2), r
    Next r
    cCte = Coluna(h, "CTE"): cData = Coluna(h, "DATA"): cNF = Coluna(h, "NOTA FISCAL")
    cRem = Coluna(h, "REMETENTE"): cDN = Coluna(h, "NOME DO DN"): cCid = Coluna(h, "CIDADE")
    cUF = Coluna(h, "UF DEST."): cPrev = Coluna(h, "PREVISAO DE ENTREGA")
    cEnt = Coluna(h, "DATA DA ENTREGA"): cResp = Coluna(h, "RESP COBRANCA")
    cTransp = Coluna(h, "TRANSP.")

    lastRow = ws.Cells(ws.Rows.Count, cCte).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 701, , "A aba " & ABA_BASE & " está sem dados."
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    Set grupos = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        If Len(Texto(data(r, cCte))) > 0 Then
            If Len(Texto(data(r, cEnt))) > 0 Then
                entregues = entregues + 1
            Else
                resp = LimparEmails(data(r, cResp))
                If Len(resp) = 0 Then
                    semResp = semResp + 1
                    If semResp <= 15 Then listaSem = listaSem & vbCrLf & "  CT-e " & Texto(data(r, cCte)) & " / NF " & Texto(data(r, cNF)) & " / " & Texto(data(r, cCid)) & "-" & Texto(data(r, cUF)) & " / " & Texto(data(r, cTransp))
                Else
                    abertas = abertas + 1
                    If Not grupos.Exists(LCase$(resp)) Then grupos.Add LCase$(resp), New Collection
                    grupos(LCase$(resp)).Add r
                End If
            End If
        End If
    Next r
    If semResp > 0 Then incluidas = IncluirTransportadoras(listaInc)
    If grupos.Count = 0 Then Err.Raise vbObjectError + 702, , "Nenhuma nota em aberto com e-mail preenchido." & _
        vbCrLf & entregues & " entregues; " & semResp & " sem e-mail." & vbCrLf & _
        "Preencha a coluna E_MAIL da aba " & ABA_CONTATOS & " (" & incluidas & " transportadoras acabaram de ser incluídas lá)."

    Set nomes = NomesResponsaveis()
    Set ccs = CopiasResponsaveis()
    For Each key In grupos.Keys
        resumo = resumo & vbCrLf & "  " & NomeGrupo(nomes, CStr(key)) & ": " & grupos(key).Count & " notas"
        If ccs.Exists(CStr(key)) Then resumo = resumo & " (cópia: " & ccs(CStr(key)) & ")"
    Next key
    If Not silencioso Then
        msg = "Serão criados " & grupos.Count & " rascunhos no Outlook. Nenhum e-mail será enviado." & vbCrLf & resumo & vbCrLf & vbCrLf & _
              entregues & " notas entregues ficam de fora."
        If semResp > 0 Then msg = msg & vbCrLf & semResp & " notas sem e-mail na aba " & ABA_CONTATOS & " ficam de fora e serão listadas no final."
        msg = msg & vbCrLf & vbCrLf & "Se já gerou hoje, os rascunhos anteriores continuarão no Outlook. Continuar?"
        If MsgBox(msg, vbYesNo + vbQuestion, TITULO) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    pasta = PastaDeSaida()
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    If ol Is Nothing Then erroOutlook = "Não foi possível abrir o Outlook clássico: os anexos foram gerados, mas sem rascunho."
    Err.Clear
    On Error GoTo Falhou

    Set usados = CreateObject("Scripting.Dictionary")
    For Each key In grupos.Keys
        nome = NomeGrupo(nomes, CStr(key))
        arquivo = NomeArquivo(pasta, nome, usados)
        ContarGrupo grupos(key), data, cCte, cPrev, qtdNF, qtdCte, qtdVenc
        GerarAnexo arquivo, grupos(key), data, cCte, cData, cNF, cRem, cDN, cCid, cUF, cPrev
        If Not ol Is Nothing Then
            ccLinha = ""
            If ccs.Exists(CStr(key)) Then ccLinha = ccs(CStr(key))
            If CriarRascunho(ol, LimparEmails(key), ccLinha, nome, arquivo, qtdNF, qtdCte, qtdVenc) Then
                feitos = feitos + 1
            Else
                falhas = falhas & vbCrLf & "  " & nome & ": " & Err.Description
            End If
        End If
    Next key

    mResultado = feitos & " rascunhos criados na pasta Rascunhos do Outlook. Nenhum e-mail foi enviado." & vbCrLf & resumo & vbCrLf & vbCrLf & _
        entregues & " notas entregues ficaram de fora." & vbCrLf & "Anexos em: " & pasta
    If Len(erroOutlook) > 0 Then mResultado = mResultado & vbCrLf & vbCrLf & erroOutlook
    If Len(falhas) > 0 Then mResultado = mResultado & vbCrLf & vbCrLf & "ATENÇÃO, rascunhos com falha:" & falhas
    If semResp > 0 Then
        mResultado = mResultado & vbCrLf & vbCrLf & "ATENÇÃO, " & semResp & " notas sem e-mail (preencha a aba " & ABA_CONTATOS & "):" & listaSem
        If incluidas > 0 Then mResultado = mResultado & vbCrLf & incluidas & " transportadoras incluídas em " & ABA_CONTATOS & " para preencher:" & listaInc
        If semResp > 15 Then mResultado = mResultado & vbCrLf & "  ... e mais " & (semResp - 15)
    End If
    GoTo Fim
Falhou:
    mResultado = "Tracking não concluído." & vbCrLf & Err.Description
    Application.ScreenUpdating = oldScreen: Application.DisplayAlerts = oldAlerts
    If Not silencioso Then MsgBox mResultado, vbCritical, TITULO
    Exit Sub
Fim:
    Application.ScreenUpdating = oldScreen: Application.DisplayAlerts = oldAlerts
    If Not silencioso Then MsgBox mResultado, vbInformation, TITULO
End Sub

Private Sub ContarGrupo(ByVal linhas As Collection, ByRef data As Variant, ByVal cCte As Long, ByVal cPrev As Long, _
                        ByRef qtdNF As Long, ByRef qtdCte As Long, ByRef qtdVenc As Long)
    Dim item As Variant, ctes As Object
    Set ctes = CreateObject("Scripting.Dictionary")
    qtdVenc = 0
    For Each item In linhas
        ctes(Texto(data(item, cCte))) = True
        If Not IsError(data(item, cPrev)) Then
            If IsNumeric(data(item, cPrev)) And Len(Texto(data(item, cPrev))) > 0 Then
                If CDbl(data(item, cPrev)) < CDbl(Date) Then qtdVenc = qtdVenc + 1
            End If
        End If
    Next item
    qtdNF = linhas.Count: qtdCte = ctes.Count
End Sub

Private Sub GerarAnexo(ByVal arquivo As String, ByVal linhas As Collection, ByRef data As Variant, _
                       ByVal cCte As Long, ByVal cData As Long, ByVal cNF As Long, ByVal cRem As Long, _
                       ByVal cDN As Long, ByVal cCid As Long, ByVal cUF As Long, ByVal cPrev As Long)
    Dim wb As Workbook, ws As Worksheet, arr() As Variant, n As Long, i As Long, item As Variant, larguras As Variant, c As Long
    n = linhas.Count
    ReDim arr(1 To n, 1 To 8)
    For Each item In linhas
        i = i + 1
        arr(i, 1) = Texto(data(item, cCte))
        arr(i, 2) = ValorOuVazio(data(item, cData))
        arr(i, 3) = Texto(data(item, cNF))
        arr(i, 4) = ValorOuVazio(data(item, cRem))
        arr(i, 5) = ValorOuVazio(data(item, cDN))
        arr(i, 6) = ValorOuVazio(data(item, cCid))
        arr(i, 7) = ValorOuVazio(data(item, cUF))
        arr(i, 8) = ValorOuVazio(data(item, cPrev))
    Next item
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo Erro
    Set ws = wb.Worksheets(1)
    ws.Name = "Tracking"
    ws.Range("A1:K1").Value = Array("Cte", "Data", "Nota fiscal", "Remetente", "Nome do DN", "Cidade", "UF Dest.", _
                                    "PREVISAO DE ENTREGA", "DT RECEBIMENTO", "DT DE ENTREGA", "STATUS")
    ws.Range("A2:A" & n + 1).NumberFormat = "@"
    ws.Range("C2:C" & n + 1).NumberFormat = "@"
    ws.Range("A2").Resize(n, 8).Value2 = arr
    ws.Range("B2:B" & n + 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("H2:J" & n + 1).NumberFormat = "dd/mm/yyyy"
    With ws.Range("A1:K" & n + 1)
        .Font.Name = "Aptos Narrow"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Range("A1:K1")
        .Interior.Color = RGB(2, 29, 73)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 12
    End With
    ws.Rows(1).RowHeight = 16
    larguras = Array(7.8, 10.1, 10.7, 31.4, 43.4, 20.5, 9, 22.5, 17.5, 15.6, 7.7)
    For c = 0 To 10
        ws.Columns(c + 1).ColumnWidth = larguras(c)
    Next c
    wb.SaveAs Filename:=arquivo, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Exit Sub
Erro:
    Dim e As String
    e = Err.Description
    wb.Close SaveChanges:=False
    Err.Raise vbObjectError + 703, , "Falha ao gerar " & arquivo & ": " & e
End Sub

Private Function CriarRascunho(ByVal ol As Object, ByVal para As String, ByVal ccLinha As String, ByVal nome As String, ByVal arquivo As String, _
                               ByVal qtdNF As Long, ByVal qtdCte As Long, ByVal qtdVenc As Long) As Boolean
    Dim m As Object, insp As Object, corpo As String, atual As String, p As Long
    On Error GoTo Erro
    Set m = ol.CreateItem(0)
    m.To = para
    m.CC = LimparEmails(Config("cfgCC", "") & ";" & ccLinha)   'CC fixo + cópia do responsável, sem repetir
    m.Subject = "Follow up | " & Config("cfgNomeEmail", "ConvaTec") & " | " & nome & " | " & Format$(Date, "dd/mm/yyyy")
    m.Attachments.Add arquivo
    Set insp = m.GetInspector   'Faz o Outlook incluir a assinatura padrão, sem abrir janela.
    corpo = CorpoEmail(nome, qtdNF, qtdCte, qtdVenc)
    'A assinatura de cada analista pode já trazer a despedida; não repetir.
    If InStr(1, m.Body, "Atenciosamente", vbTextCompare) = 0 Then corpo = corpo & "<p>Atenciosamente,</p>"
    corpo = corpo & "</div>"
    atual = m.HTMLBody
    p = InStr(1, atual, "<body", vbTextCompare)
    If p > 0 Then p = InStr(p, atual, ">")
    If p > 0 Then
        m.HTMLBody = Left$(atual, p) & corpo & Mid$(atual, p + 1)
    Else
        m.HTMLBody = corpo & atual
    End If
    m.Save
    m.Close 0
    CriarRascunho = True
    Exit Function
Erro:
    CriarRascunho = False
End Function

Private Function CorpoEmail(ByVal nome As String, ByVal qtdNF As Long, ByVal qtdCte As Long, ByVal qtdVenc As Long) As String
    Dim s As String, n As String
    n = Replace(Replace(nome, "&", "&amp;"), "<", "&lt;")
    s = "<div style=""font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt"">"
    s = s & "<p>Prezados,</p>"
    s = s & "<p>Segue em anexo o follow-up das entregas da " & Replace(Config("cfgNomeEmail", "ConvaTec"), "<", "&lt;") & " sob responsabilidade de <b>" & n & "</b>, com as notas fiscais sem entrega registrada até " & Format$(Date, "dd/mm/yyyy") & ".</p>"
    s = s & "<p><b>Resumo</b><br>Notas fiscais em aberto: " & qtdNF & "<br>CT-es: " & qtdCte & "<br>Notas com previsão vencida: " & qtdVenc & "</p>"
    s = s & "<p>Por gentileza, preencham na planilha, para cada nota:</p><ul>"
    s = s & "<li><b>DT RECEBIMENTO:</b> data em que a carga foi recebida;</li>"
    s = s & "<li><b>DT DE ENTREGA:</b> data da entrega, se já realizada;</li>"
    s = s & "<li><b>STATUS:</b> situação atual e, se houver ocorrência ou atraso, o motivo e a nova previsão de entrega.</li></ul>"
    If qtdVenc > 0 Then s = s & "<p>Pedimos prioridade para as notas com previsão vencida.</p>"
    s = s & "<p>Se alguma nota não estiver sob a responsabilidade de vocês, nos avisem na resposta. Aguardamos o retorno com a planilha atualizada.</p>"
    CorpoEmail = s
End Function

Private Function NomesResponsaveis() As Object
    Dim d As Object, ws As Worksheet, c As Long, cNome As Long, cMail As Long, r As Long, lastRow As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    Set NomesResponsaveis = d
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_CONTATOS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case Norm(ws.Cells(1, c).Value2)
            Case "RESPONSAVEL": cNome = c
            Case "E_MAIL": cMail = c
        End Select
    Next c
    If cNome = 0 Or cMail = 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, cMail).End(xlUp).Row
    For r = 2 To lastRow
        k = LCase$(LimparEmails(ws.Cells(r, cMail).Value2))
        If Len(k) > 0 And Len(Texto(ws.Cells(r, cNome).Value2)) > 0 Then d(k) = Texto(ws.Cells(r, cNome).Value2)
    Next r
End Function

' Cópia (CC) por responsável: chave = e-mails da coluna E_MAIL; junta as cópias de todas as linhas com esse e-mail.
Public Function CopiasResponsaveis() As Object
    Dim d As Object, ws As Worksheet, c As Long, cCc As Long, cMail As Long, r As Long, lastRow As Long, k As String, v As String
    Set d = CreateObject("Scripting.Dictionary")
    Set CopiasResponsaveis = d
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_CONTATOS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case Norm(ws.Cells(1, c).Value2)
            Case "E_MAIL": cMail = c
            Case "E_MAIL COPIA (CC)", "E_MAIL COPIA", "COPIA (CC)", "CC": cCc = c
        End Select
    Next c
    If cCc = 0 Or cMail = 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, cMail).End(xlUp).Row
    For r = 2 To lastRow
        k = LCase$(LimparEmails(ws.Cells(r, cMail).Value2))
        v = LimparEmails(ws.Cells(r, cCc).Value2)
        If Len(k) > 0 And Len(v) > 0 Then
            If d.Exists(k) Then d(k) = LimparEmails(d(k) & ";" & v) Else d(k) = v
        End If
    Next r
End Function

Private Function Config(ByVal nome As String, ByVal padrao As String) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(ThisWorkbook.Names(nome).RefersToRange.Value2))
    On Error GoTo 0
    If Len(v) = 0 Then v = padrao
    Config = v
End Function

Private Function NomeGrupo(ByVal nomes As Object, ByVal key As String) As String
    If nomes.Exists(key) Then
        NomeGrupo = CStr(nomes(key))
    Else
        NomeGrupo = Split(key, ";")(0)
    End If
End Function

Private Function NomeArquivo(ByVal pasta As String, ByVal nome As String, ByVal usados As Object) As String
    Dim limpo As String, ch As Variant, base As String, n As Long
    limpo = nome
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        limpo = Replace(limpo, ch, " ")
    Next ch
    base = "FOLLOW_CONVATEC_" & Format$(Date, "ddmmyyyy") & "_" & Trim$(limpo)
    NomeArquivo = base
    n = 1
    Do While usados.Exists(LCase$(NomeArquivo))
        n = n + 1
        NomeArquivo = base & "_" & n
    Loop
    usados(LCase$(NomeArquivo)) = True
    NomeArquivo = pasta & Application.PathSeparator & NomeArquivo & ".xlsx"
End Function

Private Function PastaDeSaida() As String
    Dim fso As Object, raiz As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    raiz = ThisWorkbook.Path
    If LCase$(Left$(raiz, 4)) = "http" Or Len(raiz) = 0 Then raiz = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    PastaDeSaida = raiz & Application.PathSeparator & "Tracking"
    If Not fso.FolderExists(PastaDeSaida) Then fso.CreateFolder PastaDeSaida
End Function

Private Function Coluna(ByVal h As Object, ByVal nome As String) As Long
    If Not h.Exists(nome) Then Err.Raise vbObjectError + 704, , "Não encontrei a coluna '" & nome & "' na aba " & ABA_BASE & "."
    Coluna = h(nome)
End Function

Private Function LimparEmails(ByVal v As Variant) As String
    Dim parte As Variant, vistos As Object, s As String, t As String
    If IsError(v) Then Exit Function
    Set vistos = CreateObject("Scripting.Dictionary")
    For Each parte In Split(Replace(Texto(v), ",", ";"), ";")
        t = Trim$(CStr(parte))
        If InStr(1, t, "@") > 1 And Not vistos.Exists(LCase$(t)) Then
            vistos.Add LCase$(t), True
            s = s & IIf(Len(s) > 0, ";", "") & t
        End If
    Next parte
    LimparEmails = s
End Function

Private Function ValorOuVazio(ByVal v As Variant) As Variant
    If IsError(v) Then
        ValorOuVazio = Empty
    ElseIf VarType(v) = vbString Then
        If Left$(v, 1) = "#" Then ValorOuVazio = Empty Else ValorOuVazio = v
    Else
        ValorOuVazio = v
    End If
End Function

Private Function Texto(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Texto = Trim$(Replace(CStr(v), ChrW(160), " "))
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim s As String, i As Long, de As Variant, para As Variant
    s = UCase$(Texto(v))
    de = Array("Á", "À", "Â", "Ã", "É", "Ê", "Í", "Ó", "Ô", "Õ", "Ú", "Ç")
    para = Array("A", "A", "A", "A", "E", "E", "I", "O", "O", "O", "U", "C")
    For i = 0 To UBound(de)
        s = Replace(s, de(i), para(i))
    Next i
    Norm = s
End Function
