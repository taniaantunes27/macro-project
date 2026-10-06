Attribute VB_Name = "CWPower_Importar_MASTER"
Option Explicit

' ============================================================
' CWPOWER - CRIAR PROJECT A PARTIR DO MASTER_IMPORT
' Versão 2.0 (reescrita de raiz)
'
' Executar num PROJECT NOVO / EM BRANCO (Project 2010 ou superior).
' Macro a executar: CRIAR_Project_CWPower
'
' O Excel MASTER_IMPORT_CWPOWER deve conter as folhas:
'   MASTER_CONFIG       (campo / valor)
'   MASTER_CALENDARIO   (dias da semana linhas 4-10, exceções desde linha 14)
'   MASTER_TAREFAS      (cabeçalho na linha 3, dados desde a linha 4)
'   MASTER_RECURSOS     (idem)
'   MASTER_ATRIBUICOES  (idem)
'   MASTER_FINANCEIRO   (idem)
'
' Cria no Project:
'   - calendário da obra (ex.: CAL_OBRA_CWPOWER) com horário e feriados
'   - opções de calendário (horas/dia, horas/semana, dias/mês)
'   - tarefas com estrutura (nível de destaque), datas, durações,
'     marcos, prazos, notas e restrições
'   - predecessoras (aceita FS/SS/FF/SF e TI/II/TT/IT, FI/IF, com lag
'     em dias/horas/semanas/meses, em inglês ou português)
'   - recursos (grupo, unidades máximas, calendário da obra)
'   - atribuições de recursos
'   - Custo1 = valor financeiro da tarefa; Texto30 = rubrica
'
' Pergunta inicial:
'   SIM = preserva exatamente as datas do Excel (tarefas não-resumo
'         ficam em agendamento manual).
'   NÃO = repõe o modo Manual/Auto original e o Project recalcula com
'         o calendário, predecessoras e restrições.
'
' No fim é gravado um ficheiro "..._LOG_IMPORTACAO.txt" ao lado do
' Excel com todos os avisos (linhas ignoradas, ligações falhadas, etc.).
' ============================================================

Private Const xlUp As Long = -4162
Private Const LINHA_DADOS As Long = 4

' Colunas da folha MASTER_TAREFAS
Private Const T_ID As Long = 1
Private Const T_NIVEL As Long = 4
Private Const T_NOME As Long = 5
Private Const T_RESUMO As Long = 6
Private Const T_MARCO As Long = 7
Private Const T_MODO As Long = 8
Private Const T_TIPO As Long = 9
Private Const T_EFFORT As Long = 10
Private Const T_DURACAO As Long = 11
Private Const T_INICIO As Long = 12
Private Const T_FIM As Long = 13
Private Const T_PCT As Long = 14
Private Const T_PRED As Long = 15
Private Const T_CTIPO As Long = 16
Private Const T_CDATA As Long = 17
Private Const T_DEADLINE As Long = 18
Private Const T_CAL As Long = 19
Private Const T_NOTAS As Long = 20
Private Const T_NCOLS As Long = 20

' Estado do módulo
Private mTarefas As Collection      ' ID do Excel -> Task
Private mRecursos As Collection     ' Nome do recurso -> Resource
Private mLog As String
Private mAvisos As Long

Private mCalNome As String
Private mProjInicio As Variant
Private mProjNome As String
Private mProjCodigo As String
Private mMinDia As Double
Private mMinSemana As Double
Private mDiasMes As Double

' Horário de um dia útil de referência (usado nas exceções de trabalho)
Private mTurnoIni(1 To 3) As String
Private mTurnoFim(1 To 3) As String
Private mNTurnos As Long


' ============================================================
' MACRO PRINCIPAL
' ============================================================

Public Sub CRIAR_Project_CWPower()

    Dim xlApp As Object
    Dim xlWB As Object
    Dim caminho As Variant
    Dim resposta As VbMsgBoxResult
    Dim preservarDatas As Boolean
    Dim etapa As String
    Dim calcOriginal As Long
    Dim alterouCalc As Boolean
    Dim msgFinal As String
    Dim resumo As String
    Dim ficheiroLog As String
    Dim folhas As Variant
    Dim i As Long

    Dim dT As Variant
    Dim dR As Variant
    Dim dA As Variant
    Dim dF As Variant

    Dim nTarefas As Long
    Dim nLigacoes As Long
    Dim nRecursos As Long
    Dim nAtrib As Long
    Dim nFin As Long

    On Error GoTo TrataErro

    Set mTarefas = New Collection
    Set mRecursos = New Collection
    mLog = ""
    mAvisos = 0

    etapa = "Verificar projeto ativo"

    If ContarTarefas() > 0 Then
        MsgBox "Este Project já contém tarefas." & vbCrLf & vbCrLf & _
               "Crie um Project novo/em branco (Ficheiro > Novo > Projeto em branco) " & _
               "e volte a executar a macro.", vbExclamation, "CWPower"
        Exit Sub
    End If

    resposta = MsgBox( _
        "Pretende preservar exatamente as datas do Excel?" & vbCrLf & vbCrLf & _
        "SIM (recomendado):" & vbCrLf & _
        "as tarefas não-resumo ficam em agendamento manual, com as datas do Excel." & vbCrLf & vbCrLf & _
        "NÃO:" & vbCrLf & _
        "repõe o modo Manual/Auto original e o Project recalcula com o calendário, " & _
        "predecessoras e restrições (as datas podem mudar)." & vbCrLf & vbCrLf & _
        "CANCELAR: sair sem fazer nada.", _
        vbQuestion + vbYesNoCancel, "CWPower - Modo de criação")

    If resposta = vbCancel Then Exit Sub
    preservarDatas = (resposta = vbYes)

    ' --------------------------------------------------------
    ' Abrir o Excel
    ' --------------------------------------------------------
    etapa = "Abrir Excel"
    Estado "a abrir o Excel..."

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    caminho = xlApp.GetOpenFilename( _
        "Ficheiros Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", 1, _
        "Selecionar MASTER_IMPORT_CWPOWER")

    If VarType(caminho) = vbBoolean Then GoTo SaidaLimpa

    Set xlWB = xlApp.Workbooks.Open(Filename:=CStr(caminho), UpdateLinks:=0, ReadOnly:=True)

    folhas = Array("MASTER_CONFIG", "MASTER_CALENDARIO", "MASTER_TAREFAS", _
                   "MASTER_RECURSOS", "MASTER_ATRIBUICOES", "MASTER_FINANCEIRO")

    For i = LBound(folhas) To UBound(folhas)
        If Not FolhaExiste(xlWB, CStr(folhas(i))) Then
            MsgBox "O ficheiro selecionado não é um MASTER_IMPORT_CWPOWER válido." & vbCrLf & vbCrLf & _
                   "Falta a folha: " & folhas(i), vbCritical, "CWPower"
            GoTo SaidaLimpa
        End If
    Next i

    ' --------------------------------------------------------
    ' Ler tudo para memória (rápido e sem depender do Excel depois)
    ' --------------------------------------------------------
    etapa = "Ler Excel"
    Estado "a ler o Excel..."

    LerConfig xlWB.Worksheets("MASTER_CONFIG")

    dT = LerTabela(xlWB.Worksheets("MASTER_TAREFAS"), T_NCOLS)
    dR = LerTabela(xlWB.Worksheets("MASTER_RECURSOS"), 4)
    dA = LerTabela(xlWB.Worksheets("MASTER_ATRIBUICOES"), 4)
    dF = LerTabela(xlWB.Worksheets("MASTER_FINANCEIRO"), 4)

    If IsEmpty(dT) Then
        MsgBox "A folha MASTER_TAREFAS não tem tarefas a partir da linha " & LINHA_DADOS & ".", _
               vbCritical, "CWPower"
        GoTo SaidaLimpa
    End If

    If Not IsDate(mProjInicio) Then mProjInicio = MenorData(dT, T_INICIO)

    ' --------------------------------------------------------
    ' Desligar recálculo, alertas e atualização do ecrã
    ' --------------------------------------------------------
    calcOriginal = Application.Calculation
    Application.Calculation = pjManual
    alterouCalc = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --------------------------------------------------------
    ' 1. Calendário e opções do projeto
    ' --------------------------------------------------------
    etapa = "Calendário"
    Estado "a criar o calendário " & mCalNome & "..."
    CriarCalendario xlWB.Worksheets("MASTER_CALENDARIO")

    etapa = "Informação do projeto"
    Estado "a configurar o projeto..."
    ConfigurarProjeto

    ' --------------------------------------------------------
    ' 2. Tarefas
    ' --------------------------------------------------------
    etapa = "Criar tarefas e estrutura"
    nTarefas = CriarTarefas(dT)

    etapa = "Datas e propriedades das tarefas"
    Estado "a aplicar datas e durações..."
    ConfigurarTarefas dT

    etapa = "Predecessoras"
    Estado "a criar predecessoras..."
    nLigacoes = CriarPredecessoras(dT)

    ' --------------------------------------------------------
    ' 3. Recursos e atribuições
    ' --------------------------------------------------------
    etapa = "Recursos"
    Estado "a criar recursos..."
    If Not IsEmpty(dR) Then nRecursos = CriarRecursos(dR)

    etapa = "Atribuições"
    Estado "a criar atribuições..."
    If Not IsEmpty(dA) Then nAtrib = CriarAtribuicoes(dA)

    ' --------------------------------------------------------
    ' 4. Modos, tipos de tarefa e restrições
    ' --------------------------------------------------------
    etapa = "Modos de agendamento e restrições"
    Estado "a aplicar modos e restrições..."
    FinalizarTarefas dT, preservarDatas

    etapa = "Recalcular"
    Application.CalculateProject

    etapa = "Percentagem concluída"
    Estado "a aplicar % concluída..."
    AplicarPercentagens dT

    ' --------------------------------------------------------
    ' 5. Financeiro
    ' --------------------------------------------------------
    etapa = "Financeiro"
    Estado "a importar financeiro..."
    If Not IsEmpty(dF) Then nFin = ImportarFinanceiro(dF)

    ' --------------------------------------------------------
    ' Concluir
    ' --------------------------------------------------------
    etapa = "Concluir"
    Application.Calculation = calcOriginal
    alterouCalc = False
    Application.CalculateProject

    On Error Resume Next
    OutlineShowAllTasks
    On Error GoTo TrataErro

    resumo = "CWPower - Importação MASTER" & vbCrLf & _
             "Data: " & Format$(Now, "dd/mm/yyyy hh:nn") & vbCrLf & _
             "Ficheiro: " & CStr(caminho) & vbCrLf & _
             "Projeto: " & mProjCodigo & " - " & mProjNome & vbCrLf & _
             "Modo: " & IIf(preservarDatas, "DATAS PRESERVADAS", "LÓGICA ORIGINAL (recalculado)") & vbCrLf & _
             "Tarefas: " & nTarefas & " | Ligações: " & nLigacoes & _
             " | Recursos: " & nRecursos & " | Atribuições: " & nAtrib & _
             " | Linhas financeiras: " & nFin & vbCrLf & _
             "Avisos: " & mAvisos & vbCrLf

    ficheiroLog = GravarLog(CStr(caminho), resumo)

    msgFinal = "Project criado a partir do MASTER CWPower." & vbCrLf & vbCrLf & _
               "Calendário: " & mCalNome & vbCrLf & _
               "Tarefas criadas: " & nTarefas & vbCrLf & _
               "Predecessoras criadas: " & nLigacoes & vbCrLf & _
               "Recursos criados: " & nRecursos & vbCrLf & _
               "Atribuições criadas: " & nAtrib & vbCrLf & _
               "Linhas financeiras importadas: " & nFin & vbCrLf & vbCrLf & _
               IIf(preservarDatas, _
                   "Modo: DATAS PRESERVADAS (tarefas não-resumo em agendamento manual).", _
                   "Modo: LÓGICA ORIGINAL (modos Manual/Auto repostos e recalculado).") & vbCrLf & _
               "Financeiro: Custo1 = valor; Texto30 = rubrica." & vbCrLf & vbCrLf & _
               "Avisos: " & mAvisos

    If mAvisos > 0 Then
        msgFinal = msgFinal & vbCrLf & vbCrLf & Left$(mLog, 500)
        If Len(mLog) > 500 Then msgFinal = msgFinal & "..."
    End If

    If ficheiroLog <> "" Then msgFinal = msgFinal & vbCrLf & vbCrLf & "Log completo: " & ficheiroLog

SaidaLimpa:
    On Error Resume Next
    If alterouCalc Then Application.Calculation = calcOriginal
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

    If Not xlWB Is Nothing Then xlWB.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlWB = Nothing
    Set xlApp = Nothing
    On Error GoTo 0

    If msgFinal <> "" Then
        MsgBox msgFinal, IIf(mAvisos > 0, vbExclamation, vbInformation), "CWPower - Concluído"
    End If
    Exit Sub

TrataErro:
    msgFinal = ""
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Erro ao criar o Project." & vbCrLf & vbCrLf & _
           "Etapa: " & etapa & vbCrLf & _
           "Erro " & Err.Number & " - " & Err.Description & _
           IIf(mAvisos > 0, vbCrLf & vbCrLf & "Avisos até aqui:" & vbCrLf & Left$(mLog, 500), ""), _
           vbCritical, "CWPower - Erro"

    Resume SaidaLimpa

End Sub


' ============================================================
' LEITURA DO EXCEL
' ============================================================

' Lê o bloco de dados de uma folha MASTER (linha 4 até à última
' linha preenchida da coluna A). Devolve Empty se não houver dados.
Private Function LerTabela(ByVal ws As Object, ByVal nCols As Long) As Variant

    Dim ultima As Long

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If ultima < LINHA_DADOS Then
        LerTabela = Empty
    Else
        LerTabela = ws.Range(ws.Cells(LINHA_DADOS, 1), ws.Cells(ultima, nCols)).Value
    End If

End Function


Private Sub LerConfig(ByVal ws As Object)

    mProjCodigo = Texto(ValorConfig(ws, "CODIGO"))
    mProjNome = Texto(ValorConfig(ws, "NOME DO PROJETO"))
    mProjInicio = ValorConfig(ws, "DATA INICIO")

    mCalNome = Texto(ValorConfig(ws, "CALENDARIO"))
    If mCalNome = "" Then mCalNome = "CAL_OBRA_CWPOWER"

    mMinDia = NumeroOu(ValorConfig(ws, "MINUTOS/DIA"), 0)
    If mMinDia <= 0 Then mMinDia = NumeroOu(ValorConfig(ws, "HORAS/DIA"), 9) * 60

    mMinSemana = NumeroOu(ValorConfig(ws, "MINUTOS/SEMANA"), 0)
    If mMinSemana <= 0 Then mMinSemana = NumeroOu(ValorConfig(ws, "HORAS/SEMANA"), 54) * 60

    mDiasMes = NumeroOu(ValorConfig(ws, "DIAS/MES"), 26)

End Sub


' Procura na coluna A da MASTER_CONFIG um campo que comece por "chave"
' (sem acentos, maiúsculas) e devolve o valor da coluna B.
Private Function ValorConfig(ByVal ws As Object, ByVal chave As String) As Variant

    Dim d As Variant
    Dim i As Long

    ValorConfig = Empty
    d = ws.Range("A1:B60").Value

    For i = 1 To UBound(d, 1)
        If Left$(Normalizar(d(i, 1)), Len(chave)) = chave Then
            If Not IsError(d(i, 2)) Then ValorConfig = d(i, 2)
            Exit Function
        End If
    Next i

End Function


' ============================================================
' CALENDÁRIO
' ============================================================

Private Sub CriarCalendario(ByVal ws As Object)

    Dim d As Variant
    Dim ex As Variant
    Dim i As Long
    Dim k As Long
    Dim diaProj As Long
    Dim nT As Long
    Dim f(1 To 3) As String
    Dim tt(1 To 3) As String
    Dim ultima As Long
    Dim dIni As Date
    Dim dFim As Date
    Dim nomeEx As String

    If Not CalendarioExiste(mCalNome) Then
        BaseCalendarCreate Name:=mCalNome
    End If

    ' --- Semana tipo: linhas 4 a 10 (Segunda a Domingo) ---
    d = ws.Range("A4:H10").Value
    mNTurnos = 0

    For i = 1 To 7

        diaProj = NumeroDiaProject(d(i, 1))

        If diaProj = 0 Then

            Avisar "MASTER_CALENDARIO linha " & (i + 3) & ": dia '" & Texto(d(i, 1)) & "' não reconhecido."

        ElseIf EhSim(d(i, 2)) Then

            nT = LerTurnos(d, i, f, tt)

            If nT = 0 Then
                Avisar "MASTER_CALENDARIO linha " & (i + 3) & ": dia útil sem horário - mantido o horário do calendário base."
            Else
                On Error Resume Next
                AplicarHorario diaProj, 0, 0, nT, f, tt
                VerErro "Calendário: horário de " & Texto(d(i, 1))
                On Error GoTo 0

                ' O primeiro dia útil fica como horário de referência
                If mNTurnos = 0 Then
                    mNTurnos = nT
                    For k = 1 To nT
                        mTurnoIni(k) = f(k)
                        mTurnoFim(k) = tt(k)
                    Next k
                End If
            End If

        Else

            On Error Resume Next
            BaseCalendarEditDays Name:=mCalNome, WeekDay:=diaProj, Working:=False
            VerErro "Calendário: " & Texto(d(i, 1)) & " não útil"
            On Error GoTo 0

        End If

    Next i

    If mNTurnos = 0 Then
        mNTurnos = 2
        mTurnoIni(1) = "07:00": mTurnoFim(1) = "12:00"
        mTurnoIni(2) = "13:00": mTurnoFim(2) = "17:00"
    End If

    ' --- Feriados / exceções: desde a linha 14 ---
    ultima = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > ultima Then ultima = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    If ultima >= 14 Then

        ex = ws.Range(ws.Cells(14, 1), ws.Cells(ultima, 6)).Value

        For i = 1 To UBound(ex, 1)

            If EhSim(ex(i, 1)) Then

                If IsDate(ex(i, 4)) Then

                    dIni = DateValue(CDate(ex(i, 4)))
                    If IsDate(ex(i, 5)) Then dFim = DateValue(CDate(ex(i, 5))) Else dFim = dIni
                    If dFim < dIni Then dFim = dIni

                    nomeEx = Texto(ex(i, 3))
                    If nomeEx = "" Then nomeEx = Texto(ex(i, 2))
                    If nomeEx = "" Then nomeEx = "Exceção"
                    nomeEx = Left$(nomeEx & " " & Year(dIni), 255)

                    If EhSim(ex(i, 6)) Then
                        ' Dia excecionalmente útil, com o horário de referência
                        On Error Resume Next
                        AplicarHorario 0, dIni, dFim, mNTurnos, mTurnoIni, mTurnoFim
                        VerErro "Exceção de trabalho '" & nomeEx & "'"
                        On Error GoTo 0
                    Else
                        CriarFeriado dIni, dFim, nomeEx
                    End If

                Else
                    Avisar "MASTER_CALENDARIO linha " & (i + 13) & ": data inválida - exceção ignorada."
                End If

            End If

        Next i

    End If

End Sub


' Lê até 3 turnos (colunas C:H) de uma linha da semana tipo.
Private Function LerTurnos(ByRef d As Variant, ByVal i As Long, _
                           ByRef f() As String, ByRef tt() As String) As Long

    Dim k As Long
    Dim n As Long
    Dim a As String
    Dim b As String

    For k = 1 To 3
        a = HoraTxt(d(i, 1 + 2 * k))
        b = HoraTxt(d(i, 2 + 2 * k))
        If a <> "" And b <> "" Then
            n = n + 1
            f(n) = a
            tt(n) = b
        End If
    Next k

    LerTurnos = n

End Function


' Aplica um horário (1 a 3 turnos) a um dia da semana (diaProj > 0)
' ou a um intervalo de datas (diaProj = 0).
Private Sub AplicarHorario(ByVal diaProj As Long, ByVal dIni As Date, ByVal dFim As Date, _
                           ByVal n As Long, ByRef f() As String, ByRef tt() As String)

    If diaProj > 0 Then

        ' Limpar turnos que sobram do calendário base (o "Padrão" tem 2)
        On Error Resume Next
        With ActiveProject.BaseCalendars(mCalNome).WeekDays(diaProj)
            If n < 3 Then .Shift3.Clear
            If n < 2 Then .Shift2.Clear
        End With
        Err.Clear
        On Error GoTo 0

        Select Case n
            Case 1
                BaseCalendarEditDays Name:=mCalNome, WeekDay:=diaProj, Working:=True, _
                    From1:=f(1), To1:=tt(1)
            Case 2
                BaseCalendarEditDays Name:=mCalNome, WeekDay:=diaProj, Working:=True, _
                    From1:=f(1), To1:=tt(1), From2:=f(2), To2:=tt(2)
            Case Else
                BaseCalendarEditDays Name:=mCalNome, WeekDay:=diaProj, Working:=True, _
                    From1:=f(1), To1:=tt(1), From2:=f(2), To2:=tt(2), From3:=f(3), To3:=tt(3)
        End Select

    Else

        Select Case n
            Case 1
                BaseCalendarEditDays Name:=mCalNome, StartDate:=dIni, EndDate:=dFim, Working:=True, _
                    From1:=f(1), To1:=tt(1)
            Case 2
                BaseCalendarEditDays Name:=mCalNome, StartDate:=dIni, EndDate:=dFim, Working:=True, _
                    From1:=f(1), To1:=tt(1), From2:=f(2), To2:=tt(2)
            Case Else
                BaseCalendarEditDays Name:=mCalNome, StartDate:=dIni, EndDate:=dFim, Working:=True, _
                    From1:=f(1), To1:=tt(1), From2:=f(2), To2:=tt(2), From3:=f(3), To3:=tt(3)
        End Select

    End If

End Sub


Private Sub CriarFeriado(ByVal dIni As Date, ByVal dFim As Date, ByVal nome As String)

    Dim ok As Boolean

    On Error Resume Next

    ' Preferência: exceção com nome (visível em "Alterar Período Útil")
    ' Type 1 = pjDaily (literal para compilar em qualquer versão)
    ActiveProject.BaseCalendars(mCalNome).Exceptions.Add Type:=1, _
        Start:=dIni, Finish:=dFim, Name:=nome
    ok = (Err.Number = 0)
    Err.Clear

    If Not ok Then
        BaseCalendarEditDays Name:=mCalNome, StartDate:=dIni, EndDate:=dFim, Working:=False
        VerErro "Feriado '" & nome & "' (" & Format$(dIni, "dd/mm/yyyy") & ")"
    End If

    On Error GoTo 0

End Sub


Private Sub ConfigurarProjeto()

    Dim hIni As String
    Dim hFim As String

    If IsDate(mProjInicio) Then
        ProjectSummaryInfo Start:=CDate(mProjInicio), ScheduleFrom:=pjProjectStart, Calendar:=mCalNome
    Else
        ProjectSummaryInfo ScheduleFrom:=pjProjectStart, Calendar:=mCalNome
    End If

    hIni = mTurnoIni(1)
    hFim = mTurnoFim(mNTurnos)
    If hIni = "" Then hIni = "07:00"
    If hFim = "" Then hFim = "17:00"

    On Error Resume Next

    OptionsCalendar StartWeekOn:=pjMonday, StartTime:=hIni, FinishTime:=hFim, _
        HoursPerDay:=mMinDia / 60, HoursPerWeek:=mMinSemana / 60, DaysPerMonth:=mDiasMes
    VerErro "Opções de calendário (horas/dia, horas/semana, dias/mês)"

    If mProjNome <> "" Then ActiveProject.BuiltinDocumentProperties("Title").Value = mProjNome
    If mProjCodigo <> "" Then ActiveProject.BuiltinDocumentProperties("Subject").Value = mProjCodigo
    Err.Clear

    On Error GoTo 0

End Sub


' ============================================================
' TAREFAS
' ============================================================

Private Function CriarTarefas(ByRef d As Variant) As Long

    Dim i As Long
    Dim id As Long
    Dim nivel As Long
    Dim nome As String
    Dim t As Task
    Dim n As Long

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, T_ID)) Then

            id = CLng(d(i, T_ID))

            nome = Texto(d(i, T_NOME))
            nome = Replace(Replace(nome, vbCr, " "), vbLf, " ")
            If nome = "" Then nome = "(sem nome) " & id
            If Len(nome) > 255 Then
                Avisar "Tarefa " & id & ": nome com mais de 255 caracteres - foi cortado."
                nome = Left$(nome, 255)
            End If

            Set t = ActiveProject.Tasks.Add(Name:=nome)
            t.Manual = True

            If ExisteChave(mTarefas, CStr(id)) Then
                Avisar "MASTER_TAREFAS linha " & LinhaExcel(i) & ": ID " & id & " repetido - tarefa criada mas não usada em ligações/atribuições."
            Else
                mTarefas.Add t, CStr(id)
            End If

            ' Estrutura: as linhas vêm por ordem, por isso basta definir o nível
            nivel = CLng(NumeroOu(d(i, T_NIVEL), 1))
            If nivel < 1 Then nivel = 1

            If nivel > 1 Then
                On Error Resume Next
                t.OutlineLevel = nivel
                VerErro "Tarefa " & id & ": nível de destaque " & nivel & " não aplicado"
                On Error GoTo 0
            End If

            n = n + 1
            If n Mod 20 = 0 Then Estado "a criar tarefas... " & n

        ElseIf Texto(d(i, T_NOME)) <> "" Then

            Avisar "MASTER_TAREFAS linha " & LinhaExcel(i) & ": sem ID numérico - ignorada."

        End If

    Next i

    CriarTarefas = n

End Function


Private Sub ConfigurarTarefas(ByRef d As Variant)

    Dim i As Long
    Dim id As Long
    Dim t As Task
    Dim cal As String
    Dim dur As Double

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, T_ID)) Then

            id = CLng(d(i, T_ID))
            Set t = ObterTarefa(id)

            If Not t Is Nothing Then

                If EhSim(d(i, T_RESUMO)) <> t.Summary Then
                    Avisar "Tarefa " & id & ": no Excel 'Resumo?' = " & Texto(d(i, T_RESUMO)) & _
                           " mas no Project " & IIf(t.Summary, "é", "não é") & " resumo (verifique o Outline Level)."
                End If

                On Error Resume Next

                ' Calendário próprio da tarefa (só se existir no Project)
                cal = Texto(d(i, T_CAL))
                If cal <> "" And Normalizar(cal) <> "NONE" And Normalizar(cal) <> "NENHUM" Then
                    If CalendarioExiste(cal) Then
                        t.Calendar = cal
                        VerErro "Tarefa " & id & ": calendário '" & cal & "'"
                    Else
                        Avisar "Tarefa " & id & ": calendário '" & cal & "' não existe no Project - ignorado."
                    End If
                End If

                If t.Summary Then

                    ' Resumos: manual com datas próprias, ou automático (soma das subtarefas)
                    If Normalizar(d(i, T_MODO)) = "MANUAL" Then
                        t.Manual = True
                        AplicarDatas t, id, d(i, T_INICIO), d(i, T_FIM), Empty
                    Else
                        t.Manual = False
                    End If
                    VerErro "Tarefa resumo " & id

                Else

                    t.Manual = True
                    dur = NumeroOu(d(i, T_DURACAO), -1)

                    If dur = 0 Then
                        AplicarDatas t, id, d(i, T_INICIO), Empty, 0
                    Else
                        AplicarDatas t, id, d(i, T_INICIO), d(i, T_FIM), d(i, T_DURACAO)
                    End If

                    If EhSim(d(i, T_MARCO)) Then t.Milestone = True
                    t.Estimated = False
                    VerErro "Tarefa " & id

                End If

                If IsDate(d(i, T_DEADLINE)) Then
                    t.Deadline = CDate(d(i, T_DEADLINE))
                    VerErro "Tarefa " & id & ": prazo (deadline)"
                End If

                If Texto(d(i, T_NOTAS)) <> "" Then
                    t.Notes = Texto(d(i, T_NOTAS))
                    VerErro "Tarefa " & id & ": notas"
                End If

                On Error GoTo 0

            End If

        End If

    Next i

End Sub


' Início + fim (ou início + duração em minutos).
Private Sub AplicarDatas(ByVal t As Task, ByVal id As Long, _
                         ByVal inicio As Variant, ByVal fim As Variant, ByVal dur As Variant)

    On Error Resume Next

    If IsDate(inicio) Then
        t.Start = CDate(inicio)
        VerErro "Tarefa " & id & ": data de início não aplicada"
    End If

    If IsDate(fim) Then
        t.Finish = CDate(fim)
        VerErro "Tarefa " & id & ": data de fim não aplicada"
    ElseIf EhNumero(dur) Then
        t.Duration = CDbl(dur)
        VerErro "Tarefa " & id & ": duração não aplicada"
    End If

    On Error GoTo 0

End Sub


Private Function CriarPredecessoras(ByRef d As Variant) As Long

    Dim i As Long
    Dim id As Long
    Dim txt As String
    Dim t As Task
    Dim n As Long

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, T_ID)) Then

            txt = Texto(d(i, T_PRED))

            If txt <> "" Then
                id = CLng(d(i, T_ID))
                Set t = ObterTarefa(id)
                If Not t Is Nothing Then n = n + LigarPredecessoras(t, id, txt)
            End If

        End If

    Next i

    CriarPredecessoras = n

End Function


Private Function LigarPredecessoras(ByVal t As Task, ByVal id As Long, ByVal txt As String) As Long

    Dim partes As Collection
    Dim p As Variant
    Dim predId As Long
    Dim tipo As Long
    Dim lag As Double
    Dim lagTxt As String
    Dim aviso As String
    Dim tp As Task
    Dim n As Long

    Set partes = DividirPredecessoras(txt)

    For Each p In partes

        If InterpretarLigacao(CStr(p), predId, tipo, lag, lagTxt, aviso) Then

            If aviso <> "" Then Avisar "Tarefa " & id & ": predecessora '" & p & "' - " & aviso

            Set tp = ObterTarefa(predId)

            If tp Is Nothing Then
                Avisar "Tarefa " & id & ": predecessora " & predId & " não existe - ligação ignorada."
            ElseIf predId = id Then
                Avisar "Tarefa " & id & ": tarefa ligada a si própria - ligação ignorada."
            Else
                On Error Resume Next
                If lagTxt <> "" Then
                    t.TaskDependencies.Add From:=tp, Type:=tipo, Lag:=lagTxt
                Else
                    t.TaskDependencies.Add From:=tp, Type:=tipo, Lag:=lag
                End If
                If Err.Number = 0 Then n = n + 1 Else VerErro "Tarefa " & id & ": ligação '" & p & "' não criada"
                On Error GoTo 0
            End If

        Else

            Avisar "Tarefa " & id & ": predecessora '" & p & "' não reconhecida - ignorada."

        End If

    Next p

    LigarPredecessoras = n

End Function


' Divide "24;25", "45FS+10 days;46FS+10 days" ou "24,25" em ligações.
' A vírgula só é decimal quando aparece dentro do número do lag (ex.: +1,5d).
Private Function DividirPredecessoras(ByVal txt As String) As Collection

    Dim c As Collection
    Dim i As Long
    Dim ch As String
    Dim atual As String
    Dim emLag As Boolean

    Set c = New Collection

    For i = 1 To Len(txt)

        ch = Mid$(txt, i, 1)

        Select Case ch
            Case ";"
                AdicionarParte c, atual
                atual = ""
                emLag = False
            Case ","
                If emLag And (Mid$(txt, i + 1, 1) Like "#") Then
                    atual = atual & ch
                Else
                    AdicionarParte c, atual
                    atual = ""
                    emLag = False
                End If
            Case "+", "-"
                emLag = True
                atual = atual & ch
            Case Else
                If emLag And Not (ch Like "#" Or ch = "." Or ch = " ") Then emLag = False
                atual = atual & ch
        End Select

    Next i

    AdicionarParte c, atual

    Set DividirPredecessoras = c

End Function


Private Sub AdicionarParte(ByVal c As Collection, ByVal s As String)
    s = Trim$(s)
    If s <> "" Then c.Add s
End Sub


' Interpreta uma ligação: <ID>[tipo][+/-lag[unidade]]
'   tipo:    FS SS FF SF (inglês) | TI II TT IT | FI IF (português)
'   unidade: d/day/dia, h/hr/hora, min/minuto, w/wk/sem/semana,
'            mo/month/mes, ed/edays (decorridos, aproximado), %
Private Function InterpretarLigacao(ByVal s As String, ByRef predId As Long, ByRef tipo As Long, _
                                    ByRef lag As Double, ByRef lagTxt As String, _
                                    ByRef aviso As String) As Boolean

    Dim i As Long
    Dim resto As String
    Dim codTipo As String
    Dim lagParte As String
    Dim numTxt As String
    Dim unidade As String
    Dim sinal As Double
    Dim fator As Double

    InterpretarLigacao = False
    lag = 0
    lagTxt = ""
    aviso = ""

    s = Replace(Normalizar(s), " ", "")

    ' ID da predecessora
    i = 1
    Do While i <= Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Do
        i = i + 1
    Loop
    If i = 1 Then Exit Function

    predId = CLng(Left$(s, i - 1))
    resto = Mid$(s, i)

    ' Tipo de ligação (letras até ao sinal do lag)
    i = 1
    Do While i <= Len(resto)
        If Mid$(resto, i, 1) = "+" Or Mid$(resto, i, 1) = "-" Then Exit Do
        i = i + 1
    Loop
    codTipo = Left$(resto, i - 1)
    lagParte = Mid$(resto, i)

    Select Case codTipo
        Case "", "FS", "TI", "FI": tipo = pjFinishToStart
        Case "SS", "II": tipo = pjStartToStart
        Case "FF", "TT": tipo = pjFinishToFinish
        Case "SF", "IT", "IF": tipo = pjStartToFinish
        Case Else: Exit Function
    End Select

    ' Lag
    If lagParte <> "" Then

        If Left$(lagParte, 1) = "-" Then sinal = -1 Else sinal = 1
        lagParte = Mid$(lagParte, 2)

        i = 1
        Do While i <= Len(lagParte)
            If Not (Mid$(lagParte, i, 1) Like "#" Or Mid$(lagParte, i, 1) = "." Or Mid$(lagParte, i, 1) = ",") Then Exit Do
            i = i + 1
        Loop

        numTxt = Replace(Left$(lagParte, i - 1), ",", ".")
        If numTxt = "" Then Exit Function

        unidade = Replace(Mid$(lagParte, i), "?", "")

        If unidade = "%" Then
            lagTxt = CStr(sinal * Val(numTxt)) & "%"
        Else
            fator = MinutosPorUnidade(unidade, aviso)
            If fator < 0 Then
                aviso = "unidade de lag '" & unidade & "' desconhecida - ligação criada sem lag."
            Else
                lag = sinal * Val(numTxt) * fator
            End If
        End If

    End If

    InterpretarLigacao = True

End Function


Private Function MinutosPorUnidade(ByVal u As String, ByRef aviso As String) As Double

    Select Case u
        Case "", "D", "DY", "DYS", "DAY", "DAYS", "DIA", "DIAS"
            MinutosPorUnidade = mMinDia
        Case "H", "HR", "HRS", "HOUR", "HOURS", "HORA", "HORAS"
            MinutosPorUnidade = 60
        Case "M", "MIN", "MINS", "MINUTE", "MINUTES", "MINUTO", "MINUTOS"
            MinutosPorUnidade = 1
        Case "W", "WK", "WKS", "WEEK", "WEEKS", "S", "SEM", "SEMANA", "SEMANAS"
            MinutosPorUnidade = mMinSemana
        Case "MO", "MON", "MONS", "MONTH", "MONTHS", "ME", "MES", "MESES"
            MinutosPorUnidade = mDiasMes * mMinDia
        Case "ED", "EDAY", "EDAYS", "DD", "DIASDECORRIDOS"
            MinutosPorUnidade = mMinSemana / 7
            aviso = "lag em dias decorridos convertido aproximadamente para dias úteis."
        Case "EW", "EWK", "EWEEK", "EWEEKS", "SEMD", "SEMANASDECORRIDAS"
            MinutosPorUnidade = mMinSemana
            aviso = "lag em semanas decorridas convertido aproximadamente para semanas úteis."
        Case Else
            MinutosPorUnidade = -1
    End Select

End Function


' Modos de agendamento (se pedido), tipo de tarefa e restrições.
Private Sub FinalizarTarefas(ByRef d As Variant, ByVal preservarDatas As Boolean)

    Dim i As Long
    Dim id As Long
    Dim t As Task

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, T_ID)) Then

            id = CLng(d(i, T_ID))
            Set t = ObterTarefa(id)

            If Not t Is Nothing Then

                If Not t.Summary Then

                    On Error Resume Next

                    If Not preservarDatas Then
                        If Normalizar(d(i, T_MODO)) <> "MANUAL" Then
                            t.Manual = False
                            VerErro "Tarefa " & id & ": modo automático"
                            If EhNumero(d(i, T_DURACAO)) Then
                                t.Duration = CDbl(d(i, T_DURACAO))
                                VerErro "Tarefa " & id & ": duração"
                            End If
                        End If
                    End If

                    ' Tipo original (as atribuições foram criadas em duração fixa)
                    t.Type = TipoTarefa(d(i, T_TIPO))
                    VerErro "Tarefa " & id & ": tipo de tarefa"
                    If t.Type <> pjFixedWork Then
                        t.EffortDriven = EhSim(d(i, T_EFFORT))
                        VerErro "Tarefa " & id & ": controlada pelo esforço"
                    End If

                    t.Estimated = False
                    Err.Clear

                    On Error GoTo 0

                    AplicarRestricao t, id, d(i, T_CTIPO), d(i, T_CDATA)

                End If

            End If

        End If

    Next i

End Sub


Private Sub AplicarRestricao(ByVal t As Task, ByVal id As Long, _
                             ByVal tipoTxt As Variant, ByVal dataR As Variant)

    Dim ct As Long

    ct = TipoRestricao(tipoTxt)

    If ct < 0 Then
        Avisar "Tarefa " & id & ": restrição '" & Texto(tipoTxt) & "' desconhecida - ignorada."
        Exit Sub
    End If

    On Error Resume Next

    t.ConstraintType = ct
    VerErro "Tarefa " & id & ": restrição"

    If ct <> pjASAP And ct <> pjALAP Then
        If IsDate(dataR) Then
            t.ConstraintDate = CDate(dataR)
            VerErro "Tarefa " & id & ": data da restrição"
        Else
            Avisar "Tarefa " & id & ": restrição '" & Texto(tipoTxt) & "' sem data - usada a data de início."
        End If
    End If

    On Error GoTo 0

End Sub


Private Sub AplicarPercentagens(ByRef d As Variant)

    Dim i As Long
    Dim id As Long
    Dim t As Task
    Dim p As Double

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, T_ID)) And EhNumero(d(i, T_PCT)) Then

            id = CLng(d(i, T_ID))
            Set t = ObterTarefa(id)

            If Not t Is Nothing Then
                If Not t.Summary Then

                    p = CDbl(d(i, T_PCT))
                    If p <= 1 Then p = p * 100     ' Excel em fração (0,7 = 70%)

                    If p > 0 And p <= 100 Then
                        On Error Resume Next
                        t.PercentComplete = CLng(Round(p, 0))
                        VerErro "Tarefa " & id & ": % concluída"
                        On Error GoTo 0
                    End If

                End If
            End If

        End If

    Next i

End Sub


' ============================================================
' RECURSOS E ATRIBUIÇÕES
' ============================================================

Private Function CriarRecursos(ByRef d As Variant) As Long

    Dim i As Long
    Dim nome As String
    Dim nomeLimpo As String
    Dim r As Resource
    Dim n As Long

    For i = 1 To UBound(d, 1)

        nome = Texto(d(i, 1))

        If nome <> "" Then

            nomeLimpo = LimparNomeRecurso(nome)
            If nomeLimpo <> nome Then
                Avisar "Recurso '" & nome & "': caracteres inválidos substituídos -> '" & nomeLimpo & "'."
            End If

            If ExisteChave(mRecursos, nomeLimpo) Then

                Avisar "MASTER_RECURSOS linha " & LinhaExcel(i) & ": recurso '" & nome & "' repetido - ignorado."

            Else

                Set r = ActiveProject.Resources.Add(Name:=nomeLimpo)
                mRecursos.Add r, nomeLimpo

                On Error Resume Next
                r.Type = pjResourceTypeWork
                r.Group = Texto(d(i, 2))
                If EhNumero(d(i, 3)) Then
                    If CDbl(d(i, 3)) > 0 Then r.MaxUnits = CDbl(d(i, 3))
                End If
                r.BaseCalendar = mCalNome
                If Texto(d(i, 4)) <> "" Then r.Text30 = Texto(d(i, 4))
                VerErro "Recurso '" & nome & "'"
                On Error GoTo 0

                n = n + 1

            End If

        End If

    Next i

    CriarRecursos = n

End Function


Private Function CriarAtribuicoes(ByRef d As Variant) As Long

    Dim i As Long
    Dim id As Long
    Dim nomeR As String
    Dim u As Double
    Dim t As Task
    Dim r As Resource
    Dim n As Long

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, 1)) Then

            id = CLng(d(i, 1))
            nomeR = LimparNomeRecurso(Texto(d(i, 3)))
            u = NumeroOu(d(i, 4), 0)

            Set t = ObterTarefa(id)
            Set r = ObterRecurso(nomeR)

            If t Is Nothing Then
                Avisar "MASTER_ATRIBUICOES linha " & LinhaExcel(i) & ": tarefa " & id & " não existe - ignorada."
            ElseIf r Is Nothing Then
                Avisar "MASTER_ATRIBUICOES linha " & LinhaExcel(i) & ": recurso '" & nomeR & "' não existe em MASTER_RECURSOS - ignorada."
            ElseIf u <= 0 Then
                Avisar "MASTER_ATRIBUICOES linha " & LinhaExcel(i) & ": unidades inválidas - ignorada."
            ElseIf t.Summary Then
                Avisar "MASTER_ATRIBUICOES linha " & LinhaExcel(i) & ": tarefa " & id & " é resumo - ignorada."
            Else

                On Error Resume Next

                ' Duração fixa durante a atribuição para não mexer nas datas
                t.Type = pjFixedDuration
                t.EffortDriven = False
                Err.Clear

                t.Assignments.Add ResourceID:=r.ID, Units:=u
                If Err.Number = 0 Then n = n + 1 Else VerErro "Atribuição '" & nomeR & "' à tarefa " & id

                On Error GoTo 0

            End If

        End If

    Next i

    CriarAtribuicoes = n

End Function


' ============================================================
' FINANCEIRO
' ============================================================

Private Function ImportarFinanceiro(ByRef d As Variant) As Long

    Dim i As Long
    Dim id As Long
    Dim t As Task
    Dim rubrica As String
    Dim valor As Double
    Dim vistos As Collection
    Dim n As Long

    Set vistos = New Collection

    For i = 1 To UBound(d, 1)

        If EhNumero(d(i, 1)) Then

            id = CLng(d(i, 1))
            rubrica = Texto(d(i, 3))
            valor = NumeroOu(d(i, 4), 0)

            Set t = ObterTarefa(id)

            If t Is Nothing Then

                Avisar "MASTER_FINANCEIRO linha " & LinhaExcel(i) & ": tarefa " & id & " não existe - ignorada."

            Else

                On Error Resume Next

                If ExisteChave(vistos, CStr(id)) Then
                    ' Mais do que uma linha para a mesma tarefa: soma valores
                    t.Cost1 = t.Cost1 + valor
                    If rubrica <> "" And InStr(1, t.Text30, rubrica, vbTextCompare) = 0 Then
                        t.Text30 = Left$(t.Text30 & " / " & rubrica, 255)
                    End If
                Else
                    vistos.Add id, CStr(id)
                    t.Cost1 = valor
                    t.Text30 = Left$(rubrica, 255)
                End If

                If Err.Number = 0 Then n = n + 1 Else VerErro "Financeiro da tarefa " & id

                On Error GoTo 0

            End If

        End If

    Next i

    ImportarFinanceiro = n

End Function


' ============================================================
' AUXILIARES
' ============================================================

Private Function ContarTarefas() As Long

    Dim t As Task
    Dim n As Long

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then n = n + 1
    Next t

    ContarTarefas = n

End Function


Private Function FolhaExiste(ByVal xlWB As Object, ByVal nomeFolha As String) As Boolean

    Dim ws As Object

    On Error Resume Next
    Set ws = xlWB.Worksheets(nomeFolha)
    FolhaExiste = Not ws Is Nothing
    Err.Clear

End Function


Private Function CalendarioExiste(ByVal nome As String) As Boolean

    Dim c As Calendar

    For Each c In ActiveProject.BaseCalendars
        If StrComp(c.Name, nome, vbTextCompare) = 0 Then
            CalendarioExiste = True
            Exit Function
        End If
    Next c

End Function


Private Function ObterTarefa(ByVal id As Long) As Task
    On Error Resume Next
    Set ObterTarefa = mTarefas.Item(CStr(id))
    Err.Clear
End Function


Private Function ObterRecurso(ByVal nome As String) As Resource
    If nome = "" Then Exit Function
    On Error Resume Next
    Set ObterRecurso = mRecursos.Item(nome)
    Err.Clear
End Function


Private Function ExisteChave(ByVal c As Collection, ByVal chave As String) As Boolean

    Dim tmp As Boolean

    On Error Resume Next
    Err.Clear
    tmp = IsObject(c.Item(chave))
    ExisteChave = (Err.Number = 0)
    Err.Clear

End Function


' Converte qualquer valor de célula em texto, sem rebentar com erros (#N/D, etc.)
Private Function Texto(ByVal v As Variant) As String

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsObject(v) Then
        Texto = ""
    Else
        Texto = Trim$(Replace(CStr(v), ChrW(160), " "))
    End If

End Function


' Maiúsculas, sem acentos e sem espaços nas pontas. Usa ChrW para não
' depender da codificação do ficheiro .bas.
Private Function Normalizar(ByVal v As Variant) As String

    Dim s As String

    s = UCase$(Texto(v))

    s = Replace(s, ChrW(193), "A"): s = Replace(s, ChrW(192), "A")
    s = Replace(s, ChrW(194), "A"): s = Replace(s, ChrW(195), "A")
    s = Replace(s, ChrW(196), "A")
    s = Replace(s, ChrW(201), "E"): s = Replace(s, ChrW(200), "E")
    s = Replace(s, ChrW(202), "E"): s = Replace(s, ChrW(203), "E")
    s = Replace(s, ChrW(205), "I"): s = Replace(s, ChrW(204), "I")
    s = Replace(s, ChrW(206), "I"): s = Replace(s, ChrW(207), "I")
    s = Replace(s, ChrW(211), "O"): s = Replace(s, ChrW(210), "O")
    s = Replace(s, ChrW(212), "O"): s = Replace(s, ChrW(213), "O")
    s = Replace(s, ChrW(214), "O")
    s = Replace(s, ChrW(218), "U"): s = Replace(s, ChrW(217), "U")
    s = Replace(s, ChrW(219), "U"): s = Replace(s, ChrW(220), "U")
    s = Replace(s, ChrW(199), "C")

    Normalizar = s

End Function


Private Function EhSim(ByVal v As Variant) As Boolean

    If VarType(v) = vbBoolean Then
        EhSim = v
        Exit Function
    End If

    Select Case Normalizar(v)
        Case "SIM", "S", "YES", "Y", "TRUE", "VERDADEIRO", "1"
            EhSim = True
        Case Else
            EhSim = False
    End Select

End Function


' IMPORTANTE: IsNumeric(Empty) = True em VBA, por isso testa-se IsEmpty.
Private Function EhNumero(ByVal v As Variant) As Boolean

    If IsEmpty(v) Or IsError(v) Or IsNull(v) Or IsObject(v) Then
        EhNumero = False
    ElseIf VarType(v) = vbString Then
        EhNumero = (Trim$(v) <> "" And IsNumeric(v))
    ElseIf VarType(v) = vbBoolean Then
        EhNumero = False
    Else
        EhNumero = IsNumeric(v)
    End If

End Function


Private Function NumeroOu(ByVal v As Variant, ByVal padrao As Double) As Double

    If EhNumero(v) Then
        NumeroOu = CDbl(v)
    Else
        NumeroOu = padrao
    End If

End Function


Private Function NumeroDiaProject(ByVal v As Variant) As Long

    Select Case Left$(Normalizar(v), 3)
        Case "SEG", "MON": NumeroDiaProject = pjMonday
        Case "TER", "TUE": NumeroDiaProject = pjTuesday
        Case "QUA", "WED": NumeroDiaProject = pjWednesday
        Case "QUI", "THU": NumeroDiaProject = pjThursday
        Case "SEX", "FRI": NumeroDiaProject = pjFriday
        Case "SAB", "SAT": NumeroDiaProject = pjSaturday
        Case "DOM", "SUN": NumeroDiaProject = pjSunday
        Case Else: NumeroDiaProject = 0
    End Select

End Function


' Hora do Excel (Date, fração de dia ou texto) -> "hh:mm"
Private Function HoraTxt(ByVal v As Variant) As String

    HoraTxt = ""

    If IsEmpty(v) Or IsError(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        HoraTxt = Format$(v, "hh:nn")
    ElseIf EhNumero(v) Then
        If CDbl(v) >= 0 And CDbl(v) < 1 Then HoraTxt = Format$(CDate(CDbl(v)), "hh:nn")
    ElseIf IsDate(v) Then
        HoraTxt = Format$(CDate(v), "hh:nn")
    End If

End Function


Private Function TipoTarefa(ByVal v As Variant) As Long

    Dim s As String

    s = Normalizar(v)

    If InStr(s, "DURA") > 0 Then
        TipoTarefa = pjFixedDuration
    ElseIf InStr(s, "WORK") > 0 Or InStr(s, "TRABALHO") > 0 Then
        TipoTarefa = pjFixedWork
    Else
        TipoTarefa = pjFixedUnits
    End If

End Function


Private Function TipoRestricao(ByVal v As Variant) As Long

    Select Case Normalizar(v)
        Case "", "ASAP", "AS SOON AS POSSIBLE", "O MAIS CEDO POSSIVEL", "TAO CEDO QUANTO POSSIVEL"
            TipoRestricao = pjASAP
        Case "ALAP", "AS LATE AS POSSIBLE", "O MAIS TARDE POSSIVEL", "TAO TARDE QUANTO POSSIVEL"
            TipoRestricao = pjALAP
        Case "MSO", "MUST START ON", "DEVE INICIAR EM", "TEM DE INICIAR EM"
            TipoRestricao = pjMSO
        Case "MFO", "MUST FINISH ON", "DEVE TERMINAR EM", "TEM DE TERMINAR EM"
            TipoRestricao = pjMFO
        Case "SNET", "START NO EARLIER THAN", "NAO INICIAR ANTES DE", "INICIAR NAO ANTES DE"
            TipoRestricao = pjSNET
        Case "SNLT", "START NO LATER THAN", "NAO INICIAR DEPOIS DE", "INICIAR NAO DEPOIS DE"
            TipoRestricao = pjSNLT
        Case "FNET", "FINISH NO EARLIER THAN", "NAO TERMINAR ANTES DE", "TERMINAR NAO ANTES DE"
            TipoRestricao = pjFNET
        Case "FNLT", "FINISH NO LATER THAN", "NAO TERMINAR DEPOIS DE", "TERMINAR NAO DEPOIS DE"
            TipoRestricao = pjFNLT
        Case Else
            TipoRestricao = -1
    End Select

End Function


' O Project não aceita vírgulas, ponto e vírgula nem parênteses retos
' no nome do recurso (são separadores de lista).
Private Function LimparNomeRecurso(ByVal s As String) As String

    s = Replace(s, ",", " ")
    s = Replace(s, ";", " ")
    s = Replace(s, "[", "(")
    s = Replace(s, "]", ")")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    LimparNomeRecurso = Trim$(s)

End Function


Private Function MenorData(ByRef d As Variant, ByVal coluna As Long) As Variant

    Dim i As Long

    MenorData = Empty

    For i = 1 To UBound(d, 1)
        If IsDate(d(i, coluna)) Then
            If IsEmpty(MenorData) Then
                MenorData = CDate(d(i, coluna))
            ElseIf CDate(d(i, coluna)) < MenorData Then
                MenorData = CDate(d(i, coluna))
            End If
        End If
    Next i

End Function


Private Function LinhaExcel(ByVal i As Long) As Long
    LinhaExcel = i + LINHA_DADOS - 1
End Function


Private Sub Avisar(ByVal msg As String)
    mAvisos = mAvisos + 1
    mLog = mLog & "- " & msg & vbCrLf
    Debug.Print msg
End Sub


' Regista o erro atual (se houver) como aviso e limpa-o.
Private Sub VerErro(ByVal contexto As String)
    If Err.Number <> 0 Then
        Avisar contexto & " (erro " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    End If
End Sub


Private Sub Estado(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = "CWPower: " & msg
    DoEvents
End Sub


' Grava o log ao lado do Excel (ou na pasta TEMP se não for possível).
Private Function GravarLog(ByVal caminhoExcel As String, ByVal resumo As String) As String

    Dim f As Integer
    Dim caminho As String

    On Error Resume Next

    caminho = caminhoExcel
    If InStrRev(caminho, ".") > 0 Then caminho = Left$(caminho, InStrRev(caminho, ".") - 1)
    caminho = caminho & "_LOG_IMPORTACAO.txt"

    f = FreeFile
    Open caminho For Output As #f

    If Err.Number <> 0 Then
        Err.Clear
        caminho = Environ$("TEMP") & "\CWPower_LOG_IMPORTACAO.txt"
        f = FreeFile
        Open caminho For Output As #f
    End If

    If Err.Number <> 0 Then
        Err.Clear
        GravarLog = ""
        Exit Function
    End If

    Print #f, resumo
    If mLog = "" Then
        Print #f, "Sem avisos."
    Else
        Print #f, "AVISOS:"
        Print #f, mLog
    End If
    Close #f

    GravarLog = caminho

End Function
