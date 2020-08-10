Attribute VB_Name = "ModProcessamento"

Public ContAuxTimer As Integer
Public tempoAdecorrer As Integer
Dim LinhaAtualGrid As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Iniciar_Processamento()

' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
'todo list:
'1 - percorrer o gridMonitoramento para capturar o tempo a ser usado no processo
'2 - iniciar time do processo
'3 - atualizar status
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------

    Dim QtdLinhas As Integer
    Dim IdLinhaAtual As Integer
    
    With frmPainelControleEstacionamento.griMonitoramento
            
        For I = 1 To .Rows - 1 'corre uma vez para marcar o status inicial igual para todos
            .Row = I 'linha atual
            .Col = 3 'coluna status
            .Text = "Aguardando"
        Next I
        
    End With

    Call Processar_Movimento

End Sub

Sub Processar_Movimento()
    
    With frmPainelControleEstacionamento.griMonitoramento
           
        For I = 1 To .Rows - 1 'corre a lista para checar o status
            .Row = I 'linha atual
            .Col = 3 'coluna status
            If .Text = "Aguardando" Then 'entrada para novo processamento
                   LinhaAtualGrid = I
                   .Col = 2 'captura o tempo (seg)
                   tempoAdecorrer = Val(.Text)
                   ContAuxTimer = 0 'zera o contador auxiliar do timer
                   frmPainelControleEstacionamento.timerProcessar.Enabled = True
                   Exit For 'os demais status permacem como estão até a conclusão do registro corrente
            End If
        Next I
        
    End With
    
    Call Atulizar_LabelContadores

End Sub


Sub AtualizaStatus()

    With frmPainelControleEstacionamento.griMonitoramento
        .Row = LinhaAtualGrid
        .Col = 3
        .Text = "PROCESSANDO"
        tmp = 1000 * funcIntGerarNumRND(5, 3)
        Sleep tmp '1000 * funcIntGerarNumRND(5, 3) 'gera Num. Int enttre 3 e 5 e converte para tempo em segundos
        .Text = "Concluído" 'após processado, atualiza o status da linha corrente
    End With

    Call Processar_Movimento

End Sub

Sub Atulizar_LabelContadores()

    Dim contEntrada, contSaida, contDentroEst As Integer

    With frmPainelControleEstacionamento
    
        For I = 1 To .griMonitoramento.Rows - 1
            .griMonitoramento.Row = I
            .griMonitoramento.Col = 1
                auxMovimento = .griMonitoramento.Text
            .griMonitoramento.Col = 3
                auxStatus = .griMonitoramento.Text
                
            If auxMovimento = "ENTRADA" And auxStatus = "Aguardando" Then contEntrada = contEntrada + 1
            If auxMovimento = "ENTRADA" And auxStatus = "Concluído" Then contDentroEst = contDentroEst + 1
            If auxMovimento = "SAIDA" And auxStatus = "Aguardando" Then contSaida = contSaida + 1
            If auxMovimento = "SAIDA" And auxStatus = "Concluído" Then contDentroEst = contDentroEst - 1
        
        Next I
                
        .lblQtdeFilaEntrada = CStr(contEntrada)
        .lblQtdeFilaSaida = CStr(contSaida)
        .lblQtdeDentroEstacionamento = CStr(contDentroEst)
                
    End With

End Sub
