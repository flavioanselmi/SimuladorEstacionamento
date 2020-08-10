Attribute VB_Name = "ModExibirDados"

Public Sub Carregar_gridMonitoramento(QtdeEntrada, QtdeSaida As Integer)

    Cabecalho_GridMonitoramento
    
    With frmPainelControleEstacionamento.griMonitoramento
        
        For i = 0 To QtdeEntrada - 1

            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            .Col = 0 'ID
                .Text = .Row
                .CellAlignment = 0
            .Col = 1 'Movimento E/S
                .Text = "ENTRADA"
                '.CellAlignment = 2
            .Col = 2 'Tempo (seg)
                .CellAlignment = 3
                .Text = funcIntGerarNumRND(5, 1) 'gera Nun. Int. entre 1 e 5
        Next i
        

        For i = 0 To QtdeSaida - 1

            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            .Col = 0 'ID
                .Text = .Row
                .CellAlignment = 0
            .Col = 1 'Movimento E/S
                .Text = "SAIDA"
                '.CellAlignment = 2
            .Col = 2 'Tempo (seg)
                .CellAlignment = 3
                .Text = funcIntGerarNumRND(30, 10) 'gera Num. Int enttre 10 e 30
                
        Next i
    
    End With

End Sub

Sub Cabecalho_GridMonitoramento()

    With frmPainelControleEstacionamento.griMonitoramento
        .Clear
        .Rows = 1 'QtdeLinhas 'nº de linhas
        .Cols = 4
        .FixedCols = 0
        .Row = 0
        
        'definição da largura de cada coluna
        .ColWidth(0) = 1000 'ID Sequencia Eventos
        .ColWidth(1) = 1500 'tipo Movimento [E] - Entrada  // [S] -Saída
        .ColWidth(2) = 1500  'tempo (seg)
        .ColWidth(3) = 3500  'Status
        
        'definição do nome de cada coluna
        .Col = 0: .CellAlignment = 3: .Text = "ID veículo"
        .Col = 1: .CellAlignment = 3: .Text = "Movimento E/S"
        .Col = 2: .CellAlignment = 3: .Text = "Tempo (Segundos)"
        .Col = 3: .CellAlignment = 3: .Text = "Status"

    End With


End Sub
