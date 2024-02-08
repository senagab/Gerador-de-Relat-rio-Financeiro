Sub Exportar()

    Dim linhafim As Integer, linhafim2 As Integer
    Dim workbook1 As Workbook, workbook2 As Workbook
    Dim name As String, name2 As String
    Dim oldname As String, newname As String, lastcell As String
    Dim titulofinal As String, titulo As String, titulo2 As String, Deletado As String, erro As String
    Dim voltacaminho As String, destcaminho As String, killcaminho As String
    Dim frompath As String, newnamepost As String, newnamepdfpost As String
    Dim autoapagavel As String
    Dim FSO As Object, killing As Object
    
    killcaminho = ThisWorkbook.Path
    
    name = "Relatório Financeiro - Até "
    
    If Not Dir(killcaminho & "\*Relatório*", vbDirectory) = vbNullString Then
        
        Kill killcaminho & "\*Relatório*"
        
    Else
    
        'MsgBox "Nenhum arquivo 'Relatório' encontrado na pasta"
        
    End If
    
    'Tira os alertas nativos do Excel
    Application.DisplayAlerts = False
    
    Application.ScreenUpdating = False
    
    'Seleciona um range mais escondido
    Range("V30").Select
    
    On Error GoTo erro

    'Abre a planilha com os dados recebidos do Omie, seleciona e copia
    Workbooks.Open Filename:=ActiveWorkbook.Path & "\dados.xlsx"
    
'erro1004:
    
'    MsgBox "Não há um arquivo 'Dados' na pasta. Certifique-se de que os arquivos necessários estejam próximos.", vbOKOnly, "Arquivo faltando!"
'    Exit Sub
    
    'Encontra quantas linhas totais subtraídas por 5 Atribui o valor à variável "linhafim"
    linhafim = Range("A2").End(xlDown).Row - 8

    'Seleciona o range que será copiado (no caso tudo)
    Range("A1:" & "G" & linhafim).Select

    'Copia o conteúdo
    Selection.Copy

    'Criação do arquivo
    Set workbook1 = Application.Workbooks.Add

    'Começando da célula "A1"
    Range("A1").Select

    'Cola o conteúdo copiado
    ActiveSheet.Paste

    'Cria as colunas de "Tipo" e "Valor" e as formata
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Valor"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Tipo"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]>0,RC[-2],-RC[-1])"
    Range("H2").Select
    Selection.Style = "Comma"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,""Entrada"",""Saída"")"
    Range("H2:I2").Select
    'Autofill: autocompleta as linhas da coluna, seguindo o padrão das linhas acima
    Selection.AutoFill Destination:=Range("H2:I" & linhafim)

    'Formatação da coluna "Valor"
    Range("H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0.00"

    'Seleciona as células e arruma o tamanho das colunas e linhas
    Cells.Select
    Selection.ColumnWidth = 31.86
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

    'Oculta as colunas de "Entrada" e "Saída"
    Columns("F:G").Hidden = True

    'Verifica se o filtro está ligado e se não estiver ele é ligado
    If ActiveSheet.AutoFilterMode = False Then

        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.AutoFilter
        'MsgBox "Filtros ligados.", vbApplicationModal, "Ativação de filtros."

    End If

    'Congela o topo
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Negrito
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True

    'Gets last cell
    lastcell = Format(Range("A1").End(xlDown).Value, "dd-mm-yyyy")

    'Closes the "database" window
    Windows("dados.xlsx").Activate
    Application.CutCopyMode = False
    ActiveWindow.Close

    'Variável recebe novo valor: numero total de células
    linhafim = Range("A1").End(xlDown).Row

    'Cria tabela dinâmica e linha do tempo
    Sheets.Add After:=ActiveSheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Planilha1!R1C1:R" & linhafim & "C9", Version:=6).CreatePivotTable TableDestination:= _
        "Planilha2!R1C1", TableName:="Tabela dinâmica1", DefaultVersion:=6
    Sheets("Planilha2").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("Tabela dinâmica1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Valor"), "Soma de Valor", xlSum
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Tipo")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Categoria")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Cliente/Fornecedor")
        .Orientation = xlRowField
        .Position = 3
    End With
    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 70#
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("Tabela dinâmica1"), _
        "Data", , xlTimeline).Slicers.Add ActiveSheet, , "Data", "Data", 173.25, 269.25 _
        , 262.5, 108
    ActiveSheet.Shapes.Range(Array("Data")).Select
    ActiveSheet.Shapes("Data").IncrementLeft 264.75
    ActiveSheet.Shapes("Data").IncrementTop -143.25
    ActiveSheet.Shapes("Data").ScaleWidth 1.0914285714, msoFalse, _
        msoScaleFromTopLeft

    'Forçaç formatação como número
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0.00"

    'Atualiza a tabela dinâmica
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache.Refresh

    'Renomeia as abas da planilha
    Sheets("Planilha1").Select
    Sheets("Planilha1").name = "Lançamentos"
    Sheets("Planilha2").Select
    Sheets("Planilha2").name = "Tabela Dinâmica"

    'Salva o arquivo criado - atualmente chamado "Relatório Financeiro - Até .xlsx"
    workbook1.SaveAs ThisWorkbook.Path & "\" & name & ".xlsx"

    'Alt Tab para o arquivo criado
    Windows(name & ".xlsx").Activate
    'Fecha o arquivo
    ActiveWindow.Close
    
    'Variável recebe o nome antigo
    oldname = ThisWorkbook.Path & "\" & name & ".xlsx"
    
    'Variável recebe o nome atualizado
    newname = ThisWorkbook.Path & "\" & name & lastcell & ".xlsx"
    newnamepost = name & lastcell & ".xlsx"
    
    'Variável recebe o nome atualizado em pdf
    newnamepdf = ThisWorkbook.Path & "\" & name & lastcell & ".pdf"
    newnamepdfpost = name & lastcell & ".pdf"
    'Variável recebe nome atualizado sem possuir formato Excel (- ".xlsx")
    titulo = ThisWorkbook.Path & "\" & name & lastcell
    
    On Error Resume Next
    
    'Renomeia o arquivo, adicionando a data mais recente da planilha ao nome
    Name oldname As newname
    
erro:
'    On Error Resume Next
    Select Case Err.Number
    
        Case 58
        
'            MsgBox "Já existe um arquivo com o mesmo nome. Apague-o antes de continuar.", vbOKOnly, "Arquivo Existente"
            Call OverwriteOption
            
        Case 1004
                    
            MsgBox "Não há um arquivo 'Dados' na pasta. Certifique-se de que os arquivos necessários estejam próximos.", vbOKOnly, "Arquivo Faltando"
            Exit Sub
            
        Case Else
        
            'faz nada
    
    End Select
    
    
    'Reabre o arquivo de novo
    Workbooks.Open Filename:=newname
        
    
    'Ativa a planilha para usá-la de referência
    Sheets("Lançamentos").Activate
    'Variável nômade - recebe um valor atualizado: numero total de células
    linhafim = Range("A1").End(xlDown).Row
    'Guarda a seleção de range da planilha "Lançamentos"
    
    'Ativa a planilha para usá-la de referência
    Sheets("Tabela Dinâmica").Activate
    'Variável nômade - recebe um valor atualizado: numero total de células
    linhafim2 = Range("A1").End(xlDown).Row
    'Guarda a seleção de range da planilha "Tabela Dinâmica"
    
    
    'Setup de impressão da Aba "Lançamentos"
    With Worksheets("Tabela Dinâmica").PageSetup
        .Orientation = xlPortrait
        .Zoom = 100
        .PrintArea = "A1:B" & linhafim2
    End With

    'Setup de impressão da Aba "Lançamentos"
    With Worksheets("Lançamentos").PageSetup
        .Orientation = xlLandscape
        .Zoom = 90
        .PrintGridlines = True
        .PrintArea = "A1:I" & linhafim
    End With
    
    
    'Seleciona ambas tabelas e as coloca em um array
    Sheets(Array("Tabela Dinâmica", "Lançamentos")).Select
    
    'Lembra da ultima seleção feita e exporta com formato PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        titulo, Quality:=xqualitystandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
        
    'Seleciona a aba da tela dinâmica para desagrupar as abas agrupadas durante a criação do PDF
    Sheets("Tabela Dinâmica").Select
    'Voltar a atualizar tela
    Application.ScreenUpdating = True
    'Salva as alterações feitas no arquivo
    ActiveWorkbook.Save
    'Fecha o arquivo após salvar
    ActiveWindow.Close
    
    'Volta uma pasta
    voltacaminho = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    'Insere novo caminho (concatenando)
    destcaminho = voltacaminho & "\Relatórios Acumulados\"
    
    'Instância o Objeto que vai receber o arquivo a ser copiado
    Set FSO = CreateObject("scripting.filesystemobject")
    
    'Objeto recebe arquivo a ser copiado e faz a transferência
    'Lê-se
    FSO.copyfile newname, destcaminho & newnamepost, True
    
    'Objeto recebe arquivo a ser copiado e faz a transferência
    FSO.copyfile newnamepdf, destcaminho & newnamepdfpost, True
    
    
    'Apaga todos os relatórios
'    Kill newnamepdf
    
'    Workbooks("Novo Relatorio 1.3.xlsm").Close
    ActiveWorkbook.Saved = True
    Kill ThisWorkbook.Path & "\" & name & ".xlsx"
    MsgBox "Relatório Financeiro criado com sucesso!", vbYes, "Criação de arquivo"
'    Application.Quit



End Sub

Sub OverwriteOption()
    
    If Filename = False Then
    
        Exit Sub
        
    End If
    
    If Dir(Filename) = "" Then
    
        ActiveWorkbook.SaveAs Filename
        
    Else
    
        mensagem = MsgBox("este arquivo já existe amigão. voce quer apagar?", vbYesNoCancel)
        
        If mensagem = vbYes Then
        
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs Filename
            Application.DisplayAlerts = True
            
        ElseIf mensagem = vbNo Then
        
'            GoTo filedialog
            Exit Sub
            
        Else
        
            Exit Sub
            
        End If
        
    End If

End Sub
