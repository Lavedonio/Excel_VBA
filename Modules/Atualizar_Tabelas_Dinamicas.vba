Sub RefreshPivotTables(source_sheet_name As String, Optional start_cell_position As String)
'
' Atualiza todas as Fontes de Dados das tabelas dinâmicas do arquivo e atualiza-as no final.
' O atributo opcional indica onde começa a tabela de fonte de dados.
'

'
    Dim ws As Worksheet
    Dim PT As PivotTable
    Dim SourceName As String
    Dim end_cell_position As String

    If start_cell_position = "" Then start_cell_position = "A1"

    end_cell_position = Worksheets("Base").Range(start_cell_position).End(xlDown).End(xlToRight).Address(ReferenceStyle:=xlR1C1)

    start_cell_position = Range(start_cell_position).Address(ReferenceStyle:=xlR1C1)

    SourceName = ThisWorkbook.Path & "\[" & ThisWorkbook.Name & "]" & _
        source_sheet_name & "!" & start_cell_position & ":" & end_cell_position

    'Atualiza as fontes de dados de todas as tabelas dinâmicas em todas as planilhas
    For Each ws In ActiveWorkbook.Worksheets
        For Each PT In ws.PivotTables
            PT.SourceData = SourceName
        Next PT
    Next ws

    ActiveWorkbook.RefreshAll
End Sub
