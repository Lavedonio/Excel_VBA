Attribute VB_Name = "Importar"
Sub Importar_Dados()
'
' Importa e trata uma nova base de dados para o relatório
'

'
    Dim folderPath As String
    Dim folderName As String
    Dim fileName As String
    Dim filePath As String
    Dim lista_nome_arquivos As New Collection
    Dim nome_arquivo As Variant
    Dim ha_duplicata As Boolean
    
    'Variáveis para opção Semi-Automática
    Dim data_inicio As Date
    Dim data_fim As Date
    
    Call Limpar_Mensagens_Macro
    Call Unprotect_All
    Application.ScreenUpdating = False
    Call Mensagens_Macro("Rodando Macro Importar_Dados...", "intro")
    
    'Atribui as datas dos controles nas respectivas variáveis
    Sheets("Painel de Controle").Select
    data_inicio = Range("G11").Value
    data_fim = Range("G12").Value
    
    'Vê nas configurações o nome da pasta em que estão os arquivos. Se estão na mesma pasta, deve receber uma string vazia
    folderName = Range("F22").Value

    'Adiciona o nome do arquivo no Path. Se estiver em branco, não será adicionado ou "\" no final
    folderPath = ThisWorkbook.Path & "\" & folderName
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    'Verifica se pasta existe
    If Dir(folderPath, vbDirectory) = "" Then
        Call Mensagens_Macro("Erro: pasta não encontrada!", "erro")
        Call Protect_All
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Abertura automática
    Set lista_nome_arquivos = Listar_Arquivos(folderPath, "*.csv")
    
    ' Abertura semi-automática
    If data_inicio > data_fim Then
        Call Mensagens_Macro("Erro: data final maior que data inicial!", "erro")
        Call Protect_All
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Set lista_nome_arquivos = Listar_Arquivos_Range_Data(folderPath, data_inicio, data_fim, ".csv")
        
    ' Abertura manual
    Set lista_nome_arquivos = Listar_Arquivos_Abertura_Manual()
    If lista_nome_arquivos.Count = 0 Then
        Call Mensagens_Macro("Alerta: seleção de arquivos cancelada.", "alerta")
        Call Protect_All
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call Limpar_Filtros_Base
    
    ' Com a lista de arquivos selecionados, abre e trata os dados
    ' ...
    For Each nome_arquivo In lista_nome_arquivos
        fileName = CStr(nome_arquivo)
        filePath = folderPath & fileName
        ' Com o nome do path e do arquivo, usa uma função para abrir e manipular do jeito que quiser
        'Call Alguma_Funcao(filePath, fileName)
        Call Mensagens_Macro("Arquivo " & fileName & " importado.", "info")
    Next nome_arquivo
    
    Sheets("Painel de Controle").Select
    Call Mensagens_Macro("Importação de arquivos para Base concluída!", "sucesso")
    Call Protect_All
    Application.ScreenUpdating = True
End Sub


Function Listar_Arquivos(folderPath As String, Optional fileExtension As String) As Object
'
' Lista todos os arquivos de uma pasta e os adiciona numa coleção que é então retornada
'

'
    Set Listar_Arquivos = New Collection
    Dim file As String
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    file = Dir(folderPath & fileExtension)
    
    Do While file <> ""
        Listar_Arquivos.Add file
        file = Dir
    Loop
End Function


Function Listar_Arquivos_Range_Data(folderPath As String, data_inicio As Date, data_fim As Date, Optional fileExtension As String) As Object
'
' Lista todos os arquivos de uma pasta entre um range de datas que possuem nome formatado na forma "yyyy-mm-dd"
'

'
    Dim data_investigada As Date
    Dim file As String
    Dim fileDate As String
    
    Set Listar_Arquivos_Range_Data = New Collection
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    For data_investigada = data_inicio To data_fim
        fileDate = CStr(Format(data_investigada, "yyyy-mm-dd;@")) & fileExtension
        file = Dir(folderPath & "*" & fileDate)
        If file <> "" Then
            Listar_Arquivos_Range_Data.Add file
        End If
    Next
End Function


Function Listar_Arquivos_Abertura_Manual() As Object
'
' Abre uma caixa de diálogo e retorna uma coleção com os nomes dos arquivos selecionados
' Docs: https://docs.microsoft.com/en-us/office/vba/api/excel.application.filedialog
'

'
    Dim numSelecionados As Long
    
    Set Listar_Arquivos_Abertura_Manual = New Collection
    
    'Abre caixa de diálogo
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        
        'Adiciona na coleção o nome de cada arquivo selecionado
        For numSelecionados = 1 To .SelectedItems.Count
            Listar_Arquivos_Abertura_Manual.Add Dir(.SelectedItems(numSelecionados))
        Next
    End With
End Function
