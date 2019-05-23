Attribute VB_Name = "Mensagens"
Sub Limpar_Mensagens_Macro()
Attribute Limpar_Mensagens_Macro.VB_Description = "Limpa todas as mensagens das Macros na planilha ""Painel de Controle"""
Attribute Limpar_Mensagens_Macro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpa as Mensagens das Macros na planilha "Painel de Controle"
' Depende de um "t�tulo" na coluna M para refer�ncia
'

'
    Call Unprotect_All
    
    Sheets("Painel de Controle").Select
    Range("M1").End(xlDown).Offset(2, 0).Select
    Range(ActiveCell, ActiveCell.End(xlDown)).Clear
    
    Call Protect_All
End Sub


Sub Mensagens_Macro(mensagem As String, tipo_alerta As String)
'
' Sub que imprime uma mensagem na lista de mensagens da aba "Painel de Controle"
' N�o deleta mensagens j� existentes
' Ao usar a sub, lembrar de reselecionar a aba que estava trabalhando
' Depende de um "t�tulo" na coluna M para refer�ncia
'

'
    Sheets("Painel de Controle").Select
    Range("M1").End(xlDown).End(xlDown).Offset(1, 0).Select
    ActiveCell.Value = mensagem
    
    Select Case LCase(tipo_alerta)
        Case Is = "sucesso"
            With Selection.Font
                .ColorIndex = 10
                .Bold = True
            End With
        Case Is = "alerta"
            With Selection.Font
                .ColorIndex = 45
                .Bold = True
            End With
        Case Is = "erro"
            With Selection.Font
                .ColorIndex = 3
                .Bold = True
                Call MsgBox(mensagem, vbExclamation, "Mensagem da Macro")
            End With
        Case Is = "intro"
            With Selection.Font
                .ColorIndex = 0
                .Bold = False
                .Italic = True
            End With
        Case Else
            With Selection.Font
                .ColorIndex = 0
                .Bold = False
            End With
    End Select
End Sub
