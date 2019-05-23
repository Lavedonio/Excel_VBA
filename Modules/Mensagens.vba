Sub Limpar_Mensagens_Macro()
'
' Limpa as Mensagens das Macros na planilha "Painel de Controle"
' Depende de um "título" na coluna M para referência
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
' Não deleta mensagens já existentes
' Ao usar a sub, lembrar de reselecionar a aba que estava trabalhando
' Depende de um "título" na coluna M para referência
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
