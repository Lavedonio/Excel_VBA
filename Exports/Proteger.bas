Attribute VB_Name = "Proteger"
Sub Protect_All()
Attribute Protect_All.VB_Description = "Protege todas as planilhas com uma senha padrão"
Attribute Protect_All.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Protect all Sheets and Workbook with password
'

'
    Dim ws As Worksheet
    Dim senha As String
    
    senha = "12345"
    For Each ws In Worksheets
        ws.Protect Password:=senha
    Next
    
    ActiveWorkbook.Protect Password:=senha, structure:=True, Windows:=True
End Sub


Sub Unprotect_All()
Attribute Unprotect_All.VB_Description = "Desprotege todas as planilhas com uma senha padrão"
Attribute Unprotect_All.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' Unprotect all Sheets and Workbook with password
'

'
    Dim ws As Worksheet
    Dim senha As String
    
    Application.ScreenUpdating = False
    
    senha = "12345"
    For Each ws In Worksheets
        ws.Unprotect Password:=senha
    Next
    
    ActiveWorkbook.Unprotect Password:=senha
    
    Application.ScreenUpdating = True
End Sub

