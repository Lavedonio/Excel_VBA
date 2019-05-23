Sub Protect_All()
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
