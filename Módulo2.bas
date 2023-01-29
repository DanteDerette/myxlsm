Attribute VB_Name = "Módulo2"
Sub LoopThroughFiles()
    Dim StrFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim contador As Integer
    
    contador = 1
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TodosOsOrcamentos")
    
    Do While ws.Cells(contador, 1) <> ""
        ws.Rows(contador).Delete
        contador = contador + 1
    Loop
    
    contador = 1
    
    StrFile = Dir(ThisWorkbook.path & "\OrcamentosDoSistemaDoDante\")
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        ws.Cells(contador, 1).Value = contador
        ws.Cells(contador, 2).Value = StrFile
        ws.Cells(contador, 3).Value = ThisWorkbook.path & "\OrcamentosDoSistemaDoDante\" & StrFile
        StrFile = Dir
        contador = contador + 1
    Loop
End Sub
Sub selecionar()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim estaplanilha As Workbook
    Dim clientes As Workbook
    Set estaplanilha = Workbooks("main.xlsm")
    Set clientes = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
    
    estaplanilha.Worksheets("DB_Produtos").Delete
        
    Dim source_worksheet As Worksheet
    Set source_worksheet = clientes.Worksheets("BD")
    source_worksheet.Name = "DB_Produtos"

    Dim target_worksheet As Worksheet
    Set target_worksheet = estaplanilha.Worksheets("Menu")
    
    source_worksheet.Copy After:=target_worksheet
    clientes.Close
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
        
        
    
    
End Sub

Sub AddNew()
    Workbooks.Add Template:="C:\GitHub\myxlsm\template_orcamento.xlsx"
   
    
End Sub

Public Function IsLoaded(formName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
End Function

Sub hide_menu()

    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"

End Sub

Sub show_menu()

    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"

End Sub
