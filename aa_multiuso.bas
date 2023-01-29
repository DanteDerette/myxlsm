Attribute VB_Name = "aa_multiuso"

Sub PreencheLista_Melhorado(TextoDigitado, bd, formAtual)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wsSheet As Worksheet
    Dim rnData As Range
    Dim vaData As Variant
    Dim lastrow As Integer
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\"& bd, ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets("BD")
    Set parafilter = wbBook.Worksheets("paraFilter")
    
    wsSheet.Activate
    
    If formAtual.ComboBoxCampos.ListIndex = 0 Then
        wsSheet.Range("A1").AutoFilter Field:=1, Criteria1:=(TextoDigitado)
    Else
        wsSheet.Range("A1").AutoFilter Field:=(formAtual.ComboBoxCampos.ListIndex + 1), Criteria1:="*" + TextoDigitado + "*"
    End If
    
    lastrow = wsSheet.Range("A" & wsSheet.Rows.Count).End(xlUp).Row
    Set rnData = wsSheet.Range(wsSheet.Range("A1"), wsSheet.Range("AC" & lastrow))
    
    rnData.Select
    Selection.Copy
    
    parafilter.Activate
    parafilter.Range("A1").Select
    parafilter.Paste
    
    vaData = Selection.Value
        
    With formAtual.lstLista
        .ColumnCount = 29
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
   wbBook.Close False
   
   With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub

sub init_lstLista(bd, formAtual)
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnData As Range
    Dim vaData As Variant
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\" & bd, ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets("BD")
    
    lastrow = wsSheet.Range("A" & wsSheet.Rows.Count).End(xlUp).Row
    Set rnData = wsSheet.Range(wsSheet.Range("A1"), wsSheet.Range("AC" & lastrow))
    
    vaData = rnData.Value
    
    With formAtual.lstLista
        .ColumnCount = 29
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
    wbBook.Close
end sub