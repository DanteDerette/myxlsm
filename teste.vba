Sub auxiliarDoDespacho()
    ThisWorkbook.Save
    
    Dim wb As Workbook
    Dim cronogramaDeDespacho As Worksheet
    Dim auxiliarDaNotaFiscal As Worksheet

    Set wb = ThisWorkbook
    Set cronogramaDeDespacho = wb.Sheets("cronogramaDeDespacho")

    Dim selectedColumn As Range
    Set selectedColumn = Selection.Columns(1)

    Dim firstCell As Range
    Set firstCell = selectedColumn.Cells(1)
    
    If firstCell.Value = "" Or firstCell.Value = "Cenário" Or firstCell.Value = "Código" Or firstCell.Value = "Especificação" Or firstCell.Value = "Quantidade" Or firstCell.Value = "Saldo" Then
        Exit Sub
    End If
    
    If ActiveSheet.Name <> "cronogramaDeDespacho" Then
        Exit Sub
    End If
    
    worksheetName = Replace(firstCell.Value, "/", ".")

    Dim ws As Worksheet
    
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = worksheetName Then
            MsgBox ("achou algo")
            worksheetName = Replace(firstCell.Value & "_" & i, "/", ".")
            i = 1
        End If
    Next i
    
    Sheets.Add.Name = worksheetName

    set auxiliarDaNotaFiscal = wb.Sheets(worksheetName)

    auxiliarDaNotaFiscal.range("A:C").value = cronogramaDeDespacho.range("A:C").value
    auxiliarDaNotaFiscal.range("D").value = cronogramaDeDespacho.range("F").value

    
End Sub
