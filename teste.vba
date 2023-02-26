sub auxiliarDoDespacho()
    dim wb as workbook
    dim cronogramaDeDespacho as worksheet
    dim auxiliarDaNotaFiscal as worksheet

    set wb = thisworbook
    set cronogramaDeDespacho = wb.Sheets("cronogramaDeDespacho")

    Dim selectedColumn As Range
    Set selectedColumn = Selection.Columns(1)

    Dim firstCell As Range
    Set firstCell = selectedColumn.Cells(1)

    wb.sheets.add = firstCell.value
    




end sub