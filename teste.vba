Sub impressaoDeOrcamento_F()
    impressaoDeOrcamento.Show
End Sub

Sub esconder()
    Rows("1:5").Select
    Range("A5").Activate
    Selection.EntireRow.Hidden = True
    Application.DisplayFullScreen = True
    Range("A8").Select
End Sub
Sub mostrar()
    Rows("1:5").Select
    Range("A5").Activate
    Selection.EntireRow.Hidden = False
    Application.DisplayFullScreen = False
    Range("A8").Select
End Sub

Sub abreUserForm1()
    UserForm1.Show
End Sub
Sub venda()
    ThisWorkbook.Save
    Call verificaOndeEstaOCenario
    Call geraOrcamento("Venda")
    
    
    Call GetNonRepetitiveValues("Venda")
End Sub
Sub locacao()
    ThisWorkbook.Save
    Call verificaOndeEstaOCenario
    Call geraOrcamento("Locação")
    Call GetNonRepetitiveValues("Locação")
End Sub
Sub verificaOndeEstaOCenario()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsCenarios As Worksheet
    Dim lastrow As Integer
    Dim contador As Integer
    Dim contador2 As Integer
    Dim qualLinha As Integer
    Dim qualCenario As String
        
    Set wb = ThisWorkbook
    
    Set ws = wb.Worksheets("ORÇAMENTO")
    
    Sheets.Add.Name = "cenarios"

    Set wsCenarios = wb.Worksheets("cenarios")
    
    qualLinha = 1
    
    wsCenarios.Range("A2:XFD99999").Clear
        
    wsCenarios.Cells(qualLinha, 1).Value = "Cenário"
    wsCenarios.Cells(qualLinha, 2).Value = "Código"
    wsCenarios.Cells(qualLinha, 3).Value = "Quantidade"
    wsCenarios.Cells(qualLinha, 4).Value = "Obs"
    wsCenarios.Cells(qualLinha, 5).Value = "Especificação"
    wsCenarios.Cells(qualLinha, 6).Value = "Venda Unitário"
    wsCenarios.Cells(qualLinha, 7).Value = "Venda Total"
    wsCenarios.Cells(qualLinha, 8).Value = "Locação Unitário"
    wsCenarios.Cells(qualLinha, 9).Value = "Locação Total"
    
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    contador = 1
    contador2 = 0
    qualLinha = qualLinha + 1
        
    For i = 9 To lastrow
        If ws.Cells(i, 2).Value = "" And ws.Cells(i, 5).Value <> "" Then
            contador2 = i
            qualCenario = ws.Cells(contador2, 5).Value
            contador2 = contador2 + 1
            
            Do While ws.Cells(contador2, 5).Value <> ""
                
                With wsCenarios.Cells(qualLinha, 1)
                    .Value = qualCenario
                    .Interior.Color = RGB(255, 255, 0)
                End With
                
                wsCenarios.Cells(qualLinha, 2).Value = ws.Cells(contador2, 2).Value
                wsCenarios.Cells(qualLinha, 3).Value = ws.Cells(contador2, 3).Value
                wsCenarios.Cells(qualLinha, 4).Value = ws.Cells(contador2, 4).Value
                wsCenarios.Cells(qualLinha, 5).Value = ws.Cells(contador2, 5).Value
                wsCenarios.Cells(qualLinha, 6).Value = ws.Cells(contador2, 6).Value
                wsCenarios.Cells(qualLinha, 7).Value = ws.Cells(contador2, 7).Value
                wsCenarios.Cells(qualLinha, 8).Value = ws.Cells(contador2, 8).Value
                wsCenarios.Cells(qualLinha, 9).Value = ws.Cells(contador2, 9).Value

                qualLinha = qualLinha + 1
                contador2 = contador2 + 1
            Loop
            contador = contador + 1
        End If
    Next i
    
    
    
    
        
End Sub

Sub geraOrcamento(operacao As String)
    Dim wb As Workbook
    Dim orcamentoGerado As Worksheet
    Dim cenarios As Worksheet
    Dim inventario As Worksheet
    Dim lastRowCenarios As Integer
    Dim mes As Integer
    
    Dim qualLinha As Integer
    Dim talFormula As String
    Dim orcamento As Worksheet
    Dim rowParaVar As Integer
        
    Dim searchRange As Range
    Dim foundCell As Range
        
    Set wb = ThisWorkbook
      
    Sheets.Add.Name = "OrcamentoGerado"
    
    Set orcamentoGerado = wb.Worksheets("OrcamentoGerado")
    
    Set cenarios = wb.Sheets("cenarios")
    
    Set inventario = wb.Sheets("INVENTARIO")
    Set orcamento = wb.Worksheets("ORÇAMENTO")
    
    orcamentoGerado.Range("A2:XFD99999").Clear
    
    lastRowCenarios = cenarios.Range("A" & cenarios.Rows.Count).End(xlUp).Row
    
    qualLinha = 1
    
    orcamentoGerado.Cells(qualLinha, 1).Value = "Ref"
    orcamentoGerado.Cells(qualLinha, 2).Value = "Especificação"
    
    orcamentoGerado.Cells(qualLinha, 3).Value = "Alt"
    orcamentoGerado.Cells(qualLinha, 4).Value = "Larg"
    orcamentoGerado.Cells(qualLinha, 5).Value = "Comp"
    
    orcamentoGerado.Cells(qualLinha, 6).Value = "Qtd"
    orcamentoGerado.Cells(qualLinha, 7).Value = "R$ Unit "
    orcamentoGerado.Cells(qualLinha, 8).Value = "R$ Total"
    orcamentoGerado.Cells(qualLinha, 9).Value = "Cenário"

    For i = 2 To lastRowCenarios
        orcamentoGerado.Cells(i, 1).Value = cenarios.Cells(i, 2).Value
        orcamentoGerado.Cells(i, 2).Value = cenarios.Cells(i, 5).Value
        
        Set searchRange = inventario.Range("E1:E" & inventario.Range("E" & Rows.Count).End(xlUp).Row)
        Set foundCell = searchRange.Find(cenarios.Range("B" & i).Value, LookIn:=xlValues, LookAt:=xlPart)
        
        orcamentoGerado.Cells(i, 3).Value = inventario.Cells(foundCell.Row, 10).Value
        orcamentoGerado.Cells(i, 4).Value = inventario.Cells(foundCell.Row, 11).Value
        orcamentoGerado.Cells(i, 5).Value = inventario.Cells(foundCell.Row, 12).Value
        
        orcamentoGerado.Cells(i, 6).Value = cenarios.Cells(i, 3).Value

        If operacao = "Venda" Then
            orcamentoGerado.Cells(i, 7).Value = cenarios.Cells(i, 7).Value
        ElseIf operacao = "Locação" Then
            orcamentoGerado.Cells(i, 7).Value = cenarios.Cells(i, 9).Value
        End If
        
        orcamentoGerado.Cells(i, 8).Formula = "=G" & i & "*F" & i
        orcamentoGerado.Cells(i, 9).Value = cenarios.Cells(i, 1).Value
        
    Next i
    
    orcamentoGerado.Range("G:H").Style = "Currency"

End Sub

Sub GetNonRepetitiveValues(operacao)
    Dim wb As Workbook
    Dim quaisCenarios As Worksheet
    Dim cenarios As Worksheet
    Dim orcamentoInicial As Worksheet
    
    Dim lastRowCenarios As Integer
    Dim lastRowQuaisCenarios As Integer
    Dim valorParaVerificar As String
    Dim valorParaComparar As String
    Dim escrever As Boolean
    Dim qualCenarioEstouMexer As String
    Dim descobreOComecoDosProdutos As Integer
    Dim stringQueGeraAFormulaDeSomarTudo As String
    
    stringQueGeraAFormulaDeSomarTudo = ""
     
    Set wb = ThisWorkbook
    
    Set quaisCenarios = wb.Worksheets("IMPRESSÃO DE ORÇAMENTO 2")
    Set cenarios = wb.Worksheets("orcamentoGerado")
    Set orcamentoInicial = ThisWorkbook.Worksheets("ORÇAMENTO")
    
    quaisCenarios.Range("A1:XFD99999").Clear
    
    quaisCenarios.Range("A2").Value = "ORÇAMENTO"
    With quaisCenarios.Range("A2:H2")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Font.Underline = True
    End With
        
    quaisCenarios.Range("A3").Value = " "
    quaisCenarios.Range("A4").Formula = "=today()"
    
    With quaisCenarios.Range("A4:E4")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
        .NumberFormat = """Joinville,"" d ""de"" mmmm"" de"" yyyy"
    End With
    
        
    quaisCenarios.Range("A5").Value = " "
    quaisCenarios.Range("A6").Value = " "
    quaisCenarios.Range("A7").Value = " "
    quaisCenarios.Range("A8").Value = "Cliente:"
    quaisCenarios.Range("A9").Value = "Cidade:"
    quaisCenarios.Range("A10").Value = "Telefone:"
    quaisCenarios.Range("A11").Value = "Contato:"
    quaisCenarios.Range("A12").Value = " "
    
    If operacao = "Venda" Then
        quaisCenarios.Range("A13").Value = "Pela presente, apresentamos a proposta para Venda de decoração de Páscoa conforme descrição abaixo."
    ElseIf operacao = "Locação" Then
        quaisCenarios.Range("A13").Value = "Pela presente, apresentamos a proposta para LOCAÇÃO de decoração de Páscoa conforme descrição abaixo."
    End If
        
    With quaisCenarios.Range("A13:H13")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
    End With
    
    quaisCenarios.Range("A14").Value = " "
    
    lastRowCenarios = cenarios.Range("A" & cenarios.Rows.Count).End(xlUp).Row
    lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
    
    For x = 1 To lastRowCenarios
        escrever = True
        lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
        
        For i = 0 To lastRowQuaisCenarios
            If cenarios.Cells(x + 1, 9).Value = quaisCenarios.Cells(i + 1, 1).Value Then
                escrever = False
            End If
        Next i
        
        If escrever Then
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 1, 1).Value = cenarios.Cells(x + 1, 9).Value
            
            qualCenarioEstouMexer = quaisCenarios.Cells(lastRowQuaisCenarios + 1, 1).Value
            
            With quaisCenarios.Range("A" & (lastRowQuaisCenarios + 1) & ":H" & (lastRowQuaisCenarios + 1))
                .Merge
                .HorizontalAlignment = xlCenter
                .Interior.Color = RGB(255, 255, 0)
                .Borders.LineStyle = xlContinuous
                .Font.Bold = True
            End With
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 1).Value = "Ref."
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 1).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 2).Value = "Espeficicação"
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 2).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 3).Value = "Alt"
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 3).Borders.LineStyle = xlContinuous
                        
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 4).Value = "Larg"
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 4).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 5).Value = "Comp"
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 5).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 6).Value = "Qtd."
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 6).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 7).Value = "R$ Unit."
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 7).Borders.LineStyle = xlContinuous
            
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 8).Value = "R$ Total"
            quaisCenarios.Cells(lastRowQuaisCenarios + 2, 8).Borders.LineStyle = xlContinuous
            
            
            With quaisCenarios.Range("A" & (lastRowQuaisCenarios + 2) & ":H" & (lastRowQuaisCenarios + 2))
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
            
            Dim myRange As Range
            Dim myRange2 As Range
            Dim lastRowParaUsarNoFilter As Integer
            Dim filterValue As String
            
            
            Set myRange = cenarios.Range("A1:I99999")
            
            myRange.AutoFilter
            
            filterValue = qualCenarioEstouMexer
            
            myRange.AutoFilter Field:=9, Criteria1:=Array(filterValue), Operator:=xlFilterValues
            
            lastRowCenarios = cenarios.Range("A" & cenarios.Rows.Count).End(xlUp).Row
            cenarios.Activate
            cenarios.Range("A2:" & "H" & lastRowCenarios).Select
            
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            
            Selection.Copy
            
            quaisCenarios.Activate
            
            lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
            
            descobreOComecoDosProdutos = lastRowQuaisCenarios + 1
            
            Range("A" & descobreOComecoDosProdutos).Activate
            quaisCenarios.Paste
            
            cenarios.Activate
            myRange.AutoFilter
            quaisCenarios.Activate
            
            lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row + 1
            
            quaisCenarios.Range("G" & lastRowQuaisCenarios).Value = "SubTotal:"
            quaisCenarios.Range("G" & lastRowQuaisCenarios).Borders.LineStyle = xlContinuous
            quaisCenarios.Range("G" & lastRowQuaisCenarios).Interior.Color = RGB(255, 255, 0)
            quaisCenarios.Range("G" & lastRowQuaisCenarios).HorizontalAlignment = xlRight
            quaisCenarios.Range("G" & lastRowQuaisCenarios).Font.Bold = True
            
            quaisCenarios.Range("H" & lastRowQuaisCenarios).Formula = "=SUM(H" & descobreOComecoDosProdutos & ":H" & (lastRowQuaisCenarios - 1) & ")"
            quaisCenarios.Range("H" & lastRowQuaisCenarios).Borders.LineStyle = xlContinuous
            quaisCenarios.Range("H" & lastRowQuaisCenarios).Interior.Color = RGB(255, 255, 0)
            quaisCenarios.Range("H" & lastRowQuaisCenarios).Font.Bold = True
            stringQueGeraAFormulaDeSomarTudo = "H" & lastRowQuaisCenarios & "+" & stringQueGeraAFormulaDeSomarTudo

            quaisCenarios.Range("A" & lastRowQuaisCenarios).Value = " "
            
            quaisCenarios.Range("A" & (lastRowQuaisCenarios + 1)).Value = " "
                                    
        End If
    Next x
    
    Application.DisplayAlerts = False
    'Sheets("orcamentoGerado").Delete
    'Sheets("cenarios").Delete
    Application.DisplayAlerts = True
    
    quaisCenarios.Activate
    quaisCenarios.Columns("A:H").AutoFit
    Range("A1").Select
    quaisCenarios.Range("B8").Value = ThisWorkbook.Worksheets("Orçamento").Range("C1").Value
    
    stringQueGeraAFormulaDeSomarTudo = Left(stringQueGeraAFormulaDeSomarTudo, Len(stringQueGeraAFormulaDeSomarTudo) - 1)
    stringQueGeraAFormulaDeSomarTudo = "=" & stringQueGeraAFormulaDeSomarTudo
    
    lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
    
    Range("G" & lastRowQuaisCenarios + 2).Value = "Total:"
    Range("H" & lastRowQuaisCenarios + 2).Formula = stringQueGeraAFormulaDeSomarTudo
    
    Range("A" & lastRowQuaisCenarios + 2).Value = " "
    
    With quaisCenarios.Range("G" & (lastRowQuaisCenarios + 2))
        .HorizontalAlignment = xlRight
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With
    
    With quaisCenarios.Range("H" & (lastRowQuaisCenarios + 2))
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With
    
    lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
    
    If orcamentoInicial.Range("F5").Value = "SIM" Then
        Range("A" & lastRowQuaisCenarios + 1).Value = "Frete e instalação incluso no orçamento."
    Else
        Range("A" & lastRowQuaisCenarios + 1).Value = "Frete e instalação não incluso no orçamento."
    End If
    
    Range("A" & lastRowQuaisCenarios + 1).Font.Color = RGB(255, 0, 0)
    Range("A" & lastRowQuaisCenarios + 1).Font.Bold = True
    
    Range("A" & lastRowQuaisCenarios + 3).Value = "Condições de Pagamento: A vista"
    Range("A" & lastRowQuaisCenarios + 3).Font.Bold = True
    
    Range("A" & lastRowQuaisCenarios + 5).Value = "Validade da Proposta: 30 dias"
    Range("A" & lastRowQuaisCenarios + 5).Font.Bold = True
     
    Range("A" & lastRowQuaisCenarios + 7).Value = "Data de entrega: a combinar"
    Range("A" & lastRowQuaisCenarios + 7).Font.Bold = True
    
    Range("A" & lastRowQuaisCenarios + 9).Value = "Atenciosamente,"
    Range("A" & lastRowQuaisCenarios + 9).Font.Bold = True
    
    Range("A" & lastRowQuaisCenarios + 10).Value = orcamentoInicial.Range("B2").Value
    Range("A" & lastRowQuaisCenarios + 10).Font.Bold = True
        
    lastRowQuaisCenarios = quaisCenarios.Range("A" & quaisCenarios.Rows.Count).End(xlUp).Row
    
    Range("A" & lastRowQuaisCenarios + 1).Value = "Luz e Forma Comércio e Decorações Ltda"
    Range("A" & lastRowQuaisCenarios + 2).Value = "CNPJ 02.742.361/0002-10"
    
    Range("A" & lastRowQuaisCenarios + 3).Value = "www.luzeforma.com.br"
    
    Range("A" & lastRowQuaisCenarios + 3).Font.Color = RGB(0, 0, 255)
    Range("A" & lastRowQuaisCenarios + 3).Font.Underline = xlUnderlineStyleSingle
    
    Range("A" & lastRowQuaisCenarios + 1 & ":A" & lastRowQuaisCenarios + 4).Font.Bold = True
    
    'OBSERVAÇÕES
    
    
    quaisCenarios.Columns("A:H").AutoFit
    Range("A1").Select
    
    Call CountRepetitiveValues
        
    
    
End Sub

Sub CountRepetitiveValues()
    Dim dataSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim pedidoDeVenda2 As Worksheet
    Dim lastrow As Long
    Dim valueCount As Variant
    Dim i As Long
    Dim j As Long
    Dim sumRange As Range
    Dim sequencia As Integer
    Dim orcamentoInicial As Worksheet
    
            
    Dim searchRange As Range
    Dim foundCell As Range
    
    Set pedidoDeVenda2 = ThisWorkbook.Worksheets("Pedido de Venda 2")
    Set dataSheet = ThisWorkbook.Worksheets("OrcamentoGerado")
    Set orcamentoInicial = ThisWorkbook.Worksheets("ORÇAMENTO")
    Set resultSheet = ThisWorkbook.Worksheets.Add
    
    pedidoDeVenda2.Range("A31").Value = "Séq."
    pedidoDeVenda2.Range("B31").Value = "Código"
    pedidoDeVenda2.Range("C31").Value = "NCM"
    pedidoDeVenda2.Range("D31").Value = "Especificação"
    pedidoDeVenda2.Range("E31").Value = "ALT"
    pedidoDeVenda2.Range("F31").Value = "LARG"
    pedidoDeVenda2.Range("G31").Value = "COMP"
    pedidoDeVenda2.Range("H31").Value = "QTD"
    pedidoDeVenda2.Range("I31").Value = "UNIT"
    pedidoDeVenda2.Range("J31").Value = "TOTAL"
    
    pedidoDeVenda2.Range("A32:J9999").Value = ""
    pedidoDeVenda2.Range("A32:J9999").Clear
    
    
    lastrow = dataSheet.Cells(Rows.Count, "A").End(xlUp).Row 'get the last row in column A
    
    For i = 1 To lastrow
        valueCount = Application.Match(dataSheet.Cells(i, "A"), resultSheet.Columns(1), 0)
        If IsError(valueCount) Then
            resultSheet.Cells(resultSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1, "A").Value = dataSheet.Cells(i, "A").Value
        End If
    Next i
    
    lastrow = resultSheet.Cells(Rows.Count, "A").End(xlUp).Row 'get the last row in column A
    
    resultSheet.Activate
    resultSheet.Range("A3:" & "A" & lastrow).Select
    Selection.Copy
    pedidoDeVenda2.Activate
    pedidoDeVenda2.Range("B32").Activate
    pedidoDeVenda2.Paste
    
    sequencia = 1
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 32 To lastrow
        pedidoDeVenda2.Cells(i, 1).Value = sequencia
        sequencia = sequencia + 1
    Next i
    
    With pedidoDeVenda2.Range("A31:J" & lastrow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    resultSheet.Activate
    resultSheet.Range("B3:" & "B" & lastrow).Select
    Selection.Copy
    pedidoDeVenda2.Activate
    pedidoDeVenda2.Range("H32").Activate
    
    pedidoDeVenda2.Paste
     
    Set sumRange = dataSheet.Range("F1:F99999")
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 32 To lastrow
        Set searchRange = dataSheet.Range("A1:A" & dataSheet.Range("A" & Rows.Count).End(xlUp).Row)
        Set foundCell = searchRange.Find(pedidoDeVenda2.Cells(i, 2).Value, LookIn:=xlValues, LookAt:=xlPart)
        
        pedidoDeVenda2.Cells(i, 4).Value = dataSheet.Cells(foundCell.Row, 2).Value
        pedidoDeVenda2.Cells(i, 4).Borders.LineStyle = xlContinuous
        pedidoDeVenda2.Cells(i, 4).WrapText = True
        
        pedidoDeVenda2.Cells(i, 5).Value = dataSheet.Cells(foundCell.Row, 3).Value
        pedidoDeVenda2.Cells(i, 6).Value = dataSheet.Cells(foundCell.Row, 4).Value
        pedidoDeVenda2.Cells(i, 7).Value = dataSheet.Cells(foundCell.Row, 5).Value
        
        With pedidoDeVenda2.Range("E" & i & ":G" & i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        pedidoDeVenda2.Cells(i, 9).Value = dataSheet.Cells(foundCell.Row, 7).Value
        pedidoDeVenda2.Cells(i, 10).Formula = "=H" & i & "*I" & i
        pedidoDeVenda2.Cells(i, 10).Locked = True
        
        With pedidoDeVenda2.Cells(i, 8)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        sumValue = Application.WorksheetFunction.SumIf(dataSheet.Range("A1:A99999"), pedidoDeVenda2.Cells(i, 2).Value, dataSheet.Range("F1:F99999"))
            
        pedidoDeVenda2.Cells(i, 8).Formula = sumValue
        
    Next i
    
    pedidoDeVenda2.Range("I:J").Style = "Currency"
    
    pedidoDeVenda2.Range("A31").Activate
    
    pedidoDeVenda2.Range("A31:J999").EntireRow.AutoFit
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "B").End(xlUp).Row
    
    pedidoDeVenda2.Range("A" & lastrow + 1).Value = " "
    pedidoDeVenda2.Range("I" & lastrow + 1).Value = "Total"
    pedidoDeVenda2.Range("J" & lastrow + 1).Formula = "=SUM(J32" & ":J" & (lastrow) & ")"
    pedidoDeVenda2.Range("J" & lastrow + 1).Locked = True
    
    
    With pedidoDeVenda2.Range("I" & (lastrow + 1) & ":J" & lastrow + 1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "A").End(xlUp).Row + 2
    
    pedidoDeVenda2.Range("A" & lastrow + 1).Value = "Frete:"
    pedidoDeVenda2.Range("B" & lastrow + 1 & ":J" & lastrow + 1).Merge
    pedidoDeVenda2.Range("B" & lastrow + 1 & ":J" & lastrow + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous

    pedidoDeVenda2.Range("A" & lastrow + 2).Value = "Expedição:"
    pedidoDeVenda2.Range("B" & lastrow + 2 & ":J" & lastrow + 2).Merge
    pedidoDeVenda2.Range("B" & lastrow + 2 & ":J" & lastrow + 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    pedidoDeVenda2.Range("A" & lastrow + 3).Value = "Entrega:"
    pedidoDeVenda2.Range("B" & lastrow + 3 & ":J" & lastrow + 3).Merge
    pedidoDeVenda2.Range("B" & lastrow + 3 & ":J" & lastrow + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    pedidoDeVenda2.Range("A" & lastrow + 4).Value = "End. Entrega:"
    pedidoDeVenda2.Range("B" & lastrow + 4 & ":J" & lastrow + 4).Merge
    pedidoDeVenda2.Range("B" & lastrow + 4 & ":J" & lastrow + 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    pedidoDeVenda2.Range("A" & lastrow + 5).Value = "Nfe.:"
    pedidoDeVenda2.Range("B" & lastrow + 5 & ":J" & lastrow + 5).Merge
    pedidoDeVenda2.Range("B" & lastrow + 5 & ":J" & lastrow + 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "A").End(xlUp).Row
    
    pedidoDeVenda2.Range("A" & lastrow + 2).Value = "Comprador:"
    pedidoDeVenda2.Range("B" & lastrow + 2 & ":D" & lastrow + 2).Merge
        
    pedidoDeVenda2.Range("E" & lastrow + 2).Value = "Atendente:"
    pedidoDeVenda2.Range("F" & lastrow + 2 & ":J" & lastrow + 2).Merge
    pedidoDeVenda2.Range("F" & lastrow + 2 & ":J" & lastrow + 2).HorizontalAlignment = xlCenter
    pedidoDeVenda2.Range("F" & lastrow + 2).Value = orcamentoInicial.Range("B2").Value
    
    pedidoDeVenda2.Range("A" & lastrow + 4).Value = "Ass:"
    pedidoDeVenda2.Range("E" & lastrow + 4).Value = "Ass:"
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "A").End(xlUp).Row
    
    pedidoDeVenda2.Range("A" & lastrow + 1).Value = "OBSERVAÇÕES"
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).Merge
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).HorizontalAlignment = xlCenter
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).Font.Bold = True
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "A").End(xlUp).Row
    pedidoDeVenda2.Range("A" & lastrow + 1).Value = "FAVOR CONFERIR O PEDIDO E RETONAR ASSINADO"
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).Merge
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).HorizontalAlignment = xlCenter
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).Font.Bold = True
    pedidoDeVenda2.Range("A" & lastrow + 1 & ":J" & lastrow + 1).Font.Color = RGB(255, 0, 0)
    
    lastrow = pedidoDeVenda2.Cells(Rows.Count, "A").End(xlUp).Row
    
    pedidoDeVenda2.Range("D32:D" & lastrow).HorizontalAlignment = xlLeft
    
    Application.DisplayAlerts = False
    resultSheet.Delete
    ThisWorkbook.Sheets("orcamentoGerado").Delete
    ThisWorkbook.Sheets("cenarios").Delete
    Application.DisplayAlerts = True
    
End Sub
