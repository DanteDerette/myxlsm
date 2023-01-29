VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Orcamento 
   Caption         =   "Orcamento"
   ClientHeight    =   8760.001
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   17844
   OleObjectBlob   =   "Orcamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub codigodoproduto_cenario1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    noEnterLancaNoFDozeAbre (KeyCode)
End Sub
Sub noEnterLancaNoFDozeAbre(X)
    If X = 13 Then
        
       CommandButton7_Click
    ElseIf X = vbKeyF2 Then
        If Orcamento.operacao.Value = "" Then
            MsgBox ("Favor Escolher a Opera��o")
            Orcamento.operacao.SetFocus
            Orcamento.operacao.DropDown
        ElseIf Orcamento.qualCenario.Value = "" Then
            MsgBox ("Favor Escolher um Cen�rio")
            Orcamento.qualCenario.SetFocus
            Orcamento.qualCenario.DropDown
        Else
            listaDeProdutos_paraEscolher.Show
        End If
    End If
    
End Sub

Private Sub CommandButton10_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim wsDoCenario As Worksheet
    Dim lastrow As Integer
    Dim lastrow2 As Integer
    Dim lastrow_wsDoCenario As Integer
    Dim contagem As Integer
    Dim mes As Integer
    Dim mesPorExtenso As String
    Dim X As Integer
    Dim Y As Integer
    Dim irow As Integer
    Dim indice As Integer
        
    Set wb = Workbooks.Open(urlDoOrcamento, ReadOnly = True)
    Set wbClientes = Workbooks.Open("C:\GitHub\myxlsm\clientes.xlsx", ReadOnly = True)
    Set ws = wb.Sheets("cenarios")
    Set ws2 = wb.Sheets("resultado")
    
    mes = Format(Date, "mm")
    mesPorExtenso = Switch( _
    mes = 1, "Janeiro", _
    mes = 2, "Fevereiro", _
    mes = 3, "Mar�o", _
    mes = 4, "Abril", _
    mes = 5, "Maio", _
    mes = 6, "Junho", _
    mes = 7, "Julho", _
    mes = 8, "Agosto", _
    mes = 9, "Setembro", _
    mes = 10, "Outubro", _
    mes = 11, "Novembro", _
    mes = 12, "Dezembro")
        
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    With ws2.Range("C3")
        .Value = "Or�amento"
        .HorizontalAlignment = xlCenter
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
    End With
    
    ws2.Columns("C:C").AutoFit
    ws2.Columns("B:B").AutoFit
    
    ws2.Range("B4").Value = "Joinville, " & Format(Date, "dd") & " de " & mesPorExtenso & " " & Format(Date, "yyyy")
    
    ws2.Range("A6").Value = "Cliente:"
    ws2.Range("B6").Value = wbClientes.Worksheets("BD").Cells(Orcamento.idCliente.Value, 2).Value
    
    ws2.Range("A7").Value = "Cidade:"
    irow = Orcamento.idCliente.Value
    indice = Orcamento.optDeContato.ListIndex
    
    If indice = -1 Then
        indice = 0
    End If
            
    Y = 14
    X = 16 + (Y * indice)
    
    ws2.Range("B7").Value = wbClientes.Worksheets("BD").Cells(irow, X).Value
    ws2.Range("A8").Value = "Telefone:"
    ws2.Range("B8").Value = wbClientes.Worksheets("BD").Cells(irow, X + 3).Value
    ws2.Range("A9").Value = "Contato:"
    ws2.Range("B9").Value = wbClientes.Worksheets("BD").Cells(irow, X + 1).Value
    ws2.Range("A10").Value = "Email:"
    ws2.Range("B10").Value = wbClientes.Worksheets("BD").Cells(irow, X + 4).Value
     
    wbClientes.Close False
    
    With ws2.Range("A6:A10")
       .HorizontalAlignment = xlRight
    End With

    contagem = 1
    
    ' Format(Date, "dd/mm/yyyy")
        
    For i = 2 To lastrow
        lastrow2 = ws2.Range("A" & ws.Rows.Count).End(xlUp).Row

        ws2.Cells(lastrow2 + 1, 1).Value = "Cen�rio: " & ws.Cells(i, 2).Value ' T�tulo do Cen�rio

        With ws2.Range("A" & lastrow2 + 1 & ":F" & lastrow2 + 1) ' Estilo do T�tulo
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .BorderAround
        End With

        Set wsDoCenario = wb.Sheets(ws.Cells(i, 2).Value)
        lastrow_wsDoCenario = wsDoCenario.Range("A" & ws.Rows.Count).End(xlUp).Row

        lastrow2 = ws2.Range("A" & ws.Rows.Count).End(xlUp).Row
        ws2.Cells(lastrow2 + 1, 1).Value = "Sequencia"
        ws2.Cells(lastrow2 + 1, 2).Value = "C�digo do Produto"
        ws2.Cells(lastrow2 + 1, 3).Value = "Quantidade"
        ws2.Cells(lastrow2 + 1, 4).Value = "Valor Unit�rio"

        'ws2.Cells(lastrow2 + 1, 5).Value = "Desconto"
        'ws2.Cells(lastrow2 + 1, 6).Value = "Instala��o"
        'ws2.Cells(lastrow2 + 1, 7).Value = "Frete"

        ws2.Cells(lastrow2 + 1, 5).Value = "Valor Total"
        ws2.Cells(lastrow2 + 1, 6).Value = "Descri��o"

        For X = 2 To lastrow_wsDoCenario
            lastrow2 = ws2.Range("A" & ws.Rows.Count).End(xlUp).Row
            ws2.Cells(lastrow2 + 1, 1).Value = contagem
            contagem = contagem + 1
            'ws2.Cells(lastrow2 + 1, 1).Value = wsDoCenario.Cells(x, 1).Value ' ID do Produto
            ws2.Cells(lastrow2 + 1, 2).Value = wsDoCenario.Cells(X, 2).Value ' C�digo do Produto
            ws2.Cells(lastrow2 + 1, 3).Value = wsDoCenario.Cells(X, 3).Value
            ws2.Cells(lastrow2 + 1, 4).Value = wsDoCenario.Cells(X, 4).Value
            'ws2.Cells(lastrow2 + 1, 5).Value = wsDoCenario.Cells(x, 5).Value'
            'ws2.Cells(lastrow2 + 1, 6).Value = wsDoCenario.Cells(x, 6).Value
            'ws2.Cells(lastrow2 + 1, 7).Value = wsDoCenario.Cells(x, 7).Value
            ws2.Cells(lastrow2 + 1, 5).Value = wsDoCenario.Cells(X, 8).Value
            ws2.Cells(lastrow2 + 1, 6).Value = wsDoCenario.Cells(X, 9).Value

'            With ws2.Range("A" & lastrow2 + 1 & ":I" & lastrow2 + 1) ' Estilo do T�tulo
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                .WrapText = True
'            End With

        Next X
       'lastrow2 = ws2.Range("A" & ws.Rows.Count).End(xlUp).Row


    Next i

    wb.Activate
    ws2.Activate
'    ws2.Cells(1, 1).Value = ws.Cells(2, 2).Value
'    Set wsDoCenario = wb.Worksheets(ws.Cells(2, 2).Value)
'
'    ws2.Cells(2, 1).Value = wsDoCenario.Cells(2, 2).Value
'    ws2.Cells(3, 1).Value = wsDoCenario.Cells(3, 2).Value
'
'
    ws2.Columns("C:C").AutoFit
    ws2.Columns("B:B").AutoFit
    Unload Me
'
''    For i = 2 To lastrow
''        ws2.Cells(i, 1).Value = ws.Cells(i, 2).Value
''        Set wsDoCenario = wb.Worksheets(ws2.Cells(i, 1).Value)
''
''        With ws2.Range("A" & i & ":D" & i)
''            .Merge
''            .HorizontalAlignment = xlCenter
''            .VerticalAlignment = xlCenter
''        End With
''
''        lastrow2 = wsDoCenario.Range("B" & ws.Rows.Count).End(xlUp).Row
''
''        For i = 2 To lastrow
''
''    Next i
'
'    'wb.Close False
    
End Sub

Private Sub CommandButton2_Click()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wb As Workbook
    Dim wb2 As Workbook
    Dim ws As Worksheet
    Dim ws2 As Worksheet
        
    Set wb = Workbooks.Open(urlDoOrcamento, ReadOnly = False)
    Set ws = wb.Sheets("geral")
    
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row + 1
    
    ws.Cells(lastrow, 1).Value = Me.idOrcamento.Value
    ws.Cells(lastrow, 2).Value = Me.tituloDoOrcamento.Value
    ws.Cells(lastrow, 3).Value = Me.idCliente.Value
    ws.Cells(lastrow, 4).Value = Me.nomeFantasia.Value
    ws.Cells(lastrow, 5).Value = Format(Date, "dd/mm/yyyy")
    ws.Cells(lastrow, 6).Value = Me.urlDoOrcamento.Value
    ws.Cells(lastrow, 7).Value = Me.optDeContato.Value
    ws.Cells(lastrow, 8).Value = lastrow - 1
    
    Set wb2 = Workbooks.Open("C:\GitHub\myxlsm\orcamentos.xlsx", ReadOnly = False)
    Set ws2 = wb2.Sheets("BD")
    
    ws2.Cells(Me.idOrcamento.Value, 2).Value = Me.tituloDoOrcamento.Value
    
            
    wb.Close True
    wb2.Close True
    
    
    MsgBox ("Verso " & str(lastrow - 1) & " Gerada.")
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
End Sub


Private Sub CommandButton25_Click()
    Call CommandButton1_Click
End Sub
Private Sub CommandButton3_Click()
    Unload Me
End Sub
Private Sub CommandButton4_Click()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim novaplan As Worksheet
    Dim minharesposta As String
    
    minharesposta = InputBox("Nome do novo cenario: ")
    
    If StrPtr(minharesposta) = 0 Then
        Set wb = ThisWorkbook
    ElseIf minharesposta = vbNullString Then
        Set wb = ThisWorkbook
    Else
        Set wb = Workbooks.Open(Filename:=urlDoOrcamento, ReadOnly:=False)
        Set ws = wb.Sheets("cenarios")
    
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row + 1
        ws.Range("A" & lastrow).Value = lastrow
        ws.Range("B" & lastrow).Value = minharesposta
        
        Set novaplan = wb.Worksheets.Add(After:=ws)
        novaplan.Name = minharesposta
        
        Orcamento.qualCenario.AddItem (minharesposta)

        novaplan.Cells(1, 1).Value = "Seq."
        novaplan.Cells(1, 2).Value = "Codigo"
        novaplan.Cells(1, 3).Value = "Qtd."
        novaplan.Cells(1, 4).Value = "Alt."
        novaplan.Cells(1, 5).Value = "Larg."
        novaplan.Cells(1, 6).Value = "Comp."
        novaplan.Cells(1, 7).Value = "PotUnit"
        novaplan.Cells(1, 8).Value = "ValorUnit"
        novaplan.Cells(1, 9).Value = "Desconto"
        novaplan.Cells(1, 10).Value = "Descricao"
        
        wb.Close True
        
        Orcamento.qualCenario.Value = minharesposta
        Orcamento.codigodoproduto_cenario1.SetFocus
        
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
End Sub

Private Sub CommandButton7_Click()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim irow As Integer
    Dim lastrow As Integer
    Dim rnData As Range
       
    Dim id As String
    Dim lancamento As String
    Dim codigo As String
    Dim familia As String
    Dim ncm As String
    Dim especificacao1 As String
    Dim especificacao2 As String
    Dim especificacao3 As String
    Dim tipo As String
    Dim ALT As String
    Dim LARG As String
    Dim COMP As String
    Dim POT As String
    Dim mtCorda As String
    Dim peso As String
    Dim Venda As Variant
    Dim Locacao As Variant
    
    If Me.qualCenario.Value = "" Then
        MsgBox ("Favor escolher um cenario.")
        Exit Sub
    End If
        
    If Me.id_cenario1.Value = "" Then
        Set wb = Workbooks.Open(Filename:=Me.urlDoOrcamento.Value, ReadOnly:=False)
        Set ws = wb.Worksheets(Me.qualCenario.Value)
        
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row + 1
                    
        ws.Select
        
        ws.Cells(lastrow, 1).Value = lastrow - 1
        ws.Cells(lastrow, 2).Value = Me.codigodoproduto_cenario1.Value
        ws.Cells(lastrow, 3).Value = Me.quantidade_cenario1.Value
        ws.Cells(lastrow, 4).Value = Me.alturaDoProduto.Value
        ws.Cells(lastrow, 5).Value = Me.larguraDoProduto.Value
        ws.Cells(lastrow, 6).Value = Me.comprimentoDoProduto.Value
        ws.Cells(lastrow, 7).Value = Me.potenciaUnitaria.Value
        ws.Cells(lastrow, 8).Value = Me.valorUnitario_cenario1.Value
        ws.Cells(lastrow, 9).Value = Me.desconto_cenario1.Value
        ws.Cells(lastrow, 10).Value = Me.descricaodoproduto_cenario1.Value
        
        Me.id_cenario1.Value = ""
        Me.codigodoproduto_cenario1.Value = ""
        Me.quantidade_cenario1.Value = ""
        Me.alturaDoProduto.Value = ""
        Me.larguraDoProduto.Value = ""
        Me.comprimentoDoProduto.Value = ""
        Me.potenciaUnitaria.Value = ""
        Me.valorUnitario_cenario1.Value = ""
        Me.desconto_cenario1.Value = ""
        Me.descricaodoproduto_cenario1.Value = ""
        Me.potenciaTotal.Value = ""
        Me.valorTotal_cenario1.Value = ""
        
        Me.codigodoproduto_cenario1.SetFocus
        
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        Set rnData = ws.Range(ws.Range("A1"), ws.Range("O" & lastrow))
    
        vaData = rnData.Value
        
        With Me.lstLista
            .ColumnCount = 15
            .Clear
            .List = vaData
            .ListIndex = -1
        End With
        
        wb.Close True
        
    Else
        Set wb = Workbooks.Open(Filename:=Me.urlDoOrcamento.Value, ReadOnly:=False)
        Set ws = wb.Worksheets(Me.qualCenario.Value)
        
        lastrow = Me.id_cenario1.Value + 1
                    
        ws.Select
        
        ws.Cells(lastrow, 1).Value = lastrow - 1
        ws.Cells(lastrow, 2).Value = Me.codigodoproduto_cenario1.Value
        ws.Cells(lastrow, 3).Value = Me.quantidade_cenario1.Value
        ws.Cells(lastrow, 4).Value = Me.alturaDoProduto.Value
        ws.Cells(lastrow, 5).Value = Me.larguraDoProduto.Value
        ws.Cells(lastrow, 6).Value = Me.comprimentoDoProduto.Value
        ws.Cells(lastrow, 7).Value = Me.potenciaUnitaria.Value
        ws.Cells(lastrow, 8).Value = Me.valorUnitario_cenario1.Value
        ws.Cells(lastrow, 9).Value = Me.desconto_cenario1.Value
        ws.Cells(lastrow, 10).Value = Me.descricaodoproduto_cenario1.Value
        
        Me.id_cenario1.Value = ""
        Me.codigodoproduto_cenario1.Value = ""
        Me.quantidade_cenario1.Value = ""
        Me.alturaDoProduto.Value = ""
        Me.larguraDoProduto.Value = ""
        Me.comprimentoDoProduto.Value = ""
        Me.potenciaUnitaria.Value = ""
        Me.valorUnitario_cenario1.Value = ""
        Me.desconto_cenario1.Value = ""
        Me.descricaodoproduto_cenario1.Value = ""
        Me.potenciaTotal.Value = ""
        Me.valorTotal_cenario1.Value = ""
        
        Me.codigodoproduto_cenario1.SetFocus
        
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        Set rnData = ws.Range(ws.Range("A1"), ws.Range("O" & lastrow))
    
        vaData = rnData.Value
        
        With Me.lstLista
            .ColumnCount = 15
            .Clear
            .List = vaData
            .ListIndex = -1
        End With
        
        wb.Close True
    
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
            
End Sub

Private Sub CommandButton8_Click()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim contador As Integer
    
    Set wb = Workbooks.Open(Filename:=Me.urlDoOrcamento.Value, ReadOnly:=False)
    Set ws = wb.Worksheets(Me.qualCenario.Value)
    
    lastrow = Me.id_cenario1.Value + 1
                
    ws.Select
    
    ws.Rows(lastrow).Delete
    
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    contador = 1
    
    For i = 2 To lastrow
        ws.Cells(i, 1).Value = contador
        contador = contador + 1
    Next i
    
    
    Set rnData = ws.Range(ws.Range("A1"), ws.Range("O" & lastrow))

    vaData = rnData.Value
    
    With Me.lstLista
        .ColumnCount = 15
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
    wb.Close True
    
    Me.id_cenario1.Value = ""
    Me.codigodoproduto_cenario1.Value = ""
    Me.quantidade_cenario1.Value = ""
    Me.alturaDoProduto.Value = ""
    Me.larguraDoProduto.Value = ""
    Me.comprimentoDoProduto.Value = ""
    Me.potenciaUnitaria.Value = ""
    Me.valorUnitario_cenario1.Value = ""
    Me.desconto_cenario1.Value = ""
    Me.descricaodoproduto_cenario1.Value = ""
    Me.potenciaTotal.Value = ""
    Me.valorTotal_cenario1.Value = ""
    
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub

Private Sub CommandButton9_Click()
    listaDeProdutos_paraEscolher.Show
    
End Sub

Private Sub desconto_cenario1_Change()
Call onExitNoProduto
End Sub

Private Sub frete_cenario1_Change()
Call onExitNoProduto
End Sub

Private Sub instalacao_cenario1_Change()
Call onExitNoProduto
End Sub

Private Sub Label253_Click()

End Sub

Private Sub lstLista_Click()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim irow As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    
    On Error GoTo ErrFailed
    irow = Me.lstLista.Value
    
    If irow >= 1 Then
        irow = irow + 1
        
        Set wb = Workbooks.Open(Filename:=Me.urlDoOrcamento.Value, ReadOnly:=True)
        Set ws = wb.Worksheets(Me.qualCenario.Value)
        
        Me.id_cenario1.Value = ws.Cells(irow, 1).Value
        Me.codigodoproduto_cenario1.Value = ws.Cells(irow, 2).Value
        Me.quantidade_cenario1.Value = ws.Cells(irow, 3).Value
        Me.alturaDoProduto.Value = ws.Cells(irow, 4).Value
        Me.larguraDoProduto.Value = ws.Cells(irow, 5).Value
        Me.comprimentoDoProduto.Value = ws.Cells(irow, 6).Value
        Me.potenciaUnitaria.Value = ws.Cells(irow, 7).Value
        Me.valorUnitario_cenario1.Value = ws.Cells(irow, 8).Value
        Me.desconto_cenario1.Value = ws.Cells(irow, 9).Value
        Me.descricaodoproduto_cenario1.Value = ws.Cells(irow, 10).Value
        
        wb.Close False
    End If
    
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
ErrFailed:

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub

Private Sub MultiPage1_Change()

End Sub


Private Sub operacao_Change()
    Orcamento.codigodoproduto_cenario1.SetFocus
End Sub


Private Sub optDeContato_Change()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim indice As Long
    Dim X As Integer
    Dim Y As Integer
    Dim irow As Long
    Dim wb As Workbook
    Dim ws As Worksheet
        
    Set wb = Workbooks.Open("C:\GitHub\myxlsm\clientes.xlsx", ReadOnly = True)
    Set ws = wb.Sheets("BD")
    
    irow = Orcamento.idCliente.Value
    indice = Orcamento.optDeContato.ListIndex
    
    Y = 14
    X = 16 + (Y * indice)
    
    'Contatos1
    Orcamento.cidade_contato1.Value = ws.Cells(irow, X).Value
    'Comercial
    Orcamento.comercial_nome_contato1.Value = ws.Cells(irow, X + 1).Value
    Orcamento.comercial_cargo_contato1.Value = ws.Cells(irow, X + 2).Value
    Orcamento.comercial_telefone1_contato1.Value = ws.Cells(irow, X + 3).Value
    Orcamento.comercial_email1_contato1.Value = ws.Cells(irow, X + 4).Value
    Orcamento.comercial_telefone2_contato1.Value = ws.Cells(irow, X + 5).Value
    Orcamento.comercial_email2_contato1.Value = ws.Cells(irow, X + 6).Value
    'Financeiro
    Orcamento.financeiro_nome_contato1.Value = ws.Cells(irow, X + 7).Value
    Orcamento.financeiro_cargo_contato1.Value = ws.Cells(irow, X + 8).Value
    Orcamento.financeiro_telefone1_contato1.Value = ws.Cells(irow, X + 9).Value
    Orcamento.financeiro_email1_contato1.Value = ws.Cells(irow, X + 10).Value
    Orcamento.financeiro_telefone2_contato1.Value = ws.Cells(irow, X + 11).Value
    Orcamento.financeiro_email2_contato1.Value = ws.Cells(irow, X + 12).Value
    'Observa��o
    Orcamento.observacaoDoContato_contato1.Value = ws.Cells(irow, X + 13).Value
    
    wb.Close
    
    
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    

End Sub
Private Sub CommandButton1_Click()
   With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim found As Range
    Dim irow As Integer
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx", ReadOnly:=True)
    Set ws = wb.Worksheets("BD")
        
    irow = Orcamento.idCliente.Value

    If irow > 1 Then
        cadastroDeCliente.id.Value = ws.Cells(irow, 1).Value
        cadastroDeCliente.nomeFantasia.Value = ws.Cells(irow, 2).Value
        cadastroDeCliente.cnpj.Value = ws.Cells(irow, 3).Value
        cadastroDeCliente.razaoSocial.Value = ws.Cells(irow, 4).Value
        cadastroDeCliente.atendimento.Value = ws.Cells(irow, 5).Value
        cadastroDeCliente.inscricaoEstadual.Value = ws.Cells(irow, 6).Value
        cadastroDeCliente.clienteDesde.Value = ws.Cells(irow, 7).Value

        ' Aba Endere�o
        cadastroDeCliente.cep.Value = ws.Cells(irow, 8).Value
        cadastroDeCliente.estado.Value = ws.Cells(irow, 9).Value
        cadastroDeCliente.cidade.Value = ws.Cells(irow, 10).Value
        cadastroDeCliente.bairro.Value = ws.Cells(irow, 11).Value
        cadastroDeCliente.endereco.Value = ws.Cells(irow, 12).Value
        cadastroDeCliente.regiao.Value = ws.Cells(irow, 13).Value
        cadastroDeCliente.complemento.Value = ws.Cells(irow, 14).Value

        'Aba Observa��o
        cadastroDeCliente.observacao.Value = ws.Cells(irow, 15).Value

        'Contatos
        'Contatos1
        cadastroDeCliente.cidade_contato1.Value = ws.Cells(irow, 16).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato1.Value = ws.Cells(irow, 17).Value
        cadastroDeCliente.comercial_cargo_contato1.Value = ws.Cells(irow, 18).Value
        cadastroDeCliente.comercial_telefone1_contato1.Value = ws.Cells(irow, 19).Value
        cadastroDeCliente.comercial_email1_contato1.Value = ws.Cells(irow, 20).Value
        cadastroDeCliente.comercial_telefone2_contato1.Value = ws.Cells(irow, 21).Value
        cadastroDeCliente.comercial_email2_contato1.Value = ws.Cells(irow, 22).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato1.Value = ws.Cells(irow, 23).Value
        cadastroDeCliente.financeiro_cargo_contato1.Value = ws.Cells(irow, 24).Value
        cadastroDeCliente.financeiro_telefone1_contato1.Value = ws.Cells(irow, 25).Value
        cadastroDeCliente.financeiro_email1_contato1.Value = ws.Cells(irow, 26).Value
        cadastroDeCliente.financeiro_telefone2_contato1.Value = ws.Cells(irow, 27).Value
        cadastroDeCliente.financeiro_email2_contato1.Value = ws.Cells(irow, 28).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato1.Value = ws.Cells(irow, 29).Value

        'Contatos2
        cadastroDeCliente.cidade_contato2.Value = ws.Cells(irow, 30).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato2.Value = ws.Cells(irow, 31).Value
        cadastroDeCliente.comercial_cargo_contato2.Value = ws.Cells(irow, 32).Value
        cadastroDeCliente.comercial_telefone1_contato2.Value = ws.Cells(irow, 33).Value
        cadastroDeCliente.comercial_email1_contato2.Value = ws.Cells(irow, 34).Value
        cadastroDeCliente.comercial_telefone2_contato2.Value = ws.Cells(irow, 35).Value
        cadastroDeCliente.comercial_email2_contato2.Value = ws.Cells(irow, 36).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato2.Value = ws.Cells(irow, 37).Value
        cadastroDeCliente.financeiro_cargo_contato2.Value = ws.Cells(irow, 38).Value
        cadastroDeCliente.financeiro_telefone1_contato2.Value = ws.Cells(irow, 39).Value
        cadastroDeCliente.financeiro_email1_contato2.Value = ws.Cells(irow, 40).Value
        cadastroDeCliente.financeiro_telefone2_contato2.Value = ws.Cells(irow, 41).Value
        cadastroDeCliente.financeiro_email2_contato2.Value = ws.Cells(irow, 42).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato2.Value = ws.Cells(irow, 43).Value

        'Contatos3
        cadastroDeCliente.cidade_contato3.Value = ws.Cells(irow, 44).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato3.Value = ws.Cells(irow, 45).Value
        cadastroDeCliente.comercial_cargo_contato3.Value = ws.Cells(irow, 46).Value
        cadastroDeCliente.comercial_telefone1_contato3.Value = ws.Cells(irow, 47).Value
        cadastroDeCliente.comercial_email1_contato3.Value = ws.Cells(irow, 48).Value
        cadastroDeCliente.comercial_telefone2_contato3.Value = ws.Cells(irow, 49).Value
        cadastroDeCliente.comercial_email2_contato3.Value = ws.Cells(irow, 50).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato3.Value = ws.Cells(irow, 51).Value
        cadastroDeCliente.financeiro_cargo_contato3.Value = ws.Cells(irow, 52).Value
        cadastroDeCliente.financeiro_telefone1_contato3.Value = ws.Cells(irow, 53).Value
        cadastroDeCliente.financeiro_email1_contato3.Value = ws.Cells(irow, 54).Value
        cadastroDeCliente.financeiro_telefone2_contato3.Value = ws.Cells(irow, 55).Value
        cadastroDeCliente.financeiro_email2_contato3.Value = ws.Cells(irow, 56).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato3.Value = ws.Cells(irow, 57).Value

        'Contatos4
        cadastroDeCliente.cidade_contato4.Value = ws.Cells(irow, 58).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato4.Value = ws.Cells(irow, 59).Value
        cadastroDeCliente.comercial_cargo_contato4.Value = ws.Cells(irow, 60).Value
        cadastroDeCliente.comercial_telefone1_contato4.Value = ws.Cells(irow, 61).Value
        cadastroDeCliente.comercial_email1_contato4.Value = ws.Cells(irow, 62).Value
        cadastroDeCliente.comercial_telefone2_contato4.Value = ws.Cells(irow, 63).Value
        cadastroDeCliente.comercial_email2_contato4.Value = ws.Cells(irow, 64).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato4.Value = ws.Cells(irow, 65).Value
        cadastroDeCliente.financeiro_cargo_contato4.Value = ws.Cells(irow, 66).Value
        cadastroDeCliente.financeiro_telefone1_contato4.Value = ws.Cells(irow, 67).Value
        cadastroDeCliente.financeiro_email1_contato4.Value = ws.Cells(irow, 68).Value
        cadastroDeCliente.financeiro_telefone2_contato4.Value = ws.Cells(irow, 69).Value
        cadastroDeCliente.financeiro_email2_contato4.Value = ws.Cells(irow, 70).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato4.Value = ws.Cells(irow, 71).Value

        'Contatos5
        cadastroDeCliente.cidade_contato5.Value = ws.Cells(irow, 72).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato5.Value = ws.Cells(irow, 73).Value
        cadastroDeCliente.comercial_cargo_contato5.Value = ws.Cells(irow, 74).Value
        cadastroDeCliente.comercial_telefone1_contato5.Value = ws.Cells(irow, 75).Value
        cadastroDeCliente.comercial_email1_contato5.Value = ws.Cells(irow, 76).Value
        cadastroDeCliente.comercial_telefone2_contato5.Value = ws.Cells(irow, 77).Value
        cadastroDeCliente.comercial_email2_contato5.Value = ws.Cells(irow, 78).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato5.Value = ws.Cells(irow, 79).Value
        cadastroDeCliente.financeiro_cargo_contato5.Value = ws.Cells(irow, 80).Value
        cadastroDeCliente.financeiro_telefone1_contato5.Value = ws.Cells(irow, 81).Value
        cadastroDeCliente.financeiro_email1_contato5.Value = ws.Cells(irow, 82).Value
        cadastroDeCliente.financeiro_telefone2_contato5.Value = ws.Cells(irow, 83).Value
        cadastroDeCliente.financeiro_email2_contato5.Value = ws.Cells(irow, 84).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato5.Value = ws.Cells(irow, 85).Value

        'Contatos6
        cadastroDeCliente.cidade_contato6.Value = ws.Cells(irow, 86).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato6.Value = ws.Cells(irow, 87).Value
        cadastroDeCliente.comercial_cargo_contato6.Value = ws.Cells(irow, 88).Value
        cadastroDeCliente.comercial_telefone1_contato6.Value = ws.Cells(irow, 89).Value
        cadastroDeCliente.comercial_email1_contato6.Value = ws.Cells(irow, 90).Value
        cadastroDeCliente.comercial_telefone2_contato6.Value = ws.Cells(irow, 91).Value
        cadastroDeCliente.comercial_email2_contato6.Value = ws.Cells(irow, 92).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato6.Value = ws.Cells(irow, 93).Value
        cadastroDeCliente.financeiro_cargo_contato6.Value = ws.Cells(irow, 94).Value
        cadastroDeCliente.financeiro_telefone1_contato6.Value = ws.Cells(irow, 95).Value
        cadastroDeCliente.financeiro_email1_contato6.Value = ws.Cells(irow, 96).Value
        cadastroDeCliente.financeiro_telefone2_contato6.Value = ws.Cells(irow, 97).Value
        cadastroDeCliente.financeiro_email2_contato6.Value = ws.Cells(irow, 98).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato6.Value = ws.Cells(irow, 99).Value

        'Contatos7
        cadastroDeCliente.cidade_contato7.Value = ws.Cells(irow, 100).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato7.Value = ws.Cells(irow, 101).Value
        cadastroDeCliente.comercial_cargo_contato7.Value = ws.Cells(irow, 102).Value
        cadastroDeCliente.comercial_telefone1_contato7.Value = ws.Cells(irow, 103).Value
        cadastroDeCliente.comercial_email1_contato7.Value = ws.Cells(irow, 104).Value
        cadastroDeCliente.comercial_telefone2_contato7.Value = ws.Cells(irow, 105).Value
        cadastroDeCliente.comercial_email2_contato7.Value = ws.Cells(irow, 106).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato7.Value = ws.Cells(irow, 107).Value
        cadastroDeCliente.financeiro_cargo_contato7.Value = ws.Cells(irow, 108).Value
        cadastroDeCliente.financeiro_telefone1_contato7.Value = ws.Cells(irow, 109).Value
        cadastroDeCliente.financeiro_email1_contato7.Value = ws.Cells(irow, 110).Value
        cadastroDeCliente.financeiro_telefone2_contato7.Value = ws.Cells(irow, 111).Value
        cadastroDeCliente.financeiro_email2_contato7.Value = ws.Cells(irow, 112).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato7.Value = ws.Cells(irow, 113).Value

        'Contatos8
        cadastroDeCliente.cidade_contato8.Value = ws.Cells(irow, 114).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato8.Value = ws.Cells(irow, 115).Value
        cadastroDeCliente.comercial_cargo_contato8.Value = ws.Cells(irow, 116).Value
        cadastroDeCliente.comercial_telefone1_contato8.Value = ws.Cells(irow, 117).Value
        cadastroDeCliente.comercial_email1_contato8.Value = ws.Cells(irow, 118).Value
        cadastroDeCliente.comercial_telefone2_contato8.Value = ws.Cells(irow, 119).Value
        cadastroDeCliente.comercial_email2_contato8.Value = ws.Cells(irow, 120).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato8.Value = ws.Cells(irow, 121).Value
        cadastroDeCliente.financeiro_cargo_contato8.Value = ws.Cells(irow, 122).Value
        cadastroDeCliente.financeiro_telefone1_contato8.Value = ws.Cells(irow, 123).Value
        cadastroDeCliente.financeiro_email1_contato8.Value = ws.Cells(irow, 124).Value
        cadastroDeCliente.financeiro_telefone2_contato8.Value = ws.Cells(irow, 125).Value
        cadastroDeCliente.financeiro_email2_contato8.Value = ws.Cells(irow, 126).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato8.Value = ws.Cells(irow, 127).Value

        'Contatos8
        cadastroDeCliente.cidade_contato8.Value = ws.Cells(irow, 114).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato8.Value = ws.Cells(irow, 115).Value
        cadastroDeCliente.comercial_cargo_contato8.Value = ws.Cells(irow, 116).Value
        cadastroDeCliente.comercial_telefone1_contato8.Value = ws.Cells(irow, 117).Value
        cadastroDeCliente.comercial_email1_contato8.Value = ws.Cells(irow, 118).Value
        cadastroDeCliente.comercial_telefone2_contato8.Value = ws.Cells(irow, 119).Value
        cadastroDeCliente.comercial_email2_contato8.Value = ws.Cells(irow, 120).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato8.Value = ws.Cells(irow, 121).Value
        cadastroDeCliente.financeiro_cargo_contato8.Value = ws.Cells(irow, 122).Value
        cadastroDeCliente.financeiro_telefone1_contato8.Value = ws.Cells(irow, 123).Value
        cadastroDeCliente.financeiro_email1_contato8.Value = ws.Cells(irow, 124).Value
        cadastroDeCliente.financeiro_telefone2_contato8.Value = ws.Cells(irow, 125).Value
        cadastroDeCliente.financeiro_email2_contato8.Value = ws.Cells(irow, 126).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato8.Value = ws.Cells(irow, 127).Value

        'Contatos9
        cadastroDeCliente.cidade_contato9.Value = ws.Cells(irow, 128).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato9.Value = ws.Cells(irow, 129).Value
        cadastroDeCliente.comercial_cargo_contato9.Value = ws.Cells(irow, 130).Value
        cadastroDeCliente.comercial_telefone1_contato9.Value = ws.Cells(irow, 131).Value
        cadastroDeCliente.comercial_email1_contato9.Value = ws.Cells(irow, 132).Value
        cadastroDeCliente.comercial_telefone2_contato9.Value = ws.Cells(irow, 133).Value
        cadastroDeCliente.comercial_email2_contato9.Value = ws.Cells(irow, 134).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato9.Value = ws.Cells(irow, 135).Value
        cadastroDeCliente.financeiro_cargo_contato9.Value = ws.Cells(irow, 136).Value
        cadastroDeCliente.financeiro_telefone1_contato9.Value = ws.Cells(irow, 137).Value
        cadastroDeCliente.financeiro_email1_contato9.Value = ws.Cells(irow, 138).Value
        cadastroDeCliente.financeiro_telefone2_contato9.Value = ws.Cells(irow, 139).Value
        cadastroDeCliente.financeiro_email2_contato9.Value = ws.Cells(irow, 140).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato9.Value = ws.Cells(irow, 141).Value

        'Contatos10
        cadastroDeCliente.cidade_contato10.Value = ws.Cells(irow, 142).Value
        'Comercial
        cadastroDeCliente.comercial_nome_contato10.Value = ws.Cells(irow, 143).Value
        cadastroDeCliente.comercial_cargo_contato10.Value = ws.Cells(irow, 144).Value
        cadastroDeCliente.comercial_telefone1_contato10.Value = ws.Cells(irow, 145).Value
        cadastroDeCliente.comercial_email1_contato10.Value = ws.Cells(irow, 146).Value
        cadastroDeCliente.comercial_telefone2_contato10.Value = ws.Cells(irow, 147).Value
        cadastroDeCliente.comercial_email2_contato10.Value = ws.Cells(irow, 148).Value
        'Financeiro
        cadastroDeCliente.financeiro_nome_contato10.Value = ws.Cells(irow, 149).Value
        cadastroDeCliente.financeiro_cargo_contato10.Value = ws.Cells(irow, 150).Value
        cadastroDeCliente.financeiro_telefone1_contato10.Value = ws.Cells(irow, 151).Value
        cadastroDeCliente.financeiro_email1_contato10.Value = ws.Cells(irow, 152).Value
        cadastroDeCliente.financeiro_telefone2_contato10.Value = ws.Cells(irow, 153).Value
        cadastroDeCliente.financeiro_email2_contato10.Value = ws.Cells(irow, 154).Value
        'Observa��o
        cadastroDeCliente.observacaoDoContato_contato10.Value = ws.Cells(irow, 155).Value

        cadastroDeCliente.ultimaAtualizacao.Value = ws.Cells(irow, 156).Value

        cadastroDeCliente.desc_anexo1.Value = ws.Cells(irow, 157).Value
        cadastroDeCliente.anexo1.Value = ws.Cells(irow, 158).Value

        cadastroDeCliente.desc_anexo2.Value = ws.Cells(irow, 159).Value
        cadastroDeCliente.anexo2.Value = ws.Cells(irow, 160).Value

        cadastroDeCliente.desc_anexo3.Value = ws.Cells(irow, 161).Value
        cadastroDeCliente.anexo3.Value = ws.Cells(irow, 162).Value

        cadastroDeCliente.desc_anexo4.Value = ws.Cells(irow, 163).Value
        cadastroDeCliente.anexo4.Value = ws.Cells(irow, 164).Value

        cadastroDeCliente.desc_anexo5.Value = ws.Cells(irow, 165).Value
        cadastroDeCliente.anexo5.Value = ws.Cells(irow, 166).Value

        cadastroDeCliente.desc_anexo6.Value = ws.Cells(irow, 167).Value
        cadastroDeCliente.anexo6.Value = ws.Cells(irow, 168).Value

        cadastroDeCliente.desc_anexo7.Value = ws.Cells(irow, 169).Value
        cadastroDeCliente.anexo7.Value = ws.Cells(irow, 170).Value

        cadastroDeCliente.desc_anexo8.Value = ws.Cells(irow, 171).Value
        cadastroDeCliente.anexo8.Value = ws.Cells(irow, 172).Value

        cadastroDeCliente.desc_anexo9.Value = ws.Cells(irow, 173).Value
        cadastroDeCliente.anexo9.Value = ws.Cells(irow, 174).Value

        cadastroDeCliente.desc_anexo10.Value = ws.Cells(irow, 175).Value
        cadastroDeCliente.anexo10.Value = ws.Cells(irow, 176).Value

        wb.Close
        cadastroDeCliente.qualWS.Value = Orcamento.urlDoOrcamento.Value
        cadastroDeCliente.Show
        
    End If

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    

End Sub

Private Sub qualCenario_Change()
   With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
        
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rnData As Range
    Dim lastrow As Integer
    
        
    Set wb = Workbooks.Open(Filename:=Me.urlDoOrcamento.Value, ReadOnly:=False)
    Set ws = wb.Worksheets(Me.qualCenario.Value)
    
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    Set rnData = ws.Range(ws.Range("A1"), ws.Range("O" & lastrow))
    
    vaData = rnData.Value
    
    With Me.lstLista
        .ColumnCount = 15
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
    wb.Close False
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
End Sub

Private Sub quantidade_cenario1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call onExitNoProduto
End Sub
Sub onExitNoProduto()
    If Orcamento.valorUnitario_cenario1.Value = "" Then
        Orcamento.valorUnitario_cenario1.Value = 0
    End If

    If Orcamento.quantidade_cenario1.Value = "" Then
        Orcamento.quantidade_cenario1.Value = 0
    End If

    If Orcamento.desconto_cenario1.Value = "" Then
        Orcamento.desconto_cenario1.Value = 0
    End If

    Orcamento.valorTotal_cenario1.Value = (CDbl(Orcamento.valorUnit�rio_cenario1.Value) * CDbl(Orcamento.quantidade_cenario1.Value)) _
    - CDbl(Orcamento.desconto_cenario1.Value)
End Sub

Private Sub TextBox208_Change()

End Sub

Private Sub UserForm_Initialize()
    With Me.operacao
        .AddItem "Venda"
        .AddItem "Loca��o"
    End With
    
    
    With Me.optDeContato
        .AddItem "Contato 1"
        .AddItem "Contato 2"
        .AddItem "Contato 3"
        .AddItem "Contato 4"
        .AddItem "Contato 5"
        .AddItem "Contato 6"
        .AddItem "Contato 7"
        .AddItem "Contato 8"
        .AddItem "Contato 9"
        .AddItem "Contato 10"
    End With
    
    
    
    
    ' qualCenario
    
    ' removeTudo
'    Do While ComboBox1.ListCount > 0
'        ComboBox1.RemoveItem (0)
'    Loop
    

    
    
'
'    With opera��oDoOr�amento
'        .AddItem "Venda"
'        .AddItem "Loca��o"
'        .AddItem "Manuten��o"
'    End With
    
    
End Sub

Private Sub valorTotal_cenario1_Change()
    Call onExitNoProduto
End Sub

Private Sub valorUnitario_cenario1_Change()
    Call onExitNoProduto
End Sub

