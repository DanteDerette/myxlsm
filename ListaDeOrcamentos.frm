VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} listaDeOrcamentos 
   Caption         =   "Lista de Or�amentos"
   ClientHeight    =   6816
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10488
   OleObjectBlob   =   "ListaDeOrcamentos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListaDeOrcamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optDeContato_Change()
    Dim indice As Long
    Dim X As Integer
    Dim Y As Integer
    Dim irow As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("BD")
        
    If Orcamento.idCliente.Value = "" Then
        Orcamento.optDeContato.Value = ""
       
        Orcamento.nomeFantasia.SetFocus
       ' Return goto macroS
        
    Else
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
    End If
    
End Sub

Private Sub CommandButton1_Click()
    Unload Me
    GerarOrcamento.Show
End Sub

Private Sub lstLista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim wb2 As Workbook
    Dim ws2 As Worksheet
    
    Dim found As Range
    Dim irow As Integer
    
    Dim lastrow As Integer
    
    Dim idOrcamento As String
    Dim tituloDoOrcamento As String
    Dim idCliente As String
    Dim nomeFantasia As String
    Dim urlDoOrcamento As String
    Dim dataCriacao As String
    Dim primeiroCenario As String
    
    On Error GoTo ErrFailed
    irow = lstLista.Value
    
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\orcamentos.xlsx", ReadOnly:=True)
    Set ws = wb.Worksheets("BD")
    
    urlDoOrcamento = ws.Cells(irow, 6).Value
    
    wb.Close False
    
    Set wb2 = Workbooks.Open(Filename:=urlDoOrcamento, ReadOnly:=True)
    Set ws = wb2.Worksheets("Geral")
    
    lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    idOrcamento = ws.Cells(lastrow, 1).Value
    tituloDoOrcamento = ws.Cells(lastrow, 2).Value
    idCliente = ws.Cells(lastrow, 3).Value
    nomeFantasia = ws.Cells(lastrow, 4).Value
    dataCriacao = ws.Cells(lastrow, 5).Value
    urlDoOrcamento = ws.Cells(lastrow, 6).Value
    
    If irow > 1 Then
        Orcamento.idOrcamento.Value = idOrcamento
        Orcamento.tituloDoOrcamento.Value = tituloDoOrcamento
        Orcamento.idCliente.Value = idCliente
        Orcamento.nomeFantasia.Value = nomeFantasia
'       Orcamento.dataCriacao.Value = ws.Cells(irow, 5).Value
        Orcamento.urlDoOrcamento.Value = urlDoOrcamento
        Orcamento.optDeContato.Value = ws.Cells(lastrow, 7).Value
        
        Set ws = wb2.Worksheets("cenarios")
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
                
        For i = 2 To lastrow
            With Orcamento.qualCenario
                .AddItem (ws.Cells(i, 2).Value)
            End With
        Next i
        
        primeiroCenario = ws.Cells(lastrow, 2).Value
        
        wb2.Close False
        
        If primeiroCenario <> "nomeDoCenario" Then
            Orcamento.qualCenario.Value = primeiroCenario
        End If
    End If
    
    Unload Me
    Orcamento.Show

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
'
ErrFailed:
    Set wb = ThisWorkbook
End Sub
        
Private Sub lstlista_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
            .DisplayAlerts = False
        End With
        
        Dim wb As Workbook
        Dim ws As Worksheet
        
        Dim wb2 As Workbook
        Dim ws2 As Worksheet
        
        Dim found As Range
        Dim irow As Integer
        
        Dim lastrow As Integer
        
        Dim idOrcamento As String
        Dim tituloDoOrcamento As String
        Dim idCliente As String
        Dim nomeFantasia As String
        Dim urlDoOrcamento As String
        Dim dataCriacao As String
        
        On Error GoTo ErrFailed
        irow = lstLista.Value
        
        Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\orcamentos.xlsx", ReadOnly:=True)
        Set ws = wb.Worksheets("BD")
        
        urlDoOrcamento = ws.Cells(irow, 6).Value
        
        wb.Close False
        
        Set wb2 = Workbooks.Open(Filename:=urlDoOrcamento, ReadOnly:=True)
        Set ws = wb2.Worksheets("Geral")
        
        lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        
        idOrcamento = ws.Cells(lastrow, 1).Value
        tituloDoOrcamento = ws.Cells(lastrow, 2).Value
        idCliente = ws.Cells(lastrow, 3).Value
        nomeFantasia = ws.Cells(lastrow, 4).Value
        dataCriacao = ws.Cells(lastrow, 5).Value
        urlDoOrcamento = ws.Cells(lastrow, 6).Value
        
        If irow > 1 Then
            Orcamento.idOrcamento.Value = idOrcamento
            Orcamento.tituloDoOrcamento.Value = tituloDoOrcamento
            Orcamento.idCliente.Value = idCliente
            Orcamento.nomeFantasia.Value = nomeFantasia
    '       Orcamento.dataCriacao.Value = ws.Cells(irow, 5).Value
            Orcamento.urlDoOrcamento.Value = urlDoOrcamento
            Orcamento.optDeContato.Value = ws.Cells(lastrow, 7).Value
            
            Set ws = wb2.Worksheets("cenarios")
            lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
                    
            For i = 2 To lastrow
                With Orcamento.qualCenario
                    .AddItem (ws.Cells(i, 2).Value)
                End With
            Next i
            
            wb2.Close False
        End If

        Unload Me
        Orcamento.Show
    
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
        End With
    '
ErrFailed:
        Set wb = ThisWorkbook
    End If
End Sub

Private Sub TextBoxFiltro_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    If ComboBoxCampos.Value = "" Then
        ComboBoxCampos.Value = "Titulo"
     End If
    If KeyCode = 13 Then
        If Len(TextBoxFiltro.Text) > 0 Then
            Call PreencheLista_Melhorado(TextBoxFiltro.Text, "orcamentos.xlsx", me)
            If Me.lstLista.ListCount > 1 Then
                Me.lstLista.SetFocus
                Me.lstLista.ListIndex = 1
            End If
        Else
            call init_lstLista("orcamentos.xlsx", me)
        End If
    End If
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub


Private Sub UserForm_Initialize()
    Me.ComboBoxCampos.AddItem "id"
    Me.ComboBoxCampos.AddItem "Titulo"
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    call init_lstLista("orcamentos.xlsx", me)
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub


  
