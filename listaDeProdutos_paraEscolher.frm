VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} listaDeProdutos_paraEscolher 
   Caption         =   "listaDeProdutos_paraEscolher"
   ClientHeight    =   6984
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16860
   OleObjectBlob   =   "listaDeProdutos_paraEscolher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "listaDeProdutos_paraEscolher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstLista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim found As Range
    Dim irow As Integer
    Dim id As String
    Dim lancamento As String
    Dim c�digo As String
    Dim fam�lia As String
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
    
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
    Set ws = wb.Worksheets("BD")
    
    On Error GoTo ErrFailed
    irow = lstLista.Value
    
    
    If irow > 1 Then
        On Error GoTo ErrFailed
        irow = lstLista.Value
        
        id = ws.Cells(irow, 1).Value
        lancamento = ws.Cells(irow, 2).Value
        c�digo = ws.Cells(irow, 3).Value
        fam�lia = ws.Cells(irow, 4).Value
        ncm = ws.Cells(irow, 5).Value
        especificacao1 = ws.Cells(irow, 6).Value
        especificacao2 = ws.Cells(irow, 7).Value
        especificacao3 = ws.Cells(irow, 8).Value
        tipo = ws.Cells(irow, 9).Value
        ALT = ws.Cells(irow, 10).Value
        LARG = ws.Cells(irow, 11).Value
        COMP = ws.Cells(irow, 12).Value
        POT = ws.Cells(irow, 13).Value
        mtCorda = ws.Cells(irow, 14).Value
        peso = ws.Cells(irow, 15).Value
        Venda = ws.Cells(irow, 36).Value
        Locacao = ws.Cells(irow, 37).Value
        
        Orcamento.id_cenario1.Value = ""
        Orcamento.codigodoproduto_cenario1.Value = ""
        Orcamento.quantidade_cenario1.Value = ""
        Orcamento.valorUnitario_cenario1.Value = ""
        Orcamento.desconto_cenario1.Value = ""
        Orcamento.valorTotal_cenario1.Value = ""
        Orcamento.descricaodoproduto_cenario1.Value = ""
        Orcamento.potenciaUnitaria.Value = ""
        Orcamento.comprimentoDoProduto.Value = ""
        Orcamento.larguraDoProduto.Value = ""
        Orcamento.alturaDoProduto.Value = ""
                        
        Orcamento.codigodoproduto_cenario1.Value = ws.Cells(irow, 3).Value
        Orcamento.quantidade_cenario1.Value = 0
        Orcamento.alturaDoProduto.Value = ALT
        Orcamento.larguraDoProduto.Value = LARG
        Orcamento.comprimentoDoProduto.Value = COMP
        Orcamento.potenciaUnitaria.Value = POT
        Orcamento.descricaodoproduto_cenario1.Value = ws.Cells(irow, 6).Value
        
        If Orcamento.operacao.Value = "Venda" Then
            Orcamento.valorUnitario_cenario1.Value = ws.Cells(irow, 36).Value
        ElseIf Orcamento.operacao.Value = "Locacao" Then
            Orcamento.valorUnitario_cenario1.Value = ws.Cells(irow, 37).Value
        End If
        

        
    End If
    
    Unload Me
    wb.Close False
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
ErrFailed:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub
Private Sub lstlista_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
            With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
        
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim found As Range
        Dim irow As Integer
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
        
        Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
        Set ws = wb.Worksheets("BD")
        
    
        ' Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
        ' Set wsSheet = wbBook.Worksheets(1)
            
        On Error GoTo ErrFailed
        irow = lstLista.Value
        
        
        If irow > 1 Then
            On Error GoTo ErrFailed
            irow = lstLista.Value
            
            id = ws.Cells(irow, 1).Value
            lancamento = ws.Cells(irow, 2).Value
            codigo = ws.Cells(irow, 3).Value
            familia = ws.Cells(irow, 4).Value
            ncm = ws.Cells(irow, 5).Value
            especificacao1 = ws.Cells(irow, 6).Value
            especificacao2 = ws.Cells(irow, 7).Value
            especificacao3 = ws.Cells(irow, 8).Value
            tipo = ws.Cells(irow, 9).Value
            ALT = ws.Cells(irow, 10).Value
            LARG = ws.Cells(irow, 11).Value
            COMP = ws.Cells(irow, 12).Value
            POT = ws.Cells(irow, 13).Value
            mtCorda = ws.Cells(irow, 14).Value
            peso = ws.Cells(irow, 15).Value
            Venda = ws.Cells(irow, 36).Value
            Locacao = ws.Cells(irow, 37).Value
            
            Orcamento.id_cenario1.Value = ""
            Orcamento.codigodoproduto_cenario1.Value = ""
            Orcamento.quantidade_cenario1.Value = ""
            Orcamento.valorUnitario_cenario1.Value = ""
            Orcamento.desconto_cenario1.Value = ""
            'Orcamento.instalacao_cenario1.Value = ""
            'Orcamento.frete_cenario1.Value = ""
            Orcamento.valorTotal_cenario1.Value = ""
            Orcamento.descricaodoproduto_cenario1.Value = ""
                    
            Orcamento.codigodoproduto_cenario1.Value = ws.Cells(irow, 3).Value
            Orcamento.quantidade_cenario1.Value = 0
            Orcamento.alturaDoProduto.Value = ALT
            Orcamento.larguraDoProduto.Value = LARG
            Orcamento.comprimentoDoProduto.Value = COMP
            Orcamento.potenciaUnitaria.Value = POT
            Orcamento.descricaodoproduto_cenario1.Value = ws.Cells(irow, 6).Value
            
            'If Orcamento.operacaoDoOrcamento.Value = "Venda" Then
                Orcamento.valorUnitario_cenario1.Value = ws.Cells(irow, 36).Value
            'ElseIf Orcamento.operacaoDoOrcamento.Value = "Locacao" Then
                'Orcamento.valorUnitario_cenario1.Value = ws.Cells(irow, 37).Value
            'End If
        End If
        
        Unload Me
        wb.Close False
        
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End With
        
ErrFailed:
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End With
        
    End If
    
End Sub

Private Sub TextBoxFiltro_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
         If ComboBoxCampos.Value = "" Then
            ComboBoxCampos.Value = "codigo"
        End If
        
        If Len(TextBoxFiltro.Text) > 0 Then
            Call PreencheLista_Melhorado(TextBoxFiltro.Text)
            Me.lstLista.SetFocus
            Me.lstLista.ListIndex = 1
        Else
            Dim wbBook As Workbook
            Dim wsSheet As Worksheet
            Dim rnData As Range
            Dim vaData As Variant
            
            Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
            Set wsSheet = wbBook.Worksheets(1)
            
            With wsSheet
                Set rnData = .Range(.Range("A1"), .Range("O65536").End(xlUp))
            End With
            
            vaData = rnData.Value
            
            With Me.lstLista
                .ColumnCount = 15
                .Clear
                .List = vaData
                .ListIndex = -1
            End With
            
            wbBook.Close
        End If
            
    End If
End Sub
Private Sub UserForm_Initialize()
    With ComboBoxCampos
        .AddItem "id"
        .AddItem "lancamento"
        .AddItem "codigo"
        .AddItem "familia"
        .AddItem "ncm"
        .AddItem "ESPECIFICA��O"
        .AddItem "tipo"
    End With
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnData As Range
    Dim vaData As Variant
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets(1)
    
    With wsSheet
        Set rnData = .Range(.Range("A1"), .Range("O65536").End(xlUp))
    End With
    
    vaData = rnData.Value
    
    With Me.lstLista
        .ColumnCount = 15
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
    wbBook.Close
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub
Private Sub PreencheLista_Melhorado(TextoDigitado)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    Dim wsSheet As Worksheet
    Dim rnData As Range
    Dim vaData As Variant
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets("BD")
    Set parafilter = wbBook.Worksheets("paraFilter")
    
    wsSheet.Activate
    
    If Me.ComboBoxCampos.ListIndex = 0 Then
        wsSheet.Range("A1").AutoFilter Field:=1, Criteria1:=(TextoDigitado)
    Else
        wsSheet.Range("A1").AutoFilter Field:=(Me.ComboBoxCampos.ListIndex - 1), Criteria1:="*" + TextoDigitado + "*"
    End If
    
    With wsSheet
        Set rnData = .Range(.Range("A1"), .Range("O65536").End(xlUp))
    End With
    rnData.Select
    Selection.Copy
    
    parafilter.Activate
    parafilter.Range("A1").Select
    parafilter.Paste
    
    vaData = Selection.Value
        
    With Me.lstLista
        .ColumnCount = 15
        .Clear
        .List = vaData
        .ListIndex = -1
    End With
    
   wbBook.Close
   
   With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub

