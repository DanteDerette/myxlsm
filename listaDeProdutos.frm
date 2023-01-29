VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} listaDeProdutos 
   Caption         =   "listaDeProdutos"
   ClientHeight    =   7656
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   16632
   OleObjectBlob   =   "listaDeProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "listaDeProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBoxCampos_Change()

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
        Dim found As Range
        Dim irow As Integer
        
        Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
        Set ws = wb.Worksheets("BD")
        
        On Error GoTo ErrFailed
        irow = lstLista.Value
        
        If irow > 1 Then
            cadastroDeProdutos.id.Value = ws.Cells(irow, 1).Value
            cadastroDeProdutos.lancamento.Value = ws.Cells(irow, 2).Value
            cadastroDeProdutos.codigo.Value = ws.Cells(irow, 3).Value
            cadastroDeProdutos.familia.Value = ws.Cells(irow, 4).Value
            cadastroDeProdutos.ncm.Value = ws.Cells(irow, 5).Value
            
            cadastroDeProdutos.especificacao1.Value = ws.Cells(irow, 6).Value
            cadastroDeProdutos.especificacao2.Value = ws.Cells(irow, 7).Value
            cadastroDeProdutos.especificacao3.Value = ws.Cells(irow, 8).Value
            
            cadastroDeProdutos.tipo.Value = ws.Cells(irow, 9).Value
            cadastroDeProdutos.altura.Value = ws.Cells(irow, 10).Value
            cadastroDeProdutos.largura.Value = ws.Cells(irow, 11).Value
            cadastroDeProdutos.compProf.Value = ws.Cells(irow, 12).Value
            cadastroDeProdutos.potencia.Value = ws.Cells(irow, 13).Value
            cadastroDeProdutos.mtCorda.Value = ws.Cells(irow, 14).Value
            cadastroDeProdutos.peso.Value = ws.Cells(irow, 15).Value
            
            cadastroDeProdutos.desc_anexo1.Value = ws.Cells(irow, 16).Value
            cadastroDeProdutos.anexo1.Value = ws.Cells(irow, 17).Value
            
            cadastroDeProdutos.desc_anexo2.Value = ws.Cells(irow, 18).Value
            cadastroDeProdutos.anexo2.Value = ws.Cells(irow, 19).Value
            
            cadastroDeProdutos.desc_anexo3.Value = ws.Cells(irow, 20).Value
            cadastroDeProdutos.anexo3.Value = ws.Cells(irow, 21).Value
            
            cadastroDeProdutos.desc_anexo4.Value = ws.Cells(irow, 22).Value
            cadastroDeProdutos.anexo4.Value = ws.Cells(irow, 23).Value
            
            cadastroDeProdutos.desc_anexo5.Value = ws.Cells(irow, 24).Value
            cadastroDeProdutos.anexo5.Value = ws.Cells(irow, 25).Value
            
            cadastroDeProdutos.desc_anexo6.Value = ws.Cells(irow, 26).Value
            cadastroDeProdutos.anexo6.Value = ws.Cells(irow, 27).Value
            
            cadastroDeProdutos.desc_anexo7.Value = ws.Cells(irow, 28).Value
            cadastroDeProdutos.anexo7.Value = ws.Cells(irow, 29).Value
            
            cadastroDeProdutos.desc_anexo8.Value = ws.Cells(irow, 30).Value
            cadastroDeProdutos.anexo8.Value = ws.Cells(irow, 31).Value
            
            cadastroDeProdutos.desc_anexo9.Value = ws.Cells(irow, 32).Value
            cadastroDeProdutos.anexo9.Value = ws.Cells(irow, 33).Value
            
            cadastroDeProdutos.desc_anexo10.Value = ws.Cells(irow, 34).Value
            cadastroDeProdutos.anexo10.Value = ws.Cells(irow, 35).Value
            
            cadastroDeProdutos.precoDeVenda.Value = ws.Cells(irow, 36).Value
            cadastroDeProdutos.precoDeLocacao.Value = ws.Cells(irow, 37).Value
            wb.Close
            cadastroDeProdutos.Show
        End If
        
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
        End With
    
ErrFailed:
       ThisWorkbook.Worksheets("Menu").Activate
    End If
 
End Sub

Private Sub TextBoxFiltro_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
         If ComboBoxCampos.Value = "" Then
            ComboBoxCampos.Value = "codigo"
        End If
        
        If Len(TextBoxFiltro.Text) > 0 Then
            Call PreencheLista_Melhorado(TextBoxFiltro.Text, "produtos.xlsx", me)
            
            If Me.lstLista.ListCount > 1 Then
                Me.lstLista.SetFocus
                Me.lstLista.ListIndex = 1
            Else
                Me.TextBoxFiltro.SetFocus
            End If
            
        Else
            Dim wbBook As Workbook
            Dim wsSheet As Worksheet
            Dim rnData As Range
            Dim vaData As Variant
            
            Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
            Set wsSheet = wbBook.Worksheets(1)
            
            lastrow = wsSheet.Range("A" & wsSheet.Rows.Count).End(xlUp).Row
            Set rnData = wsSheet.Range(ws.Range("A1"), wsSheet.Range("AC" & lastrow))

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
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With

    With ComboBoxCampos
        .AddItem "id"
        .AddItem "lancamento"
        .AddItem "codigo"
        .AddItem "famolia"
        .AddItem "ncm"
        .AddItem "ESPECIFICACAO"
        .AddItem "tipo"
    End With
    
    call init_lstLista("produtos.xlsx", me)
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub
Private Sub CommandButton1_Click()
    cadastroDeProdutos.Show
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
    Dim found As Range
    Dim irow As Integer
    
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
    Set ws = wb.Worksheets("BD")
    
    On Error GoTo ErrFailed
    irow = lstLista.Value
    
    
        
    
    If irow > 1 Then
        cadastroDeProdutos.id.Value = ws.Cells(irow, 1).Value
        cadastroDeProdutos.lancamento.Value = ws.Cells(irow, 2).Value
        cadastroDeProdutos.codigo.Value = ws.Cells(irow, 3).Value
        cadastroDeProdutos.familia.Value = ws.Cells(irow, 4).Value
        cadastroDeProdutos.ncm.Value = ws.Cells(irow, 5).Value
        
        cadastroDeProdutos.especificacao1.Value = ws.Cells(irow, 6).Value
        cadastroDeProdutos.especificacao2.Value = ws.Cells(irow, 7).Value
        cadastroDeProdutos.especificacao3.Value = ws.Cells(irow, 8).Value
        
        cadastroDeProdutos.tipo.Value = ws.Cells(irow, 9).Value
        cadastroDeProdutos.altura.Value = ws.Cells(irow, 10).Value
        cadastroDeProdutos.largura.Value = ws.Cells(irow, 11).Value
        cadastroDeProdutos.compProf.Value = ws.Cells(irow, 12).Value
        cadastroDeProdutos.potencia.Value = ws.Cells(irow, 13).Value
        cadastroDeProdutos.mtCorda.Value = ws.Cells(irow, 14).Value
        cadastroDeProdutos.peso.Value = ws.Cells(irow, 15).Value
        
        cadastroDeProdutos.desc_anexo1.Value = ws.Cells(irow, 16).Value
        cadastroDeProdutos.anexo1.Value = ws.Cells(irow, 17).Value
        
        cadastroDeProdutos.desc_anexo2.Value = ws.Cells(irow, 18).Value
        cadastroDeProdutos.anexo2.Value = ws.Cells(irow, 19).Value
        
        cadastroDeProdutos.desc_anexo3.Value = ws.Cells(irow, 20).Value
        cadastroDeProdutos.anexo3.Value = ws.Cells(irow, 21).Value
        
        cadastroDeProdutos.desc_anexo4.Value = ws.Cells(irow, 22).Value
        cadastroDeProdutos.anexo4.Value = ws.Cells(irow, 23).Value
        
        cadastroDeProdutos.desc_anexo5.Value = ws.Cells(irow, 24).Value
        cadastroDeProdutos.anexo5.Value = ws.Cells(irow, 25).Value
        
        cadastroDeProdutos.desc_anexo6.Value = ws.Cells(irow, 26).Value
        cadastroDeProdutos.anexo6.Value = ws.Cells(irow, 27).Value
        
        cadastroDeProdutos.desc_anexo7.Value = ws.Cells(irow, 28).Value
        cadastroDeProdutos.anexo7.Value = ws.Cells(irow, 29).Value
        
        cadastroDeProdutos.desc_anexo8.Value = ws.Cells(irow, 30).Value
        cadastroDeProdutos.anexo8.Value = ws.Cells(irow, 31).Value
        
        cadastroDeProdutos.desc_anexo9.Value = ws.Cells(irow, 32).Value
        cadastroDeProdutos.anexo9.Value = ws.Cells(irow, 33).Value
        
        cadastroDeProdutos.desc_anexo10.Value = ws.Cells(irow, 34).Value
        cadastroDeProdutos.anexo10.Value = ws.Cells(irow, 35).Value
        
        cadastroDeProdutos.precoDeVenda.Value = ws.Cells(irow, 36).Value
        cadastroDeProdutos.precoDeLocacao.Value = ws.Cells(irow, 37).Value
        wb.Close
        cadastroDeProdutos.Show
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With

ErrFailed:
   ThisWorkbook.Worksheets("Menu").Activate
End Sub

' Private Sub PreencheLista_Melhorado2(TextoDigitado)
'     With Application
'         .ScreenUpdating = False
'         .EnableEvents = False
'         .Calculation = xlCalculationManual
'         .DisplayAlerts = False
'     End With
    
'     Dim wsSheet As Worksheet
'     Dim rnData As Range
'     Dim vaData As Variant
'     Dim lastrow As Integer
    
    
'     Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=True)
'     Set wsSheet = wbBook.Worksheets("BD")
'     Set parafilter = wbBook.Worksheets("paraFilter")
    
'     wsSheet.Activate
    
'     If Me.ComboBoxCampos.ListIndex = 0 Then
'         wsSheet.Range("A1").AutoFilter Field:=1, Criteria1:=(TextoDigitado)
'     Else
'         wsSheet.Range("A1").AutoFilter Field:=(Me.ComboBoxCampos.ListIndex - 1), Criteria1:="*" + TextoDigitado + "*"
'     End If
    
'     lastrow = wsSheet.Range("A" & wsSheet.Rows.Count).End(xlUp).Row
'     Set rnData = wsSheet.Range(wsSheet.Range("A1"), wsSheet.Range("AC" & lastrow))
' '    With wsSheet
' '        Set rnData = .Range(.Range("A1"), .Range("O65536").End(xlUp))
' '    End With
'     rnData.Select
'     Selection.Copy
    
'     parafilter.Activate
'     parafilter.Range("A1").Select
'     parafilter.Paste
    
'     vaData = Selection.Value
        
'     With Me.lstLista
'         .ColumnCount = 15
'         .Clear
'         .List = vaData
'         .ListIndex = -1
'     End With
    
'    wbBook.Close
   
'    With Application
'         .ScreenUpdating = True
'         .EnableEvents = True
'         .Calculation = xlCalculationAutomatic
'         .DisplayAlerts = True
'     End With
' End Sub
