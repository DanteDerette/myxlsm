VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GerarOrcamento 
   Caption         =   "GerarOrcamento"
   ClientHeight    =   6792
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9552.001
   OleObjectBlob   =   "GerarOrcamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GerarOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
   With ComboBoxCampos
        .AddItem "id"
        .AddItem "nomeFantasia"
        .AddItem "cnpj"
        .AddItem "razaoSocial"
        .AddItem "atendimento"
        .AddItem "inscricaoEstadual"
        .AddItem "clienteDesde"
        .AddItem "cep"
        .AddItem "estado"
        .AddItem "cidade"
        .AddItem "bairro"
        .AddItem "endereço"
        .AddItem "regiao"
        .AddItem "complemento"
        .AddItem "observação"
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
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx", ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets(1)
    
    With wsSheet
        Set rnData = .Range(.Range("A1"), .Range("C65536").End(xlUp))
    End With
    
    vaData = rnData.Value
    
    With Me.lstLista
        .ColumnCount = 3
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
Private Sub lstLista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
   
    Dim estaplanilha As Workbook
    Dim outraplanilha As Workbook
    Dim novaplan As Workbook
    Dim nomeDoCliente As String
    
    Dim lastrow As Integer
    Dim irow As Integer
    Dim bd_estaplanilha As Worksheet
    Dim bd_outraplanilha As Worksheet
    Dim bd_novaplan As Worksheet
    
    minharesposta = InputBox("Título do Orçamento: ")
    
    If StrPtr(minharesposta) = 0 Then
        Set wb = ThisWorkbook
    ElseIf minharesposta = vbNullString Then
        Set wb = ThisWorkbook
    Else
        Set estaplanilha = Workbooks.Open(Filename:="C:\GitHub\myxlsm\Clientes.xlsx", ReadOnly:=True)
        Set outraplanilha = Workbooks.Open(Filename:="C:\GitHub\myxlsm\orcamentos.xlsx", ReadOnly:=False)
        
        Set bd_estaplanilha = estaplanilha.Worksheets("BD")
        bd_estaplanilha.Visible = xlSheetVisible
        Set bd_outraplanilha = outraplanilha.Worksheets("BD")
        
  '     On Error GoTo errorHandler
        irow = lstLista.Value
        
        lastrow = bd_outraplanilha.Range("A" & bd_outraplanilha.Rows.Count).End(xlUp).Row + 1
        
        nomeDoCliente = bd_estaplanilha.Range("B" & irow).Value
        
        bd_outraplanilha.Cells(lastrow, 1).Value = lastrow
        bd_outraplanilha.Cells(lastrow, 2).Value = minharesposta
        bd_outraplanilha.Cells(lastrow, 3).Value = bd_estaplanilha.Range("A" & irow).Value
        bd_outraplanilha.Cells(lastrow, 4).Value = nomeDoCliente
        bd_outraplanilha.Cells(lastrow, 5).Value = Format(Date, "dd/mm/yyyy")
        bd_outraplanilha.Cells(lastrow, 6).Value = "C:\GitHub\myxlsm\orcamentos\" _
        & nomeDoCliente & "_" & minharesposta & ".xlsx"
           
        Set novaplan = Workbooks.Add("C:\GitHub\myxlsm\template_orcamento.xlsx")
        novaplan.SaveAs ("C:\GitHub\myxlsm\orcamentos\" & nomeDoCliente & "_" & minharesposta & ".xlsx")
        Set bd_novaplan = novaplan.Worksheets("geral")
        
        bd_novaplan.Range("A2").Value = lastrow
        bd_novaplan.Range("B2").Value = minharesposta
        bd_novaplan.Range("C2").Value = bd_estaplanilha.Range("A" & irow).Value
        bd_novaplan.Range("D2").Value = nomeDoCliente
        bd_novaplan.Range("E2").Value = Format(Date, "dd/mm/yyyy")
        bd_novaplan.Range("F2").Value = "C:\GitHub\myxlsm\orcamentos\" _
        & nomeDoCliente & "_" & minharesposta & ".xlsx"
               
        novaplan.Close True
        outraplanilha.Close True
        estaplanilha.Close
        
        Unload Me
        ListaDeOrcamentos.Show
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
            ComboBoxCampos.Value = "nomeFantasia"
        End If
        
        If Len(TextBoxFiltro.Text) > 0 Then
            Call PreencheLista_Melhorado(TextBoxFiltro.Text)
        End If
    End If
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
    Dim lastrow As Integer
    
    Set wbBook = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx", ReadOnly:=True)
    Set wsSheet = wbBook.Worksheets("BD")
    Set parafilter = wbBook.Worksheets("paraFilter")
    
    wsSheet.Activate
    
    If Me.ComboBoxCampos.ListIndex = 0 Then
        wsSheet.Range("A1").AutoFilter Field:=1, Criteria1:=(TextoDigitado)
    Else
        wsSheet.Range("A1").AutoFilter Field:=(Me.ComboBoxCampos.ListIndex + 1), Criteria1:="*" + TextoDigitado + "*"
    End If
    
    lastrow = wsSheet.Range("A" & wsSheet.Rows.Count).End(xlUp).Row
    Set rnData = wsSheet.Range(wsSheet.Range("A1"), wsSheet.Range("AC" & lastrow))
    
    rnData.Select
    Selection.Copy
    
    parafilter.Activate
    parafilter.Range("A1").Select
    parafilter.Paste
    
    vaData = Selection.Value
        
    With Me.lstLista
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


