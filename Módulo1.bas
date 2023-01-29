Attribute VB_Name = "Modulo1"
Sub abrirMenu()
    menu.Show 
End Sub
Sub abrirForm()
    listaDeClientes.Show
End Sub
Sub abrirFormListaDeProdutos()
    listaDeProdutos.Show
End Sub
Sub abrirGerarOrcamento()
    
End Sub
Sub abrirListaDeOrcamentos()
    ListaDeOrcamentos.Show
End Sub
Sub abrirOrcamentoDesteCliente()
    orcamentoDesteCliente.Show
End Sub
Sub CurUserNames()
    Dim str As String
    str = "Users currently online:" & Chr(10)
    For i = 1 To UBound(ThisWorkbook.UserStatus)
         str = str & ThisWorkbook.UserStatus(i, 1) & ", "
    Next
    Range("F2").Value = Mid(str, 1, Len(str) - 2)
End Sub
Sub linkParaClientes()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim wbAqui As Workbook
    Dim myfilename As String
    
    Set wbAqui = ThisWorkbook
    
    myfilename = "C:\Users\sdant\OneDrive\Desktop\Nova pasta\clientes.xlsm"
    Set wb = Workbooks.Open(myfilename)
    wbAqui.Close savechanges:=True
    Application.DisplayAlerts = True
    
End Sub
Sub test()
    Dim mes As Integer
    Dim mesPorExtenso As String
    
    mes = Format(Date, "mm")
    mesPorExtenso = Switch( _
    mes = 1, "Janeiro", _
    mes = 2, "Fevereiro", _
    mes = 3, "Marco", _
    mes = 4, "Abril", _
    mes = 5, "Maio", _
    mes = 6, "Junho", _
    mes = 7, "Julho", _
    mes = 8, "Agosto", _
    mes = 9, "Setembro", _
    mes = 10, "Outubro", _
    mes = 11, "Novembro", _
    mes = 12, "Dezembro")

    MsgBox (mesPorExtenso)
End Sub















