VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} listaDeClientes 
   Caption         =   "Lista de Clientes - Luz & Forma"
   ClientHeight    =   7440
   ClientLeft      =   36
   ClientTop       =   96
   ClientWidth     =   16344
   OleObjectBlob   =   "listaDeClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "listaDeClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LinhaCabecalho As Integer = 1


Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cadastroDeCliente.qualWS.Value = "C:\GitHub\myxlsm\clientes.xlsx"
    cadastroDeCliente.Show
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
        .AddItem "endereco"
        .AddItem "regiao"
        .AddItem "complemento"
        .AddItem "observacao"
    End With
    
    call init_lstLista("clientes.xlsx", me)
        
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub
Private Sub CommandButton1_Click()
    cadastroDeCliente.qualWS.Value = "C:\GitHub\myxlsm\clientes.xlsx"
    cadastroDeCliente.Show
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
            
            Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx", ReadOnly:=True)
            Set ws = wb.Worksheets("BD")
            
            
            On Error GoTo ErrFailed
            irow = lstLista.Value
            
            If irow > 1 Then
                cadastroDeCliente.id.Value = ws.Cells(irow, 1).Value
                cadastroDeCliente.nomeFantasia.Value = ws.Cells(irow, 2).Value
                cadastroDeCliente.cnpj.Value = ws.Cells(irow, 3).Value
                cadastroDeCliente.razaoSocial.Value = ws.Cells(irow, 4).Value
                cadastroDeCliente.atendimento.Value = ws.Cells(irow, 5).Value
                cadastroDeCliente.inscricaoEstadual.Value = ws.Cells(irow, 6).Value
                cadastroDeCliente.clienteDesde.Value = ws.Cells(irow, 7).Value
                
                
                cadastroDeCliente.cep.Value = ws.Cells(irow, 8).Value
                cadastroDeCliente.estado.Value = ws.Cells(irow, 9).Value
                cadastroDeCliente.cidade.Value = ws.Cells(irow, 10).Value
                cadastroDeCliente.bairro.Value = ws.Cells(irow, 11).Value
                cadastroDeCliente.endereco.Value = ws.Cells(irow, 12).Value
                cadastroDeCliente.regiao.Value = ws.Cells(irow, 13).Value
                cadastroDeCliente.complemento.Value = ws.Cells(irow, 14).Value
                
                
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
                
                cadastroDeCliente.qualWS.Value = "C:\GitHub\myxlsm\clientes.xlsx"
                wb.Close
                cadastroDeCliente.Show
                
            End If
            
            With Application
                .ScreenUpdating = True
                .EnableEvents = True
                .Calculation = xlCalculationAutomatic
                .DisplayAlerts = True
            End With
            
ErrFailed:
            ThisWorkbook.Sheets("Menu").Select
    End If
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
    
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx", ReadOnly:=True)
    Set ws = wb.Worksheets("BD")
    
    
    On Error GoTo ErrFailed
    irow = lstLista.Value
    
    If irow > 1 Then
        cadastroDeCliente.id.Value = ws.Cells(irow, 1).Value
        cadastroDeCliente.nomeFantasia.Value = ws.Cells(irow, 2).Value
        cadastroDeCliente.cnpj.Value = ws.Cells(irow, 3).Value
        cadastroDeCliente.razaoSocial.Value = ws.Cells(irow, 4).Value
        cadastroDeCliente.atendimento.Value = ws.Cells(irow, 5).Value
        cadastroDeCliente.inscricaoEstadual.Value = ws.Cells(irow, 6).Value
        cadastroDeCliente.clienteDesde.Value = ws.Cells(irow, 7).Value
        
        
        cadastroDeCliente.cep.Value = ws.Cells(irow, 8).Value
        cadastroDeCliente.estado.Value = ws.Cells(irow, 9).Value
        cadastroDeCliente.cidade.Value = ws.Cells(irow, 10).Value
        cadastroDeCliente.bairro.Value = ws.Cells(irow, 11).Value
        cadastroDeCliente.endereco.Value = ws.Cells(irow, 12).Value
        cadastroDeCliente.regiao.Value = ws.Cells(irow, 13).Value
        cadastroDeCliente.complemento.Value = ws.Cells(irow, 14).Value
        
        
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
        
        cadastroDeCliente.qualWS.Value = "C:\GitHub\myxlsm\clientes.xlsx"
        wb.Close
        cadastroDeCliente.Show
        
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
ErrFailed:
    ThisWorkbook.Sheets("Menu").Select
    
End Sub

Private Sub TextBoxFiltro_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
         If ComboBoxCampos.Value = "" Then
            ComboBoxCampos.Value = "nomeFantasia"
        End If
        
        If Len(TextBoxFiltro.Text) > 0 Then
            Call PreencheLista_Melhorado(TextBoxFiltro.Text, "clientes.xlsx", me)
    
            If Me.lstLista.ListCount > 1 Then
                Me.lstLista.SetFocus
                Me.lstLista.ListIndex = 1
            End If

        Else
            MsgBox ("Aqui")
        End If
    End If
    
End Sub


