VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastroDeCliente 
   Caption         =   "cadastroDeCliente"
   ClientHeight    =   9600.001
   ClientLeft      =   36
   ClientTop       =   96
   ClientWidth     =   19980
   OleObjectBlob   =   "cadastroDeCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastroDeCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo1.Value)
End Sub

Private Sub CommandButton5_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo2.Value)
End Sub

Private Sub CommandButton6_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo3.Value)
End Sub

Private Sub CommandButton7_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo4.Value)
End Sub

Private Sub CommandButton8_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo5.Value)
End Sub

Private Sub CommandButton9_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo6.Value)
End Sub

Private Sub CommandButton10_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo7.Value)
End Sub

Private Sub CommandButton11_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo8.Value)
End Sub

Private Sub CommandButton12_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo9.Value)
End Sub

Private Sub CommandButton13_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeCliente.anexo10.Value)
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim irow As Integer
    
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\clientes.xlsx")
    Set ws = Sheets("BD")
    
    irow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, _
    SearchDirection:=xlPrevious, LookIn:=xlValues).Row
       
    If id.Value = "" Then
        irow = irow + 1
        id.Value = irow
    Else
        irow = id.Value
    End If
        
    ' Parte Superior
    ws.Cells(irow, 1).Value = id.Value
    ws.Cells(irow, 2).Value = nomeFantasia.Value
    ws.Cells(irow, 3).Value = cnpj.Value
    ws.Cells(irow, 4).Value = razaoSocial.Value
    ws.Cells(irow, 5).Value = atendimento.Value
    ws.Cells(irow, 6).Value = inscricaoEstadual.Value
    ws.Cells(irow, 7).Value = clienteDesde.Value

    ' Aba Endereço
    ws.Cells(irow, 8).Value = cep.Value
    ws.Cells(irow, 9).Value = estado.Value
    ws.Cells(irow, 10).Value = cidade.Value
    ws.Cells(irow, 11).Value = bairro.Value
    ws.Cells(irow, 12).Value = endereco.Value
    ws.Cells(irow, 13).Value = regiao.Value
    ws.Cells(irow, 14).Value = complemento.Value

    'Aba Observação
    ws.Cells(irow, 15).Value = observacao.Value

    'Contatos
    '''''''''''''''''''''''''''''''''''Contatos1'''''''''''''''''''''''''''''''
    ws.Cells(irow, 16).Value = cidade_contato1.Value
    'Comercial
    ws.Cells(irow, 17).Value = comercial_nome_contato1.Value
    ws.Cells(irow, 18).Value = comercial_cargo_contato1.Value
    ws.Cells(irow, 19).Value = comercial_telefone1_contato1.Value
    ws.Cells(irow, 20).Value = comercial_email1_contato1.Value
    ws.Cells(irow, 21).Value = comercial_telefone2_contato1.Value
    ws.Cells(irow, 22).Value = comercial_email2_contato1.Value
    'Financeiro
    ws.Cells(irow, 23).Value = financeiro_nome_contato1.Value
    ws.Cells(irow, 24).Value = financeiro_cargo_contato1.Value
    ws.Cells(irow, 25).Value = financeiro_telefone1_contato1.Value
    ws.Cells(irow, 26).Value = financeiro_email1_contato1.Value
    ws.Cells(irow, 27).Value = financeiro_telefone2_contato1.Value
    ws.Cells(irow, 28).Value = financeiro_email2_contato1.Value
    'Observação
    ws.Cells(irow, 29).Value = observacaoDoContato_contato1.Value
    
    '''''''''''''''''''''''''''''''''''Contatos2'''''''''''''''''''''''''''''''
    ws.Cells(irow, 30).Value = cidade_contato2.Value
    'Comercial
    ws.Cells(irow, 31).Value = comercial_nome_contato2.Value
    ws.Cells(irow, 32).Value = comercial_cargo_contato2.Value
    ws.Cells(irow, 33).Value = comercial_telefone1_contato2.Value
    ws.Cells(irow, 34).Value = comercial_email1_contato2.Value
    ws.Cells(irow, 35).Value = comercial_telefone2_contato2.Value
    ws.Cells(irow, 36).Value = comercial_email2_contato2.Value
    'Financeiro
    ws.Cells(irow, 37).Value = financeiro_nome_contato2.Value
    ws.Cells(irow, 38).Value = financeiro_cargo_contato2.Value
    ws.Cells(irow, 39).Value = financeiro_telefone1_contato2.Value
    ws.Cells(irow, 40).Value = financeiro_email1_contato2.Value
    ws.Cells(irow, 41).Value = financeiro_telefone2_contato2.Value
    ws.Cells(irow, 42).Value = financeiro_email2_contato2.Value
    'Observação
    ws.Cells(irow, 43).Value = observacaoDoContato_contato2.Value
    
    '''''''''''''''''''''''''''''''Contatos3''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 44).Value = cidade_contato3.Value
    'Comercial
    ws.Cells(irow, 45).Value = comercial_nome_contato3.Value
    ws.Cells(irow, 46).Value = comercial_cargo_contato3.Value
    ws.Cells(irow, 47).Value = comercial_telefone1_contato3.Value
    ws.Cells(irow, 48).Value = comercial_email1_contato3.Value
    ws.Cells(irow, 49).Value = comercial_telefone2_contato3.Value
    ws.Cells(irow, 50).Value = comercial_email2_contato3.Value
    'Financeiro
    ws.Cells(irow, 51).Value = financeiro_nome_contato3.Value
    ws.Cells(irow, 52).Value = financeiro_cargo_contato3.Value
    ws.Cells(irow, 53).Value = financeiro_telefone1_contato3.Value
    ws.Cells(irow, 54).Value = financeiro_email1_contato3.Value
    ws.Cells(irow, 55).Value = financeiro_telefone2_contato3.Value
    ws.Cells(irow, 56).Value = financeiro_email2_contato3.Value
    'Observação
    ws.Cells(irow, 57).Value = observacaoDoContato_contato3.Value
    
    '''''''''''''''''''''''''''''''Contatos4''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 58).Value = cidade_contato4.Value
    'Comercial
    ws.Cells(irow, 59).Value = comercial_nome_contato4.Value
    ws.Cells(irow, 60).Value = comercial_cargo_contato4.Value
    ws.Cells(irow, 61).Value = comercial_telefone1_contato4.Value
    ws.Cells(irow, 62).Value = comercial_email1_contato4.Value
    ws.Cells(irow, 63).Value = comercial_telefone2_contato4.Value
    ws.Cells(irow, 64).Value = comercial_email2_contato4.Value
    'Financeiro
    ws.Cells(irow, 65).Value = financeiro_nome_contato4.Value
    ws.Cells(irow, 66).Value = financeiro_cargo_contato4.Value
    ws.Cells(irow, 67).Value = financeiro_telefone1_contato4.Value
    ws.Cells(irow, 68).Value = financeiro_email1_contato4.Value
    ws.Cells(irow, 69).Value = financeiro_telefone2_contato4.Value
    ws.Cells(irow, 70).Value = financeiro_email2_contato4.Value
    'Observação
    ws.Cells(irow, 71).Value = observacaoDoContato_contato4.Value
    
    '''''''''''''''''''''''''''''''Contatos5''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 72).Value = cidade_contato5.Value
    'Comercial
    ws.Cells(irow, 73).Value = comercial_nome_contato5.Value
    ws.Cells(irow, 74).Value = comercial_cargo_contato5.Value
    ws.Cells(irow, 75).Value = comercial_telefone1_contato5.Value
    ws.Cells(irow, 76).Value = comercial_email1_contato5.Value
    ws.Cells(irow, 77).Value = comercial_telefone2_contato5.Value
    ws.Cells(irow, 78).Value = comercial_email2_contato5.Value
    'Financeiro
    ws.Cells(irow, 79).Value = financeiro_nome_contato5.Value
    ws.Cells(irow, 80).Value = financeiro_cargo_contato5.Value
    ws.Cells(irow, 81).Value = financeiro_telefone1_contato5.Value
    ws.Cells(irow, 82).Value = financeiro_email1_contato5.Value
    ws.Cells(irow, 83).Value = financeiro_telefone2_contato5.Value
    ws.Cells(irow, 84).Value = financeiro_email2_contato5.Value
    'Observação
    ws.Cells(irow, 85).Value = observacaoDoContato_contato5.Value

    '''''''''''''''''''''''''''''''Contatos5''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 72).Value = cidade_contato5.Value
    'Comercial
    ws.Cells(irow, 73).Value = comercial_nome_contato5.Value
    ws.Cells(irow, 74).Value = comercial_cargo_contato5.Value
    ws.Cells(irow, 75).Value = comercial_telefone1_contato5.Value
    ws.Cells(irow, 76).Value = comercial_email1_contato5.Value
    ws.Cells(irow, 77).Value = comercial_telefone2_contato5.Value
    ws.Cells(irow, 78).Value = comercial_email2_contato5.Value
    'Financeiro
    ws.Cells(irow, 79).Value = financeiro_nome_contato5.Value
    ws.Cells(irow, 80).Value = financeiro_cargo_contato5.Value
    ws.Cells(irow, 81).Value = financeiro_telefone1_contato5.Value
    ws.Cells(irow, 82).Value = financeiro_email1_contato5.Value
    ws.Cells(irow, 83).Value = financeiro_telefone2_contato5.Value
    ws.Cells(irow, 84).Value = financeiro_email2_contato5.Value
    'Observação
    ws.Cells(irow, 85).Value = observacaoDoContato_contato5.Value
    
    '''''''''''''''''''''''''''''''Contatos6''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 86).Value = cidade_contato6.Value
    'Comercial
    ws.Cells(irow, 87).Value = comercial_nome_contato6.Value
    ws.Cells(irow, 88).Value = comercial_cargo_contato6.Value
    ws.Cells(irow, 89).Value = comercial_telefone1_contato6.Value
    ws.Cells(irow, 90).Value = comercial_email1_contato6.Value
    ws.Cells(irow, 91).Value = comercial_telefone2_contato6.Value
    ws.Cells(irow, 92).Value = comercial_email2_contato6.Value
    'Financeiro
    ws.Cells(irow, 93).Value = financeiro_nome_contato6.Value
    ws.Cells(irow, 94).Value = financeiro_cargo_contato6.Value
    ws.Cells(irow, 95).Value = financeiro_telefone1_contato6.Value
    ws.Cells(irow, 96).Value = financeiro_email1_contato6.Value
    ws.Cells(irow, 97).Value = financeiro_telefone2_contato6.Value
    ws.Cells(irow, 98).Value = financeiro_email2_contato6.Value
    'Observação
    ws.Cells(irow, 99).Value = observacaoDoContato_contato6.Value
    
    '''''''''''''''''''''''''''''''Contatos7''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 100).Value = cidade_contato7.Value
    'Comercial
    ws.Cells(irow, 101).Value = comercial_nome_contato7.Value
    ws.Cells(irow, 102).Value = comercial_cargo_contato7.Value
    ws.Cells(irow, 103).Value = comercial_telefone1_contato7.Value
    ws.Cells(irow, 104).Value = comercial_email1_contato7.Value
    ws.Cells(irow, 105).Value = comercial_telefone2_contato7.Value
    ws.Cells(irow, 106).Value = comercial_email2_contato7.Value
    'Financeiro
    ws.Cells(irow, 107).Value = financeiro_nome_contato7.Value
    ws.Cells(irow, 108).Value = financeiro_cargo_contato7.Value
    ws.Cells(irow, 109).Value = financeiro_telefone1_contato7.Value
    ws.Cells(irow, 110).Value = financeiro_email1_contato7.Value
    ws.Cells(irow, 111).Value = financeiro_telefone2_contato7.Value
    ws.Cells(irow, 112).Value = financeiro_email2_contato7.Value
    'Observação
    ws.Cells(irow, 113).Value = observacaoDoContato_contato7.Value
    
    '''''''''''''''''''''''''''''''Contatos8''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 114).Value = cidade_contato8.Value
    'Comercial
    ws.Cells(irow, 115).Value = comercial_nome_contato8.Value
    ws.Cells(irow, 116).Value = comercial_cargo_contato8.Value
    ws.Cells(irow, 117).Value = comercial_telefone1_contato8.Value
    ws.Cells(irow, 118).Value = comercial_email1_contato8.Value
    ws.Cells(irow, 119).Value = comercial_telefone2_contato8.Value
    ws.Cells(irow, 120).Value = comercial_email2_contato8.Value
    'Financeiro
    ws.Cells(irow, 121).Value = financeiro_nome_contato8.Value
    ws.Cells(irow, 122).Value = financeiro_cargo_contato8.Value
    ws.Cells(irow, 123).Value = financeiro_telefone1_contato8.Value
    ws.Cells(irow, 124).Value = financeiro_email1_contato8.Value
    ws.Cells(irow, 125).Value = financeiro_telefone2_contato8.Value
    ws.Cells(irow, 126).Value = financeiro_email2_contato8.Value
    'Observação
    ws.Cells(irow, 127).Value = observacaoDoContato_contato8.Value
    
    '''''''''''''''''''''''''''''''Contatos9''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 128).Value = cidade_contato9.Value
    'Comercial
    ws.Cells(irow, 129).Value = comercial_nome_contato9.Value
    ws.Cells(irow, 130).Value = comercial_cargo_contato9.Value
    ws.Cells(irow, 131).Value = comercial_telefone1_contato9.Value
    ws.Cells(irow, 132).Value = comercial_email1_contato9.Value
    ws.Cells(irow, 133).Value = comercial_telefone2_contato9.Value
    ws.Cells(irow, 134).Value = comercial_email2_contato9.Value
    'Financeiro
    ws.Cells(irow, 135).Value = financeiro_nome_contato9.Value
    ws.Cells(irow, 136).Value = financeiro_cargo_contato9.Value
    ws.Cells(irow, 137).Value = financeiro_telefone1_contato9.Value
    ws.Cells(irow, 138).Value = financeiro_email1_contato9.Value
    ws.Cells(irow, 139).Value = financeiro_telefone2_contato9.Value
    ws.Cells(irow, 140).Value = financeiro_email2_contato9.Value
    'Observação
    ws.Cells(irow, 141).Value = observacaoDoContato_contato9.Value
    
    '''''''''''''''''''''''''''''''Contatos10''''''''''''''''''''''''''''''''''''
    ws.Cells(irow, 142).Value = cidade_contato10.Value
    'Comercial
    ws.Cells(irow, 143).Value = comercial_nome_contato10.Value
    ws.Cells(irow, 144).Value = comercial_cargo_contato10.Value
    ws.Cells(irow, 145).Value = comercial_telefone1_contato10.Value
    ws.Cells(irow, 146).Value = comercial_email1_contato10.Value
    ws.Cells(irow, 147).Value = comercial_telefone2_contato10.Value
    ws.Cells(irow, 148).Value = comercial_email2_contato10.Value
    'Financeiro
    ws.Cells(irow, 149).Value = financeiro_nome_contato10.Value
    ws.Cells(irow, 150).Value = financeiro_cargo_contato10.Value
    ws.Cells(irow, 151).Value = financeiro_telefone1_contato10.Value
    ws.Cells(irow, 152).Value = financeiro_email1_contato10.Value
    ws.Cells(irow, 153).Value = financeiro_telefone2_contato10.Value
    ws.Cells(irow, 154).Value = financeiro_email2_contato10.Value
    'Observação
    ws.Cells(irow, 155).Value = observacaoDoContato_contato10.Value
    
    'ws.Cells(irow, 156).Value = cadastroDeCliente.ultimaAtualização.Value
    
    ws.Cells(irow, 157).Value = cadastroDeCliente.desc_anexo1.Value
    ws.Cells(irow, 158).Value = cadastroDeCliente.anexo1.Value
    
    ws.Cells(irow, 159).Value = cadastroDeCliente.desc_anexo2.Value
    ws.Cells(irow, 160).Value = cadastroDeCliente.anexo2.Value
    
    ws.Cells(irow, 161).Value = cadastroDeCliente.desc_anexo3.Value
    ws.Cells(irow, 162).Value = cadastroDeCliente.anexo3.Value
    
    ws.Cells(irow, 163).Value = cadastroDeCliente.desc_anexo4.Value
    ws.Cells(irow, 164).Value = cadastroDeCliente.anexo4.Value
    
    ws.Cells(irow, 165).Value = cadastroDeCliente.desc_anexo5.Value
    ws.Cells(irow, 166).Value = cadastroDeCliente.anexo5.Value
    
    ws.Cells(irow, 167).Value = cadastroDeCliente.desc_anexo6.Value
    ws.Cells(irow, 168).Value = cadastroDeCliente.anexo6.Value
    
    ws.Cells(irow, 169).Value = cadastroDeCliente.desc_anexo7.Value
    ws.Cells(irow, 170).Value = cadastroDeCliente.anexo7.Value
    
    ws.Cells(irow, 171).Value = cadastroDeCliente.desc_anexo8.Value
    ws.Cells(irow, 172).Value = cadastroDeCliente.anexo8.Value
    
    ws.Cells(irow, 173).Value = cadastroDeCliente.desc_anexo9.Value
    ws.Cells(irow, 174).Value = cadastroDeCliente.anexo9.Value
    
    ws.Cells(irow, 175).Value = cadastroDeCliente.desc_anexo10.Value
    ws.Cells(irow, 176).Value = cadastroDeCliente.anexo10.Value
    
    wb.Close True
        
    Unload Me

    If IsLoaded("listaDeClientes") Then
        Unload listaDeClientes
        listaDeClientes.Show
    End If
    
End Sub

Private Sub MultiPage1_Change()

End Sub
