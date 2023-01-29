VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastroDeProdutos 
   Caption         =   "cadastroDeProdutos"
   ClientHeight    =   6948
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16236
   OleObjectBlob   =   "cadastroDeProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastroDeProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim irow As Integer
        
    Set wb = Workbooks.Open(Filename:="C:\GitHub\myxlsm\produtos.xlsx", ReadOnly:=False)
    Set ws = wb.Sheets("BD")
    
    irow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, _
    SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    
    If id.Value = "" Then
        irow = irow + 1
        id.Value = irow
    Else
        irow = id.Value
    End If
    
    ws.Cells(irow, 1).Value = cadastroDeProdutos.id.Value
    ws.Cells(irow, 2).Value = cadastroDeProdutos.lancamento.Value
    ws.Cells(irow, 3).Value = cadastroDeProdutos.codigo.Value
    ws.Cells(irow, 4).Value = cadastroDeProdutos.familia.Value
    ws.Cells(irow, 5).Value = cadastroDeProdutos.ncm.Value
    
    ws.Cells(irow, 6).Value = cadastroDeProdutos.especificacao1.Value
    ws.Cells(irow, 7).Value = cadastroDeProdutos.especificacao2.Value
    ws.Cells(irow, 8).Value = cadastroDeProdutos.especificacao3.Value
    
    ws.Cells(irow, 9).Value = cadastroDeProdutos.tipo.Value
    ws.Cells(irow, 10).Value = cadastroDeProdutos.altura.Value
    ws.Cells(irow, 11).Value = cadastroDeProdutos.largura.Value
    ws.Cells(irow, 12).Value = cadastroDeProdutos.compProf.Value
    ws.Cells(irow, 13).Value = cadastroDeProdutos.potencia.Value
    
    ws.Cells(irow, 14).Value = cadastroDeProdutos.mtCorda.Value
    ws.Cells(irow, 15).Value = cadastroDeProdutos.peso.Value
    
    ws.Cells(irow, 16).Value = cadastroDeProdutos.desc_anexo1.Value
    ws.Cells(irow, 17).Value = cadastroDeProdutos.anexo1.Value
    
    ws.Cells(irow, 18).Value = cadastroDeProdutos.desc_anexo2.Value
    ws.Cells(irow, 19).Value = cadastroDeProdutos.anexo2.Value
    
    ws.Cells(irow, 20).Value = cadastroDeProdutos.desc_anexo3.Value
    ws.Cells(irow, 21).Value = cadastroDeProdutos.anexo3.Value
    
    ws.Cells(irow, 22).Value = cadastroDeProdutos.desc_anexo4.Value
    ws.Cells(irow, 23).Value = cadastroDeProdutos.anexo4.Value
    
    ws.Cells(irow, 24).Value = cadastroDeProdutos.desc_anexo5.Value
    ws.Cells(irow, 25).Value = cadastroDeProdutos.anexo5.Value
    
    ws.Cells(irow, 26).Value = cadastroDeProdutos.desc_anexo6.Value
    ws.Cells(irow, 27).Value = cadastroDeProdutos.anexo6.Value
    
    ws.Cells(irow, 28).Value = cadastroDeProdutos.desc_anexo7.Value
    ws.Cells(irow, 29).Value = cadastroDeProdutos.anexo7.Value
    
    ws.Cells(irow, 30).Value = cadastroDeProdutos.desc_anexo8.Value
    ws.Cells(irow, 31).Value = cadastroDeProdutos.anexo8.Value
    
    ws.Cells(irow, 32).Value = cadastroDeProdutos.desc_anexo9.Value
    ws.Cells(irow, 33).Value = cadastroDeProdutos.anexo9.Value
    
    ws.Cells(irow, 34).Value = cadastroDeProdutos.desc_anexo10.Value
    ws.Cells(irow, 35).Value = cadastroDeProdutos.anexo10.Value
    
    ws.Cells(irow, 36).Value = cadastroDeProdutos.precoDeVenda.Value
    ws.Cells(irow, 37).Value = cadastroDeProdutos.precoDeLocacao.Value
    
    wb.Close True
    
    Unload Me
    Unload listaDeProdutos
    listaDeProdutos.Show
    
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo1.Value)
End Sub

Private Sub CommandButton5_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo2.Value)
End Sub

Private Sub CommandButton6_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo3.Value)
End Sub

Private Sub CommandButton7_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo4.Value)
End Sub

Private Sub CommandButton8_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo5.Value)
End Sub

Private Sub CommandButton9_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo6.Value)
End Sub

Private Sub CommandButton10_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo7.Value)
End Sub

Private Sub CommandButton11_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo8.Value)
End Sub

Private Sub CommandButton12_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo9.Value)
End Sub

Private Sub CommandButton13_Click()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink (cadastroDeProdutos.anexo10.Value)
End Sub





