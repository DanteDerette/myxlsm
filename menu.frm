VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menu 
   Caption         =   "Menu"
   ClientHeight    =   3504
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10416
   OleObjectBlob   =   "menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    ListaDeOrcamentos.Show
End Sub
Private Sub CommandButton2_Click()

    listaDeClientes.Show
End Sub
Private Sub CommandButton3_Click()

    listaDeProdutos.Show
End Sub

Private Sub CommandButton7_Click()
 Call show_menu
End Sub

Private Sub UserForm_Click()

End Sub
