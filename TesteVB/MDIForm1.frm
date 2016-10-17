VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Cad Clientes"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9270
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuclientes 
      Caption         =   "Clientes"
   End
   Begin VB.Menu mnuCconsulta 
      Caption         =   "Consulta"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCconsulta_Click()
frmConsulta.Show

End Sub

Private Sub mnuclientes_Click()
frmCli.Show

End Sub

Private Sub mnuSair_Click()
Unload MDIForm1

End Sub
