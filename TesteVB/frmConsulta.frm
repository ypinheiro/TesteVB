VERSION 5.00
Begin VB.Form frmConsulta 
   Caption         =   "Consulta"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9915
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "Consultar"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox LstView 
      Height          =   1425
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   8895
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label lblnome 
      Caption         =   "Nome:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsulta_Click()
    Dim Cliente As New clsCliente
    Dim item As clsDCliente
    Dim lista As Collection
    Dim Consulta As New clsDCliente
    
    On Error GoTo erro
      
    Consulta.Nome = txtNome.Text
    Set lista = Cliente.Buscar(Consulta)
    
    For Each item In lista
        LstView.AddItem item.Nome
    Next
    
fim:
    
    Set item = Nothing
    Set Cliente = Nothing
    Set lista = Nothing
   
    Exit Sub
    
erro:
    MsgBox Err.Description
    
    GoTo fim
End Sub



