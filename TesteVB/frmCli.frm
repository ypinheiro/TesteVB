VERSION 5.00
Begin VB.Form frmCli 
   Caption         =   "Clientes"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   9675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCad 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblnome 
      Caption         =   "Nome:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblemail 
      Caption         =   "E-mail:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCad_Click()
Dim Cliente As New clsCliente
Dim dCliente As New clsDCliente
    
    On Error GoTo erro
    If Len(txtNome.Text) = 0 Then
        MsgBox "Nome Obrigatorio!", vbCritical
    Else
        dCliente.Nome = txtNome.Text
    End If
    
    If Len(txtEmail.Text) = 0 Then
        MsgBox "Email Obrigatorio!", vbCritical
    Else
        dCliente.Email = txtEmail.Text
    End If
    
    If Not Cliente.Inserir(dCliente) Then
        MsgBox "Erro Cadastro", vbCritical
    Else
        MsgBox "Cadstro OK", vbInformation
    End If

fim:
    Set Cliente = Nothing
    Set dCliente = Nothing
    
    Exit Sub
    
erro:
    MsgBox Err.Description

    GoTo fim

End Sub
