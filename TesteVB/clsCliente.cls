VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const pUrlBase As String = "http://localhost:64228/APITeste/usuario"

Public Function Buscar(dCliente As clsDCliente) As Collection
    Dim Ans As String
    Dim oHTTP As MSXML2.XMLHTTP, sBuffer As String
        
    On Error GoTo erro
  
    Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    oHTTP.Open "GET", pUrlBase & "?" & "Nome=" & dCliente.Nome, False
    oHTTP.send
    sBuffer = oHTTP.responseText
    Set oHTTP = Nothing
    
    Dim jsonCli As Object
    Set jsonCli = JSON.parse(sBuffer)
    Dim clienteJson As Object
    Dim ConsCli As New Collection


    For Each clienteJson In jsonCli
        Dim lista As New clsDCliente
        lista.ID = CInt(clienteJson.item("Id"))
        lista.Nome = clienteJson.item("Nome")
        lista.Email = clienteJson.item("Email")
        ConsCli.Add lista
        Set lista = Nothing
    Next
    
    Set Buscar = ConsCli
    
fim:
    Set jsonCli = Nothing
    Set clienteJson = Nothing
    
    Exit Function

erro:
    MsgBox Err.Description
    
    Set Buscar = Nothing
    
    GoTo fim
End Function
    
Public Function Inserir(dCliente As clsDCliente) As Boolean
    Dim Ans As String
    Dim oHTTP As MSXML2.XMLHTTP, sBuffer As String
    Dim param As String
    
    On Error GoTo erro
    
    Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    oHTTP.Open "POST", pUrlBase, False
    
        param = "Nome:" & dCliente.Nome
        param = param & ",Email:" & dCliente.Email
    
    oHTTP.send "{" & param & "}"
    sBuffer = oHTTP.responseText
    Inserir = True
    
fim:
    Set oHTTP = Nothing
    Exit Function
erro:
    Inserir = False
    MsgBox Err.Description, vbCritical
    GoTo fim
End Function








