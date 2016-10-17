VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intID As Integer
Private strNome As String
Private strEmail As String

Public Property Get ID() As String
    ID = intID
End Property

Public Property Let ID(ByVal vID As String)
    intID = vID
End Property

Public Property Get Nome() As String
    Nome = strNome
End Property

Public Property Let Nome(ByVal vNome As String)
    strNome = vNome
End Property

Public Property Get Email() As String
    Email = strEmail
End Property

Public Property Let Email(ByVal vEmail As String)
    strEmail = vEmail
End Property
