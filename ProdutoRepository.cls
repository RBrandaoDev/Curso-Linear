VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProdutoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RecuperarPrdutos() As ADODB.Recordset
    Dim sql As String

    sql = "Select * from produto;"

    Set RecuperarProdutos = RetornaConsulta(sql)


End Function

Public Sub ExcluirProdutos(Codigo_ As Long)
  Dim sql As String
    
     Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    
102     cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=127.0.0.1;" & "DATABASE=linear1;" & "UID=adminlinear;" & "PWD=@2013linear;" & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
103     cnn.CursorLocation = adUseClient
104     cnn.Open
        
        sql = "Delete from produto "
        sql = sql & " Where codigo = " & val(Codigo_)
        cnn.Execute (sql)
End Sub

Public Sub Gravar(Nome_, Preco_, Estoque_, Codigo_)

  Dim sql As String

    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
        
    cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=127.0.0.1;" & "DATABASE=linear1;" & "UID=adminlinear;" & "PWD=@2013linear;" & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    cnn.CursorLocation = adUseClient
    cnn.Open
    
    If Codigo_ = 0 Then
        sql = "INSERT INTO produto (nome, preco, estoque) VALUES ( '" & Nome_ & "', '" & Preco_ & "', '" & Estoque_ & "');"
    Else
        sql = " UPDATE produto set nome = '" & Nome_ & "' , preco = '" & Preco_ & "' , estoque = '" & Estoque_ & "', where codigo = '" & Codigo_ & "'"
    
    End If
        
        cnn.Execute (sql)
End Sub
