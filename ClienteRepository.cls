VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RecuperarClientes() As ADODB.Recordset
    Dim sql As String
    
    sql = "Select * from clientes;"
    
    Set RecuperarClientes = RetornaConsulta(sql)
        
End Function

Public Sub Excluir(Codigo_ As Long) ' (codigo_)Convenção, para saber que é um parametro

        Dim sql As String
    
        Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    
102     cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=127.0.0.1;" & "DATABASE=linear1;" & "UID=adminlinear;" & "PWD=@2013linear;" & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
103     cnn.CursorLocation = adUseClient
104     cnn.Open
        
        sql = "Delete from clientes "
        sql = sql & " Where codigo = " & val(Codigo_)
        cnn.Execute (sql)

End Sub

Public Sub Gravar(Nome_, telefone_, limiteCredito_, Codigo_)

        Dim sql As String
    
        Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    
102     cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=127.0.0.1;" & "DATABASE=linear1;" & "UID=adminlinear;" & "PWD=@2013linear;" & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
103     cnn.CursorLocation = adUseClient
104     cnn.Open
        
        If Codigo_ = 0 Then
            sql = "INSERT INTO `clientes` (`nome`, `telefone`, `limite_credito`) "
            sql = sql & " VALUES ('" & Nome_ & "', '" & telefone_ & "', '" & limiteCredito_ & "');"
        Else
            sql = " UPDATE clientes SET nome = '" & Nome_
            sql = sql & "', telefone = '" & telefone_
            sql = sql & "', limite_credito = " & limiteCredito_
            sql = sql & " Where codigo = " & Codigo_
        
        End If
        
        cnn.Execute (sql)
        
End Sub

