Attribute VB_Name = "Module1"
Global usuarioLogin As String

Public Function RetornaConsulta(sql_ As String) As ADODB.Recordset
        On Error GoTo Exception
        Dim rst As ADODB.Recordset
100     Set rst = New ADODB.Recordset
        Dim cnn As ADODB.Connection
    
101     Set cnn = New ADODB.Connection
    
102     cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=127.0.0.1;" & "DATABASE=linear1;" & "UID=adminlinear;" & "PWD=@2013linear;" & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
103     cnn.CursorLocation = adUseClient
104     cnn.Open
105     rst.Open sql_, cnn, adOpenForwardOnly, adLockReadOnly
106     Set RetornaConsulta = rst
        
        'Try
Fim:
        Exit Function
  
        'Catch
Exception:
107     'MsgBox "Erro Mysql: " & Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description

End Function
