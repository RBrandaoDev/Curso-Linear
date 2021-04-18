VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbousuario 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtSenha 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogar 
      Caption         =   "Logar"
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   990
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   540
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   450
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLogar_Click()

        On Error GoTo Exception

        Dim rs As ADODB.Recordset
100     Set rs = New ADODB.Recordset
        Dim sql As String
    
101     sql = " Select * from usuarios where nome = '" & cbousuario.Text & "' and password = '" & txtSenha.Text & "'"
    
102     Set rs = RetornaConsulta(sql)

        'Se não for o final do arquivo (end of File)
103     If Not rs.EOF Then
            usuarioLogin = rs!nome
104         Menu.Show
            
        Else
105         MsgBox "Login não efetuado!"
        
        End If

Fim:
        Unload Me

Exception:

        If Err.Number = "-2147217900" Then
            MsgBox "Erro o executar SQL no banco. " & Err.Description
        End If
    
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim rs  As ADODB.Recordset

    On Error GoTo Exception

    sql = "Select * from usuarios"

    Set rs = RetornaConsulta(sql)

    Do While Not (rs.EOF)
        cbousuario.AddItem rs!nome
        rs.MoveNext
    Loop

    cbousuario.ListIndex = 0
    
    'End If
    'Try
Fim:
    'call FecharRs(rs, False)
    Exit Sub
    
    'Exception
Exception:
    MsgBox "Erro: " & Err.Description & " - " & Err.Number
End Sub

