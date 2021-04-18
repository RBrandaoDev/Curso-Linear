VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Menu"
   ClientHeight    =   4035
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7965
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
   ScaleHeight     =   4035
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblUsuarioLogado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UsuarioLogado"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   600
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuProdutos 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnuFornecedores 
         Caption         =   "Fornecedores"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblUsuarioLogado = usuarioLogin
    lblUsuarioLogado.ForeColor = vbRed

End Sub


Private Sub mnuClientes_Click()
Cliente.Show
End Sub

Private Sub mnuFornecedores_Click()
Fornecedor.Show
End Sub

Private Sub mnuProdutos_Click()
Produto.Show
End Sub

