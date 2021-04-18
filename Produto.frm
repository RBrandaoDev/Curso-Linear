VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Produto 
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
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
   ScaleHeight     =   4155
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   360
      Left            =   4200
      TabIndex        =   12
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   360
      Left            =   2400
      TabIndex        =   11
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   360
      Left            =   480
      TabIndex        =   10
      Top             =   3600
      Width           =   990
   End
   Begin MSDataGridLib.DataGrid grdProduto 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameDados 
      Caption         =   "Dados"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtEstoque 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtPreco 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lbllEstoque 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblLPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preco"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblLNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum enuProduto
    codigo = 0
    nome = 1
    preco = 2
    estoque = 3
    
End Enum

Private Sub cmdExcluir_Click()
    ExcluirProduto
      PreencheGrid
End Sub

Private Sub ExcluirProduto()
 Dim Produto As ProdutoRepository
 Set Produto = New ProdutoRepository
 
 Produto.ExcluirProdutos (Codigo_)
 LimpaCampos
  
End Sub

Private Sub cmdGravar_Click()
    GravarProduto
    PreencheGrid
End Sub

Private Sub GravarProduto()
  Dim Produto As ProdutoRepository
  Set Produto = New ProdutoRepository
  Call Produto.Gravar(txtNome.Text, txtPreco.Text, txtEstoque.Text, val(txtCodigo.Text))
  
End Sub


Private Sub cmdSair_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
  PreencheGrid
  
End Sub

Private Sub PreencheGrid()
    Dim listaProdutos As ProdutoRepository
    Set listaProdutos = New ProdutoRepository
    Set grdProduto.DataSource = listaProdutos.RecuperarPrdutos
          
End Sub

Private Sub grdProduto_Click()

txtCodigo.Text = grdProduto.Columns(enuProduto.codigo)
txtNome.Text = grdProduto.Columns(1)
txtPreco.Text = grdProduto.Columns(2)
txtEstoque.Text = grdProduto.Columns(3)

End Sub

Private Sub LimpaCampos()
txtCodigo.Text = ""
txtNome.Text = ""
txtPreco.Text = ""
txtEstoque.Text = ""

End Sub
