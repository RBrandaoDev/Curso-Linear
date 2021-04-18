VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Cliente 
   Caption         =   "Cadastro de cliente"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
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
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   360
      Left            =   4320
      TabIndex        =   12
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   360
      Left            =   3000
      TabIndex        =   11
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   360
      Left            =   1680
      TabIndex        =   10
      Top             =   3600
      Width           =   990
   End
   Begin MSDataGridLib.DataGrid grdCliente 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2566
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtLimiteCredito 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLimiteCredito 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite Credito"
         Height          =   195
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
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
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum enuCliente
    codigo = 0
    nome = 1
    telefone = 2
    limiteCredito = 3
End Enum

Private Sub cmdExcluir_Click()
    ExcluirCliente
        PreencheGrid

End Sub

' VC -> Model, View, Controller

Private Sub cmdGravar_Click()
    GravarCliente
     PreencheGrid
End Sub

Private Sub GravarCliente()
    Dim Clientes As ClienteRepository
    Set Clientes = New ClienteRepository
    
    Call Clientes.Gravar(txtNome.Text, txtTelefone.Text, txtLimiteCredito.Text, Val(txtCodigo.Text))
    LimpaCampos
End Sub

Private Sub cmdSair_Click()
 Unload Me
        
End Sub

Private Sub Form_Load()
    PreencheGrid

End Sub

Private Sub PreencheGrid()
    'listaClientes = variavel(Objeto) do tipo ClienteRepository
    ' declarando o objeto
    Dim listaClientes As ClienteRepository
    ' Instanciando o Objeto na memoria (Se fosse somente uma váriavel não seria necessário instanciar na memoria pois ele não possui comportamentos)
    ' min 09:00 videoaula1
    Set listaClientes = New ClienteRepository
    
    ' DataSource é uma propriedade do gridCliente ele precisa de um Recordset
    Set grdCliente.DataSource = listaClientes.RecuperarClientes
    
End Sub
Private Sub ExcluirCliente()
    Dim Clientes As ClienteRepository
    Set Clientes = New ClienteRepository
    Call Clientes.Excluir(txtCodigo.Text)
    LimpaCampos
        
End Sub

Private Sub grdCliente_Click()
txtCodigo.Text = grdCliente.Columns(enuCliente.codigo)
txtNome.Text = grdCliente.Columns(1)
txtTelefone = grdCliente.Columns(2)
txtLimiteCredito = grdCliente.Columns(3)
End Sub

Private Sub LimpaCampos()
txtCodigo.Text = ""
txtNome.Text = ""
txtLimiteCredito.Text = ""
End Sub
