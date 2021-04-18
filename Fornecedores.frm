VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Fornecedor 
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
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
   ScaleHeight     =   4305
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   360
      Index           =   2
      Left            =   4080
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Excluir"
      Height          =   360
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Gravar"
      Height          =   360
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      Top             =   3840
      Width           =   990
   End
   Begin MSDataGridLib.DataGrid grdFornecedor 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2990
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
      Width           =   6495
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtRepresentante 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000002&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   195
         Left            =   4200
         TabIndex        =   9
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblLRepresentante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante"
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1080
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
      Begin VB.Label lblLNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblLCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Fornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    PreencheGrid

End Sub


Private Sub PreencheGrid()
    Dim sql As String

    sql = "Select * from fornecedor"

    Set grdFornecedor.DataSource = RetornarConsulta(sql)

End Sub

Private Sub grdFornecedor_Click()

    txtCodigo.Text = grdFornecedor.Colluns(0)
    txtNome.Text = grdFornecedor.Colluns(1)
    txtTelefone.Text = grdFornecedor.Colluns(2)
    txtRepresentante.Text = grdFornecedor.Colluns(3)
    txtEmail.Text = grdFornecedor.Colluns(4)

End Sub

