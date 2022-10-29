VERSION 5.00
Begin VB.Form CadastroPessoas 
   Caption         =   "Cadastro Pessoas"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   6015
   End
   Begin VB.CommandButton btVoltar 
      Caption         =   "Voltar"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox tfCpf 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox tfEmail 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton btGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox tfNome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nome"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "E-mail"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Frame Frame3 
      Caption         =   "CPF"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Frame Frame5 
      Caption         =   "Telefone"
      Height          =   735
      Left            =   3960
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "CadastroPessoas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btGravar_Click()
    Call Module1.InserirPessoa(tfNome)
End Sub

