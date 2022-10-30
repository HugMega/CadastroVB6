VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ListaPessoas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Pessoas"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7815
   DrawMode        =   1  'Blackness
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstPessoas 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483628
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btNovo 
      Caption         =   "Cadastrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "ListaPessoas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btExcluir_Click()
If lstPessoas.ListItems.Count > 0 Then
    If MsgBox("Deseja realmente excluir a Pessoa '" & lstPessoas.SelectedItem.ListSubItems(1).text & "'?", vbYesNo, "Excluir") = vbYes Then
        Call Module1.ExcluirPessoa(lstPessoas.SelectedItem.text)
        Call ListarPessoa
    End If
End If
End Sub

Private Sub btNovo_Click()
CadastroPessoas.Show
CadastroPessoas.LimparCampos
Me.Hide
End Sub

Private Sub Command1_Click()
If lstPessoas.ListItems.Count > 0 Then
    Call PesquisarContato(lstPessoas.SelectedItem.text)
End If
CadastroPessoas.Show
Me.Hide
End Sub


Private Sub PesquisarContato(Codigo As String)
Dim lst As ListItem
If Module1.PesquisarPessoa(Codigo) = True Then
    lstPessoas.ListItems.Clear
    If Not rs.EOF Then
    With CadastroPessoas
        .tfCodigo = rs.Fields("codigo")
        .tfNome = rs.Fields("nome")
        .tfCpf = rs.Fields("CPF")
        .tfEndereco = rs.Fields("endereco")
        .tfTelefone = rs.Fields("telefone")
        .tfEmail = rs.Fields("email")
    End With
    End If
    Call Desconectar
End If
End Sub


Private Sub ListarPessoa()
Dim lst As ListItem
If Module1.PesquisarPessoa("") = True Then
    lstPessoas.ListItems.Clear
    'enquanto houver registros inclui os registros no ListView
    Do While Not rs.EOF
        Set lst = lstPessoas.ListItems.Add(, , rs.Fields("codigo"))
        lst.SubItems(1) = rs.Fields("nome")
        lst.SubItems(2) = rs.Fields("cpf")
        rs.MoveNext
    Loop
    Call Desconectar
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
Call ListarPessoa
End Sub

Private Sub Form_Load()
    lstPessoas.ColumnHeaders.Add , , "Código"
    lstPessoas.ColumnHeaders.Add , , "Nome", 2460
    lstPessoas.ColumnHeaders.Add , , "CPF", 1900
End Sub
