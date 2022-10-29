VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ListaPessoas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Pessoas"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11100
   DrawMode        =   1  'Blackness
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstPessoas 
      Height          =   2670
      Left            =   240
      TabIndex        =   0
      Top             =   705
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4710
      View            =   2
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton btNovo 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "ListaPessoas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call ListarPessoa
End Sub
Private Sub ListarPessoa()
Dim lst As ListItem
If Module1.PesquisarPessoa("") = True Then
    lstPessoas.ListItems.Clear
    lstPessoas.ColumnHeaders.Add , , "Código"
    lstPessoas.ColumnHeaders.Add , , "Nome"
    'enquanto houver registros inclui os registros no ListView
    Do While Not rs.EOF
        Set lst = lstPessoas.ListItems.Add(, , rs.Fields("nome"))
        lst.SubItems(1) = rs.Fields("codigo")
        rs.MoveNext
    Loop
    Call Desconectar
End If
End Sub

