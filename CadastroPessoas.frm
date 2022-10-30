VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CadastroPessoas 
   Caption         =   "Cadastro Pessoas"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Código"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
      Begin VB.TextBox tfCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox tfEndereco 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   7215
   End
   Begin VB.CommandButton btVoltar 
      Caption         =   "Voltar"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox tfEmail 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton btGravar 
      Caption         =   "Gravar"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox tfNome 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nome"
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "E-mail"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "CPF"
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin MSMask.MaskEdBox tfCpf 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   7455
   End
   Begin VB.Frame Frame5 
      Caption         =   "Telefone"
      Height          =   735
      Left            =   4080
      TabIndex        =   11
      Top             =   1800
      Width           =   3495
      Begin VB.TextBox tfTelefone 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "CadastroPessoas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btGravar_Click()
If txtCodigo = "" Then
    Call Module1.InserirPessoa(tfNome, tfCpf, tfEndereco, tfEmail, tfTelefone)
    Else
    Call Module1.AtualizarPessoa(tfCodigo, tfNome, tfCpf, tfEndereco, tfEmail, tfTelefone)
End Sub

Private Sub btVoltar_Click()
Me.Hide
ListaPessoas.Show
End Sub

Public Sub LimparCampos()
tfCodigo = ""
tfNome = ""
tfCpf = ""
tfEndereco = ""
tfEmail = ""
tfTelefone = ""
End Sub

Private Sub tfCpf_GotFocus()
tfCpf.Mask = "##############"
End Sub

Private Sub tfCpf_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub

Private Sub tfCpf_LostFocus()
If Len(tfCpf.Text) > 0 Then
tfCpf.Mask = "###.###.###-##"
    If Not calculacpf(tfCpf.Text) Then
            MsgBox "CPF com DV incorreto !!!"
            tfCpf = ""
            tfCpf.Mask = "##########"
            tfCpf.SetFocus
        End If
End If
End Sub

Private Sub tfEmail_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub

Private Sub tfEndereco_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub


Private Sub tfNome_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub

Private Sub tfTelefone_KeyPress(KeyAscii As Integer)
  'se teclar enter envia um TAB
  If KeyAscii = 13 Then
     SendKeys "{TAB}"
     KeyAscii = 0
  End If
End Sub
Function calculacpf(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
On Error GoTo Err_CPF
Dim I As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divisão
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1º digito verificador
Dim intDig2 As Integer 'armazena o 2º digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia cálculos do 1º dígito
For I = 2 To 10
    strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For I = 2 To 11
     strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    calculacpf = False
Else
    calculacpf = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function
