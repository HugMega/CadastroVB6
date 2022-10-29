Attribute VB_Name = "Module1"
Private con As ADODB.Connection
Public rs As ADODB.Recordset

Private Sub Conectar()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Pessoas.mdb"
End Sub

Public Sub Desconectar()
con.Close
Set con = Nothing
End Sub

Public Function PesquisarPessoa(Valor As String) As Boolean
Dim Criterio As String
Call Conectar
Set rs = New ADODB.Recordset
If Trim(Valor) <> "" Then
    Criterio = " WHERE Codigo=" & Valor
End If
rs.Open "SELECT * FROM Pessoas" & Criterio, con, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
    PesquisarPessoa = True
Else
    PesquisarPessoa = False
End If
End Function

Public Sub InserirPessoa(Nome As String, CPF As String, Endereco As String, Email As String, Telefone As String)
Call Conectar
con.Execute "INSERT INTO Pessoas(Nome, CPF, Endereco, Email, Telefone) " & _
            "VALUES ('" & Nome & "', '" & CPF & "', '" & Endereco & "', '" & Email & "', '" & Telefone & "')"
Call Desconectar
End Sub

Public Sub AtualizarPessoa(Codigo As String, Nome As String, Endereco As String, Telefone As String, Email As String)
Call Conectar
con.Execute "UPDATE Pessoas " & _
            "SET Nome='" & Nome & "', " & _
            "Endereco='" & Endereco & "', " & _
            "Telefone='" & Telefone & "', " & _
            "Email='" & Email & "' " & _
            "WHERE Codigo=" & Codigo
Call Desconectar
End Sub

Public Sub ExcluirContato(Codigo As String)
Call Conectar
con.Execute "DELETE FROM Pessoas WHERE Codigo=" & Codigo
Call Desconectar
End Sub

