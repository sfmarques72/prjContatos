Attribute VB_Name = "Mod_Main"
Option Explicit

' variavel que guardara a conexao com o banco de dados
Public ADOConnDB As ADODB.Connection

Public MaxNome As Integer
Public MaxSobrenome As Integer
Public MaxEmail As Integer
Public MaxTelefone As Integer

Sub Main()

'   On Error GoTo MostraErro
   
'   ADOConnDB = ConectarAoPostgreSQL
   
   ' define tamanhos maximos de campos
   MaxNome = 50
   MaxSobrenome = 100
   MaxEmail = 255
   MaxTelefone = 15
   
   ' chama o formulario principal'
   Form2.Show
'MostraErro:
'   MsgBox "Erro: " & Err.Description & " - " & Err.Source, vbApplicationModal + vbCritical + vbOKOnly, "Erro Critico"
'   ' Sair do sistema
'   SaidaSistema
End Sub

' Conexao com banco de dados Postgres
Function ConectarAoPostgreSQL() As ADODB.Connection
    On Error Resume Next
    Dim conn As ADODB.Connection
    
    Set conn = New ADODB.Connection
    
    ' Iniciando conexão com o banco de dados
    conn.ConnectionString = "DRIVER={PostgreSQL ODBC Driver(ANSI)};" & _
                            "Server=127.0.0.1;" & _
                            "Port=5432;" & _
                            "Database=postgres;" & _
                            "Uid=postgres;" & _
                            "Pwd=mysecretpassword;"
    
    conn.Open
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao conectar ao banco de dados: " & Err.Description, vbExclamation
        Set conn = Nothing
    End If
    
'    On Error GoTo ErrorHandler
    
    Set ConectarAoPostgreSQL = conn
    Exit Function
End Function

Private Sub FecharConexao(ByRef conn As ADODB.Connection)
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub
   
Private Sub cmdAdicionar_Click()
    On Error GoTo ErrorHandler
    Dim conn As ADODB.Connection
    Set conn = ConectarAoPostgreSQL()
    
    ' Nomeando o tipo da variavel
    Dim nome As String
    Dim email As String
    Dim telefone As String
    
    nome = txtNome.Text
    email = txtEmail.Text
    telefone = txtTelefone.Text
    
    ' Chame a função para adicionar um contato
    adicionar_contato nome, email, telefone
    
    ' Feche a conexão
    FecharConexao conn
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao adicionar contato: " & Err.Description, vbExclamation
    ' Trate o erro de acordo com suas necessidades

    ' Feche a conexão
    FecharConexao conn
End Sub

Private Sub cmdEditar_Click()
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Set conn = ConectarAoPostgreSQL()
    
    Dim id_contato As Integer
    Dim novo_nome As String
    Dim novo_email As String
    Dim novo_telefone As String
    
    id_contato = CInt(txtID.Text)
    novo_nome = txtNovoNome.Text
    novo_email = txtNovoEmail.Text
    novo_telefone = txtNovoTelefone.Text
    
    ' Call
    editar_contato id_contato, novo_nome, novo_email, novo_telefone
    
    ' Feche a conexão
    FecharConexao conn
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao editar contato: " & Err.Description, vbExclamation
    ' Trate o erro de acordo com suas necessidades

    ' Feche a conexão
    FecharConexao conn
End Sub

Private Sub cmdVisualizar_Click()
    On Error GoTo ErrorHandler
    Dim conn As ADODB.Connection
    Set conn = ConectarAoPostgreSQL()
    
    ' Crie um objeto de Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Abra o Recordset com a consulta SQL
    rs.Open "SELECT * FROM contatos", conn
    
    ' Verifique se o Recordset está vazio
    If rs.EOF Then
        MsgBox "Nenhum contato encontrado.", vbInformation
    Else
        ' Inicialize uma variável para armazenar os dados
        Dim mensagem As String
        mensagem = "Contatos:" & vbCrLf & vbCrLf
        
        ' Percorra o Recordset e adicione os dados à mensagem
        Do While Not rs.EOF
            mensagem = mensagem & "ID: " & rs("id").Value & vbCrLf & _
                       "Nome: " & rs("nome").Value & vbCrLf & _
                       "Email: " & rs("email").Value & vbCrLf & _
                       "Telefone: " & rs("telefone").Value & vbCrLf & vbCrLf
            rs.MoveNext
        Loop
        
        ' Exiba a mensagem com os dados dos contatos
        MsgBox mensagem, vbInformation
    End If
    
    ' Feche o Recordset
    rs.Close
    
    ' Feche a conexão
    FecharConexao conn
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao visualizar contatos: " & Err.Description, vbExclamation
    ' Trate o erro de acordo com suas necessidades

    ' Feche a conexão
    FecharConexao conn
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Set conn = ConectarAoPostgreSQL()
    
    Dim id_contato As Integer
    id_contato = CInt(txtIDExclusao.Text) ' Supondo que haja um campo de texto para inserir o ID do contato a ser excluído
    
    ' Chame a função para excluir um contato
    excluir_contato id_contato
    
    ' Feche a conexão
    FecharConexao conn
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao excluir contato: " & Err.Description, vbExclamation
    ' Trate o erro de acordo com suas necessidades

    ' Feche a conexão
    FecharConexao conn
End Sub

Private Sub excluir_contato(ByVal id_contato As Integer)
    On Error Resume Next
    
    Dim conn As ADODB.Connection
    Set conn = ConectarAoPostgreSQL()
    
    ' Execute uma consulta SQL para excluir o contato com o ID fornecido
    conn.Execute "DELETE FROM contatos WHERE id = " & id_contato
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao excluir o contato: " & Err.Description, vbExclamation
    Else
        MsgBox "Contato excluído com sucesso!"
    End If
    
    ' Feche a conexão
    FecharConexao conn
End Sub

' encerramento do sistema
Public Function SaidaSistema() As Integer
   On Error Resume Next
   
   ADOConnDB.Close
   End
   
End Function
