VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cadastro de Contatos"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4245
      TabIndex        =   10
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   1275
      TabIndex        =   9
      Top             =   3000
      Width           =   2000
   End
   Begin VB.TextBox txtTelefone 
      Height          =   300
      Left            =   1425
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1995
      Width           =   2000
   End
   Begin VB.TextBox txtEmail 
      Height          =   300
      Left            =   1425
      MaxLength       =   100
      TabIndex        =   7
      Top             =   1590
      Width           =   5800
   End
   Begin VB.TextBox txtSobrenome 
      Height          =   300
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1200
      Width           =   5800
   End
   Begin VB.TextBox txtNome 
      Height          =   300
      Left            =   1425
      MaxLength       =   30
      TabIndex        =   5
      Top             =   825
      Width           =   5800
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefone:"
      Height          =   195
      Left            =   465
      TabIndex        =   4
      Top             =   2025
      Width           =   870
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1665
      Width           =   870
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Sobrenome:"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1260
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   195
      Left            =   495
      TabIndex        =   1
      Top             =   855
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dados do Contato"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   195
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cnxn As ADODB.Connection
Public rs As ADODB.Recordset
Public strCnxn As String

' inicializa o formulario de cadastro
Private Sub Form_Load()
   ' centralizar o formulario na tela
   Top = (Screen.Height - Height) / 2
   Left = (Screen.Width - Width) / 2
   
   
   ' define o tamanho maximo permitido nos campos
   MaxNome = 30
   MaxSobrenome = 100
   MaxEmail = 255
   MaxTelefone = 15
      
   ' define propriedade MaxLength
   txtNome.MaxLength = MaxNome
   txtSobrenome.MaxLength = MaxSobrenome
   txtEmail.MaxLength = MaxEmail
   txtTelefone.MaxLength = MaxTelefone
   On Error GoTo MostraErro
   
    Dim strSql As String
     'record variables
    Dim strID As String
    Dim strFirstName As String
    Dim strLastName As String
   
   ' abrir o banco de dados
    Set Cnxn = New ADODB.Connection
    ' Crie uma conexao ODBC de 32 bits com nome "PostgreSQL30"
    strCnxn = "DSN=PostgreSQL30;"
    ' abre o banco de dados
    Cnxn.Open strCnxn
    ' prepara o RecordSet dos contatos
    Set rs = New ADODB.Recordset
    strSql = "contatos"
    rs.Open strSql, strCnxn, adOpenKeyset, adLockOptimistic, adCmdTable
    Exit Sub

MostraErro:
   MsgBox "Erro: " & Err.Description & " - " & Err.Source, vbApplicationModal + vbCritical + vbOKOnly, "Erro Critico"
   

End Sub

' Botao cancelar
' Fecha a janela
Private Sub cmdCancelar_Click()

'libera os objetos
   Set rs = Nothing
   Set db = Nothing
   Unload Me
End Sub

' Botao Salvar
Private Sub cmdSalvar_Click()
   Dim strSql As String
   Dim nome As String
   Dim sobrenome As String
   Dim email As String
   Dim telefone As String
   Dim codigo As Long
   
   ' remove espacos a esquerda e a direita dos campos
   txtNome.Text = Trim(txtNome.Text)
   txtSobrenome.Text = Trim(txtSobrenome.Text)
   txtEmail.Text = Trim(txtEmail.Text)
   txtTelefone.Text = Trim(txtTelefone.Text)
   
   ' Salva os dados da tela
   nome = txtNome.Text
   sobrenome = txtSobrenome.Text
   email = txtEmail.Text
   telefone = txtEmail.Text

   If Consistencia() Then
      ' Pegar ultimo ID. Quando tiver mais tempo implementar Sequence no banco de dados
      Set rsContato = New ADODB.Recordset
      strSql = "SELECT MAX(codigo) + 1 FROM contatos"
      rsContato.Open strSql, strCnxn, adOpenDynamic, adLockOptimistic, adCmdText
      codigo = rsContato.Fields(0).Value
      rsContato.Close
      ' inclui o novo registro
      rs.AddNew
      rs!codigo = codigo
      rs!nome = nome
      rs!sobrenome = sobrenome
      rs!email = email
      rs!telefone = telefone
      rs.Update
      ' rs.Close
      ' mostra mensagem de sucesso na tela
      MsgBox "Contato salvo", vbApplicationModal + vbExclamation + vbInformation, "Informação"
   End If
End Sub

' Realiza a consitencia das informações digitadas
Public Function Consistencia() As Boolean
   ' retorno padrao da consitentecia
   Consistencia = False
   
   nome = txtNome.Text
   sobrenome = txtSobrenome.Text
   email = txtEmail.Text
   telefone = txtTelefone.Text
   
   ' Se Nome nao for preenchido
   If Len(nome) = 0 Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Nome é obrigatório", vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtNome.SetFocus
      Exit Function
   End If
   ' Se nome for maior que o tamanho definido no banco de dados. Mesmo que eu tenha
   ' definido a propriedade MasLength
   If Len(nome) > MaxNome Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Nome deve ter no máximo " & CStr(MaxNome), vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtNome.SetFocus
      Exit Function
   End If
   
   ' Se sobrenome nao for preenchido
   If Len(sobrenome) = 0 Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Sobrenome é obrigatório", vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do sobrenome para preenchimento
      txtSobrenome.SetFocus
      Exit Function
   End If
   ' Se sobrenome for maior que o tamanho definido no banco de dados. Mesmo que eu tenha
   ' definido a propriedade MasLength
   If Len(sobrenome) > MaxSobrenome Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Sobrenome deve ter no máximo " & CStr(MaxSobrenome), vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtSobrenome.SetFocus
      Exit Function
   End If
   
   ' Se o e-mail nao for digitado
   If Len(email) = 0 Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Endereço de e-mail é obrigatório", vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do sobrenome para preenchimento
      txtEmail.SetFocus
      Exit Function
   End If
   ' Se e-mail for maior que o tamanho definido no banco de dados. Mesmo que eu tenha
   ' definido a propriedade MasLength
   If Len(email) > MaxEmail Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Endereço de e-mail deve ter no máximo " & CStr(MaxEmail), vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtEmail.SetFocus
      Exit Function
   End If
   If Not isEmail(email) Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "Endereço de e-mail inválido", vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtEmail.SetFocus
      Exit Function
   End If
   
   ' Se o numero do telefone nao for digitado
   If Len(telefone) = 0 Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "O número do telefone é obrigatório", vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do sobrenome para preenchimento
      txtTelefone.SetFocus
      Exit Function
   End If
   ' Se o telefone for maior que o tamanho definido no banco de dados. Mesmo que eu tenha
   ' definido a propriedade MasLength
   If Len(telefone) > MaxTelefone Then
      ' a mensagem ficara na tela impedindo de continuar
      MsgBox "O número do telefone deve ter no máximo " & CStr(MaxTelefone), vbApplicationModal + vbExclamation + vbOKOnly, "Erro"
      ' coloca o cursor no textbox do Nome para preenchimento
      txtTelefone.SetFocus
      Exit Function
   End If
   
   ' Não entrou em nenhum IF então os dados são consistentes e pode salvar
   Consistencia = True
End Function

' Função para checar e-mail
' Obtida de https://www.clubedohardware.com.br/forums/topic/248420-verificar-se-%C3%A9-e-mail-no-vb6/
 Public Function isEmail(ByVal pEmail As String) As Boolean
        
    Dim Conta As Integer, Flag As Integer, cValido As String
    isEmail = False
    If Len(pEmail) < 5 Then
        Exit Function
    End If
    'Verifica se existe caracter inválido
    For Conta = 1 To Len(pEmail)
        cValido = Mid(pEmail, Conta, 1)
        If Not (LCase(cValido) Like "[a-z]" Or cValido = _
            "@" Or cValido = "." Or cValido = "-" Or _
            cValido = "_") Then
            Exit Function
        End If
    Next
    'Verifica a existência de (@)
    If InStr(pEmail, "@") = 0 Then
        Exit Function
    Else
        Flag = 0
        
        For Conta = 1 To Len(pEmail)
            If Mid(pEmail, Conta, 1) = "@" Then
                Flag = Flag + 1
            End If
        Next
        
        If Flag > 1 Then Exit Function
    End If
    
    If Left(pEmail, 1) = "@" Then
        Exit Function
    ElseIf Right(pEmail, 1) = "@" Then
        Exit Function
    ElseIf InStr(pEmail, ".@") > 0 Then
        Exit Function
    ElseIf InStr(pEmail, "@.") > 0 Then
        Exit Function
    End If
  
  
    'Verifica a existência de (.)
    If InStr(pEmail, ".") = 0 Then
        Exit Function
    ElseIf Left(pEmail, 1) = "." Then
        Exit Function
    ElseIf Right(pEmail, 1) = "." Then
        Exit Function
    ElseIf InStr(pEmail, "..") > 0 Then
        Exit Function
    End If
    
    isEmail = True
End Function

' coloca nome em maicusculas
Private Sub txtNome_LostFocus()
   txtNome.Text = UCase(txtNome.Text)
End Sub

' coloca sobre nome em maiusculas
Private Sub txtSobrenome_LostFocus()
   txtSobrenome.Text = UCase(txtSobrenome.Text)
End Sub


' transforma email em minusculas
Private Sub txtEmail_LostFocus()
   txtEmail.Text = LCase(txtEmail.Text)
End Sub

' so permite caracteres especificos no campo telefone
Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
   Dim strValid As String
   strValid = "() -0123456789" & Chr(8)
   ' se nao for caracter valido, ignora o caracter digitado
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub
