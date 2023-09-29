VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Contatos1"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12255
   LinkTopic       =   "Form2"
   ScaleHeight     =   8670
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   9285
      TabIndex        =   6
      Top             =   7700
      Width           =   2000
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   495
      Left            =   6495
      TabIndex        =   5
      Top             =   7700
      Width           =   2000
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   495
      Left            =   3795
      TabIndex        =   4
      Top             =   7700
      Width           =   2000
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   825
      TabIndex        =   3
      Top             =   7700
      Width           =   2000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contatos"
      Height          =   5085
      Left            =   300
      TabIndex        =   1
      Top             =   2400
      Width           =   11430
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4710
         Left            =   105
         TabIndex        =   2
         Top             =   195
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   8308
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar"
      Height          =   1755
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   11430
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   9195
         TabIndex        =   11
         Top             =   645
         Width           =   2000
      End
      Begin VB.TextBox txtSobrenome 
         Height          =   300
         Left            =   1185
         MaxLength       =   100
         TabIndex        =   10
         Top             =   990
         Width           =   5580
      End
      Begin VB.TextBox txtNome 
         Height          =   300
         Left            =   840
         MaxLength       =   50
         TabIndex        =   8
         Top             =   480
         Width           =   3180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sobrenome:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   525
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cnxn As ADODB.Connection
Public rsContato As ADODB.Recordset

' inicializa o formulario
Private Sub Form_Load()
   On Error GoTo MostraErro
   
    'recordset and connection variables
    'Dim Cnxn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strCnxn As String
    Dim strSql As String
     'record variables
    Dim strID As String
    Dim strFirstName As String
    Dim strLastName As String
   
   ' centralizar o formulario na tela
   Top = (Screen.Height - Height) / 2
   Left = (Screen.Width - Width) / 2

   ' abrir o banco de dados aqui
   MSFlexGrid1.FocusRect = flexFocusNone
   MSFlexGrid1.HighLight = flexHighlightAlways
   MSFlexGrid1.SelectionMode = flexSelectionByRow

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
    ' mostra um ampuleta para mostrar que está tabalhando
    Screen.MousePointer = vbHourglass
    MSFlexGrid1.Refresh
    'define o numero de linhas e colunas e configura o grid
   MSFlexGrid1.Rows = rs.RecordCount + 1
   MSFlexGrid1.Cols = rs.Fields.Count - 1
   MSFlexGrid1.ColWidth(0) = 1000
   MSFlexGrid1.ColWidth(1) = 3000
   MSFlexGrid1.ColWidth(2) = 4000
   MSFlexGrid1.ColWidth(3) = 3500
   MSFlexGrid1.Row = 0
   MSFlexGrid1.Col = 0
   MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
   MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
MSFlexGrid1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
MSFlexGrid1.Row = 1
MSFlexGrid1.Visible = True

'libera os objetos
Set rs = Nothing
Set db = Nothing

Screen.MousePointer = vbDefault
Exit Sub

MostraErro:
   MsgBox "Erro: " & Err.Description & " - " & Err.Source, vbApplicationModal + vbCritical + vbOKOnly, "Erro Critico"
   
End Sub

' chama o forumario de cadastro
' passando dados
Private Sub cmdAlterar_Click()
   ' Parametros aqui
   Form1.Show
End Sub

' chama funcao para excluir registro
Private Sub cmdExcluir_Click()
   On Error GoTo MostraErro
   
   Dim strSql As String
   Dim codigo As Long
   Dim linhaSelecionada As Integer
   Dim confirmacao As Integer
   
   ' Pega numero da linha selecionada
   linhaSelecionada = MSFlexGrid1.RowSel
   ' pegar o codigo do registro clicado
   codigo = MSFlexGrid1.TextMatrix(linhaSelecionada, 0)
   ' monta a query para excluir o codigo encontrado
   strSql = "delete from contatos where codigo = " & CStr(codigo)
   ' solicitar confirmacao
   confirmacao = MsgBox("Confirma a exclusão deste contato", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação")
   If confirmacao = vbYes Then
      ' executa
      Cnxn.Execute strSql
      ' mostra mensagem com resultado
      MsgBox "Contato excluido!", vbApplicationModal + vbExclamation + vbOKOnly, "Informação"
   End If
   Exit Sub

MostraErro:
   MsgBox "Erro: " & Err.Description & " - " & Err.Source, vbApplicationModal + vbCritical + vbOKOnly, "Erro"
   
End Sub

' fecha o formulario
Private Sub cmdFechar_Click()
   Unload Me
End Sub

' chama o formulario de cadastro
Private Sub cmdIncluir_Click()
   Form1.Show
   MSFlexGrid1.Refresh
   DoEvents
End Sub

' remove o formulario da memoria
Private Sub Form_Unload(Cancel As Integer)
   Set Form2 = Nothing
End Sub
