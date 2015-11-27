VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificação"
   ClientHeight    =   3930
   ClientLeft      =   3795
   ClientTop       =   2910
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2321.975
   ScaleMode       =   0  'User
   ScaleWidth      =   6168.875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datUsuarios 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome FROM SN ORDER BY Nome"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo cmbUsuarios 
      Bindings        =   "frmLogin.frx":0442
      Height          =   330
      Left            =   2115
      TabIndex        =   6
      Top             =   885
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Nome"
      BoundColumn     =   "Nome"
      Text            =   "DBCombo1"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2835
      TabIndex        =   2
      ToolTipText     =   "Clique para confirmar a senha"
      Top             =   3030
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   390
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Clique para finalizar a aplicação"
      Top             =   3030
      Width           =   1245
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3555
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Digite a senha do sistema"
      Top             =   2325
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      Height          =   210
      Left            =   1275
      TabIndex        =   5
      Top             =   825
      Width           =   600
   End
   Begin VB.Label lblTentativas 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   135
      TabIndex        =   4
      Top             =   1155
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   -15
      Picture         =   "frmLogin.frx":045C
      Stretch         =   -1  'True
      Top             =   45
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   2370
      TabIndex        =   0
      Top             =   2355
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SENHA As String
Dim TENTATIVAS As Integer
Dim DB As Database
Dim RC As Recordset


Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim SSS As String
    If Trim(Me.cmbUsuarios.BoundText) = "" Then
        MsgBox "Selecione um usuário na lista.", vbInformation, "Identificação"
        Me.cmbUsuarios.SetFocus
        Exit Sub
    End If
    If Not PegaSenhaDoUsuario(Me.cmbUsuarios.BoundText) Then
        MsgBox "O usuário não foi localizado, talvez tenha sido excluído.", vbInformation, "Identificação"
        Me.cmbUsuarios.SetFocus
        Exit Sub
    End If
    If Me.txtPassword.Text = SENHA Then
        MsgBox "Senha Incorreta.", vbCritical, "Identificação"
        TENTATIVAS = TENTATIVAS + 1
        Me.lblTentativas.Caption = "Tentativa: " & TENTATIVAS + 1
        If TENTATIVAS = 3 Then End
        Me.txtPassword.SetFocus
        Me.txtPassword.SelStart = 0
        Me.txtPassword.SelLength = Len(Me.txtPassword.Text)
    Else
        EscreverIni "Identificação", "Usuário", Me.cmbUsuarios.BoundText
        Unload Me
        frmMenu.Show
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    AtribuiBanco Me
    Set DB = OpenDatabase(CaminhoBanco)
    Set RC = DB.OpenRecordset("SN", dbOpenTable)
    If Err = 3011 Then
        MsgBox "O REGISTRO QUE CONTÉM AS SENHAS DO SISTEMA" & vbCrLf & _
            "FOI APAGADO. SEM SEGURANÇA O PROGRAMA NÃO PODE RODAR!", vbExclamation, "ATENÇÃO"
        DB.Close
        End
    End If
    If RC.RecordCount = 0 Then
        MsgBox "O REGISTRO QUE CONTÉM AS SENHAS DO SISTEMA" & vbCrLf & _
            "FOI APAGADO. SEM SEGURANÇA O PROGRAMA NÃO PODE RODAR!", vbExclamation, "ATENÇÃO"
        RC.Close
        DB.Close
        End
    End If
    On Error GoTo 0
    Me.lblTentativas.Caption = "Tentativa: " & TENTATIVAS + 1
    On Error Resume Next
    Me.cmbUsuarios.BoundText = LerIni("Identificação", "Usuário", "ADMINISTRADOR")
    If Err <> 0 Then
        Me.cmbUsuarios.BoundText = "ADMINISTRADOR"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RC.Close
    DB.Close
End Sub

Private Sub txtPassword_GotFocus()
    Seleciona Me.txtPassword
End Sub

Private Function PegaSenhaDoUsuario(Usuario As String) As Boolean
    On Error GoTo Erro
    PegaSenhaDoUsuario = True
    RC.Index = "PrimaryKey"
    RC.Seek "=", Me.cmbUsuarios.BoundText
    If RC.NoMatch Then
        PegaSenhaDoUsuario = False
    Else
        SENHA = EncriptarDescriptar(RC("Senha") & "", 3, 10, 80)
        AtualUsuario = Usuario
        PegaSenhaDoUsuario = True
    End If
    Exit Function
Erro:
    PegaSenhaDoUsuario = False
End Function
