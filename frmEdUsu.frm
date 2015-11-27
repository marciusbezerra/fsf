VERSION 5.00
Begin VB.Form frmEdUsu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2445
      TabIndex        =   5
      Top             =   2580
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   570
      TabIndex        =   4
      Top             =   2595
      Width           =   1290
   End
   Begin VB.TextBox txtSenha 
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Top             =   1020
      Width           =   2400
   End
   Begin VB.TextBox txtNome 
      Height          =   465
      Left            =   1665
      TabIndex        =   2
      Top             =   345
      Width           =   2460
   End
   Begin VB.Label Label2 
      Caption         =   "Senha:"
      Height          =   180
      Left            =   345
      TabIndex        =   1
      Top             =   1185
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      Height          =   375
      Left            =   450
      TabIndex        =   0
      Top             =   375
      Width           =   750
   End
End
Attribute VB_Name = "frmEdUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    frmUsuarios.datUsuarios.Recordset.CancelUpdate
    With frmUsuarios.Toolbar1
        If frmUsuarios.datUsuarios.Recordset.RecordCount = 0 Then
            .Buttons("ADICIONAR").Enabled = True
            .Buttons("DELETAR").Enabled = False
            .Buttons("EDITAR").Enabled = False
            .Buttons("MUDARSENHA").Enabled = False
        Else
            .Buttons("ADICIONAR").Enabled = True
            .Buttons("DELETAR").Enabled = True
            .Buttons("EDITAR").Enabled = True
            .Buttons("MUDARSENHA").Enabled = True
        End If
    End With
    frmUsuarios.Editando = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Salvar
End Sub


Private Sub Salvar()
    If Not Valida Then Exit Sub
    frmUsuarios.datUsuarios.Recordset("Nome") = Me.txtNome.Text
    frmUsuarios.datUsuarios.Recordset("Senha") = _
                EncriptarDescriptar(Me.txtSenha.Text & "", 3, 10, 80)
    frmUsuarios.datUsuarios.UpdateRecord
    With frmUsuarios.Toolbar1
        If frmUsuarios.datUsuarios.Recordset.RecordCount = 0 Then
            .Buttons("ADICIONAR").Enabled = True
            .Buttons("DELETAR").Enabled = False
            .Buttons("EDITAR").Enabled = False
            .Buttons("MUDARSENHA").Enabled = False
        Else
            .Buttons("ADICIONAR").Enabled = True
            .Buttons("DELETAR").Enabled = True
            .Buttons("EDITAR").Enabled = True
            .Buttons("MUDARSENHA").Enabled = True
        End If
    End With
    frmUsuarios.Editando = False
    Unload Me
End Sub



Private Function Valida() As Boolean
    Valida = True
    Me.txtNome.Text = Trim(UCase(Me.txtNome.Text))
    Me.txtSenha.Text = Trim(Me.txtSenha.Text)
    If Me.txtNome.Text = "" Then
        MsgBox "O nome do usuário é requerido.", vbInformation, "Editar usuário"
        Valida = False
        Me.txtNome.SetFocus
        Exit Function
    End If
    If Me.txtSenha.Text = "" Then
        MsgBox "Para editar, você precisa informa a senha.", vbInformation, "Editar usuário"
        Valida = False
        Me.txtSenha.SetFocus
        Exit Function
    End If
    If Me.txtSenha.Text <> EncriptarDescriptar(frmUsuarios.datUsuarios.Recordset("Senha") & "", 3, 10, 80) Then
        MsgBox "A senha atual é inválida.", vbCritical, "Editar usuário"
        Valida = False
        Me.txtSenha.SetFocus
        Exit Function
    End If
    If InStr(Me.txtNome.Text, "'") > 0 Then
        MsgBox "Pela segurança das consultas na base de dados, o caractere ' não é aceito.", vbInformation, "Editar usuário"
        Me.txtNome.SetFocus
        Valida = False
        Exit Function
    End If
    
    Dim RC As Recordset
    
    Set RC = frmUsuarios.datUsuarios.Recordset.Clone
    
    RC.Index = "PrimaryKey"
    RC.Seek "=", Me.txtNome.Text
    If Not RC.NoMatch Then
        If frmUsuarios.datUsuarios.EditMode = dbEditInProgress Then
            If RC("Nome") = frmUsuarios.datUsuarios.Recordset("Nome") Then
                MsgBox "Já existe um usuário com o mesmo nome.", vbInformation, "Editar usuário"
                Valida = False
                RC.Close
                Exit Function
            End If
        End If
    End If
    RC.Close
End Function

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

