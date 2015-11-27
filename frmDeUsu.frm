VERSION 5.00
Begin VB.Form frmDeUsu 
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Não"
      Height          =   300
      Left            =   1665
      TabIndex        =   6
      Top             =   2625
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Sim"
      Height          =   330
      Left            =   255
      TabIndex        =   5
      Top             =   2580
      Width           =   1185
   End
   Begin VB.TextBox txtSenha 
      Height          =   405
      Left            =   1290
      TabIndex        =   3
      Top             =   1020
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Deseja realmente deletar ?"
      Height          =   285
      Left            =   255
      TabIndex        =   4
      Top             =   1950
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Senha:"
      Height          =   225
      Left            =   270
      TabIndex        =   2
      Top             =   1110
      Width           =   780
   End
   Begin VB.Label lblUsuario 
      Height          =   405
      Left            =   1395
      TabIndex        =   1
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Usuário:"
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   375
      Width           =   810
   End
End
Attribute VB_Name = "frmDeUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Deletar
End Sub


Private Sub Deletar()
    If Not Valida Then Exit Sub
    With frmUsuarios.datUsuarios.Recordset
        .Delete
        If .RecordCount <> 0 Then .MoveNext
        If .EOF Then .MoveLast
        If .RecordCount = 0 Then frmUsuarios.datUsuarios.Refresh
    End With
    Unload Me
End Sub



Private Function Valida() As Boolean
    Valida = True
    Me.txtSenha.Text = Trim(Me.txtSenha.Text)
    If Me.txtSenha.Text = "" Then
        MsgBox "Para deletar um usuário, você precisa informa a senha atual.", vbInformation, "Deletar usuário"
        Valida = False
        Me.txtSenha.SetFocus
        Exit Function
    End If
    If Me.txtSenha.Text <> EncriptarDescriptar(frmUsuarios.datUsuarios.Recordset("Senha") & "", 3, 10, 80) Then
        MsgBox "A senha atual é inválida.", vbCritical, "Deletar usuário"
        Valida = False
        Me.txtSenha.SetFocus
        Exit Function
    End If
End Function

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub




