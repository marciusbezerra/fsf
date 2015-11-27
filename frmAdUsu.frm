VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAdUsu 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConfirma 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2910
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3465
      Width           =   4275
   End
   Begin VB.TextBox txtSenha 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2910
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2805
      Width           =   4275
   End
   Begin VB.TextBox txtNome 
      Height          =   330
      Left            =   2910
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2145
      Width           =   4275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1275
      ScaleHeight     =   690
      ScaleWidth      =   10620
      TabIndex        =   3
      Top             =   0
      Width           =   10620
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "NOVO USUÁRIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   225
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   120
         X2              =   10515
         Y1              =   90
         Y2              =   90
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   3  'Align Left
      Height          =   5880
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   10372
      BandCount       =   1
      Picture         =   "frmAdUsu.frx":0000
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   1260
      _CBHeight       =   5880
      _Version        =   "6.0.8169"
      BandForeColor1  =   -2147483628
      BandBackColor1  =   -2147483636
      Child1          =   "Toolbar1"
      MinHeight1      =   1200
      Width1          =   1350
      UseCoolbarColors1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1058
         ButtonWidth     =   1931
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "     Salvar    "
               Key             =   "SALVAR"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "    Cancelar    "
               Key             =   "CANCELAR"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o seu nome, a sua senha e confirme-a depois:"
      Height          =   210
      Left            =   1800
      TabIndex        =   10
      Top             =   1155
      Width           =   3795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmação:"
      Height          =   210
      Left            =   1800
      TabIndex        =   9
      Top             =   3525
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      Height          =   210
      Left            =   1800
      TabIndex        =   8
      Top             =   2865
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   210
      Left            =   1800
      TabIndex        =   7
      Top             =   2205
      Width           =   450
   End
End
Attribute VB_Name = "frmAdUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CANCELAR": Unload Me
        Case "SALVAR"
            Salvar
    End Select
End Sub



Private Sub Salvar()
    If Trim(Me.txtNome.Text) = "" Then
        MsgBox "O Nome é requerido.", vbCritical, "Novo usuário"
        Me.txtNome.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtSenha.Text) = "" Then
        MsgBox "É obrigatório o uso da senha.", vbInformation, "Novo usuário"
        Me.txtSenha.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtSenha.Text) <> Trim(Me.txtConfirma.Text) Then
        MsgBox "A senha não confirma.", vbCritical, "Novo usuário"
        Me.txtConfirma.SetFocus
        Exit Sub
    End If
    Me.txtNome.Text = Trim(UCase(Me.txtNome.Text))
    Me.txtSenha.Text = Trim(Me.txtSenha.Text)
    If InStr(Me.txtNome.Text, "'") > 0 Then
        MsgBox "Pela segurança das consultas na base de dados, o caractere ' não é aceito.", vbInformation, "Novo usuário"
        Me.txtNome.SetFocus
        Exit Sub
    End If
    If frmUsuarios.datUsuarios.Recordset.RecordCount <> 0 Then
        frmUsuarios.datUsuarios.Recordset.Index = "PrimaryKey"
        frmUsuarios.datUsuarios.Recordset.Seek "=", Me.txtNome.Text
        If Not frmUsuarios.datUsuarios.Recordset.NoMatch Then
            MsgBox "Já existe usuário com este nome.", vbInformation, "Novo usuário"
            Me.txtNome.SetFocus
            Exit Sub
        End If
    End If
    With frmUsuarios.datUsuarios.Recordset
        .AddNew
            !Nome = Me.txtNome
            !SENHA = _
                EncriptarDescriptar(Me.txtSenha.Text & "", 3, 10, 80)
        frmUsuarios.datUsuarios.UpdateRecord
        .Bookmark = .LastModified
        If .RecordCount = 0 Then
            frmUsuarios.Toolbar1.Buttons("ADICIONAR").Enabled = True
            frmUsuarios.Toolbar1.Buttons("DELETAR").Enabled = False
            frmUsuarios.Toolbar1.Buttons("EDITAR").Enabled = False
            frmUsuarios.Toolbar1.Buttons("MUDARSENHA").Enabled = False
        Else
            frmUsuarios.Toolbar1.Buttons("ADICIONAR").Enabled = True
            frmUsuarios.Toolbar1.Buttons("DELETAR").Enabled = True
            frmUsuarios.Toolbar1.Buttons("EDITAR").Enabled = True
            frmUsuarios.Toolbar1.Buttons("MUDARSENHA").Enabled = True
        End If
    End With
    Unload Me
End Sub
