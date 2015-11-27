VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmUsuarios 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmUsuarios.frx":0000
      Height          =   3300
      Left            =   1815
      OleObjectBlob   =   "frmUsuarios.frx":001A
      TabIndex        =   5
      Top             =   1260
      Width           =   3405
   End
   Begin VB.Data datUsuarios 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "SN"
      Top             =   4005
      Visible         =   0   'False
      Width           =   1140
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
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   120
         X2              =   10515
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "USUÁRIOS"
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
         TabIndex        =   1
         Top             =   225
         Width           =   1470
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   3  'Align Left
      Height          =   5040
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   8890
      BandCount       =   1
      Picture         =   "frmUsuarios.frx":0849
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   1260
      _CBHeight       =   5040
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
         TabIndex        =   3
         Top             =   30
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1058
         ButtonWidth     =   2064
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adicionar"
               Key             =   "ADICIONAR"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mudar nome"
               Key             =   "EDITAR"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Deletar"
               Key             =   "DELETAR"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mudar a senha"
               Key             =   "MUDARSENHA"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Fechar"
               Key             =   "SAIR"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listagem de usuários cadastrados:"
      Height          =   210
      Left            =   1845
      TabIndex        =   4
      Top             =   885
      Width           =   2550
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Editando As Boolean

Private Sub datUsuarios_Error(DataErr As Integer, Response As Integer)
    MsgBox "Ocorreu o erro " & DataErr & ":" & Error$(DataErr) & _
        vbCrLf & vbCrLf & "Entre em contato com o fornecedor.", vbCritical, "Usuários"
    Response = 0
End Sub

Private Sub datUsuarios_Reposition()
    With Me.Toolbar1
        If Me.datUsuarios.Recordset.RecordCount = 0 Then
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

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    AtribuiBanco Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "SAIR"
        Screen.MousePointer = 0
        Unload Me
    Case "ADICIONAR"
        If AtualUsuario <> "ADMINISTRADOR" Then
            MsgBox "Você precisa entrar como o usuário ADMINISTRADOR para ter permissões de incluir um novo usuário.", vbCritical, "Novo usuário"
            Exit Sub
        End If
        Screen.MousePointer = 11
        MostraForm frmAdUsu
    Case "EDITAR"
        If Me.datUsuarios.Recordset.RecordCount = 0 Then _
            Exit Sub
        If UCase(Me.datUsuarios.Recordset("Nome")) = "ADMINISTRADOR" Then
            MsgBox "O usuário administrador não pode ser editado.", vbInformation, "Editar"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Me.datUsuarios.Recordset.Edit
        Editando = True
        frmEdUsu.txtNome.Text = Me.datUsuarios.Recordset("Nome")
        frmEdUsu.Show vbModal 'A
    Case "DELETAR"
        If Me.datUsuarios.Recordset.RecordCount = 0 Then _
            Exit Sub
        If UCase(Me.datUsuarios.Recordset("Nome")) = "ADMINISTRADOR" Then
            MsgBox "O usuário administrador não pode ser deletado.", vbInformation, "Deletar"
            Exit Sub
        End If
        Screen.MousePointer = 11
        frmDeUsu.lblUsuario.Caption = Me.datUsuarios.Recordset("Nome")
        frmDeUsu.Show vbModal 'A
    Case "MUDARSENHA"
        If Me.datUsuarios.Recordset.RecordCount = 0 Then _
            Exit Sub
        Screen.MousePointer = 11
        Me.datUsuarios.Recordset.Edit
        Editando = True
        frmSenha.lblUsuario.Caption = Me.datUsuarios.Recordset("Nome")
        frmSenha.Show vbModal   'A
    End Select
End Sub


