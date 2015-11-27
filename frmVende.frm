VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVendendores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "tblVendedores"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7425
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   180
      ScaleHeight     =   1170
      ScaleWidth      =   3585
      TabIndex        =   31
      Top             =   5100
      Visible         =   0   'False
      Width           =   3585
      Begin VB.TextBox txtBuscNome 
         Height          =   330
         Left            =   1290
         TabIndex        =   32
         Top             =   255
         Width           =   2190
      End
      Begin MSMask.MaskEdBox txtBuscCPF 
         Height          =   330
         Left            =   1335
         TabIndex        =   36
         ToolTipText     =   "Cadastro de pessoa física"
         Top             =   630
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   315
         Left            =   225
         TabIndex        =   34
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "CPF:"
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   645
         Width           =   1215
      End
   End
   Begin VB.Data datListEstado 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Estados ORDER BY Estado"
      Top             =   3225
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7425
      TabIndex        =   24
      Top             =   6495
      Width           =   7425
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   30
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   2385
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   4755
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   6045
         TabIndex        =   27
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   210
         Left            =   3675
         TabIndex        =   26
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   240
         Left            =   1380
         TabIndex        =   25
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Data datVendedores 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblVendedores"
      Top             =   5550
      Width           =   3285
   End
   Begin VB.TextBox txtVizinho 
      DataField       =   "Vizinho"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2475
      TabIndex        =   11
      Top             =   4650
      Width           =   3375
   End
   Begin VB.TextBox txtRG 
      DataField       =   "RG"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2505
      TabIndex        =   2
      Top             =   885
      Width           =   3375
   End
   Begin VB.TextBox txtPai 
      DataField       =   "Pai"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2490
      TabIndex        =   9
      Top             =   3615
      Width           =   3375
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   105
      Width           =   3375
   End
   Begin VB.TextBox txtMae 
      DataField       =   "Mae"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2490
      TabIndex        =   10
      Top             =   3990
      Width           =   3375
   End
   Begin VB.TextBox txtEndereco 
      DataField       =   "Endereco"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2430
      TabIndex        =   5
      Top             =   1995
      Width           =   3375
   End
   Begin VB.TextBox txtCidade 
      DataField       =   "Cidade"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2430
      TabIndex        =   7
      Top             =   2790
      Width           =   3375
   End
   Begin VB.TextBox txtBairro 
      DataField       =   "Bairro"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2415
      TabIndex        =   6
      Top             =   2445
      Width           =   3375
   End
   Begin VB.TextBox txtApelido 
      DataField       =   "Apelido"
      DataSource      =   "datVendedores"
      Height          =   285
      Left            =   2460
      TabIndex        =   3
      Top             =   1290
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtpDataNasc 
      CausesValidation=   0   'False
      DataField       =   "Nascimento"
      DataSource      =   "datVendedores"
      Height          =   330
      Left            =   2505
      TabIndex        =   4
      ToolTipText     =   "Data da entrada"
      Top             =   1605
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24576001
      CurrentDate     =   36510
   End
   Begin MSMask.MaskEdBox txtCPF 
      DataField       =   "CPF"
      DataSource      =   "datVendedores"
      Height          =   330
      Left            =   2505
      TabIndex        =   1
      ToolTipText     =   "Cadastro de pessoa física"
      Top             =   450
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin MSDBCtls.DBCombo cmbEstado 
      Bindings        =   "frmVende.frx":0000
      DataField       =   "UF"
      DataSource      =   "datVendedores"
      Height          =   315
      Left            =   2430
      TabIndex        =   8
      Top             =   3165
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Estado"
      BoundColumn     =   "Sigla"
      Text            =   "DBCombo1"
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   3585
      TabIndex        =   35
      Top             =   5175
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vizinho:"
      Height          =   255
      Index           =   12
      Left            =   555
      TabIndex        =   23
      Top             =   4650
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UF:"
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   22
      Top             =   3195
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RG:"
      Height          =   255
      Index           =   10
      Left            =   585
      TabIndex        =   21
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pai:"
      Height          =   255
      Index           =   9
      Left            =   570
      TabIndex        =   20
      Top             =   3615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   19
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nascimento:"
      Height          =   255
      Index           =   7
      Left            =   570
      TabIndex        =   18
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Mae:"
      Height          =   255
      Index           =   6
      Left            =   570
      TabIndex        =   17
      Top             =   3990
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Endereco:"
      Height          =   255
      Index           =   5
      Left            =   510
      TabIndex        =   16
      Top             =   1995
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CPF:"
      Height          =   255
      Index           =   4
      Left            =   570
      TabIndex        =   15
      Top             =   510
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   2
      Left            =   495
      TabIndex        =   14
      Top             =   2775
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   1
      Left            =   495
      TabIndex        =   13
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Apelido:"
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   12
      Top             =   1290
      Width           =   1815
   End
End
Attribute VB_Name = "frmVendendores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Editando As Boolean
Const TITULO As String = "Cadastro de vendedores"

Private Sub cmdUpdate_Click()
    SalvarReg
End Sub

Private Sub datVendedores_Reposition()
    If Me.datVendedores.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    
    Me.datVendedores.Caption = "Registro " & Trim(Str(Me.datVendedores.Recordset.AbsolutePosition + 1)) & " de " & Trim(Str(Me.datVendedores.Recordset.RecordCount))
End Sub

Private Sub datVendedores_Validate(Action As Integer, Save As Integer)
    If Action = vbDataActionUnload And Editando Then
        Save = False
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And Not Editando Then
        If Me.datVendedores.Recordset.RecordCount = 0 Then
            MsgBox "Não existem registros para serem localizados.", vbInformation, TITULO
        Else
            Me.Picture1.Visible = True
            Me.Picture2.Visible = True
            Me.txtBuscNome.SetFocus
        End If
    End If
    
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA
                If Not Editando Then AdicionarReg
            Case vbKeyE
                If Not Editando Then EditarReg
            Case vbKeyD
                If Not Editando Then DeletarReg
            Case vbKeyS
                If Editando Then SalvarReg
            Case vbKeyZ
                If Editando Then CancelarReg
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Screen.ActiveControl Is Nothing Then
            If TypeOf Screen.ActiveControl Is TextBox Or _
                TypeOf Screen.ActiveControl Is DTPicker Or _
                TypeOf Screen.ActiveControl Is MaskEdBox Or _
                TypeOf Screen.ActiveControl Is DBCombo Then
                KeyAscii = 0
                SendKeys "{TAB}"
            End If
        End If
    End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    
End Sub

Private Sub Form_Load()
    AtribuiBanco Me
    Trava True
End Sub

Private Sub Trava(Travar As Boolean)
    txtNome.Locked = Travar
    txtCPF.Enabled = Not Travar
    txtRG.Locked = Travar
    txtApelido.Locked = Travar
    dtpDataNasc.Enabled = Not Travar
    txtEndereco.Locked = Travar
    txtBairro.Locked = Travar
    txtCidade.Locked = Travar
    cmbEstado.Locked = Travar
    txtPai.Locked = Travar
    txtMae.Locked = Travar
    txtVizinho.Locked = Travar
End Sub

Private Function Valida() As Boolean
    Valida = True
    If Trim(Me.txtNome.Text) = "" Then
        MsgBox "O nome do vendedor é obrigatório.", vbInformation, TITULO
        Me.txtNome.SetFocus
        Valida = False
        Exit Function
    End If
    If Trim(Me.txtCPF.Text) <> "" Then
        If Not VerificaCPF(Me.txtCPF.Text) Then
            MsgBox "O C.P.F. não está correto.", vbInformation, TITULO
            Valida = False
            Me.txtCPF.SetFocus
            Exit Function
        End If
    End If
    If Trim(Me.cmbEstado.BoundText) = "" Then
        MsgBox "Selecione pelo menos o estado.", vbInformation, TITULO
        Me.cmbEstado.SetFocus
        Valida = False
        Exit Function
    End If
    
    
    Dim C As Control
    
    For Each C In Me
        If (TypeOf C Is TextBox) Or (TypeOf C Is MaskEdBox) Then
            If Trim(C.DataField) <> "" Then
                If InStr(1, C.Text, "'") > 0 Then
                    MsgBox "Para a integridade da base de dados, o caractere ' não é válido.", vbInformation, TITULO
                    Valida = False
                    C.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next C
    
    Dim RC As Recordset
    
    Set RC = Me.datVendedores.Recordset.Clone
    
    RC.FindFirst "Nome = '" & Me.txtNome.Text & "'"
    If Not RC.NoMatch Then
        If Me.datVendedores.Recordset.EditMode = dbEditInProgress Then
            If RC("Codigo") <> Me.datVendedores.Recordset("Codigo") Then
                If MsgBox("Já existe um vendedor com este nome, não é aconselhável repetir nomes, deseja continuar salvando ?", vbQuestion + vbYesNo, TITULO) = vbNo Then
                    Valida = False
                    RC.Close
                    Exit Function
                Else
                    Valida = True
                    RC.Close
                End If
            Else
                RC.Close
            End If
        Else
            If MsgBox("Já existe um vendedor com este nome, não é aconselhável repetir nomes, deseja continuar salvando ?", vbQuestion + vbYesNo, TITULO) = vbNo Then
                Valida = False
                RC.Close
                Exit Function
            Else
                Valida = True
                RC.Close
            End If
        End If
    Else
        RC.Close
    End If
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Screen.ActiveControl Is Nothing Then
        If Screen.ActiveControl.Name <> "txtBuscNome" And _
            Screen.ActiveControl.Name <> "txtBuscCPF" And _
                Me.Picture1.Visible Then
            Me.Picture1.Visible = False
            Me.Picture2.Visible = False
        End If
    End If
End Sub

Private Sub txtBuscCPF_Change()
    On Error Resume Next
    If Me.datVendedores.Recordset.RecordCount = 0 Then Exit Sub
    If Editando Then Exit Sub
    Me.datVendedores.Recordset.FindFirst "CPF like '" & Me.txtBuscCPF.Text & "*'"
End Sub

Private Sub txtBuscCPF_GotFocus()
    Me.txtBuscCPF.Text = ""
End Sub

Private Sub txtBuscNome_Change()
    On Error Resume Next
    If Me.datVendedores.Recordset.RecordCount = 0 Then Exit Sub
    If Editando Then Exit Sub
    Me.datVendedores.Recordset.FindFirst "Nome like '*" & Me.txtBuscNome.Text & "*'"
End Sub

Private Sub txtBuscNome_GotFocus()
    Me.txtBuscNome.Text = ""
End Sub


Private Sub AdicionarReg()
    Trava False
    Me.datVendedores.Recordset.AddNew
    Me.cmdAdd.Enabled = False
    Me.cmdCancel.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.txtNome.SetFocus
    Editando = True
    Me.datVendedores.Visible = False
End Sub

Private Sub EditarReg()
    Trava False
    Me.datVendedores.Recordset.Edit
    Me.cmdAdd.Enabled = False
    Me.cmdCancel.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.txtNome.SetFocus
    Editando = True
    Me.datVendedores.Visible = False
End Sub

Private Sub DeletarReg()
    Dim MSG As String
    Dim DB As Database
    Dim RC As Recordset
    If Me.datVendedores.Recordset.RecordCount = 0 Then Exit Sub
    Set DB = Me.datVendedores.Database
    Set RC = DB.OpenRecordset("SELECT Codigo FROM tblAssociados WHERE Vendedor = " & Me.datVendedores.Recordset("Codigo"))
    If RC.RecordCount <> 0 Then
        MsgBox "Para excluir este vendedor, exclua primeiro os associados que ele atendeu.", vbInformation, TITULO
        RC.Close
        Set DB = Nothing
        Exit Sub
    End If
    RC.Close
    Set DB = Nothing
    MSG = "Deseja excluir o vendedor '" & Me.txtNome.Text & "' ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    With Me.datVendedores.Recordset
        .Delete
        If .RecordCount <> 0 Then .MoveNext
        If .EOF Then .MoveLast
        If .RecordCount = 0 Then Me.datVendedores.Refresh
    End With
End Sub

Private Sub SalvarReg()
    If Not Valida Then Exit Sub
    Me.datVendedores.UpdateRecord
    Me.datVendedores.Recordset.Bookmark = _
        Me.datVendedores.Recordset.LastModified
    Trava True
    If Me.datVendedores.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    Editando = False
    Me.datVendedores.Visible = True
End Sub

Private Sub CancelarReg()
    Dim MSG As String
    MSG = "Deseja cancelar a operação ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    Me.datVendedores.Recordset.CancelUpdate
    Trava True
    If Me.datVendedores.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancel.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    Editando = False
    Me.datVendedores.Visible = True
End Sub

Private Sub FecharReg()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    AdicionarReg
End Sub

Private Sub cmdCancel_Click()
    CancelarReg
End Sub

Private Sub cmdClose_Click()
    FecharReg
End Sub

Private Sub cmdDelete_Click()
    DeletarReg
End Sub

Private Sub cmdEditar_Click()
    EditarReg
End Sub

