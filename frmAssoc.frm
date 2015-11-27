VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmAssociados 
   Caption         =   "tblAssociados"
   ClientHeight    =   8190
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10185
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10185
      TabIndex        =   39
      Top             =   7890
      Width           =   10185
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   2385
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   4755
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   6045
         TabIndex        =   42
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   210
         Left            =   3675
         TabIndex        =   41
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   240
         Left            =   1380
         TabIndex        =   40
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   5955
      ScaleHeight     =   1170
      ScaleWidth      =   3585
      TabIndex        =   33
      Top             =   450
      Visible         =   0   'False
      Width           =   3585
      Begin VB.TextBox txtBuscApelido 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   615
         Width           =   2160
      End
      Begin VB.TextBox txtBuscNome 
         Height          =   330
         Left            =   1290
         TabIndex        =   36
         Top             =   255
         Width           =   2190
      End
      Begin VB.Label Label2 
         Caption         =   "Apelido:"
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   315
         Left            =   225
         TabIndex        =   34
         Top             =   210
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   6045
      ScaleHeight     =   1200
      ScaleWidth      =   3585
      TabIndex        =   32
      Top             =   525
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Data datListEstado 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5955
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Estados ORDER BY Estado"
      Top             =   4470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datListVendedor 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Codigo, Nome & ' (' & CPF & ')' As NomeCPF FROM tblVendedores ORDER BY Nome & ' (' & CPF & ')'"
      Top             =   795
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSDBGrid.DBGrid grdDependentes 
      Bindings        =   "frmAssoc.frx":0000
      Height          =   1350
      Left            =   75
      OleObjectBlob   =   "frmAssoc.frx":001D
      TabIndex        =   31
      Top             =   5340
      Width           =   6720
   End
   Begin VB.Data datDependentes 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblDependentes"
      Top             =   6450
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datAssociados 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4140
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblAssociados"
      Top             =   2625
      Width           =   3225
   End
   Begin VB.CommandButton cmdExcluiGrade 
      Caption         =   "Excluir"
      Height          =   300
      Left            =   2955
      TabIndex        =   30
      Top             =   6840
      Width           =   1680
   End
   Begin VB.CommandButton cmdEditGrade 
      Caption         =   "Editar"
      Height          =   255
      Left            =   1545
      TabIndex        =   29
      Top             =   6900
      Width           =   1170
   End
   Begin VB.CommandButton cmdAddGrade 
      Caption         =   "Adicionar"
      Height          =   225
      Left            =   150
      TabIndex        =   28
      Top             =   6915
      Width           =   1155
   End
   Begin MSDBCtls.DBCombo cmbVendedor 
      Bindings        =   "frmAssoc.frx":10CB
      DataField       =   "Vendedor"
      DataSource      =   "datAssociados"
      Height          =   315
      Left            =   2025
      TabIndex        =   0
      Top             =   780
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "NomeCPF"
      BoundColumn     =   "Codigo"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo cmbEstado 
      Bindings        =   "frmAssoc.frx":10E9
      DataField       =   "Uf"
      DataSource      =   "datAssociados"
      Height          =   315
      Left            =   2070
      TabIndex        =   11
      Top             =   4410
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
   Begin VB.TextBox txtVizinho 
      DataField       =   "Vizinho"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2145
      TabIndex        =   12
      Top             =   4755
      Width           =   3375
   End
   Begin VB.TextBox txtProfissao 
      DataField       =   "Profissao"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2100
      TabIndex        =   7
      Top             =   3165
      Width           =   3375
   End
   Begin VB.TextBox txtPlano 
      DataField       =   "Plano"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2085
      TabIndex        =   4
      Top             =   2115
      Width           =   3375
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2070
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtEndereco 
      DataField       =   "Endereco"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2145
      TabIndex        =   8
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtContrato 
      DataField       =   "Contrato"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2115
      TabIndex        =   3
      Top             =   1830
      Width           =   3375
   End
   Begin VB.TextBox txtCidade 
      DataField       =   "Cidade"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2055
      TabIndex        =   10
      Top             =   4125
      Width           =   3375
   End
   Begin VB.TextBox txtBairro 
      DataField       =   "Bairro"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   3810
      Width           =   3375
   End
   Begin VB.TextBox txtApelido 
      DataField       =   "Apelido"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2055
      TabIndex        =   2
      Top             =   1530
      Width           =   3375
   End
   Begin VB.TextBox txtOperador 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "Operador"
      DataSource      =   "datAssociados"
      Height          =   285
      Left            =   2055
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   380
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtpDataInicio 
      CausesValidation=   0   'False
      DataField       =   "DataInicio"
      DataSource      =   "datAssociados"
      Height          =   330
      Left            =   2085
      TabIndex        =   5
      ToolTipText     =   "Data da entrada"
      Top             =   2415
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24510465
      CurrentDate     =   36510
   End
   Begin MSComCtl2.DTPicker dtpDataNasc 
      CausesValidation=   0   'False
      DataField       =   "Nascimento"
      DataSource      =   "datAssociados"
      Height          =   330
      Left            =   2130
      TabIndex        =   6
      ToolTipText     =   "Data da entrada"
      Top             =   2790
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24510465
      CurrentDate     =   36510
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      Height          =   300
      Left            =   5565
      TabIndex        =   38
      Top             =   2130
      Width           =   405
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vizinho:"
      Height          =   255
      Index           =   14
      Left            =   225
      TabIndex        =   27
      Top             =   4755
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Vendedor:"
      Height          =   255
      Index           =   13
      Left            =   225
      TabIndex        =   26
      Top             =   750
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Uf:"
      Height          =   255
      Index           =   12
      Left            =   135
      TabIndex        =   25
      Top             =   4410
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Profissao:"
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   24
      Top             =   3165
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Plano:"
      Height          =   255
      Index           =   10
      Left            =   165
      TabIndex        =   23
      Top             =   2115
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   9
      Left            =   135
      TabIndex        =   22
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nascimento:"
      Height          =   255
      Index           =   8
      Left            =   195
      TabIndex        =   21
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Endereco:"
      Height          =   255
      Index           =   7
      Left            =   210
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DataInicio:"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   19
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Contrato:"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   18
      Top             =   1830
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   4125
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3810
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Apelido:"
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   15
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Operador:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   380
      Width           =   1815
   End
End
Attribute VB_Name = "frmAssociados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Editando As Boolean
Const TITULO As String = "Cadastro de associados"

Private Sub cmdAdd_Click()
    AdicionarReg
End Sub

Private Sub cmdAddGrade_Click()
    Screen.MousePointer = 11
    Me.datDependentes.Recordset.AddNew
    Editando = True
    frmDependentes.dtpDataIni.Value = Date
    frmDependentes.Show vbModal
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

Private Sub cmdEditGrade_Click()
    Screen.MousePointer = 11
    Me.datDependentes.Recordset.Edit
    Editando = True
    With frmDependentes
        .txtNome.Text = Me.datDependentes.Recordset("Dependente") & ""
        .txtCausa.Text = Me.datDependentes.Recordset("Causa") & ""
        .dtpDataNasc.Value = Me.datDependentes.Recordset("Nascimento") & ""
        .dtpDataIni.Value = Me.datDependentes.Recordset("Inicio") & ""
        .chkObito.Value = IIf(Me.datDependentes.Recordset("ObitoSN").Value = True, 1, 0)
        .dtpDataObito.Value = Me.datDependentes.Recordset("DtObito") & ""
    End With
    frmDependentes.Show vbModal
End Sub

Private Sub cmdExcluiGrade_Click()
    Dim MSG As String
    If Me.datDependentes.Recordset.RecordCount = 0 Then Exit Sub
    MSG = "Deseja excluir o dependente '" & Me.datDependentes.Recordset("Dependente") & "' ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    With Me.datDependentes.Recordset
        .Delete
        If .RecordCount <> 0 Then .MoveNext
        If .EOF Then .MoveLast
        If .RecordCount = 0 Then Me.datDependentes.Refresh
    End With
    If Me.datDependentes.Recordset.RecordCount = 0 Then
        Me.cmdAddGrade.Enabled = True
        Me.cmdEditGrade.Enabled = False
        Me.cmdExcluiGrade.Enabled = False
    Else
        Me.cmdAddGrade.Enabled = True
        Me.cmdEditGrade.Enabled = True
        Me.cmdExcluiGrade.Enabled = True
    End If
End Sub

Private Sub cmdUpdate_Click()
    SalvarReg
End Sub

Private Sub datAssociados_Reposition()
    If Me.datAssociados.Recordset.RecordCount = 0 Then
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
    
    Me.datDependentes.RecordSource = "select * from tblDependentes where CodAssociado = " & Me.datAssociados.Recordset("Codigo")
    Me.datDependentes.Refresh
    
    If Me.datDependentes.Recordset.RecordCount = 0 Then
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = False
        Me.cmdExcluiGrade.Enabled = False
    Else
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdExcluiGrade.Enabled = Me.cmdEditar.Enabled
    End If
    Me.datAssociados.Caption = "Registro " & Trim(Str(Me.datAssociados.Recordset.AbsolutePosition + 1)) & " de " & Trim(Str(Me.datAssociados.Recordset.RecordCount))
End Sub

Private Sub datAssociados_Validate(Action As Integer, Save As Integer)
    If Action = vbDataActionUnload And Editando Then
        Save = False
    End If
End Sub

Private Sub Trava(Travar As Boolean)
    cmbVendedor.Locked = Travar
    txtNome.Locked = Travar
    txtApelido.Locked = Travar
    txtContrato.Locked = Travar
    txtPlano.Locked = Travar
    dtpDataInicio.Enabled = Not Travar
    dtpDataNasc.Enabled = Not Travar
    txtProfissao.Locked = Travar
    txtEndereco.Locked = Travar
    txtBairro.Locked = Travar
    txtCidade.Locked = Travar
    cmbEstado.Locked = Travar
    txtVizinho.Locked = Travar
End Sub

Private Function Valida() As Boolean
    Valida = True
    If Trim(Me.cmbVendedor.BoundText) = "" Then
        MsgBox "Selecione um vendedor na lista.", vbInformation, TITULO
        Me.cmbVendedor.SetFocus
        Valida = False
        Exit Function
    End If
    If Trim(Me.txtNome.Text) = "" Then
        MsgBox "Digite o nome do associado.", vbInformation, TITULO
        Me.txtNome.SetFocus
        Valida = False
        Exit Function
    End If
    If Trim(Me.txtContrato.Text) = "" Then
        MsgBox "Digite o código do contrato.", vbCritical, TITULO
        Me.txtContrato.SetFocus
        Valida = False
        Exit Function
    End If
    If Not IsNumeric(Me.txtPlano) Then
        MsgBox "Digite um valor numérico para o plano.", vbCritical, TITULO
        Me.txtPlano.SetFocus
        Valida = False
        Exit Function
    End If
    If IsNull(Me.dtpDataInicio.Value) Then
        MsgBox "Informe a data de início.", vbCritical, TITULO
        Valida = False
        Me.dtpDataInicio.SetFocus
        Exit Function
    End If
    If Frac(Me.txtPlano) > 0 Then
        MsgBox "Informe um valor inteiro para o plano.", vbCritical, TITULO
        Valida = False
        Me.txtPlano.SetFocus
        Exit Function
    End If
    If Trim(Me.cmbEstado.BoundText) = "" Then
        MsgBox "Selecione um estado na lista.", vbCritical, TITULO
        Me.cmbEstado.SetFocus
        Valida = False
        Exit Function
    End If
    
    
    Dim C As Control
    
    For Each C In Me
        If TypeOf C Is TextBox Then
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
    
    Set RC = Me.datAssociados.Recordset.Clone
    
    RC.FindFirst "Nome = '" & Me.txtNome.Text & "'"
    If Not RC.NoMatch Then
        If Me.datAssociados.Recordset.EditMode = dbEditInProgress Then
            If RC("Codigo") <> Me.datAssociados.Recordset("Codigo") Then
                If MsgBox("Já existe um associado com este nome, não é aconselhável repetir nomes, deseja continuar salvando ?", vbQuestion + vbYesNo, TITULO) = vbNo Then
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
            If MsgBox("Já existe um associado com este nome, não é aconselhável repetir nomes, deseja continuar salvando ?", vbQuestion + vbYesNo, TITULO) = vbNo Then
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And Not Editando Then
        If Me.datAssociados.Recordset.RecordCount = 0 Then
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Screen.ActiveControl Is Nothing Then
        If Screen.ActiveControl.Name <> "txtBuscNome" And _
            Screen.ActiveControl.Name <> "txtBuscApelido" And _
                Me.Picture1.Visible Then
            Me.Picture1.Visible = False
            Me.Picture2.Visible = False
        End If
    End If
End Sub

Private Sub grdDependentes_DblClick()
    If Me.cmdEditGrade.Enabled Then
        cmdEditGrade_Click
    End If
End Sub

Private Sub txtBuscApelido_Change()
    On Error Resume Next
    If Me.datAssociados.Recordset.RecordCount = 0 Then Exit Sub
    If Editando Then Exit Sub
    Me.datAssociados.Recordset.FindFirst "Apelido like '*" & Me.txtBuscApelido.Text & "*'"
End Sub

Private Sub txtBuscApelido_GotFocus()
    Me.txtBuscNome.Text = ""
End Sub

Private Sub txtBuscNome_Change()
    On Error Resume Next
    If Me.datAssociados.Recordset.RecordCount = 0 Then Exit Sub
    If Editando Then Exit Sub
    Me.datAssociados.Recordset.FindFirst "Nome like '*" & Me.txtBuscNome.Text & "*'"
End Sub

Private Sub txtBuscNome_GotFocus()
    Me.txtBuscApelido.Text = ""
End Sub


Private Sub AdicionarReg()
    Trava False
    Me.datAssociados.Recordset.AddNew
    Me.cmdAdd.Enabled = False
    Me.cmdCancel.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.dtpDataInicio.Value = Date
    Me.cmbVendedor.SetFocus
    Editando = True
    Me.datAssociados.Visible = False
    Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
    Me.cmdEditGrade.Enabled = Me.cmdEditar.Enabled
    Me.cmdExcluiGrade.Enabled = Me.cmdEditar.Enabled
    Me.txtOperador.Text = AtualUsuario
End Sub

Private Sub EditarReg()
    On Error GoTo Erro
    Trava False
    Me.datAssociados.Recordset.Edit
    Me.cmdAdd.Enabled = False
    Me.cmdCancel.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.cmbVendedor.SetFocus
    Editando = True
    Me.datAssociados.Visible = False
    Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
    Me.cmdEditGrade.Enabled = Me.cmdEditar.Enabled
    Me.cmdExcluiGrade.Enabled = Me.cmdEditar.Enabled
    Exit Sub
Erro:
    Dim X As Long
    If Err = 3197 Then
        MsgBox "Este registro acabou de ser modificado por outro usuário na rede.", vbInformation, TITULO
        Resume
        Me.datAssociados.Recordset.Move 0
        Resume
    End If
    If Err = 3167 Then
        MsgBox "Este registro já foi excluído por outro usuário na rede.", vbInformation, TITULO
        With Me.datAssociados.Recordset
            .Move 0
            If .RecordCount <> 0 Then .MoveNext
            If .EOF Then .MoveLast
            If .RecordCount = 0 Then Me.datAssociados.Refresh
        End With
    Else
        MsgBox Error$ & " " & Err
        Stop
    End If
End Sub

Private Sub DeletarReg()
    Dim MSG As String
    If Me.datAssociados.Recordset.RecordCount = 0 Then Exit Sub
    If Me.datDependentes.Recordset.RecordCount <> 0 Then
        MsgBox "Para excluir este associados, exclua primeiro os seus dependentes.", vbInformation, TITULO
        Exit Sub
    End If
    MSG = "Deseja excluir o associado '" & Me.txtNome.Text & "' ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    With Me.datAssociados.Recordset
        .Delete
        If .RecordCount <> 0 Then .MoveNext
        If .EOF Then .MoveLast
        If .RecordCount = 0 Then Me.datAssociados.Refresh
    End With
End Sub

Private Sub SalvarReg()
    If Not Valida Then Exit Sub
    Me.datAssociados.UpdateRecord
    Me.datAssociados.Recordset.Bookmark = _
        Me.datAssociados.Recordset.LastModified
    Trava True
    If Me.datAssociados.Recordset.RecordCount = 0 Then
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
    Me.datAssociados.Visible = True
    If Me.datDependentes.Recordset.RecordCount = 0 Then
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = False
        Me.cmdExcluiGrade.Enabled = False
    Else
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdExcluiGrade.Enabled = Me.cmdEditar.Enabled
    End If
End Sub

Private Sub CancelarReg()
    Dim MSG As String
    MSG = "Deseja cancelar a operação ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    Me.datAssociados.Recordset.CancelUpdate
    Trava True
    If Me.datAssociados.Recordset.RecordCount = 0 Then
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
    Me.datAssociados.Visible = True
    Me.datDependentes.RecordSource = "select * from tblDependentes where CodAssociado = " & Me.datAssociados.Recordset("Codigo")
    Me.datDependentes.Refresh
    If Me.datDependentes.Recordset.RecordCount = 0 Then
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = False
        Me.cmdExcluiGrade.Enabled = False
    Else
        Me.cmdAddGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdEditGrade.Enabled = Me.cmdEditar.Enabled
        Me.cmdExcluiGrade.Enabled = Me.cmdEditar.Enabled
    End If
End Sub

Private Sub FecharReg()
    Screen.MousePointer = 0
    Unload Me
End Sub
