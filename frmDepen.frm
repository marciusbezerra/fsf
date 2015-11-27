VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDependentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "tblDependentes"
   ClientHeight    =   4935
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6885
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   2265
      ScaleHeight     =   2115
      ScaleWidth      =   3585
      TabIndex        =   15
      Top             =   2265
      Visible         =   0   'False
      Width           =   3585
      Begin VB.CommandButton cmdAceita 
         Caption         =   "Aceitar"
         Height          =   360
         Left            =   60
         TabIndex        =   18
         Top             =   1725
         Width           =   1350
      End
      Begin VB.Data datListDependentes 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Temp\CONTROLE E GERENCIAMENTO DE ESTOQUE\Banco.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT DISTINCT Dependente FROM tblDependentes ORDER BY Dependente"
         Top             =   1575
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBList lstDenpendentes 
         Bindings        =   "frmDepen.frx":0000
         Height          =   1620
         Left            =   45
         TabIndex        =   17
         Top             =   60
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2858
         _Version        =   393216
         ListField       =   "Dependente"
         BoundColumn     =   "Dependente"
      End
      Begin VB.Label Label1 
         Caption         =   "ESC para cancelar"
         Height          =   240
         Left            =   1560
         TabIndex        =   19
         Top             =   1770
         Width           =   1890
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6885
      TabIndex        =   12
      Top             =   4635
      Width           =   6885
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1350
         TabIndex        =   14
         Top             =   -15
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkObito 
      DataField       =   "ObitoSN"
      Height          =   285
      Left            =   2460
      TabIndex        =   3
      Top             =   1950
      Width           =   3375
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Dependente"
      Height          =   285
      Left            =   2340
      TabIndex        =   0
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtCausa 
      DataField       =   "Causa"
      Height          =   285
      Left            =   2355
      TabIndex        =   5
      Top             =   2670
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtpDataNasc 
      CausesValidation=   0   'False
      DataField       =   "DataInicio"
      DataSource      =   "datAssociados"
      Height          =   330
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Data da entrada"
      Top             =   1095
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24444929
      CurrentDate     =   36510
   End
   Begin MSComCtl2.DTPicker dtpDataIni 
      CausesValidation=   0   'False
      DataField       =   "DataInicio"
      DataSource      =   "datAssociados"
      Height          =   330
      Left            =   2340
      TabIndex        =   2
      ToolTipText     =   "Data da entrada"
      Top             =   1515
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24444929
      CurrentDate     =   36510
   End
   Begin MSComCtl2.DTPicker dtpDataObito 
      CausesValidation=   0   'False
      DataField       =   "DataInicio"
      DataSource      =   "datAssociados"
      Height          =   330
      Left            =   2445
      TabIndex        =   4
      ToolTipText     =   "Data da entrada"
      Top             =   2325
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24444929
      CurrentDate     =   36510
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   2355
      ScaleHeight     =   2160
      ScaleWidth      =   3585
      TabIndex        =   16
      Top             =   2340
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label lblLabels 
      Caption         =   "ObitoSN:"
      Height          =   255
      Index           =   7
      Left            =   540
      TabIndex        =   11
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nascimento:"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   1185
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Inicio:"
      Height          =   255
      Index           =   5
      Left            =   420
      TabIndex        =   9
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DtObito:"
      Height          =   255
      Index           =   4
      Left            =   435
      TabIndex        =   8
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   7
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Causa:"
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   6
      Top             =   2670
      Width           =   1815
   End
End
Attribute VB_Name = "frmDependentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TITULO As String = "Adicionar ou editar dependente"

Private Sub cmdAceita_Click()
    If Trim(Me.lstDenpendentes.Text) <> "" Then
        Me.txtNome.Text = Me.lstDenpendentes.Text
    End If
    Me.Picture1.Visible = False
    Me.Picture2.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Dim MSG As String
    MSG = "Deseja cancelar a operação ?"
    If MsgBox(MSG, vbQuestion + vbYesNo, TITULO) = vbNo Then Exit Sub
    frmAssociados.datDependentes.Recordset.CancelUpdate
    frmAssociados.datDependentes.RecordSource = "select * from tblDependentes where CodAssociado = " & frmAssociados.datAssociados.Recordset("Codigo")
    frmAssociados.datDependentes.Refresh
    If frmAssociados.datDependentes.Recordset.RecordCount = 0 Then
        frmAssociados.cmdAddGrade.Enabled = True
        frmAssociados.cmdEditGrade.Enabled = False
        frmAssociados.cmdExcluiGrade.Enabled = False
    Else
        frmAssociados.cmdAddGrade.Enabled = True
        frmAssociados.cmdEditGrade.Enabled = True
        frmAssociados.cmdExcluiGrade.Enabled = True
    End If
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    If Not Valida Then Exit Sub
    With frmAssociados
        .datDependentes.Recordset("CodAssociado") = .datAssociados.Recordset("Codigo")
        .datDependentes.Recordset("Dependente") = Me.txtNome.Text
        .datDependentes.Recordset("Causa") = Me.txtCausa.Text
        .datDependentes.Recordset("Nascimento") = Me.dtpDataNasc.Value
        .datDependentes.Recordset("Inicio") = Me.dtpDataIni.Value
        .datDependentes.Recordset("ObitoSN").Value = IIf(Me.chkObito.Value = 1, True, False)
        .datDependentes.Recordset("DtObito") = Me.dtpDataObito.Value
        .datDependentes.UpdateRecord
        .datDependentes.Recordset.Bookmark = _
            .datDependentes.Recordset.LastModified
        If .datDependentes.Recordset.RecordCount = 0 Then
            .cmdAddGrade.Enabled = True
            .cmdEditGrade.Enabled = False
            .cmdExcluiGrade.Enabled = False
        Else
            .cmdAddGrade.Enabled = True
            .cmdEditGrade.Enabled = True
            .cmdExcluiGrade.Enabled = True
        End If
    End With
    Unload Me
End Sub


Private Function Valida() As Boolean
    Valida = True
    If Trim(Me.txtNome.Text) = "" Then
        MsgBox "Digite o nome do dependente ou pressione F2.", vbInformation, TITULO
        Me.txtNome.SetFocus
        Valida = False
        Exit Function
    End If
    If IsNull(Me.dtpDataIni.Value) Then
        MsgBox "A data de início é obrigatória.", vbInformation, TITULO
        Me.dtpDataIni.SetFocus
        Valida = False
        Exit Function
    End If
    If Me.chkObito.Value = 1 And IsNull(Me.dtpDataObito.Value) Then
        MsgBox "Em caso de óbito, a data do óbito é obrigatória.", vbCritical, TITULO
        Me.dtpDataObito.SetFocus
        Valida = False
        Exit Function
    End If
    If Not IsNull(Me.dtpDataObito.Value) And Me.chkObito.Value = 0 Then
        MsgBox "Se você inseriu uma data de óbito, você deve marcar a caixa de seleção Óbito.", vbCritical, TITULO
        Me.chkObito.SetFocus
        Valida = False
        Exit Function
    End If
End Function

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    AtribuiBanco Me
End Sub

Private Sub lstDenpendentes_DblClick()
    cmdAceita_Click
End Sub

Private Sub txtNome_Change()
    If Me.Picture1.Visible Then
        If Me.datListDependentes.Recordset.RecordCount <> 0 Then
            Me.datListDependentes.Recordset.FindFirst "Dependente like '" & Me.txtNome.Text & "*'"
            If Me.datListDependentes.Recordset.NoMatch = False Then
                Me.lstDenpendentes.Text = _
                    Me.datListDependentes.Recordset("Dependente")
            End If
        End If
    End If
End Sub

Private Sub txtNome_LostFocus()
    If Me.Picture1.Visible Then
        Me.Picture1.Visible = False
        Me.Picture2.Visible = False
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
            Me.Picture1.Visible = True
            Me.Picture2.Visible = True
            Me.txtNome.SetFocus
    End If
    
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyS
                cmdUpdate_Click
            Case vbKeyZ
                cmdCancel_Click
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.Picture1.Visible Then
        cmdAceita_Click
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Not Screen.ActiveControl Is Nothing Then
            If TypeOf Screen.ActiveControl Is TextBox Or _
                TypeOf Screen.ActiveControl Is DTPicker Then
                KeyAscii = 0
                SendKeys "{TAB}"
            End If
        End If
    End If
    If KeyAscii = 27 And Me.Picture1.Visible Then
        Me.Picture1.Visible = False
        Me.Picture2.Visible = False
        KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    
End Sub
