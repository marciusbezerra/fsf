VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   1725
   ClientTop       =   1530
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      Begin MSComctlLib.ProgressBar PB 
         Height          =   165
         Left            =   360
         TabIndex        =   8
         Top             =   2385
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Max             =   40
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog Dialogo 
         Left            =   255
         Top             =   5460
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "SETOR DE ALMOXARIFADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   780
         Width           =   6000
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2550
         Left            =   75
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   180
         Width           =   6195
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2535
         TabIndex        =   4
         Top             =   4125
         Width           =   3720
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   4335
         Width           =   6150
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   4845
         Width           =   6165
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5610
         TabIndex        =   5
         Top             =   3765
         Width           =   660
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 95/98/NT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   4350
         TabIndex        =   6
         Top             =   3405
         Width           =   1920
      End
      Begin VB.Label lblLicenseTo 
         BackStyle       =   0  'Transparent
         Caption         =   "Licenciado para:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   105
         TabIndex        =   1
         Top             =   2790
         Width           =   6195
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   3030
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    On Error Resume Next
    lblVersion.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
    Me.lblCompany = App.CompanyName
    Me.lblCompanyProduct = "UNIVERSIDADE ESTADUAL VALE DO ACARAÚ"
    Me.lblCopyright.Caption = App.LegalCopyright
    Me.lblWarning.Caption = "Atenção: " & App.LegalTrademarks
    ''Err.Clear
    ''On Error GoTo 0
    ''Dim DB As Database
    ''Dim RC As Recordset
    'Refresh
    'On Error Resume Next
    'PB.Min = 1
    'PB.Max = 8
    'Set DB = OpenDatabase(CaminhoBanco)
    'Set RC = DB.OpenRecordset("Entradas")
    'RC.Close
    'PB.Value = 1
    'Set RC = DB.OpenRecordset("Fornecedores")
    'RC.Close
    'PB.Value = 2
    'Set RC = DB.OpenRecordset("Grupos")
    'RC.Close
    'PB.Value = 3
    'Set RC = DB.OpenRecordset("Itens")
    'RC.Close
    'PB.Value = 4
    'Set RC = DB.OpenRecordset("Saidas")
    'RC.Close
    'PB.Value = 5
    'Set RC = DB.OpenRecordset("Setores")
    'RC.Close
    'PB.Value = 6
    'Set RC = DB.OpenRecordset("SubGrupos")
    'RC.Close
    'PB.Value = 7
    'Set RC = DB.OpenRecordset("SN")
    'RC.Close
    'PB.Value = 8
    'DB.Close
    'If Err <> 0 Then
    '    MsgBox "A sua base de dados foi verificada e não está funcionando corretamente." & _
            vbCrLf & vbCrLf & "Restaure o último backup que você realizou e reinicie o sistema.", vbInformation, "Alerta"
    '    frmRest.Show vbModal
    '    End
    'End If
    frmPirat.Show
    Unload Me
End Sub

Private Sub Form_Load()
    PegaEnderecoDoBanco Me.Dialogo
    CaminhoApp = App.Path
    If Right(CaminhoApp, 1) <> "\" Then CaminhoApp = CaminhoApp & "\"
    If App.PrevInstance Then
        MsgBox "A aplicação já está rodando.", vbInformation, "Atenção"
        AppActivate App.Title
        End
    End If
End Sub
