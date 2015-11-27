VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPirat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANTI-PIRATA"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPirat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   4515
      TabIndex        =   8
      ToolTipText     =   "Clique para terminar a aplicação"
      Top             =   2055
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2730
      TabIndex        =   7
      ToolTipText     =   "Clique para confirmar a contra senha"
      Top             =   2055
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contra-senha"
      Height          =   795
      Left            =   1365
      TabIndex        =   2
      Top             =   1140
      Width           =   4350
      Begin MSMask.MaskEdBox T1 
         Height          =   330
         Left            =   225
         TabIndex        =   3
         ToolTipText     =   "Digite a contra-senha dada pelo programador"
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox T2 
         Height          =   330
         Left            =   1125
         TabIndex        =   4
         ToolTipText     =   "Digite a contra-senha dada pelo programador"
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox T3 
         Height          =   330
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Digite a contra-senha dada pelo programador"
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox T4 
         Height          =   330
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Digite a contra-senha dada pelo programador"
         Top             =   300
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2790
         TabIndex        =   11
         Top             =   285
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1830
         TabIndex        =   10
         Top             =   285
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   9
         Top             =   285
         Width           =   105
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmPirat.frx":0442
      Top             =   315
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Consulte o programador sobre qual é a contra-senha informando o número do seu CIC."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1335
      TabIndex        =   1
      Top             =   525
      Width           =   4380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Controle ante-pirataria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmPirat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TENTATIVAS As Integer

Private Sub Command1_Click()
'    If (T1.Text & T2.Text & T3.Text & T4.Text) <> GeraSenha Then
'        MsgBox "Contra-senha incorreta.", vbCritical, Caption
'        TENTATIVAS = TENTATIVAS + 1
'        If TENTATIVAS = 3 Then End
'    Else
        Unload Me
        EscreverIni "BANCO", "Erros na base de dados", "1"
        On Error Resume Next
        frmLogin.Show
'    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Secreto As String
    If Trim(LerIni("BANCO", "Erros na base de dados", "")) = "1" Then
        frmLogin.Show
        Unload Me
    End If
End Sub

Private Sub T1_GotFocus()
    Seleciona Me.T1
End Sub

Private Sub T2_GotFocus()
    Seleciona Me.T2
End Sub

Private Sub T3_GotFocus()
    Seleciona Me.T3
End Sub

Private Sub T4_GotFocus()
    Seleciona Me.T4
End Sub
