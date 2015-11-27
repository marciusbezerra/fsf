VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H8000000E&
   Caption         =   "Sistema de Controle e Gerenciamento de Associados da Funerária São Francisco"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   645
      Left            =   4125
      TabIndex        =   6
      Top             =   3960
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   4095
      TabIndex        =   5
      Top             =   3060
      Width           =   1035
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   9705
      TabIndex        =   2
      Top             =   0
      Width           =   9705
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUNERÁRIA SÃO FRANCISCO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   1650
         TabIndex        =   3
         Top             =   285
         Width           =   7455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":031C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   3  'Align Left
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   1155
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   8017
      BandCount       =   2
      Picture         =   "frmMenu.frx":0638
      ImageList       =   "ImageList1"
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   1650
      _CBHeight       =   4545
      _Version        =   "6.0.8169"
      BandForeColor1  =   -2147483628
      BandBackColor1  =   -2147483636
      Caption1        =   "&Estoque"
      Image1          =   "1"
      Child1          =   "Toolbar1"
      MinHeight1      =   1590
      Width1          =   4110
      UseCoolbarColors1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      BandForeColor2  =   -2147483634
      BandBackColor2  =   -2147483636
      Caption2        =   "Segurança"
      Image2          =   "2"
      Child2          =   "Toolbar2"
      MinHeight2      =   1590
      Width2          =   2565
      UseCoolbarColors2=   0   'False
      UseCoolbarPicture2=   0   'False
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   600
         Left            =   30
         TabIndex        =   4
         Top             =   4515
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1058
         ButtonWidth     =   2699
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "        Usuários          "
               Key             =   "USUARIOS"
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   2220
         Left            =   30
         TabIndex        =   1
         Top             =   450
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   3916
         ButtonWidth     =   2699
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clientes em estoque"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Testes de sistema"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nada"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Tudo"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmAssociados.Show vbModal
End Sub

Private Sub Command2_Click()
    frmVendendores.Show vbModal
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Dim I As Integer
    For I = 1 To CoolBar1.Bands.Count
        If CoolBar1.Bands(I).NewRow Then
            CoolBar1.Bands(I).NewRow = False
        End If
    Next
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    MostraForm frmUsuarios
    Screen.MousePointer = vbDefault
End Sub
