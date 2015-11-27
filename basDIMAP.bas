Attribute VB_Name = "basDIMAP"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function NadaFaz Lib "areceber.rpt" ()
Private Declare Function NadaFaz1 Lib "poestoqu.rpt" ()
Private Declare Function NadaFaz2 Lib "infminim.rpt" ()
Private Declare Function NadaFaz3 Lib "naoestoc.rpt" ()
Private Declare Function NadaFaz4 Lib "balfin.rpt" ()
Private Declare Function NadaFaz5 Lib "ultent.rpt" ()
Private Declare Function NadaFaz6 Lib "conger.rpt" ()
Private Declare Function NadaFaz7 Lib "forneced.rpt" ()
Private Declare Function NadaFaz8 Lib "setores.rpt" ()
Private Declare Function NadaFaz9 Lib "estper.rpt" ()
Private Declare Function NadaFaz10 Lib "lissub.rpt" ()
Private Declare Function NadaFaz11 Lib "conper.rpt" ()
Private Declare Function NadaFaz12 Lib "recomse.rpt" ()
Private Declare Function NadaFaz13 Lib "itemper.rpt" ()
Private Declare Function NadaFaz14 Lib "conseord.rpt" ()

Enum Orientação
    Vertical = 1
    Horizontal = 2
End Enum

Public AtualUsuario As String 'USADA
Public CaminhoApp As String 'USADA
Public CaminhoBanco As String 'USADA
Public Mensagens(1 To 10)  As String

Public Sub Mens(Texto As String)
    frmMain.StatusBar1.Panels(1).Text = Texto
End Sub

Public Sub Seleciona(Objeto As Object)
    Objeto.SelStart = 0
    Objeto.SelLength = Len(Objeto.Text)
End Sub

Public Sub CentraNaTela(F As Form)
 With frmMain
  If F.WindowState = vbNormal Then
   If TypeOf F Is MDIForm Then
    F.Top = (Screen.Height - F.Height) / 2
    F.Left = (Screen.Width - F.Width) / 2
   Else
    If F.MDIChild = True Then
     F.Top = (.ScaleHeight - F.Height) / 2
     F.Left = (.ScaleWidth - F.Width) / 2
    Else
     F.Top = (Screen.Height - F.Height) / 2
     F.Left = (Screen.Width - F.Width) / 2
    End If
   End If
  End If
 End With
End Sub

Public Sub CentraPanFundo(F As Form)
 Dim L As Long, T As Long
 If F.WindowState <> vbMinimized Then
  T = ((F.ScaleHeight - F.Painel.Height) / 2)
  L = (F.ScaleWidth - F.Painel.Width) / 2
  F.Painel.Move L, T
 End If
End Sub

'X
Public Sub MostraForm(F As Form)
    frmMenu.CoolBar1.Visible = False
    F.Left = 0
    F.Top = frmMenu.CoolBar1.Top + 300
    F.Height = frmMenu.CoolBar1.Height
    F.Width = frmMenu.Picture2.Width
    F.Show vbModal
    frmMenu.CoolBar1.Visible = True
End Sub

'Usada
Function LerIni(Secao As String, Chave As String, Padrao As String) As String
    Dim Valor As String
    Dim Tamanho As Long
    Valor = String(255, " ")
    Tamanho = GetPrivateProfileString(Secao, Chave, _
    Padrao, Valor, 255, App.EXEName & ".ini")
    LerIni = Left(Valor, Tamanho)
End Function

'Usada
Function EscreverIni(Secao As String, Chave As String, Valor As String)
     WritePrivateProfileString Secao, Chave, Valor, App.EXEName & ".ini"
End Function

Public Sub GravaParametrosDoForm(F As Form)
    If F.WindowState = vbMinimized Then Exit Sub
    If F.WindowState = vbMaximized Then
        EscreverIni F.Name, "Maximizado", "1"
    Else
        EscreverIni F.Name, "Maximizado", "0"
        EscreverIni F.Name, "Top", F.Top
        EscreverIni F.Name, "Left", F.Left
        EscreverIni F.Name, "Width", F.Width
        EscreverIni F.Name, "Height", F.Height
    End If
End Sub

Public Sub AtribuiParametrosDoForm(F As Form)
    Dim Topo As String
    Dim Esquerda As String
    Dim Altura As String
    Dim Largura As String
    Dim Maximizado As String
    Maximizado = LerIni(F.Name, "Maximizado", "")
    Topo = LerIni(F.Name, "Top", "")
    Esquerda = LerIni(F.Name, "Left", "")
    Largura = LerIni(F.Name, "Width", "")
    Altura = LerIni(F.Name, "Height", "")
    If Trim(Maximizado) = "" Then Exit Sub
    If Maximizado = "1" Then
        F.WindowState = vbMaximized
    Else
        F.Top = CDbl(Topo)
        F.Left = CDbl(Esquerda)
        F.Height = CDbl(Altura)
        F.Width = CDbl(Largura)
    End If
End Sub

'USADA
Public Sub PegaEnderecoDoBanco(CaixaDeDialogo As CommonDialog)
    On Error GoTo Erro
    CaminhoBanco = LerIni("Banco De Dados", "Endereço", "")
    If CaminhoBanco = "" Or Dir(CaminhoBanco) = "" Then
        With CaixaDeDialogo
            MsgBox "Por favor, localize na rede a base de dados (Banco.mdb) desta aplicação. Para que ela possa funcionar.", vbInformation, "Base de Dados"
            .DialogTitle = "Selecione o Bando de Dados Banco.mdb"
            .Filter = "Base de Dados da aplicação|Banco.mdb"
            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
            .CancelError = True
            .ShowOpen
            EscreverIni "Banco De Dados", "Endereço", .FileName
            CaminhoBanco = .FileName
        End With
    End If
    Exit Sub
Erro:
    If Err = cdlCancel Then
        End
    Else
        MsgBox Error$ & " - Informe o número de erro " & Err & " ao programador do sistema.", vbCritical, "ERRO"
    End If
    Exit Sub
End Sub

'usada
Public Sub AtribuiBanco(F As Form)
    Dim DATA_C As Control
    For Each DATA_C In F
        If TypeOf DATA_C Is Data Then
            DATA_C.DatabaseName = CaminhoBanco
            DATA_C.Refresh
        End If
    Next
End Sub




Public Sub AtualizaForms(ParamArray QForms())
    Dim F As Form
    Dim C As Control
    Dim X As Variant
    For Each X In QForms
        For Each F In Forms
            If F.Name = X Then
                If Not F.Editando Then
                    For Each C In F.Controls
                        If TypeOf C Is Data Then
                                C.Refresh
                        End If
                    Next
                End If
            End If
        Next
    Next
End Sub


Public Function PodeAtualizar(ParamArray QForms()) As Boolean
    Dim F As Form
    Dim Cont As Integer
    Dim X As Variant
    Dim C As Control
    Cont = 0
    For Each X In QForms
        For Each F In Forms
            If F.Name = X Then
                If F.Editando Then
                    MsgBox "Existem pedências de dados no formulário " & UCase(F.Caption) & "." & _
                        vbCrLf & vbCrLf & "Já que este formulário tem ligação com o formulário" & _
                        vbCrLf & "que você está usando, salve ou cancele primeiro essas" & _
                        vbCrLf & "pendências para depois continuar o processo atual.", vbInformation, "Atenção"
                    MostraForm F
                    Cont = Cont + 1
                End If
            End If
        Next
    Next
    If Cont > 0 Then
        PodeAtualizar = False
    Else
        PodeAtualizar = True
    End If
End Function


Public Sub Degradê(Formulário As Form, CorInicial As Long, _
    CorFinal As Long, Optional Sentido As Orientação = Vertical)

    Dim EscalaAntiga As Integer
    
    Dim PTamanho As Double
    
    Dim VermelhoIni As Integer, VerdeIni As Integer, AzulIni As Integer
    Dim VermelhoFin As Integer, VerdeFin As Integer, AzulFin As Integer
    Dim VermelhoMeio As Double, VerdeMeio As Double, AzulMeio As Double
    
    Dim I As Long, Cor As Long
    
    EscalaAntiga = Formulário.ScaleMode
    Formulário.ScaleMode = vbPixels
    
    If Sentido = Horizontal Then
        PTamanho = Formulário.ScaleWidth
    Else
        PTamanho = Formulário.ScaleHeight
    End If
    
    HEX_Para_RGB CorInicial, VermelhoIni, VerdeIni, AzulIni
    HEX_Para_RGB CorFinal, VermelhoFin, VerdeFin, AzulFin
    
    VermelhoMeio = (VermelhoFin - VermelhoIni) / PTamanho
    VerdeMeio = (VerdeFin - VerdeIni) / PTamanho
    AzulMeio = (AzulFin - AzulIni) / PTamanho
    
    Formulário.AutoRedraw = True
    
    For I = 0 To PTamanho - 1
        Cor = RGB(VermelhoIni + VermelhoMeio * I, _
            VerdeIni + VerdeMeio * I, AzulIni + AzulMeio * I)
        If Sentido = Horizontal Then
            Formulário.Line (I, 0)-(I, Formulário.Height - 1), Cor
        Else
            Formulário.Line (0, I)-(Formulário.Width - 1, I), Cor
        End If
    Next
    
    Formulário.ScaleMode = EscalaAntiga
    
End Sub

Public Sub HEX_Para_RGB(ByVal CorHEX As Long, _
    RetornoVermelho As Integer, RetornoVerde As Integer, _
    RetornoAzul As Integer)
    
    RetornoVermelho = CorHEX Mod 256
    RetornoVerde = ((CorHEX And &HFF00FF00) / 256&)
    RetornoAzul = ((CorHEX And &HFF0000) / 65536)
    
End Sub

Public Function VerificaCGC(CGC As String) As Boolean
    On Error GoTo Err_CGC
    If Len(Trim(CGC)) <> 14 Then
        VerificaCGC = False
        Exit Function
    End If
    Dim I As Integer
    Dim strCGC As String
    Dim strcampo As String
    Dim strCaracter As String
    Dim intNumero As Integer
    Dim intMais As Integer
    Dim intSoma As Long
    Dim intSoma1 As Long
    Dim intSoma2 As Long
    Dim dblDivisao As Double
    Dim intInteiro As Long
    Dim intResto As Integer
    Dim intDig1 As Integer
    Dim intDig2 As Integer
    Dim strConf As String
    
    intSoma = 0
    intSoma1 = 0
    intSoma2 = 0
    intNumero = 0
    intMais = 0
    
    strCGC = Right(CGC, 6)
    strCGC = Left(strCGC, 4)
    strcampo = Left(CGC, 8)
    strcampo = Right(strcampo, 4) & strCGC
    For I = 2 To 9
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        intSoma1 = intSoma1 + intMais
    Next I
    strcampo = Left(CGC, 4)
    For I = 2 To 5
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        intSoma2 = intSoma2 + intMais
    Next I
    intSoma = intSoma1 + intSoma2
    dblDivisao = intSoma / 11
    intInteiro = Int(dblDivisao) * 11
    intResto = intSoma - intInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig1 = 0
    Else
        intDig1 = 11 - intResto
    End If
    intSoma = 0
    intSoma1 = 0
    intSoma2 = 0
    intNumero = 0
    intMais = 0
    strCGC = Right(CGC, 6)
    strCGC = Left(strCGC, 4)
    strcampo = Left(CGC, 8)
    strcampo = Right(strcampo, 3) & strCGC & intDig1
    For I = 2 To 9
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        intSoma1 = intSoma1 + intMais
    Next I
    strcampo = Left(CGC, 5)
    For I = 2 To 6
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        intSoma2 = intSoma2 + intMais
    Next I
    intSoma = intSoma1 + intSoma2
    dblDivisao = intSoma / 11
    intInteiro = Int(dblDivisao) * 11
    intResto = intSoma - intInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig2 = 0
    Else
        intDig2 = 11 - intResto
    End If
    strConf = intDig1 & intDig2
    If strConf <> Right(CGC, 2) Then
        VerificaCGC = False
    Else
        VerificaCGC = True
    End If
    Exit Function
Exit_CGC:
        Exit Function
Err_CGC:
        VerificaCGC = False
        Resume Exit_CGC
End Function

Public Function VerificaCPF(CPF As String) As Boolean
    On Error GoTo Err_CPF
    If Len(Trim(CPF)) <> 11 Then
        VerificaCPF = False
        Exit Function
    End If
    Dim I As Integer
    Dim strcampo As String
    Dim strCaracter As String
    Dim intNumero As Integer
    Dim intMais As Integer
    Dim lngSoma As Long
    Dim dblDivisao As Double
    Dim lngInteiro As Long
    Dim intResto As Integer
    Dim intDig1 As Integer
    Dim intDig2 As Integer
    Dim strConf As String
    
    lngSoma = 0
    intNumero = 0
    intMais = 0
    strcampo = Left(CPF, 9)
    
    For I = 2 To 10
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig1 = 0
    Else
        intDig1 = 11 - intResto
    End If
    
    strcampo = strcampo & intDig1
    lngSoma = 0
    intNumero = 0
    intMais = 0
    For I = 2 To 11
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig2 = 0
    Else
        intDig2 = 11 - intResto
    End If
    strConf = intDig1 & intDig2
    If strConf <> Right(CPF, 2) Then
        VerificaCPF = False
    Else
        VerificaCPF = True
    End If
    Exit Function
Exit_CPF:
        Exit Function
Err_CPF:
        VerificaCPF = False
        Resume Exit_CPF
End Function

'Usada
Public Function Frac(Número As Double) As Double
    Frac = Número - Fix(Número)
End Function

'Usada
Public Function EncriptarDescriptar(Texto As String, ParamArray Chaves1a255()) As String
    If Len(Texto) = 0 Then EncriptarDescriptar = "": Exit Function
    Dim I As Long
    Dim J As Variant
    Dim Encript As String
    Dim Resultado As String
    Encript = Texto
    For Each J In Chaves1a255
        For I = 1 To Len(Encript)
            Resultado = Resultado & Chr$(Asc(Mid(Encript, I, 1)) Xor J)
        Next
        Encript = Resultado
        Resultado = ""
    Next
    EncriptarDescriptar = Encript
End Function

'usada
Public Function GeraSenha() As String
    Dim DT As String
    Dim TL As String
    Dim I As Integer
    Dim J As Integer
    Dim Result As String
    Result = ""
    DT = Format(Date, "ddmmyyyy")
    TL = "9619643"
    For I = 1 To Len(TL)
        Result = Result & Mid(DT, I, 1) & Mid(TL, I, 1)
    Next
    Result = Result & Mid(DT, I, 1)
    GeraSenha = Result
End Function

Public Function Curto(Longo As String) As String
    Dim sCurto As String * 255
    Dim Tam As Long
    Tam = GetShortPathName(Longo, sCurto, 255)
    Curto = Left(sCurto, Tam)
End Function

Public Sub ReparaBD()
    Dim vgNomeBAK As String, vgNomePAK As String, X As String
    Beep
    If MsgBox("Esta operação pode ser demorada, Continuar ?", vbQuestion + vbYesNo + _
        vbDefaultButton2, "REPARAR E COMPACTAR A BASE DE DADOS") = vbYes Then
        Screen.MousePointer = vbHourglass
        vgNomeBAK = Left(CaminhoBanco, Len(CaminhoBanco) - 3) + "BAK"
        vgNomePAK = Left(CaminhoBanco, Len(CaminhoBanco) - 3) + "PAK"
        On Error GoTo Erro
        CompactDatabase CaminhoBanco, vgNomePAK
        If Existe(vgNomeBAK) Then Kill vgNomeBAK
        Name CaminhoBanco As vgNomeBAK
        Name vgNomePAK As CaminhoBanco
        RepairDatabase CaminhoBanco
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "A base de dados foi reparada.", vbExclamation, "REPARAR E COMPACTAR A BASE DE DADOS"
    End If
    Exit Sub
Erro:
    If Err = 3356 Or Err = 3196 Then
        X = "Banco de dados bloqueado por outro usuário."
    Else
        X = Error$
    End If
    Beep
    MsgBox X, vbCritical, "ERRO"
End Sub

Public Function Existe(NomeArq As String) As Integer
    Existe = Len(Dir(NomeArq)) > 0
End Function

Public Sub FechaForms(Formulário As Form, ParamArray QForms())
    Dim F As Form
    Dim X
    For Each X In QForms
        For Each F In Forms
            If F.Name = X Then
                Unload F
                MostraForm F
            End If
        Next
    Next
    Formulário.SetFocus
End Sub
