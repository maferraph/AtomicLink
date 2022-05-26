VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Tela_Navegador 
   AutoRedraw      =   -1  'True
   Caption         =   "Navegador"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   ScaleHeight     =   3720
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet INet 
      Left            =   720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      URL             =   "http://"
   End
   Begin MSComctlLib.ImageList LI 
      Left            =   120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":0F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":1B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Navegador.frx":2144
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "LI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abre um arquivo de internet"
            Object.Tag             =   "Abrir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salva esta página de internet no seu computador"
            Object.Tag             =   "Salvar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recarregar"
            Object.ToolTipText     =   "Recarrega a página de internet atual"
            Object.Tag             =   "Recarregar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Parar"
            Object.ToolTipText     =   "Para o carregamento da página de internet atual"
            Object.Tag             =   "Parar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Chupinhar"
            Object.ToolTipText     =   "Chupinhador de HTML"
            Object.Tag             =   "Chupinhar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Voltar"
            Object.ToolTipText     =   "Volta ao Gerenciador do Atomic Link"
            Object.Tag             =   "Voltar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar BPC 
         Height          =   200
         Left            =   4080
         TabIndex        =   4
         Top             =   320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox TXT_L 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar BS 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Sites Encontrados"
            TextSave        =   "Sites Encontrados"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Número de links encontrados pelo Chupinha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Sites Ignorados"
            TextSave        =   "Sites Ignorados"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Número de links já cadastrados pelo Chupinha"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Erros"
            TextSave        =   "Erros"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Número de erros executados pelo Chupinha"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "14/04/00"
            Object.ToolTipText     =   "Data do Sistema"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "16:54"
            Object.ToolTipText     =   "Hora do Sistema"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "Status"
            Object.Tag             =   "Status"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   1215
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      ExtentX         =   5106
      ExtentY         =   2143
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Tela_Navegador.frx":2A20
   End
End
Attribute VB_Name = "Tela_Navegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RespMsg, I As Long, J As Long, K As Long, M As Long, R As Long, AchoLink As Boolean, ArquivoHTML As String
Private Sub BF_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "Abrir" Then
        'abre dialogo
        DLG.DialogTitle = "Selecione um arquivo HTML"
        DLG.Filter = "Arquivos HTML|*.htm;*.html;*.shtml;"
        If Mid(TXT_L.Text, 2, 2) = ":\" Then 'é um arquivo
            DLG.InitDir = TXT_L.Text
        Else
            DLG.InitDir = "C:"
        End If
        DLG.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
        DLG.ShowOpen
        'pega nome do arquivo
        TXT_L.Text = DLG.FileName
        'carrega pagina
        If TXT_L.Text <> "" Then
            On Error GoTo ERRO_NAOEXISTEARQ
            Open TXT_L.Text For Input As #1
            Close #1
            CarregaHTML
        End If
    ElseIf Button.Tag = "Salvar" Then
        DLG.DialogTitle = "Salvar arquivo HTML"
        DLG.Filter = "Arquivos HTML|*.htm;*.html;*.shtml;"
        DLG.InitDir = "C:"
        DLG.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
        DLG.ShowSave
        If DLG.FileName <> "" Then
            RTB.SaveFile DLG.FileName, rtfText
        End If
    ElseIf Button.Tag = "Recarregar" Then
        WB.Refresh
    ElseIf Button.Tag = "Parar" Then
        WB.Stop
    ElseIf Button.Tag = "Chupinhar" Then
        RespMsg = MsgBox("Você está prestes à chupinhar este arquivo. Deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, "Chupinhar arquivo")
        If RespMsg = vbYes Then
            TelaEmEspera (True)
            Chupinha
            TelaEmEspera (False)
        End If
    ElseIf Button.Tag = "Voltar" Then
        Unload Me
    End If
    Exit Sub
ERRO_NAOEXISTEARQ: If Err Then MsgBox "Não foi possível carregar o arquivo na página ou ele não existe. Tente novamente.", vbCritical + vbOKOnly, "Erro de carregamento"
End Sub
Private Sub Form_Load()
    WB.Navigate App.Path & "/AtomicLink.html"
End Sub
Private Sub Form_Resize()
    ArrumaTela
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Tela_Principal.Show
End Sub
Private Sub TXT_L_GotFocus()
    TXT_L.SelLength = Len(TXT_L)
End Sub
Private Sub TXT_L_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CarregaHTML
    End If
End Sub
Private Sub TXT_L_LostFocus()
    TXT_L.SelLength = Len(TXT_L)
    If UCase(Left(TXT_L.Text, 3)) = "www" Then TXT_L.Text = "http://" & TXT_L.Text
    If UCase(Left(TXT_L.Text, 4)) = "HTTP" Then WB.Offline = False
    WB.Navigate TXT_L
End Sub
Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    BS.Panels(9).Text = "Abrindo página de internet..."
    TXT_L.Text = WB.LocationURL
End Sub
Private Sub WB_DownloadComplete()
    BS.Panels(9).Text = "Documento Pronto"
End Sub
Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    TXT_L.Text = WB.LocationURL
    'le arquivo html
    If UCase(Left(TXT_L.Text, 3)) = "WWW" Or UCase(Left(TXT_L.Text, 4)) = "HTTP" Then
        Dim b() As Byte, txt As String
        b() = INet.OpenURL(TXT_L.Text, 1)
        txt = ""
        For I = 0 To UBound(b) - 1
            txt = txt + Chr(b(I))
        Next
        RTB.Text = txt
    End If
    
    BF.Buttons(5).Enabled = False
    TXT_L.Text = WB.LocationURL
    Me.Caption = WB.LocationName
    ArquivoHTML = URL
End Sub
Private Sub WB_StatusTextChange(ByVal Text As String)
    BS.Panels(9).Text = Text
End Sub
Private Sub WB_TitleChange(ByVal Text As String)
    Me.Caption = Text
End Sub



'****************************************************************
'                           F U N Ç Õ E S
'****************************************************************
Private Sub ArrumaTela()
    WB.Top = 0 + BF.Height
    WB.Left = 0
    If Me.WindowState = vbMaximized Then
        WB.Height = Screen.Height - BS.Height - Screen.TwipsPerPixelY * 85
        WB.Width = Screen.Width
        TXT_L.Width = Screen.Width - TXT_L.Left - 200
    Else
        If Me.WindowState = vbMinimized Then Exit Sub
        WB.Height = Me.Height - BS.Height
        WB.Width = Me.Width
        TXT_L.Width = Me.Width - TXT_L.Width
    End If
    BPC.Width = TXT_L.Width
End Sub
Private Static Sub TelaEmEspera(Espera As Boolean)
    If Espera = True Then
        Me.Enabled = False
        Me.MousePointer = vbHourglass
    Else
        Me.Enabled = True
        Me.MousePointer = vbDefault
    End If
End Sub
Private Sub CarregaHTML()
    On Error GoTo ERRO_CARREGAHTML
    TXT_L.SelLength = Len(TXT_L)
    WB.Offline = True
    If UCase(Left(TXT_L.Text, 3)) = "WWW" Then TXT_L.Text = "http://" & TXT_L.Text
    If UCase(Left(TXT_L.Text, 4)) = "HTTP" Then WB.Offline = False
    BF.Buttons(5).Enabled = True
    'abre pagina
    WB.Navigate TXT_L.Text
    'le arquivo html
    If UCase(Left(TXT_L.Text, 3)) = "WWW" Or UCase(Left(TXT_L.Text, 4)) = "HTTP" Then
        Dim b() As Byte, txt As String
        b() = INet.OpenURL(TXT_L.Text, 1)
        txt = ""
        For I = 0 To UBound(b) - 1
            txt = txt + Chr(b(I))
        Next
        RTB.Text = txt
    End If
ERRO_CARREGAHTML:
    If Err Then MsgBox "Ocorreu algum erro ao carregar o arquivo HTML.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
