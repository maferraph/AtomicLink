VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Tela_Principal 
   AutoRedraw      =   -1  'True
   Caption         =   "Atomic Link"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9600
   Icon            =   "Tela_Principal.frx":0000
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar BS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5970
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FR_BA 
      Height          =   3615
      Left            =   360
      TabIndex        =   15
      Top             =   4080
      Width           =   9495
      Begin VB.Frame FR_BO 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         Width           =   5535
         Begin VB.CommandButton BT_Backup 
            Caption         =   "Destino do Backup"
            Height          =   1095
            Left            =   840
            Picture         =   "Tela_Principal.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Indicar o local onde será realizado o backup"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton RB_PR 
            Caption         =   "Fazer backup do banco de dados principal"
            Height          =   255
            Left            =   840
            TabIndex        =   8
            ToolTipText     =   "Para fazer backup do banco de dados principal"
            Top             =   480
            Width           =   3735
         End
         Begin VB.OptionButton RB_NO 
            Caption         =   "Fazer backup do banco de dados de novos links"
            Height          =   255
            Left            =   840
            TabIndex        =   9
            ToolTipText     =   "Para fazer backup do banco de dados de novos links"
            Top             =   840
            Width           =   3855
         End
         Begin VB.Label LB_DE 
            Caption         =   "C:\"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   2160
            TabIndex        =   29
            ToolTipText     =   "Local onde será realizado o backup"
            Top             =   1560
            Width           =   3060
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   2160
            TabIndex        =   28
            Top             =   1320
            Width           =   585
         End
      End
   End
   Begin VB.Frame FR_AT 
      Height          =   5055
      Left            =   600
      TabIndex        =   14
      Top             =   600
      Width           =   9495
      Begin VB.Frame FR_AS 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   480
         TabIndex        =   30
         Top             =   600
         Width           =   5895
         Begin VB.Frame Frame3 
            Caption         =   "Local de origem do backup:"
            Height          =   1455
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   5775
            Begin VB.CommandButton BT_AT 
               Caption         =   "Atualizar"
               Height          =   1095
               Left            =   4560
               Picture         =   "Tela_Principal.frx":1194
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Atualizar sistema"
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton BT_LB 
               Caption         =   "Local do Backup"
               Height          =   1095
               Left            =   120
               Picture         =   "Tela_Principal.frx":149E
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Para localalizar onde se encontra o arquivo de backup"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label LB_LO 
               Caption         =   "C:\"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   840
               Left            =   1320
               TabIndex        =   37
               ToolTipText     =   "Local onde se encontra o arquivo de backup"
               Top             =   360
               Width           =   3180
            End
         End
         Begin VB.OptionButton RB_AN 
            Caption         =   "Fazer atualização do banco de dados de novos links"
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            ToolTipText     =   "Para fazer a atualização de links novos no banco de dados principal"
            Top             =   720
            Width           =   4095
         End
         Begin VB.OptionButton RB_AP 
            Caption         =   "Fazer atualização do banco de dados principal"
            Height          =   255
            Left            =   1080
            TabIndex        =   4
            ToolTipText     =   "Para fazer atualização do banco de dados principal"
            Top             =   360
            Width           =   3735
         End
         Begin VB.Frame Frame1 
            Height          =   975
            Left            =   120
            TabIndex        =   31
            Top             =   2640
            Width           =   5775
            Begin MSComctlLib.StatusBar BSE 
               Height          =   255
               Left            =   2880
               TabIndex        =   32
               ToolTipText     =   "Número de links novos encontrados"
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               Style           =   1
               _Version        =   393216
               BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
                  NumPanels       =   1
                  BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.StatusBar BSI 
               Height          =   255
               Left            =   2880
               TabIndex        =   33
               ToolTipText     =   "Número de links novos ignorados"
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               Style           =   1
               _Version        =   393216
               BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
                  NumPanels       =   1
                  BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Links já catalogados:"
               Height          =   195
               Left            =   1200
               TabIndex        =   35
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Links encontrados:"
               Height          =   195
               Left            =   1200
               TabIndex        =   34
               Top             =   240
               Width           =   1350
            End
         End
      End
   End
   Begin VB.Frame FR_CH 
      Height          =   4575
      Left            =   600
      TabIndex        =   17
      Top             =   1560
      Width           =   9495
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
         Left            =   6840
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FR_RE 
         Height          =   1575
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   5895
         Begin VB.CommandButton BT_CH 
            Caption         =   "Chupinhar"
            Height          =   975
            Left            =   4680
            Picture         =   "Tela_Principal.frx":1D68
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Chupinhador de HTML"
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton BT_AB 
            Caption         =   "Abrir página"
            Height          =   975
            Left            =   3600
            Picture         =   "Tela_Principal.frx":2632
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Abre uma página de internet para ser chupinhada"
            Top             =   360
            Width           =   975
         End
         Begin MSComctlLib.StatusBar BS_EN 
            Height          =   255
            Left            =   1920
            TabIndex        =   20
            ToolTipText     =   "Número de link encontrados"
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            Style           =   1
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.StatusBar BS_IG 
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "Número de links ignorados"
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            Style           =   1
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.StatusBar BS_ER 
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            ToolTipText     =   "Número de erros ocorridos"
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            Style           =   1
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Erros:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Links já catalogados:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Links encontrados:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   1350
         End
      End
      Begin SHDocVwCtl.WebBrowser WB 
         Height          =   1215
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   9255
         ExtentX         =   16325
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
   End
   Begin VB.Frame FR_CA 
      Height          =   3615
      Left            =   600
      TabIndex        =   12
      ToolTipText     =   "Cadastro de Categorias"
      Top             =   1200
      Width           =   9495
      Begin VB.Frame FR_C2 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   9375
         Begin VB.Frame FR_CA2 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   4560
            TabIndex        =   40
            Top             =   480
            Width           =   4770
            Begin VB.TextBox TXT_Categorias 
               Height          =   285
               Left            =   120
               TabIndex        =   42
               ToolTipText     =   "Nome da categoria"
               Top             =   345
               Width           =   4575
            End
            Begin VB.TextBox TXT_CatInd 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   41
               ToolTipText     =   "Número  da categoria"
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox TXT_CatDes 
               Height          =   285
               Left            =   1080
               TabIndex        =   44
               ToolTipText     =   "Descrição desta categoria"
               Top             =   960
               Width           =   3615
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nome da Categoria"
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   120
               Width           =   1365
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Índice:"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   720
               Width           =   480
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   1080
               TabIndex        =   43
               Top             =   720
               Width           =   765
            End
         End
         Begin VB.ListBox LT_Categorias 
            ForeColor       =   &H80000007&
            Height          =   1815
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Lista de categorias"
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Categorias"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   750
         End
      End
   End
   Begin VB.Frame FR_LICA 
      Height          =   4695
      Left            =   600
      TabIndex        =   16
      ToolTipText     =   "Categorias deste link"
      Top             =   840
      Width           =   9615
      Begin VB.Frame FR_L2 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   9495
         Begin VB.CommandButton BT_Fechar 
            Caption         =   "Fechar Categorias do Link"
            Height          =   255
            Left            =   3240
            TabIndex        =   54
            ToolTipText     =   "Volta à edição do link"
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox TXT_Lica 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   53
            Text            =   "www.atomicfind.com.br/engenharia/mecanica"
            ToolTipText     =   "Site em questão"
            Top             =   840
            Width           =   9135
         End
         Begin VB.ListBox LT_CA 
            Height          =   2400
            Left            =   120
            TabIndex        =   52
            ToolTipText     =   "Lista de categorias disponíveis"
            Top             =   1440
            Width           =   4215
         End
         Begin VB.ListBox LT_LICA 
            Height          =   2400
            Left            =   5040
            TabIndex        =   51
            ToolTipText     =   "Lista de categorias registradas para este link"
            Top             =   1440
            Width           =   4215
         End
         Begin VB.CommandButton BT_P 
            Caption         =   ">>"
            Height          =   495
            Left            =   4440
            TabIndex        =   50
            ToolTipText     =   "Inseri"
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton BT_T 
            Caption         =   "<<"
            Height          =   495
            Left            =   4440
            TabIndex        =   49
            ToolTipText     =   "Remove"
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "En&dereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Categorias Disponiveis:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Categorias deste link:"
            Height          =   195
            Left            =   5040
            TabIndex        =   55
            Top             =   1200
            Width           =   1515
         End
      End
   End
   Begin MSComctlLib.ImageList LI 
      Left            =   0
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":2EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":3218
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":3AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":3E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":46F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":4A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":4D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":5048
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":54A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":57C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":5AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":5DFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "LI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo"
            Object.Tag             =   "Novo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Editar"
            Object.Tag             =   "Editar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            Object.Tag             =   "Salvar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Apagar"
            Object.ToolTipText     =   "Apagar"
            Object.Tag             =   "Apagar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   "Sair"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar BP 
         Height          =   195
         Left            =   4920
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   0
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FR_BD 
      Height          =   5415
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   "Dados sobre o link"
      Top             =   600
      Width           =   9495
      Begin VB.Frame FR_B2 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   9495
         Begin VB.CommandButton BT_Net 
            Caption         =   "Ir para página"
            Height          =   255
            Left            =   3240
            TabIndex        =   80
            ToolTipText     =   "Inicia seu navegador padrão e vai até a página selecionada na lista abaixo"
            Top             =   120
            Width           =   1335
         End
         Begin VB.Frame FR_BD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2595
            Left            =   120
            TabIndex        =   63
            Top             =   2520
            Width           =   9405
            Begin VB.TextBox TXT_Descricao 
               Height          =   525
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   70
               ToolTipText     =   "Descrição deste site (palavras-chave)"
               Top             =   1440
               Width           =   9255
            End
            Begin VB.TextBox TXT_Titulo 
               Height          =   285
               Left            =   2280
               TabIndex        =   69
               ToolTipText     =   "Título deste site"
               Top             =   840
               Width           =   6975
            End
            Begin VB.TextBox TXT_Link 
               Height          =   285
               Left            =   0
               MaxLength       =   200
               TabIndex        =   68
               Text            =   "www.atomicfind.com.br/engenharia/mecanica"
               ToolTipText     =   "Endereço do link"
               Top             =   240
               Width           =   7935
            End
            Begin VB.TextBox TXT_Data 
               Enabled         =   0   'False
               Height          =   285
               Left            =   8160
               MaxLength       =   10
               TabIndex        =   67
               ToolTipText     =   "Data de inserção do site"
               Top             =   240
               Width           =   1095
            End
            Begin VB.ComboBox CB_Lingua 
               Height          =   315
               ItemData        =   "Tela_Principal.frx":625C
               Left            =   0
               List            =   "Tela_Principal.frx":6266
               TabIndex        =   66
               ToolTipText     =   "Língua deste site"
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox TXT_Procedencia 
               Height          =   285
               Left            =   0
               TabIndex        =   65
               ToolTipText     =   "Origem ou proprietário do documento no link"
               Top             =   2280
               Width           =   4335
            End
            Begin VB.TextBox TXT_Email 
               Height          =   285
               Left            =   4680
               MaxLength       =   200
               TabIndex        =   64
               ToolTipText     =   "E-mail para contato"
               Top             =   2280
               Width           =   4575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Left            =   0
               TabIndex        =   77
               Top             =   0
               Width           =   735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Descrição da Página:"
               Height          =   195
               Left            =   0
               TabIndex        =   76
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Título da Página:"
               Height          =   195
               Left            =   2280
               TabIndex        =   75
               Top             =   600
               Width           =   1230
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Data:"
               Height          =   195
               Left            =   8160
               TabIndex        =   74
               Top             =   0
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Língua:"
               Height          =   195
               Left            =   0
               TabIndex        =   73
               Top             =   600
               Width           =   555
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Procedência:"
               Height          =   195
               Left            =   0
               TabIndex        =   72
               Top             =   2040
               Width           =   945
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "e-mail:"
               Height          =   195
               Left            =   4680
               TabIndex        =   71
               Top             =   2040
               Width           =   450
            End
         End
         Begin VB.ComboBox CB_Exibir 
            Height          =   315
            ItemData        =   "Tela_Principal.frx":627D
            Left            =   120
            List            =   "Tela_Principal.frx":6293
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   360
            Width           =   2895
         End
         Begin VB.ListBox LT_Link 
            ForeColor       =   &H80000007&
            Height          =   1620
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   9255
         End
         Begin VB.CommandButton BT_CA 
            Caption         =   "Categorias"
            Height          =   255
            Left            =   3240
            TabIndex        =   60
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TXT_CA 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            MaxLength       =   200
            TabIndex        =   59
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label LB_Exibir 
            AutoSize        =   -1  'True
            Caption         =   "Exibir por:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   120
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categorias:"
            Height          =   195
            Left            =   4800
            TabIndex        =   78
            Top             =   120
            Width           =   795
         End
      End
   End
   Begin VB.Menu Menu_Principal 
      Caption         =   "&Principal"
      Begin VB.Menu Menu_Principal_Novo 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu_Principal_Editar 
         Caption         =   "&Editar"
         Shortcut        =   ^E
      End
      Begin VB.Menu Menu_Principal_Salvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu_Principal_Apagar 
         Caption         =   "&Apagar"
         Shortcut        =   ^A
      End
      Begin VB.Menu Menu_Principal_Cancelar 
         Caption         =   "Cance&lar"
         Shortcut        =   ^R
      End
      Begin VB.Menu Menu_Principal_L1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Principal_Finalizar 
         Caption         =   "&Finalizar"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Menu_Editar 
      Caption         =   "E&ditar"
      Begin VB.Menu Menu_Editar_Copiar 
         Caption         =   "Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu_Editar_Recortar 
         Caption         =   "Recortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu Menu_Editar_Colar 
         Caption         =   "Colar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Menu_Exibir 
      Caption         =   "E&xibir"
      Begin VB.Menu Menu_Exibir_Links 
         Caption         =   "Links"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu Menu_Exibir_Categorias 
         Caption         =   "Categorias"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu Menu_Exibir_L1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exibir_Chupinhador 
         Caption         =   "Chupinhador"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
      Begin VB.Menu Menu_Exibir_L2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exibir_Atualizacao 
         Caption         =   "Atualização"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
      Begin VB.Menu Menu_Exibir_Backup 
         Caption         =   "Backups"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu Menu_Chupinador 
      Caption         =   "Chupin&hador"
      Begin VB.Menu Menu_Chupinador_Abrir 
         Caption         =   "Abrir Página"
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu_Chupinador_L1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Chupinador_Chupinhar 
         Caption         =   "Chupinhar"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu Menu_Sobre 
      Caption         =   "S&obre"
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ModoEdicao As Boolean, LinkNovo As Boolean, RespMsg, ArquivoHTML As String
Private Sub BF_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ERRO_ATOMICLINK
    If Button.Tag = "Sair" Then
        Menu_Principal_Finalizar_Click
    ElseIf Button.Tag = "Novo" Then
        Menu_Principal_Novo_Click
    ElseIf Button.Tag = "Editar" Then
        Menu_Principal_Editar_Click
    ElseIf Button.Tag = "Apagar" Then
        Menu_Principal_Apagar_Click
    ElseIf Button.Tag = "Cancelar" Then
        Menu_Principal_Cancelar_Click
    ElseIf Button.Tag = "Salvar" Then
        Menu_Principal_Salvar_Click
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_AB_Click()
    On Error GoTo ERRO_ATOMICLINK
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
        If ArquivoExiste(TXT_L.Text) = False Then Exit Sub
        'abre pagina no navegador
        WB.Offline = True
        WB.Navigate TXT_L.Text
        'abre arquivo para leitura
        Open TXT_L.Text For Input As #1
        ArquivoHTML = Input(LOF(1), 1)
        Close #1
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_AT_Click()
    'On Error GoTo ERRO_ATOMICLINK
    If RB_AP.Value = False And RB_AN.Value = False Then
        MsgBox "Você deve primeiro selecionar qual arquivo será atualizado.", vbCritical + vbOKOnly, "Selecione o tipo"
        RB_AP.SetFocus
        Exit Sub
    End If
    If ArquivoExiste(LB_LO.Caption) = False Then
        MsgBox "Este arquivo não está sendo acessado ou encontrado.", vbCritical + vbOKOnly, "Erro de acesso"
        Exit Sub
    End If
    If RB_AP.Value = True Then
        RespMsg = MsgBox("Você está prestes à atualizar o banco de dados final do Atomic Link neste computador - se você prosseguir, o arquivo de backup que você selecionou irá substituir o banco de dados final, então só faça isso se for o caso. Você tem certeza que deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, "Atualização Final")
    ElseIf RB_AN.Value = True Then
        RespMsg = MsgBox("Você está prestes à atualizar o banco de dados de novos links do Atomic Link neste computador - se você prosseguir, o banco de dados final será alterado com as novas informações se necessário e o banco de dados de novos links será limpo, sem possibilidade de desfazer as alterações. Você tem certeza que deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, "Atualização Novos")
    End If
    If RespMsg = vbNo Then Exit Sub
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... atualizando o Atomic Link."
    Dim DirWinTmp As String
    DirWinTmp = Environ$("temp")
    If DirWinTmp = "" Then DirWinTmp = "C:\WINDOWS\TEMP"
    Dim DosDirLoc As String, Tamanho As Long, DosDirArq As String, ShortDir As String, DosDirApp As String
    ShortDir = Space$(1024)
    Tamanho = GetShortPathName(Trim(LB_LO.Caption), ShortDir, Len(ShortDir))
    DosDirLoc = Left$(ShortDir, Tamanho)
    ShortDir = Space$(1024)
    Tamanho = GetShortPathName(DirWinTmp, ShortDir, Len(ShortDir))
    DosDirArq = Left$(ShortDir, Tamanho)
    ShortDir = Space$(1024)
    Tamanho = GetShortPathName(App.path, ShortDir, Len(ShortDir))
    DosDirApp = Left$(ShortDir, Tamanho)
    If RB_AP.Value = True Then
        If LCase(Right(DosDirLoc, 11)) <> "al_bkp_f.af" Then
            MsgBox "O arquivo especificado para fazer atualização do banco de dados principal não é o al_bkp_f.af - Localize-o e tente novamente.", vbCritical + vbOKOnly, "Arquivo errado"
            GoTo ERRO_ATOMICLINK
        End If
        Shell App.path & "/arj e -y " & DosDirLoc & " " & DosDirArq & "\", vbMinimizedNoFocus
        FechaArj
        'verifica novo bd
        If ArquivoExiste(DosDirArq & "\BDAF.af") = False Then GoTo ERRO_FALTAARQUIVO
        'fecha banco de dados atual
        FechaBD
        'apaga banco de dados velho
        Kill DosDirApp & "\BDAF.af"
        'copia bd novo para dir do atomic link
        FileCopy DosDirArq & "\BDAF.af", DosDirApp & "\BDAF.af"
        'apaga bd temporario
        Kill DosDirArq & "\BDAF.af"
        'abre banco de dados novo
        If AbreBD = False Then End
    ElseIf RB_AN.Value = True Then
        If LCase(Right(DosDirLoc, 11)) <> "al_bkp_n.af" Then
            MsgBox "O arquivo especificado para fazer atualização do banco de dados de novos links não é o al_bkp_n.af - Localize-o e tente novamente.", vbCritical + vbOKOnly, "Arquivo errado"
            GoTo ERRO_ATOMICLINK
        End If
        Shell App.path & "/arj e -y " & DosDirLoc & " " & DosDirArq & "\", vbMinimizedNoFocus
        FechaArj
        'abre bd
        BS.SimpleText = "Abrindo banco de dados de atualização..."
        If ArquivoExiste(DosDirArq & "\BDNovos.af") = False Then GoTo ERRO_FALTAARQUIVO
        If AbreBD_Atualizacao_Novos(DirWinTmp & "\BDNovos.af") = False Then GoTo ERRO_FALTAARQUIVO
        'le novas categorias
        If BDATU_TBCAT.RecordCount > 0 Then 'existe novas categorias
            BS.SimpleText = "Atualizando tabela de categorias..."
            BSE.SimpleText = 0
            BSI.SimpleText = 0
            BDATU_TBCAT.MoveFirst
            Do While Not BDATU_TBCAT.EOF
                BDATF_TBCAT.Seek "=", BDATU_TBCAT_CPCAT.Value 'procura se a categoria nova já existe
                If BDATF_TBCAT.NoMatch Then
                    'nao existe entao ira inserir na tab.final
                    BDATF_TBCAT.AddNew
                    BDATF_TBCAT_CPCAT.Value = BDATU_TBCAT_CPCAT.Value
                    BDATF_TBCAT_CPDES.Value = BDATU_TBCAT_CPDES.Value
                    BDATF_TBCAT.Update
                    BSE.SimpleText = BSE.SimpleText + 1
                Else
                    BSI.SimpleText = BSI.SimpleText + 1
                End If
                BDATU_TBCAT.MoveNext
            Loop
        End If
        'le novos links
        If BDATU_TBLIN.RecordCount > 0 Then 'existe novos links
            BS.SimpleText = "Atualizando tabela de links..."
            BSE.SimpleText = 0
            BSI.SimpleText = 0
            BDATU_TBLIN.MoveFirst
            Do While Not BDATU_TBLIN.EOF
                BDATF_TBLIN.Seek "=", BDATU_TBLIN_CPLIN.Value 'procura se o link novo já existe
                If BDATF_TBCAT.NoMatch Then
                    'nao existe entao ira inserir na tab.final
                    BDATF_TBLIN.AddNew
                    BDATF_TBLIN_CPLIN.Value = BDATU_TBLIN_CPLIN.Value
                    BDATF_TBLIN_CPDAT.Value = BDATU_TBLIN_CPDAT.Value
                    BDATF_TBLIN_CPTIT.Value = BDATU_TBLIN_CPTIT.Value
                    BDATF_TBLIN_CPDES.Value = BDATU_TBLIN_CPDES.Value
                    BDATF_TBLIN_CPCAT.Value = BDATU_TBLIN_CPCAT.Value
                    BDATF_TBLIN_CPLGA.Value = BDATU_TBLIN_CPLGA.Value
                    BDATF_TBLIN_CPPRO.Value = BDATU_TBLIN_CPPRO.Value
                    BDATF_TBLIN_CPEMA.Value = BDATU_TBLIN_CPEMA.Value
                    BDATF_TBLIN.Update
                    BSE.SimpleText = BSE.SimpleText + 1
                Else
                    BSI.SimpleText = BSI.SimpleText + 1
                End If
                BDATU_TBLIN.MoveNext
            Loop
        End If
        'se tiver novas categorias, corrigir numero dos links
        If BDATU_TBCAT.RecordCount > 0 Then
            BS.SimpleText = "Corrigindo índices de categorias..."
            BSE.SimpleText = 0
            BSI.SimpleText = 0
            BDATU_TBCAT.MoveFirst
            Dim OldCat As String, Cat As String, NewCat As String, NumCat As Integer, NumCarCat As Integer
            Cat = ""
            NewCat = ""
            Do While Not BDATU_TBCAT.EOF
                If BDATU_TBCAT_CPLIN.Value <> "" Then
                    OldCat = BDATU_TBCAT_CPLIN.Value
                    For I = 1 To Len(OldCat)
                        If Mid(OldCat, I, 1) = ";" Then
                            If Left(Cat, 1) = "N" Then 'é categoria nova
                                BDATF_TBCAT.Seek "=", BDATU_TBCAT_CPCAT.Value
                                If Not BDATF_TBCAT.NoMatch Then NewCat = BDATF_TBLIN_CPIND.Value
                                'substitui a nova categoria
                                For J = I To Len(OldCat)
                                    If Mid(OldCat, J, 1) = ";" Then
                                        NumCarCat = J - I
                                        BDATU_TBCAT.Edit
                                        OldCat = Mid(OldCat, 1, I - 1) & NewCat & Mid(OldCat, J, Len(OldCat))
                                        BDATU_TBCAT_CPLIN.Value = OldCat
                                        BDATU_TBCAT.Update
                                        Exit For
                                    End If
                                Next J
                            End If
                            Cat = ""
                        Else
                            Cat = Cat & Mid(OldCat, I, 1)
                        End If
                    Next I
                    BSE.SimpleText = BSE.SimpleText + 1
                Else
                    BSI.SimpleText = BSI.SimpleText + 1
                End If
                BDATU_TBCAT.MoveNext
            Loop
            BDATU_TBCAT.MoveFirst
            'atualiza categorias consertadas
            BS.SimpleText = "Atualizando categorias corrigidas..."
            Do While Not BDATU_TBCAT.EOF
                BDATF_TBCAT.Seek "=", BDATU_TBCAT_CPCAT.Value
                If Not BDATF_TBCAT.NoMatch Then
                    BDATF_TBCAT.Edit
                    BDATF_TBCAT_CPCAT.Value = BDATU_TBCAT_CPCAT.Value
                    BDATF_TBCAT.Update
                End If
                BDATU_TBCAT.MoveNext
            Loop
        End If
        'apaga dados do banco de dados novos
        BS.SimpleText = "Apagando informações do banco de dados de novos..."
        'apaga bdnocos temporario
        BS.SimpleText = "Apagando banco de dados de novos links temporário..."
        FechaBD_Atualizacao_Novos
        Kill DosDirArq & "\BDNovos.af"
    End If
    'atualiza programa
    BS.SimpleText = "Atualizando Atomic Link..."
    CarregaCategorias
ERRO_ATOMICLINK:
    TelaEmEspera False
    LimpaTextos
    Exit Sub
ERRO_FALTAARQUIVO:
    MsgBox "Falta algum arquivo para completar a atualização do sistema.", vbCritical + vbOKOnly, "Falta arquivo"
    Exit Sub
End Sub
Private Sub BT_Backup_Click()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera True
    Dim NovoDir As String
    NovoDir = Diretorio
    If NovoDir <> "" Then
        LB_DE.Caption = NovoDir
    Else
        LB_DE.Caption = App.path
    End If
    TelaEmEspera False
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_CA_Click()
    If TXT_Link.Text = "" Then
        MsgBox "Digite primeiro o nome do link", vbInformation + vbOKOnly, "Falta link"
        TXT_Link.SetFocus
        Exit Sub
    End If
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando lista de categorias."
    BP.Max = LT_Categorias.ListCount + 1
    BP.Value = 0
    'carrega categorias
    LT_CA.Clear
    LT_LICA.Clear
    For I = 0 To LT_Categorias.ListCount - 1
        LT_CA.AddItem (LT_Categorias.List(I))
        BP.Value = BP.Value + 1
    Next I
    'carrega categorias do link
    BS.SimpleText = "Aguarde... carregando categorias deste link."
    CarregaCategoriasLink
    BP.Value = BP.Value + 1
    
    FR_BD.Visible = False
    FR_LICA.Visible = True
    TelaEmEspera (False)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_CH_Click()
    RespMsg = MsgBox("Você está prestes à chupinhar este arquivo. Deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, "Chupinhar arquivo")
    If RespMsg = vbYes Then
        TelaEmEspera (True)
        BS.SimpleText = "Aguarde... chupinhando o arquivo HTML"
        Chupinha
        TelaEmEspera (False)
        LimpaTextos
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_Fechar_Click()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    'pega categorias
    Dim Cat As String
    Cat = ""
    If LT_LICA.ListCount <> 0 Then
        For I = 0 To LT_LICA.ListCount - 1
            BDATF_TBCAT.Seek "=", LT_LICA.List(I)
            If Not BDATF_TBCAT.NoMatch Then Cat = Cat & BDATF_TBCAT_CPIND.Value & ";"
            BDAFN_TBCAT.Seek "=", LT_LICA.List(I)
            If Not BDAFN_TBCAT.NoMatch Then Cat = Cat & BDAFN_TBCAT_CPINO.Value & ";"
        Next I
    End If
    TXT_CA.Text = Cat
    FR_BD.Visible = True
    FR_LICA.Visible = False
    TelaEmEspera (False)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_LB_Click()
    On Error GoTo ERRO_ATOMICLINK
    If RB_AP.Value = False And RB_AN.Value = False Then
        MsgBox "Você deve primeiro selecionar qual arquivo será atualizado.", vbCritical + vbOKOnly, "Selecione o tipo"
        RB_AP.SetFocus
        Exit Sub
    End If
    'abre dialogo
    DLG.DialogTitle = "Indique o caminho do arquivo de backup"
    If RB_AP.Value = True Then
        DLG.Filter = "Arquivos de Backup do Atomic Link|al_bkp_f.af;"
    ElseIf RB_AN.Value = True Then
        DLG.Filter = "Arquivos de Backup do Atomic Link|al_bkp_n.af;"
    End If
    DLG.InitDir = App.path
    DLG.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
    DLG.ShowOpen
    'pega nome do arquivo
    LB_LO.Caption = DLG.FileName
    'carrega pagina
    If LB_LO.Caption <> "" Then If ArquivoExiste(LB_LO.Caption) = False Then Exit Sub
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_Net_Click()
    If UCase(Left(TXT_Link.Text, 4)) = "HTTP" Then Shell "start.exe " & Trim(TXT_Link.Text), vbHide
End Sub
Private Sub BT_P_Click()
    On Error GoTo ERRO_ATOMICLINK
    If LT_CA.ListIndex = -1 Then Exit Sub
    TelaEmEspera (True)
    'adiciona na lista nova
    LT_LICA.AddItem (LT_CA.List(LT_CA.ListIndex))
    LT_CA.RemoveItem (LT_CA.ListIndex)
    TelaEmEspera (False)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub BT_T_Click()
    On Error GoTo ERRO_ATOMICLINK
    If LT_LICA.ListIndex = -1 Then Exit Sub
    TelaEmEspera (True)
    'adiciona na lista nova
    LT_CA.AddItem (LT_LICA.List(LT_LICA.ListIndex))
    LT_LICA.RemoveItem (LT_LICA.ListIndex)
    TelaEmEspera (False)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub

Private Sub Menu_Chupinador_Abrir_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 3, False
    BT_AB_Click
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Chupinador_Chupinhar_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 3, False
    BT_AB_Click
    BT_CH_Click
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Exibir_Chupinhador_Click()
    ModoTela 3, False
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_CA_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then LT_Link.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub CB_Exibir_Click()
    On Error GoTo ERRO_ATOMICLINK
    TXT_CA.Enabled = False
    LT_Link.Clear
    LimpaTextos
    If CB_Exibir.Text = "Todos os Links" Then
        CarregaLinks
    ElseIf CB_Exibir.Text = "Novos" Then
        CarregaLinksNovos
    ElseIf CB_Exibir.Text = "Editados" Then
        CarregaLinksEditados
    ElseIf CB_Exibir.Text = "Chupinhados" Then
        CarregaLinksChupinhados
    ElseIf CB_Exibir.Text = "Registrados" Then
        CarregaLinksRegistrados
    ElseIf CB_Exibir.Text = "Re-Editados" Then
        CarregaLinksReEditados
    Else
        LT_Link.Clear
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub CB_Exibir_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then LT_Link.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub CB_Lingua_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    CB_Lingua.SelLength = Len(CB_Lingua.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub CB_Lingua_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then TXT_Titulo.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Form_Load()
    'On Error GoTo ERRO_ATOMICLINK
    'Verifica se os arquivos necessarios estão no diretorio
    If ArquivoExiste(App.path & "\BDAF.af") = False Then GoTo ERRO_FALTAARQUIVO
    If ArquivoExiste(App.path & "\BDNovos.af") = False Then GoTo ERRO_FALTAARQUIVO
    If ArquivoExiste(App.path & "\AtomicLink.html") = False Then GoTo ERRO_FALTAARQUIVO
    If ArquivoExiste(App.path & "\arj.exe") = False Then GoTo ERRO_FALTAARQUIVO
    If ArquivoExiste(App.path & "\Logotipo.jpg") = False Then GoTo ERRO_FALTAARQUIVO
    TelaEmEspera (True)
    If AbreBD = False Then End
    LimpaTextos
    ModoTela 1, False
    CarregaCategorias
    DeleteMenu GetSystemMenu(Me.hwnd, False), 6, MF_BYPOSITION 'desabilita X
    TelaEmEspera (False)
    Exit Sub
ERRO_FALTAARQUIVO:
    MsgBox "Está faltando algum arquivo necessário para o bom funcionamento do Atomic Link. Verifique o diretório.", vbCritical + vbOKOnly, "Faltam arquivos"
    End
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Form_Resize()
    On Error GoTo ERRO_ATOMICLINK
    ArrumaTela
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_ATOMICLINK
    FechaBD
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub LT_Categorias_Click()
    On Error GoTo ERRO_ATOMICLINK
    'Procura dados das categorias
    If LT_Categorias.ListIndex = -1 Then GoTo ERRO_ATOMICLINK
    LimpaTextos
    If BDATF_TBCAT.RecordCount > 0 Then 'Registrados
        BDATF_TBCAT.Seek "=", LT_Categorias.Text
        If Not BDATF_TBCAT.NoMatch Then
            TXT_Categorias.ForeColor = &H0&
            TXT_CatInd.ForeColor = &H0&
            TXT_CatDes.ForeColor = &H0&
            If BDATF_TBCAT_CPCAT.Value <> "" Then TXT_Categorias.Text = BDATF_TBCAT_CPCAT.Value
            If BDATF_TBCAT_CPDES.Value <> "" Then TXT_CatDes.Text = BDATF_TBCAT_CPDES.Value
            If BDATF_TBCAT_CPIND.Value <> "" Then TXT_CatInd.Text = BDATF_TBCAT_CPIND.Value
            LinkNovo = False
            GoTo ERRO_ATOMICLINK
        End If
    End If
    If BDAFN_TBCAT.RecordCount > 0 Then 'Novos
        BDAFN_TBCAT.Seek "=", LT_Categorias.Text
        If Not BDAFN_TBCAT.NoMatch Then
            TXT_Categorias.ForeColor = &HFF&
            TXT_CatDes.ForeColor = &HFF&
            TXT_CatInd.ForeColor = &HFF&
            If BDAFN_TBCAT_CPCAT.Value <> "" Then TXT_Categorias.Text = BDAFN_TBCAT_CPCAT.Value
            If BDAFN_TBCAT_CPDES.Value <> "" Then TXT_CatDes.Text = BDAFN_TBCAT_CPDES.Value
            If BDAFN_TBCAT_CPINO.Value <> "" Then TXT_CatInd.Text = BDAFN_TBCAT_CPINO.Value
            LinkNovo = True
            GoTo ERRO_ATOMICLINK
        End If
    End If
    'Se chegar até aqui não tem categorias
    MsgBox "Não foi possível encontrar a categoria correspondente.", vbCritical + vbOKOnly, "Falha na procura"
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi possível procurar dados sobre a categoria. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub LT_Link_Click()
    On Error GoTo ERRO_ATOMICLINK
    'Procura dados dos links
    If LT_Link.ListIndex = -1 Then Exit Sub
    TelaEmEspera True
    BS.SimpleText = "Aguarde... procurando dados sobre o link selecionado."
    LimpaTextos
    If CB_Exibir.Text = "Registrados" Then GoTo PEGA_REGISTRADOS
    If BDAFN_TBLIN.RecordCount > 0 Then 'Novos
        BDAFN_TBLIN.Seek "=", LT_Link.Text
        If Not BDAFN_TBLIN.NoMatch Then
            If BDAFN_TBLIN_CPEDI.Value = True Then
                TXT_CA.ForeColor = &HFF&
                TXT_Link.ForeColor = &HFF&
                TXT_Data.ForeColor = &HFF&
                CB_Lingua.ForeColor = &HFF&
                TXT_Titulo.ForeColor = &HFF&
                TXT_Descricao.ForeColor = &HFF&
                TXT_Procedencia.ForeColor = &HFF&
                TXT_Email.ForeColor = &HFF&
            ElseIf BDAFN_TBLIN_CPCHU.Value = True Then
                TXT_CA.ForeColor = &HC00000
                TXT_Link.ForeColor = &HC00000
                TXT_Data.ForeColor = &HC00000
                CB_Lingua.ForeColor = &HC00000
                TXT_Titulo.ForeColor = &HC00000
                TXT_Descricao.ForeColor = &HC00000
                TXT_Procedencia.ForeColor = &HC00000
                TXT_Email.ForeColor = &HC00000
            ElseIf BDAFN_TBLIN_CPREE.Value = True Then
                TXT_CA.ForeColor = &HC000C0
                TXT_Link.ForeColor = &HC000C0
                TXT_Data.ForeColor = &HC000C0
                CB_Lingua.ForeColor = &HC000C0
                TXT_Titulo.ForeColor = &HC000C0
                TXT_Descricao.ForeColor = &HC000C0
                TXT_Procedencia.ForeColor = &HC000C0
                TXT_Email.ForeColor = &HC000C0
            Else 'é novo
                TXT_CA.ForeColor = &H8000&
                TXT_Link.ForeColor = &H8000&
                TXT_Data.ForeColor = &H8000&
                CB_Lingua.ForeColor = &H8000&
                TXT_Titulo.ForeColor = &H8000&
                TXT_Descricao.ForeColor = &H8000&
                TXT_Procedencia.ForeColor = &H8000&
                TXT_Email.ForeColor = &H8000&
            End If
            If BDAFN_TBLIN_CPLIN.Value <> "" Then TXT_Link.Text = BDAFN_TBLIN_CPLIN.Value
            If BDAFN_TBLIN_CPDAT.Value <> "" Then TXT_Data.Text = BDAFN_TBLIN_CPDAT.Value
            If BDAFN_TBLIN_CPLGA.Value <> "" Then CB_Lingua.Text = BDAFN_TBLIN_CPLGA.Value
            If BDAFN_TBLIN_CPTIT.Value <> "" Then TXT_Titulo.Text = BDAFN_TBLIN_CPTIT.Value
            If BDAFN_TBLIN_CPDES.Value <> "" Then TXT_Descricao.Text = BDAFN_TBLIN_CPDES.Value
            If BDAFN_TBLIN_CPPRO.Value <> "" Then TXT_Procedencia.Text = BDAFN_TBLIN_CPPRO.Value
            If BDAFN_TBLIN_CPEMA.Value <> "" Then TXT_Email.Text = BDAFN_TBLIN_CPEMA.Value
            If BDAFN_TBLIN_CPCAT.Value <> "" Then TXT_CA.Text = BDAFN_TBLIN_CPCAT.Value
            LinkNovo = True
            GoTo ERRO_ATOMICLINK
        End If
    End If
PEGA_REGISTRADOS:
    If BDATF_TBLIN.RecordCount > 0 Then 'Registrados
        BDATF_TBLIN.Seek "=", LT_Link.Text
        If Not BDATF_TBLIN.NoMatch Then
            TXT_CA.ForeColor = &H0&
            TXT_Link.ForeColor = &H0&
            TXT_Data.ForeColor = &H0&
            CB_Lingua.ForeColor = &H0&
            TXT_Titulo.ForeColor = &H0&
            TXT_Descricao.ForeColor = &H0&
            TXT_Procedencia.ForeColor = &H0&
            TXT_Email.ForeColor = &H0&
            If BDATF_TBLIN_CPLIN.Value <> "" Then TXT_Link.Text = BDATF_TBLIN_CPLIN.Value
            If BDATF_TBLIN_CPDAT.Value <> "" Then TXT_Data.Text = BDATF_TBLIN_CPDAT.Value
            If BDATF_TBLIN_CPLGA.Value <> "" Then CB_Lingua.Text = BDATF_TBLIN_CPLGA.Value
            If BDATF_TBLIN_CPTIT.Value <> "" Then TXT_Titulo.Text = BDATF_TBLIN_CPTIT.Value
            If BDATF_TBLIN_CPDES.Value <> "" Then TXT_Descricao.Text = BDATF_TBLIN_CPDES.Value
            If BDATF_TBLIN_CPPRO.Value <> "" Then TXT_Procedencia.Text = BDATF_TBLIN_CPPRO.Value
            If BDATF_TBLIN_CPEMA.Value <> "" Then TXT_Email.Text = BDATF_TBLIN_CPEMA.Value
            If BDATF_TBLIN_CPCAT.Value <> "" Then TXT_CA.Text = BDATF_TBLIN_CPCAT.Value
            LinkNovo = False
            GoTo ERRO_ATOMICLINK
        End If
    End If
    'Se chegar até aqui não tem link
    MsgBox "Não foi possível encontrar o link correspondente.", vbCritical + vbOKOnly, "Falha na procura"
ERRO_ATOMICLINK:
    TelaEmEspera False
    BS.SimpleText = ""
    If Err Then MsgBox "Não foi possível procurar dados sobre link. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Exibir_Atualizacao_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 4, False
    LB_LO.Caption = App.path
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Exibir_Backup_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 5, False
    LB_DE.Caption = App.path
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Exibir_Categorias_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 2, False
    LT_Categorias.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Exibir_Links_Click()
    On Error GoTo ERRO_ATOMICLINK
    ModoTela 1, False
    CB_Exibir.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Internet_Click()
    On Error GoTo ERRO_ATOMICLINK
    Me.Hide
    Tela_Navegador.Show
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Apagar_Click()
    On Error GoTo ERRO_ATOMICLINK
    If FR_BD.Visible = True Then
        If LT_Link.ListIndex = -1 Then
            MsgBox "Selecione primeiro um link na lista para ser apagado.", vbInformation + vbOKOnly, "Falta link"
            LT_Link.SetFocus
            Exit Sub
        End If
        If LinkNovo = False Then
            MsgBox "Para deletar links que já estão registrados no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
            Exit Sub
        End If
        RespMsg = MsgBox("Você tem certeza que deseja remover do banco de dados o link:" & vbCr & vbCr & Trim(LT_Link.Text), vbInformation + vbYesNo + vbDefaultButton1, "Apagar link")
        If RespMsg = vbYes Then
            If LinkNovo = True Then 'Esta no BD de links novos
                BDAFN_TBLIN.Delete
                LT_Link.RemoveItem (LT_Link.ListIndex)
            Else
                MsgBox "Para apagar links que já estão registrados no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
            End If
        End If
    ElseIf FR_CA.Visible = True Then
        If LT_Categorias.ListIndex = -1 Then
            MsgBox "Selecione primeiro uma categoria na lista para ser apagada.", vbInformation + vbOKOnly, "Falta categoria"
            LT_Categorias.SetFocus
            Exit Sub
        End If
        If LinkNovo = False Then
            MsgBox "Para deletar categorias que já estão registradas no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
            Exit Sub
        End If
        RespMsg = MsgBox("Você tem certeza que deseja remover do banco de dados o categoria:" & vbCr & vbCr & Trim(LT_Categorias.Text), vbInformation + vbYesNo + vbDefaultButton1, "Apagar categoria")
        If RespMsg = vbYes Then
            If LinkNovo = True Then
                BDAFN_TBCAT.Delete
                LT_Categorias.RemoveItem (LT_Categorias.ListIndex)
            Else
                MsgBox "Para apagar categorias que já estão registrados no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
            End If
        End If
    End If
    LimpaTextos
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Cancelar_Click()
    On Error GoTo ERRO_ATOMICLINK
    LimpaTextos
    If FR_BD.Visible = True Or FR_LICA.Visible = True Then
        ModoTela 1, False
        LT_Link.ListIndex = -1
        LT_Link.SetFocus
    ElseIf FR_CA.Visible = True Then
        ModoTela 2, False
        TXT_Categorias.Enabled = True
        LT_Categorias.ListIndex = -1
        LT_Categorias.SetFocus
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Editar_Click()
    On Error GoTo ERRO_ATOMICLINK
    If FR_BD.Visible = True Or FR_LICA.Visible = True Then
        If LT_Link.ListIndex = -1 Then
            MsgBox "Selecione primeiro um link na lista para ser editado.", vbInformation + vbOKOnly, "Falta link"
            LT_Link.SetFocus
            Exit Sub
        End If
    ElseIf FR_CA.Visible = True Then
        If LT_Categorias.ListIndex = -1 Then
            MsgBox "Selecione primeiro uma categoria na lista para ser editada.", vbInformation + vbOKOnly, "Falta categoria"
            LT_Categorias.SetFocus
            Exit Sub
        End If
    End If
    ModoEdicao = True
    If FR_BD.Visible = True Or FR_LICA.Visible = True Then
        ModoTela 1, True
        If LinkNovo = False Then 'Re-Editado
            TXT_CA.ForeColor = &HC000C0
            TXT_Link.ForeColor = &HC000C0
            TXT_Data.ForeColor = &HC000C0
            CB_Lingua.ForeColor = &HC000C0
            TXT_Titulo.ForeColor = &HC000C0
            TXT_Descricao.ForeColor = &HC000C0
            TXT_Procedencia.ForeColor = &HC000C0
            TXT_Email.ForeColor = &HC000C0
        Else 'Editado
            TXT_CA.ForeColor = &HFF&
            TXT_Link.ForeColor = &HFF&
            TXT_Data.ForeColor = &HFF&
            CB_Lingua.ForeColor = &HFF&
            TXT_Titulo.ForeColor = &HFF&
            TXT_Descricao.ForeColor = &HFF&
            TXT_Procedencia.ForeColor = &HFF&
            TXT_Email.ForeColor = &HFF&
        End If
        TXT_Link.Enabled = False
        TXT_CA.Enabled = True
        CB_Lingua.SetFocus
    ElseIf FR_CA.Visible = True Then
        ModoTela 2, True
        If LinkNovo = True Then
            TXT_Categorias.Enabled = False
            TXT_CatDes.SetFocus
        Else
            MsgBox "Para alterar categorias que já estão registradas no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
            Menu_Principal_Cancelar_Click
        End If
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Finalizar_Click()
    On Error GoTo ERRO_ATOMICLINK
    RespMsg = MsgBox("Você tem certeza que deseja finalizar o Atomic Link ?", vbInformation + vbYesNo + vbDefaultButton1, "Finalizar")
    If RespMsg = vbYes Then
        End
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Novo_Click()
    On Error GoTo ERRO_ATOMICLINK
    LinkNovo = True
    If FR_BD.Visible = True Then
        ModoTela 1, True
        LT_Link.ListIndex = -1
        LimpaTextos
        TXT_CA.Enabled = True
        TXT_CA.ForeColor = &H8000&
        TXT_Link.ForeColor = &H8000&
        TXT_Data.ForeColor = &H8000&
        CB_Lingua.ForeColor = &H8000&
        TXT_Titulo.ForeColor = &H8000&
        TXT_Descricao.ForeColor = &H8000&
        TXT_Procedencia.ForeColor = &H8000&
        TXT_Email.ForeColor = &H8000&
        TXT_Data.Text = Format(Date, "dd/mm/yyyy")
        TXT_Link.Enabled = True
        TXT_Link.SetFocus
    ElseIf FR_CA.Visible = True Then 'Categorias
        ModoTela 2, True
        LT_Categorias.ListIndex = -1
        LimpaTextos
        If BDAFN_TBCAT.RecordCount = 0 Then
            TXT_CatInd.Text = "N0"
        Else
            BDAFN_TBNCT.MoveFirst
            TXT_CatInd.Text = "N" & (BDAFN_TBNCT_CPNCT.Value + 1)
        End If
        TXT_Categorias.SetFocus
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub Menu_Principal_Salvar_Click()
    On Error GoTo ERRO_ATOMICLINK
    If FR_BD.Visible = True Or FR_LICA.Visible = True Then
        TelaEmEspera (True)
        BS.SimpleText = "Salvando link..."
        If TXT_Link.Text = "" Then
            MsgBox "Não foi digitado o link.", vbExclamation + vbOKOnly, "Falta link"
            TXT_Link.SetFocus
            GoTo SAIDA
        End If
        If LinkNovo = False Then 'Edicao de link registrado
            RespMsg = MsgBox("Você está prestes à alterar os dados de um link já registrado. Você tem certeza que deseja continuar ?", vbExclamation + vbYesNo + vbDefaultButton1, "Editar Link Registrado")
            If RespMsg = vbNo Then GoTo SAIDA
        End If
        'Comeca a salvar dados
        If ModoEdicao = False And LinkNovo = True Then
            BDAFN_TBLIN.Seek "=", TXT_Link.Text
            If Not BDAFN_TBLIN.NoMatch Then
                MsgBox "Este link já existe no banco de dados de novos links. Digite outro.", vbExclamation + vbOKOnly, "Link já existe"
                GoTo SAIDA
            End If
        End If
        If ModoEdicao = False And LinkNovo = True Then
            BDAFN_TBLIN.AddNew
        ElseIf ModoEdicao = True And LinkNovo = False Then
            BDAFN_TBLIN.AddNew
        Else
            BDAFN_TBLIN.Edit
        End If
        BDAFN_TBLIN_CPLIN.Value = TXT_Link.Text
        BDAFN_TBLIN_CPDAT.Value = TXT_Data.Text
        BDAFN_TBLIN_CPTIT.Value = TXT_Titulo.Text
        BDAFN_TBLIN_CPDES.Value = TXT_Descricao.Text
        BDAFN_TBLIN_CPCAT.Value = TXT_CA.Text
        BDAFN_TBLIN_CPLGA.Value = CB_Lingua.Text
        BDAFN_TBLIN_CPPRO.Value = TXT_Procedencia.Text
        BDAFN_TBLIN_CPEMA.Value = TXT_Email.Text
        If ModoEdicao = False And LinkNovo = True Then 'Novo
            BDAFN_TBLIN_CPEDI.Value = False
            BDAFN_TBLIN_CPREE.Value = False
        ElseIf ModoEdicao = True And LinkNovo = False Then 'Edicao Registrado
            BDAFN_TBLIN_CPEDI.Value = False
            BDAFN_TBLIN_CPREE.Value = True
        Else 'Edicao novos
            BDAFN_TBLIN_CPEDI.Value = True
            BDAFN_TBLIN_CPREE.Value = False
        End If
        BDAFN_TBLIN_CPCHU.Value = False
        BDAFN_TBLIN.Update
        ModoTela 1, False
        Dim CBInd As Integer
        CBInd = CB_Exibir.ListIndex
        CB_Exibir.ListIndex = -1
        CB_Exibir.ListIndex = CBInd 'Esta rotina recarrega a lista de links
        LT_Link.ListIndex = -1
    ElseIf FR_CA.Visible = True Then 'Categorias
        TelaEmEspera (True)
        BS.SimpleText = "Salvando categoria..."
        If TXT_Categorias.Text = "" Then
            MsgBox "Não foi digitado a categoria.", vbExclamation + vbOKOnly, "Falta categoria"
            TXT_Categorias.SetFocus
            GoTo SAIDA
        End If
        If LinkNovo = True Then 'Esta no BD de categorias novas
            If ModoEdicao = False Then
                'verifica se categoria registrada ja existe
                BDATF_TBCAT.Seek "=", TXT_Categorias.Text
                If Not BDATF_TBCAT.NoMatch Then
                    MsgBox "Esta categoria já existe no banco de dados de categorias registradas. Digite outra.", vbExclamation + vbOKOnly, "Categoria já existe"
                    GoTo SAIDA
                End If
                'verifica se categoria nova ja existe
                BDAFN_TBCAT.Seek "=", TXT_Categorias.Text
                If Not BDAFN_TBCAT.NoMatch Then
                    MsgBox "Esta categoria já existe no banco de dados de categorias novas. Digite outra.", vbExclamation + vbOKOnly, "Categoria já existe"
                    GoTo SAIDA
                End If
            End If
            If ModoEdicao = False Then
                BDAFN_TBCAT.AddNew
                LT_Categorias.AddItem (TXT_Categorias.Text)
            Else
                BDAFN_TBCAT.Edit
            End If
            BDAFN_TBCAT_CPINO.Value = TXT_CatInd.Text
            BDAFN_TBCAT_CPCAT.Value = TXT_Categorias.Text
            BDAFN_TBCAT_CPDES.Value = TXT_CatDes.Text
            BDAFN_TBCAT.Update
            'altera indice
            BDAFN_TBNCT.MoveFirst
            Dim III As Integer
            III = BDAFN_TBNCT_CPNCT.Value + 1
            BDAFN_TBNCT.Edit
            BDAFN_TBNCT_CPNCT.Value = III
            BDAFN_TBNCT.Update
        ElseIf LinkNovo = False Then
            MsgBox "Para alterar categorias que já estão registradas no banco de dados principal, entre em contato com o administrador.", vbCritical + vbOKOnly, "Erro de Procedimento"
        End If
        ModoTela 2, False
    ElseIf FR_BA.Visible = True Then 'Backup
        If RB_PR.Value = False And RB_NO.Value = False Then
            MsgBox "Você deve selecionar pelo menos uma opção de backup abaixo.", vbCritical + vbOKOnly, "Falta tipo de backup"
            GoTo SAIDA
        End If
        RespMsg = MsgBox("Você está prestes à fazer o backup da banco de dados do Atomic Link. Deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, "Backup")
        If RespMsg = vbYes Then
            TelaEmEspera (True)
            Dim DosDir As String, Tamanho As Long, DosDirArq As String, ShortDir As String
            ShortDir = Space$(1024)
            Tamanho = GetShortPathName(Trim(LB_DE.Caption), ShortDir, Len(ShortDir))
            DosDir = Left$(ShortDir, Tamanho)
            ShortDir = Space$(1024)
            Tamanho = GetShortPathName(App.path, ShortDir, Len(ShortDir))
            DosDirArq = Left$(ShortDir, Tamanho)
            Dim DirWinTmp As String
            DirWinTmp = Environ$("temp")
            If DirWinTmp = "" Then DirWinTmp = "c:\Windows\Temp"
            FechaBD
            If RB_PR.Value = True Then
                'repara, compacta, encripta e joga senha no banco de dados novo
                BS.SimpleText = "Reparando e compactando banco de dados..."
                DBEngine.CompactDatabase App.path & "\BDAF.af", App.path & "\xyz.af", , dbEncrypt
                'renomeia arquivo
                Kill DosDirArq & "\BDAF.af"
                FileCopy DosDirArq & "\xyz.af", DosDirArq & "\BDAF.af"
                Kill DosDirArq & "\xyz.af"
                'cria backup
                BS.SimpleText = "Aguarde... criando arquivo de backup - al_bkp_f.af"
                Shell App.path & "/arj a -jm -y " & DosDir & "\al_bkp_f.af " & DosDirArq & "\BDAF.af", vbMinimizedNoFocus
            ElseIf RB_NO.Value = True Then
                'repara, compacta, encripta e joga senha no banco de dados novo
                BS.SimpleText = "Reparando e compactando banco de dados..."
                DBEngine.CompactDatabase App.path & "\BDNovos.af", App.path & "\xyz.af", , dbEncrypt
                'renomeia arquivo
                Kill DosDirArq & "\BDNovos.af"
                FileCopy DosDirArq & "\xyz.af", DosDirArq & "\BDNovos.af"
                Kill DosDirArq & "\xyz.af"
                'cria backup
                BS.SimpleText = "Aguarde... criando arquivo de backup - al_bkp_n.af"
                Shell App.path & "/arj a -jm -y " & DosDir & "\al_bkp_n.af " & DosDirArq & "\BDNovos.af", vbMinimizedNoFocus
            End If
            FechaArj
            If AbreBD = False Then End
            'apaga dados do banco de dados novos
            BS.SimpleText = "Removendo informações do banco de dados novos..."
            If BDAFN_TBCAT.RecordCount > 0 Then
                BDAFN_TBCAT.MoveFirst
                Do While Not BDAFN_TBCAT.EOF
                    BDAFN_TBCAT.Delete
                    BDAFN_TBCAT.MoveNext
                Loop
            End If
            If BDAFN_TBLIN.RecordCount > 0 Then
                BDAFN_TBLIN.MoveFirst
                Do While Not BDAFN_TBLIN.EOF
                    BDAFN_TBLIN.Delete
                    BDAFN_TBLIN.MoveNext
                Loop
            End If
        End If
    End If
    LimpaTextos
    If FR_BD.Visible = True Then
        TXT_Link.Enabled = True
        LT_Link.ListIndex = -1
        LT_Link.SetFocus
    ElseIf FR_CA.Visible = True Then
        LT_Categorias.ListIndex = -1
        LT_Categorias.SetFocus
    End If
SAIDA:
    LimpaTextos
    TelaEmEspera (False)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_CatDes_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    TXT_CatDes.SelLength = Len(TXT_CatDes.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Categorias_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    TXT_Categorias.SelLength = Len(TXT_Categorias.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Categorias_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then TXT_CatDes.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Descricao_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    TXT_Descricao.SelLength = Len(TXT_Descricao.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Link_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    TXT_Link.SelLength = Len(TXT_Link.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Link_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then CB_Lingua.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Titulo_GotFocus()
    On Error GoTo ERRO_ATOMICLINK
    TXT_Titulo.SelLength = Len(TXT_Titulo.Text)
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Sub TXT_Titulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_ATOMICLINK
    If KeyAscii = vbKeyReturn Then TXT_Descricao.SetFocus
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub



'****************************************************************
'                           F U N Ç Õ E S
'****************************************************************
Private Static Sub TelaEmEspera(Espera As Boolean)
    If Espera = True Then
        Tela_Principal.Enabled = False
        Tela_Principal.MousePointer = vbHourglass
    Else
        Tela_Principal.Enabled = True
        Tela_Principal.MousePointer = vbDefault
    End If
End Sub
Private Static Sub ArrumaTela()
    On Error GoTo ERRO_ATOMICLINK
    If Tela_Principal.WindowState = vbMinimized Then Exit Sub
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... organizando a tela."
    FR_BD.Top = BF.Height
    FR_CA.Top = BF.Height
    FR_AT.Top = BF.Height
    FR_BA.Top = BF.Height
    FR_CH.Top = BF.Height
    FR_LICA.Top = BF.Height
    FR_BD.Left = -50
    FR_CA.Left = -50
    FR_AT.Left = -50
    FR_BA.Left = -50
    FR_CH.Left = -50
    FR_LICA.Left = -50
    If Me.WindowState = vbMaximized Then
        FR_BD.Width = Screen.Width + 100
        FR_CA.Width = Screen.Width + 100
        FR_AT.Width = Screen.Width + 100
        FR_BA.Width = Screen.Width + 100
        FR_CH.Width = Screen.Width + 100
        FR_LICA.Width = Screen.Width + 100
        FR_BD.Height = Screen.Height - BS.Height
        FR_CA.Height = Screen.Height - BS.Height
        FR_AT.Height = Screen.Height - BS.Height
        FR_BA.Height = Screen.Height - BS.Height
        FR_CH.Height = Screen.Height - BS.Height
        FR_LICA.Height = Screen.Height - BS.Height
    Else
        FR_BD.Width = Me.Width
        FR_CA.Width = Me.Width
        FR_AT.Width = Me.Width
        FR_BA.Width = Me.Width
        FR_CH.Width = Me.Width
        FR_LICA.Width = Me.Width
        FR_BD.Height = Me.Height - BS.Height
        FR_CA.Height = Me.Height - BS.Height
        FR_AT.Height = Me.Height - BS.Height
        FR_BA.Height = Me.Height - BS.Height
        FR_CH.Height = Me.Height - BS.Height
        FR_LICA.Height = Me.Height - BS.Height
    End If
    'FR_CH
    WB.Left = 200
    WB.Height = FR_CH.Height - WB.Top - 1500
    WB.Width = FR_CH.Width - 400
    'Reposiciona frames
    FR_B2.Left = (FR_BD.Width - FR_B2.Width) / 2
    FR_L2.Left = (FR_LICA.Width - FR_L2.Width) / 2
    FR_C2.Left = (FR_CA.Width - FR_C2.Width) / 2
    FR_RE.Left = (FR_CH.Width - FR_RE.Width) / 2
    FR_BO.Left = (FR_BA.Width - FR_BO.Width) / 2
    FR_AS.Left = (FR_AT.Width - FR_AS.Width) / 2

ERRO_ATOMICLINK:
    BS.SimpleText = ""
    TelaEmEspera False
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Function IL(Estado As Boolean) As Boolean
    If Estado = False Then
        IL = True
    Else
        IL = False
    End If
End Function
Private Static Sub ModoTela(Tela As Integer, Edicao As Boolean)
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    
    RB_AN.Enabled = False
    
    BS.SimpleText = "Organizando exibição de telas..."
    Me.Menu_Exibir_Links.Checked = False
    Me.Menu_Exibir_Categorias.Checked = False
    Me.Menu_Exibir_Atualizacao.Checked = False
    Me.Menu_Exibir_Backup.Checked = False
    Me.Menu_Exibir_Chupinhador.Checked = False
    FR_BD.Visible = False
    FR_CA.Visible = False
    FR_AT.Visible = False
    FR_BA.Visible = False
    FR_CH.Visible = False
    FR_LICA.Visible = False
    If Tela = 1 Then 'FR_BD
        Me.Menu_Exibir_Links.Checked = True
        FR_BD.Visible = True
        FR_BD2.Enabled = Edicao
        CB_Exibir.Enabled = IL(Edicao)
        LT_Link.Enabled = IL(Edicao)
        BT_CA.Enabled = Edicao
        TXT_CA.Enabled = False
    ElseIf Tela = 2 Then 'FR_CA
        Me.Menu_Exibir_Categorias.Checked = True
        FR_CA.Visible = True
        FR_CA2.Enabled = Edicao
    ElseIf Tela = 3 Then 'FR_CH
        Me.Menu_Exibir_Chupinhador.Checked = True
        FR_CH.Visible = True
    ElseIf Tela = 4 Then 'FR_AT
        Me.Menu_Exibir_Atualizacao.Checked = True
        FR_AT.Visible = True
    ElseIf Tela = 5 Then 'FR_BA
        Me.Menu_Exibir_Backup.Checked = True
        FR_BA.Visible = True
    End If
    If Edicao = True Then
        BF.Buttons(1).Enabled = False
        BF.Buttons(2).Enabled = False
        BF.Buttons(3).Enabled = True
        BF.Buttons(4).Enabled = False
        BF.Buttons(5).Enabled = True
        BF.Buttons(7).Enabled = False
        Me.Menu_Principal_Novo.Enabled = False
        Me.Menu_Principal_Editar.Enabled = False
        Me.Menu_Principal_Salvar.Enabled = True
        Me.Menu_Principal_Apagar.Enabled = False
        Me.Menu_Principal_Cancelar.Enabled = True
        Me.Menu_Principal_Finalizar.Enabled = False
        Me.Menu_Editar.Enabled = True
        Me.Menu_Exibir.Enabled = False
        Me.Menu_Chupinador.Enabled = False
        Me.Menu_Sobre.Enabled = False
    Else
        BF.Buttons(1).Enabled = True
        BF.Buttons(2).Enabled = True
        BF.Buttons(3).Enabled = False
        BF.Buttons(4).Enabled = True
        BF.Buttons(5).Enabled = False
        BF.Buttons(7).Enabled = True
        Me.Menu_Principal_Novo.Enabled = True
        Me.Menu_Principal_Editar.Enabled = True
        Me.Menu_Principal_Salvar.Enabled = False
        Me.Menu_Principal_Apagar.Enabled = True
        Me.Menu_Principal_Cancelar.Enabled = False
        Me.Menu_Principal_Finalizar.Enabled = True
        Me.Menu_Editar.Enabled = False
        Me.Menu_Exibir.Enabled = True
        Me.Menu_Chupinador.Enabled = True
        Me.Menu_Sobre.Enabled = True
        ModoEdicao = False
    End If
    If Tela = 3 Or Tela = 4 Or Tela = 5 Then
        BF.Buttons(1).Enabled = False
        BF.Buttons(2).Enabled = False
        BF.Buttons(3).Enabled = True
        BF.Buttons(4).Enabled = False
        BF.Buttons(5).Enabled = False
        BF.Buttons(7).Enabled = True
        Me.Menu_Principal_Novo.Enabled = False
        Me.Menu_Principal_Editar.Enabled = False
        Me.Menu_Principal_Salvar.Enabled = True
        Me.Menu_Principal_Apagar.Enabled = False
        Me.Menu_Principal_Cancelar.Enabled = False
        Me.Menu_Principal_Finalizar.Enabled = True
        Me.Menu_Editar.Enabled = False
        Me.Menu_Exibir.Enabled = True
        Me.Menu_Chupinador.Enabled = True
        Me.Menu_Sobre.Enabled = True
    End If
    If Tela = 3 Or Tela = 4 Then
        BF.Buttons(3).Enabled = False
        Me.Menu_Principal_Salvar.Enabled = False
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
    ArrumaTela
    BS.SimpleText = ""
    TelaEmEspera (False)
End Sub
Private Static Sub CarregaLinks()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links catalogados."
    BP.Max = 1
    BP.Value = 0
    If BDATF_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDATF_TBLIN.RecordCount
    If BDAFN_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBLIN.RecordCount
    'Carrega Banco Dados Principal
    If BDATF_TBLIN.RecordCount > 0 Then
        BDATF_TBLIN.MoveFirst
        Do While Not BDATF_TBLIN.EOF
            LT_Link.AddItem BDATF_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDATF_TBLIN.MoveNext
        Loop
    End If
    'Carrega Banco Dados Novos
    If BDAFN_TBLIN.RecordCount > 0 Then
        BDAFN_TBLIN.MoveFirst
        Do While Not BDAFN_TBLIN.EOF
            If BDAFN_TBLIN_CPEDI.Value = True Then 'Editado
                LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            Else
                LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            End If
            BP.Value = BP.Value + 1
            BDAFN_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaLinksNovos()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links novos."
    BP.Max = 1
    BP.Value = 0
    If BDAFN_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBLIN.RecordCount
    'Carrega Banco Dados Novos
    If BDAFN_TBLIN.RecordCount > 0 Then
        BDAFN_TBLIN.MoveFirst
        Do While Not BDAFN_TBLIN.EOF
            If BDAFN_TBLIN_CPEDI.Value = False And BDAFN_TBLIN_CPCHU.Value = False And BDAFN_TBLIN_CPREE.Value = False Then LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDAFN_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links novos no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaLinksEditados()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links editados."
    BP.Max = 1
    BP.Value = 0
    If BDAFN_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBLIN.RecordCount
    'Carrega Banco Dados Novos
    If BDAFN_TBLIN.RecordCount > 0 Then
        BDAFN_TBLIN.MoveFirst
        Do While Not BDAFN_TBLIN.EOF
            If BDAFN_TBLIN_CPEDI.Value = True And BDAFN_TBLIN_CPCHU.Value = False And BDAFN_TBLIN_CPREE.Value = False Then LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDAFN_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links editados no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaLinksRegistrados()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links registrados."
    BP.Max = 1
    BP.Value = 0
    If BDATF_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDATF_TBLIN.RecordCount
    'Carrega Banco Dados Principal
    If BDATF_TBLIN.RecordCount > 0 Then
        BDATF_TBLIN.MoveFirst
        Do While Not BDATF_TBLIN.EOF
            LT_Link.AddItem BDATF_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDATF_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links registrados no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaLinksChupinhados()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links chupinhados."
    BP.Max = 1
    BP.Value = 0
    If BDAFN_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBLIN.RecordCount
    'Carrega Banco Dados Novos
    If BDAFN_TBLIN.RecordCount > 0 Then
        BDAFN_TBLIN.MoveFirst
        Do While Not BDAFN_TBLIN.EOF
            If BDAFN_TBLIN_CPEDI.Value = False And BDAFN_TBLIN_CPCHU.Value = True And BDAFN_TBLIN_CPREE.Value = False Then LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDAFN_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links chupinhados no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaLinksReEditados()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos os links re-editados."
    BP.Max = 1
    BP.Value = 0
    If BDAFN_TBLIN.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBLIN.RecordCount
    'Carrega Banco Dados Novos
    If BDAFN_TBLIN.RecordCount > 0 Then
        BDAFN_TBLIN.MoveFirst
        Do While Not BDAFN_TBLIN.EOF
            If BDAFN_TBLIN_CPEDI.Value = False And BDAFN_TBLIN_CPCHU.Value = False And BDAFN_TBLIN_CPREE.Value = True Then LT_Link.AddItem BDAFN_TBLIN_CPLIN.Value
            BP.Value = BP.Value + 1
            BDAFN_TBLIN.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then MsgBox "Não foi possível carregar lista de links editados no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub LimpaTextos()
    On Error GoTo ERRO_ATOMICLINK
    BS.SimpleText = ""
    'FR_BD
    TXT_CA.ForeColor = &H0&
    TXT_Link.ForeColor = &H0&
    TXT_Data.ForeColor = &H0&
    CB_Lingua.ForeColor = &H0&
    TXT_Titulo.ForeColor = &H0&
    TXT_Descricao.ForeColor = &H0&
    TXT_Procedencia.ForeColor = &H0&
    TXT_Email.ForeColor = &H0&
    TXT_CA.Text = ""
    TXT_Link.Text = ""
    TXT_Data.Text = ""
    TXT_Titulo.Text = ""
    TXT_Data.Text = ""
    CB_Lingua.Text = ""
    TXT_Descricao.Text = ""
    TXT_Procedencia.Text = ""
    TXT_Email.Text = ""
    'FR_CA
    TXT_Categorias.Text = ""
    TXT_CatInd.Text = ""
    TXT_CatDes.Text = ""
    'FR_CH
    BS_EN.SimpleText = 0
    BS_IG.SimpleText = 0
    BS_ER.SimpleText = 0
    ArquivoHTML = ""
    TXT_L.Text = ""
    WB.Navigate App.path & "/AtomicLink.html"
    BS.SimpleText = ""
    BSE.SimpleText = 0
    BSI.SimpleText = 0
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaCategorias()
    On Error GoTo ERRO_ATOMICLINK
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... carregando todos as categorias."
    BP.Max = 1
    BP.Value = 0
    If BDATF_TBCAT.RecordCount > 0 Then BP.Max = BP.Max + BDATF_TBCAT.RecordCount
    If BDAFN_TBCAT.RecordCount > 0 Then BP.Max = BP.Max + BDAFN_TBCAT.RecordCount
    'apaga lista atual
    LT_Categorias.Clear
    'carrega categorias
    If BDATF_TBCAT.RecordCount > 0 Then
        BDATF_TBCAT.MoveFirst
        Do While Not BDATF_TBCAT.EOF
            LT_Categorias.AddItem BDATF_TBCAT_CPCAT.Value
            BP.Value = BP.Value + 1
            BDATF_TBCAT.MoveNext
        Loop
    End If
    If BDAFN_TBCAT.RecordCount > 0 Then
        BDAFN_TBCAT.MoveFirst
        Do While Not BDAFN_TBCAT.EOF
            LT_Categorias.AddItem BDAFN_TBCAT_CPCAT.Value
            BP.Value = BP.Value + 1
            BDAFN_TBCAT.MoveNext
        Loop
    End If
ERRO_ATOMICLINK:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera False
    If Err Then MsgBox "Não foi possível carregar lista de categorias no banco de dados. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub CarregaCategoriasLink()
    On Error GoTo ERRO_ATOMICLINK
    If TXT_CA.Text = "" Then Exit Sub
    Dim Cat As String
    Cat = ""
    BDATF_TBCAT.Index = "Índice"
    BDAFN_TBCAT.Index = "Índice Novo"
    For I = 1 To Len(TXT_CA.Text)
        If Mid(TXT_CA.Text, I, 1) = ";" Then
            'procura categorias registradas
            If BDATF_TBCAT.RecordCount > 0 Then
                BDATF_TBCAT.Seek "=", Cat
                If Not BDATF_TBCAT.NoMatch Then
                    LT_LICA.AddItem (BDATF_TBCAT_CPCAT.Value)
                    For J = 0 To LT_CA.ListCount = -1
                        If LT_CA.List(J) = BDATF_TBCAT_CPCAT.Value Then
                            LT_CA.RemoveItem (J)
                            Exit For
                        End If
                    Next J
                Else
                    MsgBox "Não foi possível encontrar a categorias de índice: " & Cat, vbCritical + vbOKOnly, "Erro na procura"
                End If
            End If
            'procura por categorias novas
            If BDAFN_TBCAT.RecordCount > 0 Then
                BDAFN_TBCAT.Seek "=", Cat
                If Not BDAFN_TBCAT.NoMatch Then
                    LT_LICA.AddItem (BDAFN_TBCAT_CPCAT.Value)
                    For J = 0 To LT_CA.ListCount = -1
                        If LT_CA.List(J) = BDAFN_TBCAT_CPCAT.Value Then
                            LT_CA.RemoveItem (J)
                            Exit For
                        End If
                    Next J
                Else
                    MsgBox "Não foi possível encontrar a categorias de índice: " & Cat, vbCritical + vbOKOnly, "Erro na procura"
                End If
            End If
            Cat = ""
        Else
            Cat = Cat & Mid(TXT_CA.Text, I, 1)
        End If
    Next I
    'remove categrias
    For I = 0 To LT_LICA.ListCount - 1
        For J = 0 To LT_CA.ListCount - 1
            If LT_CA.List(J) = LT_LICA.List(I) Then
                LT_CA.RemoveItem (J)
                Exit For
            End If
        Next J
    Next I
ERRO_ATOMICLINK:
    BDATF_TBCAT.Index = "Categoria"
    BDAFN_TBCAT.Index = "Categoria"
    If Err Then MsgBox "Não foi possível carregar lista de categorias deste link.", vbCritical + vbOKOnly, "Erro de Acesso"
End Sub
Private Static Sub Chupinha()
    On Error GoTo ERRO_CHUPINHA
    Dim Str1 As String, Texto As Variant, Num As Long, N2 As Long, N3 As Long, ExiReg As Boolean
    Dim Link As String, Titulo As String, Descricao As String, Diretorio As String, Fonte As String
    Texto = ArquivoHTML
    Str1 = ""
    Num = 0
    BP.Max = Len(ArquivoHTML)
    BP.Value = 0
    BS_EN.SimpleText = 0
    BS_IG.SimpleText = 0
    BS_ER.SimpleText = 0
    Diretorio = ""
    For I = 1 To Len(Texto) - 8
        On Error GoTo ERRO_FOR
        Link = ""
        AchoLink = False
        'procura diretorio
        If Diretorio = "" Then
            If Mid(Texto, I, 7) = "<TITLE>" Then
                    For J = (I + 8) To Len(Texto) - 8
                        If (J - I - 9) > 2000 Then Exit For
                        If Mid(Texto, J, 8) = "</TITLE>" Then
                            Diretorio = Trim(Mid(Texto, (I + 8), (J - 10)))
                            Exit For
                        End If
                    Next J
            End If
            'Se nao achou o titulo
            If Diretorio = "" Then
                Do While RespMsg = ""
                    RespMsg = InputBox("Não foi possível achar o nome desta página - para poder prosseguir, por favor digite-o abaixo:", "Falta Nome", "")
                Loop
                Diretorio = RespMsg
            End If
            Me.Caption = Diretorio
        End If
        'Procura links
        If Mid(Texto, I, 8) = "<A HREF=" Then
            If Mid(Texto, I + 9, 4) = "http" Then 'achou um link
                AchoLink = True
                'pega endereço
                For J = (I + 9) To Len(Texto) - 8
                    If (J - I - 9) > 2000 Then GoTo ERRO_FOR
                    If Asc(Mid(Texto, J, 1)) = 34 Then
                        Link = Str1
                        N2 = J + 1
                        Num = Num + 1
                        BS_EN.SimpleText = BS_EN.SimpleText + 1
                        Str1 = ""
                        Exit For
                    End If
                    Str1 = Str1 + Mid(Texto, J, 1)
                Next J
                'pega titulo
                For K = N2 To Len(Texto) - 8
                    If (K - N2) > 2000 Then GoTo ERRO_FOR
                    If Mid(Texto, K, 1) = ">" Then
                        For M = (K + 1) To Len(Texto) - 8
                            If (M - K - 1) > 1000 Then GoTo ERRO_FOR
                            If Mid(Texto, M, 7) = "</A> - " Then
                                Titulo = Trim(Mid(Texto, (K + 1), (M - K - 1)))
                                N3 = M + 7
                                Exit For
                            End If
                        Next M
                        Exit For
                    End If
                Next K
                'pega descrição
                For R = N3 To Len(Texto) - 8
                    If (R - N3) > 2000 Then GoTo ERRO_FOR
                    If Mid(Texto, R, 4) = "<BR>" Then
                        Descricao = Trim(Mid(Texto, N3, (R - N3)))
                        Exit For
                    End If
                Next R
            End If
            AchoLink = True
        End If
        
        Titulo = CorrigeHTML(Titulo)
        Descricao = CorrigeHTML(Descricao)
        'Salva dados
        If Link <> "" And AchoLink = True Then
            If BDAFN_TBLIN.RecordCount > 0 Then
                BDAFN_TBLIN.Seek "=", Link
                If Not BDAFN_TBLIN.NoMatch Then
                    BS_IG.SimpleText = BS_IG.SimpleText + 1
                    ExiReg = True
                Else
                    ExiReg = False
                    BDAFN_TBLIN.AddNew
                End If
            Else
                ExiReg = False
                BDAFN_TBLIN.AddNew
            End If
            If ExiReg = False Then
                BDAFN_TBLIN_CPLIN.Value = Link
                BDAFN_TBLIN_CPDAT.Value = Format(Date, "dd/mm/yyyy")
                BDAFN_TBLIN_CPTIT.Value = Titulo
                BDAFN_TBLIN_CPDES.Value = Descricao
                BDAFN_TBLIN_CPCAT.Value = Diretorio
                BDAFN_TBLIN_CPLGA.Value = ""
                BDAFN_TBLIN_CPEDI.Value = False
                BDAFN_TBLIN_CPCHU.Value = True
                BDAFN_TBLIN.Update
            End If
        End If
ERRO_FOR:
        If Err Then BS_ER.SimpleText = BS_ER.SimpleText + 1
        BP.Value = BP.Value + 1
    Next I
    BP.Value = 0
    Texto = ""
    ArquivoHTML = ""
    Exit Sub
ERRO_CHUPINHA:
    MsgBox "Ocorreu algum erro enquanto o sistema chupinhava o arquivo.", vbCritical + vbOKOnly, "Erro do Chupinhador"
End Sub
Private Static Function CorrigeHTML(Str1 As String) As String
    On Error GoTo ERRO_ATOMICLINK
    If Str1 = "" Then Exit Function
    CorrigeHTML = ""
    Dim X As Integer, Num
    X = 0
    Do
    X = X + 1
    For Num = X To Len(Str1)
        If Mid(Str1, Num, 8) = "&aacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "á" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Aacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Á" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&eacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "é" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Eacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "É" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&iacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "í" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Iacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Í" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&oacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ó" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Oacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ó" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&uacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ú" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Uacute;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ú" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&agrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "â" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Agrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Â" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 7) = "&ecirc;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ê" & Right(Str1, Len(Str1) - (Num + 6))
        ElseIf Mid(Str1, Num, 6) = "&Ecirc;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ê" & Right(Str1, Len(Str1) - (Num + 5))
        ElseIf Mid(Str1, Num, 8) = "&igrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "î" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Igrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Î" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&ograve;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ô" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Ograve;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ô" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&ugrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "û" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Ugrave;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Û" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 7) = "&acirc;" Or Mid(Str1, Num, 8) = "&atilde;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ã" & Right(Str1, Len(Str1) - (Num + 6))
        ElseIf Mid(Str1, Num, 7) = "&Acirc;" Or Mid(Str1, Num, 8) = "&Atilde;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ã" & Right(Str1, Len(Str1) - (Num + 6))
        ElseIf Mid(Str1, Num, 7) = "&ocirc;" Or Mid(Str1, Num, 8) = "&otilde;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "õ" & Right(Str1, Len(Str1) - (Num + 6))
        ElseIf Mid(Str1, Num, 7) = "&Ocirc;" Or Mid(Str1, Num, 8) = "&Otilde;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Õ" & Right(Str1, Len(Str1) - (Num + 6))
        ElseIf Mid(Str1, Num, 6) = "&uuml;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ü" & Right(Str1, Len(Str1) - (Num + 5))
        ElseIf Mid(Str1, Num, 6) = "&Uuml;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ü" & Right(Str1, Len(Str1) - (Num + 5))
        ElseIf Mid(Str1, Num, 8) = "&ccedil;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "ç" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 8) = "&Ccedil;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "Ç" & Right(Str1, Len(Str1) - (Num + 7))
        ElseIf Mid(Str1, Num, 4) = "&lt;" Then
            CorrigeHTML = Left(Str1, Num - 1) & ">" & Right(Str1, Len(Str1) - (Num + 3))
        ElseIf Mid(Str1, Num, 4) = "&gt;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "<" & Right(Str1, Len(Str1) - (Num + 3))
        ElseIf Mid(Str1, Num, 5) = "&amp;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "&" & Right(Str1, Len(Str1) - (Num + 4))
        ElseIf Mid(Str1, Num, 6) = "&quot;" Then
            CorrigeHTML = Left(Str1, Num - 1) & Chr(34) & Right(Str1, Len(Str1) - (Num + 5))
        ElseIf Mid(Str1, Num, 5) = "&reg;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "®" & Right(Str1, Len(Str1) - (Num + 4))
        ElseIf Mid(Str1, Num, 6) = "&copy;" Then
            CorrigeHTML = Left(Str1, Num - 1) & "©" & Right(Str1, Len(Str1) - (Num + 5))
        End If
        If CorrigeHTML <> "" And CorrigeHTML <> Str1 Then Str1 = CorrigeHTML
        Exit For
    Next Num
    If X = Len(Str1) Then Exit Do
    Loop
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Function
Private Static Function Diretorio() As String
    On Error GoTo ERRO_ATOMICLINK
    Diretorio = ""
    Dim BD As BROWSEINFO
    Dim idl As ITEMIDLIST
    Dim rtn&, pidl&, path$, pos%
    BD.hOwner = Me.hwnd
    rtn& = SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal 17, idl)
    BD.pidlRoot = idl.mkid.cb
    BD.lpszTitle = "Indique o caminho de onde será feito o backup:"
    BD.ulFlags = BIF_RETURNONLYFSDIRS
    pidl& = SHBrowseForFolder(BD)
    path$ = Space$(512)
    lResult = SHGetPathFromIDList(ByVal pidl&, ByVal path$)
    If pidl& Then
       pos% = InStr(path$, Chr$(0))
       Diretorio = Left(path$, pos - 1)
    End If
ERRO_ATOMICLINK:
    If Err Then MsgBox "Não foi concluir esta operação pois ocorreu algum erro. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso"
End Function
Private Static Function ArquivoExiste(Arquivo As String) As Boolean
    On Error GoTo ERRO_ARQUIVO
    Open Arquivo For Input As #1
    Close #1
    ArquivoExiste = True
ERRO_ARQUIVO:
    If Err Then ArquivoExiste = False
End Function
