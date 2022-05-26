Attribute VB_Name = "BancoDados"
Public WKSP As Workspace
'Banco de Dados
Public BDATF As Database
Public BDAFN As Database
Public BDATU As Database

'Tabelas BDATF
Public BDATF_TBLIN As Recordset
Public BDATF_TBCAT As Recordset

'Tabelas BDAFN
Public BDAFN_TBLIN As Recordset
Public BDAFN_TBCAT As Recordset
Public BDAFN_TBNCT As Recordset

'Tabelas BDBDATU
Public BDATU_TBLIN As Recordset
Public BDATU_TBCAT As Recordset
Public BDATU_TBNCT As Recordset

'Campos da Tabela BDATF_TBLIN
Public BDATF_TBLIN_CPIND As Field 'Índice
Public BDATF_TBLIN_CPLIN As Field
Public BDATF_TBLIN_CPDAT As Field
Public BDATF_TBLIN_CPTIT As Field
Public BDATF_TBLIN_CPDES As Field
Public BDATF_TBLIN_CPCAT As Field
Public BDATF_TBLIN_CPLGA As Field
Public BDATF_TBLIN_CPPRO As Field 'Procedência
Public BDATF_TBLIN_CPEMA As Field 'e-mail
'Campos da Tabela BDATF_TBCAT
Public BDATF_TBCAT_CPIND As Field
Public BDATF_TBCAT_CPCAT As Field
Public BDATF_TBCAT_CPDES As Field
Public BDATF_TBCAT_CPLIN As Field


'Campos da Tabela BDAFN_TBLIN
Public BDAFN_TBLIN_CPIND As Field 'Índice Novo
Public BDAFN_TBLIN_CPLIN As Field 'Link
Public BDAFN_TBLIN_CPDAT As Field 'Data
Public BDAFN_TBLIN_CPTIT As Field 'Título
Public BDAFN_TBLIN_CPDES As Field 'Descricao
Public BDAFN_TBLIN_CPCAT As Field 'Categoria
Public BDAFN_TBLIN_CPLGA As Field 'Lingua
Public BDAFN_TBLIN_CPPRO As Field 'Procedência
Public BDAFN_TBLIN_CPEMA As Field 'e-mail
Public BDAFN_TBLIN_CPEDI As Field 'Editado
Public BDAFN_TBLIN_CPCHU As Field 'Chupinhado
Public BDAFN_TBLIN_CPREE As Field 'Re-editado
'Campos da Tabela BDAFN_TBCAT
Public BDAFN_TBCAT_CPINO As Field
Public BDAFN_TBCAT_CPCAT As Field
Public BDAFN_TBCAT_CPDES As Field
Public BDAFN_TBCAT_CPLIN As Field
'Campos da Tabela BDAFN_TBNCT
Public BDAFN_TBNCT_CPNCT As Field


'Campos da Tabela BDATU_TBLIN
Public BDATU_TBLIN_CPIND As Field 'Índice Novo
Public BDATU_TBLIN_CPLIN As Field 'Link
Public BDATU_TBLIN_CPDAT As Field 'Data
Public BDATU_TBLIN_CPTIT As Field 'Título
Public BDATU_TBLIN_CPDES As Field 'Descricao
Public BDATU_TBLIN_CPCAT As Field 'Categoria
Public BDATU_TBLIN_CPLGA As Field 'Lingua
Public BDATU_TBLIN_CPPRO As Field 'Procedência
Public BDATU_TBLIN_CPEMA As Field 'e-mail
Public BDATU_TBLIN_CPEDI As Field 'Editado
Public BDATU_TBLIN_CPCHU As Field 'Chupinhado
Public BDATU_TBLIN_CPREE As Field 'Re-editado
'Campos da Tabela BDATU_TBCAT
Public BDATU_TBCAT_CPINO As Field
Public BDATU_TBCAT_CPCAT As Field
Public BDATU_TBCAT_CPDES As Field
Public BDATU_TBCAT_CPLIN As Field
'Campos da Tabela BDATU_TBNCT
Public BDATU_TBNCT_CPNCT As Field


'****************************************************************
'                           F U N Ç Õ E S
'****************************************************************
Public Function AbreBD() As Boolean
    AbreBD = False
    'On Error GoTo ERRO_ABREBD
    'Abre Workspace
    Set WKSP = DBEngine.Workspaces(0)
    'Abre Bancos de Dados
    Set BDATF = WKSP.OpenDatabase(App.path & "\BDAF.af", , False)
    Set BDAFN = WKSP.OpenDatabase(App.path & "\BDNovos.af", , False)
    'Abre tabelas de BDATF
    Set BDATF_TBLIN = BDATF.OpenRecordset("Links")
    Set BDATF_TBCAT = BDATF.OpenRecordset("Categorias")
    'Abre tabelas de BDAFN
    Set BDAFN_TBLIN = BDAFN.OpenRecordset("Links")
    Set BDAFN_TBCAT = BDAFN.OpenRecordset("Categorias")
    Set BDAFN_TBNCT = BDAFN.OpenRecordset("NumCat")
    'Abre Índices de BDATF
    BDATF_TBLIN.Index = "Link"
    BDATF_TBCAT.Index = "Categoria"
    'Abre Índices de BDAFN
    BDAFN_TBLIN.Index = "Link"
    BDAFN_TBCAT.Index = "Categoria"
    'Abre campos da tabela BDATF_TBLIN
    Set BDATF_TBLIN_CPIND = BDATF_TBLIN.Fields("Índice")
    Set BDATF_TBLIN_CPLIN = BDATF_TBLIN.Fields("Link")
    Set BDATF_TBLIN_CPDAT = BDATF_TBLIN.Fields("Data")
    Set BDATF_TBLIN_CPTIT = BDATF_TBLIN.Fields("Título")
    Set BDATF_TBLIN_CPDES = BDATF_TBLIN.Fields("Descrição")
    Set BDATF_TBLIN_CPCAT = BDATF_TBLIN.Fields("Categoria")
    Set BDATF_TBLIN_CPLGA = BDATF_TBLIN.Fields("Língua")
    Set BDATF_TBLIN_CPPRO = BDATF_TBLIN.Fields("Procedência")
    Set BDATF_TBLIN_CPEMA = BDATF_TBLIN.Fields("e-mail")
    'Abre campos da tabela BDATF_TBCAT
    Set BDATF_TBCAT_CPIND = BDATF_TBCAT.Fields("Índice")
    Set BDATF_TBCAT_CPCAT = BDATF_TBCAT.Fields("Categoria")
    Set BDATF_TBCAT_CPDES = BDATF_TBCAT.Fields("Descrição")
    Set BDATF_TBCAT_CPLIN = BDATF_TBCAT.Fields("Links")
    'Abre campos da tabela BDAFN_TBLIN
    Set BDAFN_TBLIN_CPIND = BDAFN_TBLIN.Fields("Índice Novo")
    Set BDAFN_TBLIN_CPLIN = BDAFN_TBLIN.Fields("Link")
    Set BDAFN_TBLIN_CPDAT = BDAFN_TBLIN.Fields("Data")
    Set BDAFN_TBLIN_CPTIT = BDAFN_TBLIN.Fields("Título")
    Set BDAFN_TBLIN_CPDES = BDAFN_TBLIN.Fields("Descrição")
    Set BDAFN_TBLIN_CPCAT = BDAFN_TBLIN.Fields("Categoria")
    Set BDAFN_TBLIN_CPLGA = BDAFN_TBLIN.Fields("Língua")
    Set BDAFN_TBLIN_CPPRO = BDAFN_TBLIN.Fields("Procedência")
    Set BDAFN_TBLIN_CPEMA = BDAFN_TBLIN.Fields("e-mail")
    Set BDAFN_TBLIN_CPEDI = BDAFN_TBLIN.Fields("Editado")
    Set BDAFN_TBLIN_CPCHU = BDAFN_TBLIN.Fields("Chupinhado")
    Set BDAFN_TBLIN_CPREE = BDAFN_TBLIN.Fields("Re-editado")
    'Abre campos da tabela BDAFN_TBCAT
    Set BDAFN_TBCAT_CPINO = BDAFN_TBCAT.Fields("Índice Novo")
    Set BDAFN_TBCAT_CPCAT = BDAFN_TBCAT.Fields("Categoria")
    Set BDAFN_TBCAT_CPDES = BDAFN_TBCAT.Fields("Descrição")
    Set BDAFN_TBCAT_CPLIN = BDAFN_TBCAT.Fields("Links")
    'Abre campos da tabela BDAFN_TBNCT
    Set BDAFN_TBNCT_CPNCT = BDAFN_TBNCT.Fields("NumCat")
    
    AbreBD = True
ERRO_ABREBD:
    If Err Then MsgBox "Não foi possível acessar os bancos de dados do Atomic Link. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Abertura"
End Function
Public Function FechaBD() As Boolean
    FechaBD = False
    On Error GoTo ERRO_FECHABD
    BDATF_TBLIN.Close
    BDATF_TBCAT.Close
    BDAFN_TBLIN.Close
    BDAFN_TBCAT.Close
    BDAFN_TBNCT.Close
    BDATF.Close
    BDAFN.Close
    FechaBD = True
ERRO_FECHABD:
    If Err Then MsgBox "Aconteceu algum erro ao fechar o banco de dados do Atomic Link.", vbCritical + vbOKOnly, "Erro de Encerramento"
End Function
Public Function AbreBD_Atualizacao_Novos(ArquivoAtualizacaoNovo As String) As Boolean
    AbreBD_Atualizacao_Novos = False
    On Error GoTo ERRO_ABREBD
    'Abre Bancos de Dados
    Set BDATU = WKSP.OpenDatabase(ArquivoAtualizacaoNovo, , False)
    'Abre tabelas de BDATU
    Set BDATU_TBLIN = BDATU.OpenRecordset("Links")
    Set BDATU_TBCAT = BDATU.OpenRecordset("Categorias")
    Set BDATU_TBNCT = BDATU.OpenRecordset("NumCat")
    'Abre Índices de BDATU
    BDATU_TBLIN.Index = "Link"
    BDATU_TBCAT.Index = "Categoria"
    'Abre campos da tabela BDATU_TBLIN
    Set BDATU_TBLIN_CPIND = BDATU_TBLIN.Fields("Índice Novo")
    Set BDATU_TBLIN_CPLIN = BDATU_TBLIN.Fields("Link")
    Set BDATU_TBLIN_CPDAT = BDATU_TBLIN.Fields("Data")
    Set BDATU_TBLIN_CPTIT = BDATU_TBLIN.Fields("Título")
    Set BDATU_TBLIN_CPDES = BDATU_TBLIN.Fields("Descrição")
    Set BDATU_TBLIN_CPCAT = BDATU_TBLIN.Fields("Categoria")
    Set BDATU_TBLIN_CPLGA = BDATU_TBLIN.Fields("Língua")
    Set BDATU_TBLIN_CPPRO = BDATU_TBLIN.Fields("Procedência")
    Set BDATU_TBLIN_CPEMA = BDATU_TBLIN.Fields("e-mail")
    Set BDATU_TBLIN_CPEDI = BDATU_TBLIN.Fields("Editado")
    Set BDATU_TBLIN_CPCHU = BDATU_TBLIN.Fields("Chupinhado")
    Set BDATU_TBLIN_CPREE = BDATU_TBLIN.Fields("Re-editado")
    'Abre campos da tabela BDATU_TBCAT
    Set BDATU_TBCAT_CPINO = BDATU_TBCAT.Fields("Índice Novo")
    Set BDATU_TBCAT_CPCAT = BDATU_TBCAT.Fields("Categoria")
    Set BDATU_TBCAT_CPDES = BDATU_TBCAT.Fields("Descrição")
    Set BDATU_TBCAT_CPLIN = BDATU_TBCAT.Fields("Links")
    'Abre campos da tabela BDATU_TBNCT
    Set BDATU_TBNCT_CPNCT = BDATU_TBNCT.Fields("NumCat")
    AbreBD_Atualizacao_Novos = True
ERRO_ABREBD:
    If Err Then MsgBox "Não foi possível acessar o banco de dados de novos links para atualização do Atomic Link. Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Abertura"
End Function
Public Function FechaBD_Atualizacao_Novos() As Boolean
    FechaBD_Atualizacao_Novos = False
    On Error GoTo ERRO_FECHABD
    BDATU_TBLIN.Close
    BDATU_TBCAT.Close
    BDATU_TBNCT.Close
    BDATU.Close
    FechaBD_Atualizacao_Novos = True
ERRO_FECHABD:
    If Err Then MsgBox "Aconteceu algum erro ao fechar o banco de dados do Atomic Link.", vbCritical + vbOKOnly, "Erro de Encerramento"
End Function
