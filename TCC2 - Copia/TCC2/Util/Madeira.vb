'orientação objeto para uma classe de madeira
'criando um objeto para definir propriedades e metódos.
'Objeto é definido por uma classe que descreve as variaveis.
'Classe onde espera as caracteristicas da madeira, que confrome a escolha que o usuario faz da madeira, manda as informações atraves dessa classe
Public Class Madeira
    'Public Shared id As Integer = 0
    'Public Shared name As String = ""
    'Public Shared resistCompParalela As Double = 0 'serve para receber valor e mandar quando pedir 
    'Public Shared resistTracaoParalela As Double = 0
    'Public Shared resistTracaoNormal As Double = 0
    'Public Shared resistAoCisalhamento As Double = 0
    'Public Shared modElasticidade As Double = 0
    'Public Shared densAparente As Double = 0
    'Public Shared propriedadesGeometricas As PropriedadesGeometricas = New PropriedadesGeometricas()
    'Public Shared tracao As Tracao = New Tracao()
    'Public Shared compressao As Compressao = New Compressao()
    'Public Shared cisalhamento As Cisalhamento = New Cisalhamento()
    'Public Shared flexao As Flexao = New Flexao()
    'Public Shared resistenciaCalculo As ResistenciaCalculo = New ResistenciaCalculo()
    Private _id As Integer = 0
    Private _name As String = ""
    Private _resistCompParalela As Double = 0 'serve para receber valor e mandar quando pedir 
    Private _resistTracaoParalela As Double = 0
    Private _resistTracaoNormal As Double = 0
    Private _resistAoCisalhamento As Double = 0
    Private _modElasticidade As Double = 0
    Private _densAparente As Double = 0
    Private _propriedadesGeometricas As PropriedadesGeometricas = New PropriedadesGeometricas()
    Private _tracao As Tracao = New Tracao()
    Private _compressao As Compressao = New Compressao()
    Private _cisalhamento As Cisalhamento = New Cisalhamento()
    Private _flexao As Flexao = New Flexao()
    Private _resistenciaCalculo As ResistenciaCalculo = New ResistenciaCalculo()

    'ENUM: fornece uma forma de tratar um conjunto relacionado de constantes e de associar os valores destas constantes com os nomes atribuídos a elas.
    Enum TipoSecao
        Retangular
        Circular
        SecaoT
        SecaoI
        Caixao
        ElementosJustaposto2
        ElementosJustaposto3
    End Enum

    'função retorna string.
    '"As" precisa de um retorno
    Public Sub New()

    End Sub
    Public Sub New(ByVal id As Integer,
                   ByVal name As String,
                   ByVal resistCompParalela As Double,
                   ByVal resistTracaoParalela As Double,
                   ByVal resistTracaoNormal As Double,
                   ByVal resistAoCisalhamento As Double,
                   ByVal modElasticidade As Double,
                   ByVal densAparente As Double)
        _id = id
        _name = name
        _resistCompParalela = resistCompParalela
        _resistTracaoParalela = resistTracaoParalela
        _resistTracaoNormal = resistTracaoNormal
        _resistAoCisalhamento = resistAoCisalhamento
        _modElasticidade = modElasticidade
        _densAparente = densAparente
    End Sub

    Public Shared ListaEspecies = New List(Of Madeira) From {
                New Madeira(0, "Angelim araroba", 50.5, 69.2, 3.1, 7.1, 12876, 688),
                New Madeira(1, "Angelim ferro", 79.9, 117.8, 3.7, 11.8, 208727, 1170),
                New Madeira(2, "Angelim pedra", 59.8, 75.5, 3.5, 8.8, 12912, 694),
                New Madeira(3, "Angelim pedra verdadeiro", 76.7, 104.9, 4.8, 11.3, 16694, 1170),
                New Madeira(4, "Branquilho", 48.1, 87.9, 3.2, 9.8, 13481, 803),
                New Madeira(5, "Cafearana", 59.1, 79.7, 3.0, 5.9, 14098, 677),
                New Madeira(6, "Canafístula", 52.0, 84.9, 6.2, 11.1, 14613, 871),
                New Madeira(7, "Casca Grossa", 56.0, 120.2, 4.1, 8.2, 16224, 801),
                New Madeira(8, "Castelo", 54.8, 99.5, 7.5, 12.8, 11105, 759),
                New Madeira(9, "Cedro amargo", 39.0, 58.1, 3.0, 6.1, 9839, 504),
                New Madeira(10, "Cedo doce", 31.5, 71.4, 3.0, 5.6, 8058, 500),
                New Madeira(11, "Champagne", 93.2, 133.5, 2.9, 10.7, 23002, 1090),
                New Madeira(12, "Cupiúna", 54.4, 62.1, 3.3, 10.4, 13627, 838),
                New Madeira(13, "Catiúna", 83.8, 86.2, 3.3, 11.1, 1942, 1221),
                New Madeira(14, "Eucalipto Alba", 47.3, 69.4, 4.6, 9.5, 13409, 705),
                New Madeira(15, "Eucalipto Camaldulensis", 48.0, 78.1, 4.6, 9.0, 13286, 899),
                New Madeira(16, "Eucalipto Citriodora", 62.0, 123.6, 3.9, 10.7, 18421, 999),
                New Madeira(17, "Eucalipto Cloeziana ", 51.8, 90.8, 4.0, 10.5, 13963, 822),
                New Madeira(18, "Eucalipto Dunni", 48.9, 139.2, 6.9, 9.8, 18029, 690),
                New Madeira(19, "Eucalipto Grandis", 40.3, 70.2, 2.6, 7.0, 12813, 640),
                New Madeira(20, "Eucalipto Maculata", 63.5, 115.6, 4.1, 10.6, 18099, 931),
                New Madeira(21, "Eucalipto Maidene", 48.3, 83.7, 4.8, 10.3, 14431, 924),
                New Madeira(22, "Eucalipto Microcorys", 54.9, 118.6, 4.5, 10.3, 16782, 929),
                New Madeira(23, "Eucalipto Paniculata", 72.7, 147.4, 4.7, 12.4, 19881, 1087),
                New Madeira(24, "Eucalipto Propinqua", 51.6, 89.1, 4.7, 9.7, 15561, 952),
                New Madeira(25, "Eucalipto Punctata", 78.5, 125.6, 6.0, 12.9, 19360, 948),
                New Madeira(26, "Eucalipto Saligna", 46.8, 95.5, 4.0, 8.2, 14933, 731),
                New Madeira(27, "Eucalipto Terticornis", 57.7, 115.9, 4.6, 9.7, 17198, 899),
                New Madeira(28, "Eucalipto Triantha", 53.9, 100.9, 2.7, 9.2, 14617, 755),
                New Madeira(29, "Eucalipto Umbra", 42.7, 90.4, 3.0, 9.4, 14577, 889),
                New Madeira(30, "Eucalipto Urophylla", 46.0, 85.1, 4.1, 8.3, 13166, 739),
                New Madeira(31, "Garapa Roraima", 78.4, 108.0, 6.9, 11.9, 18359, 892),
                New Madeira(32, "Guaiçara", 71.4, 115.6, 4.2, 12.5, 14624, 825),
                New Madeira(33, "Guarucaia", 62.4, 70.9, 5.5, 15.5, 17212, 919),
                New Madeira(34, "Ipê", 76.0, 96.8, 3.1, 13.1, 18011, 1068),
                New Madeira(35, "Jatobá", 93.3, 157.5, 3.2, 15.7, 23607, 1074),
                New Madeira(36, "Louro Preto", 56.5, 111.9, 3.3, 9.0, 14185, 684),
                New Madeira(37, "Maçaranduba", 82.9, 138.5, 5.4, 14.9, 22733, 1143),
                New Madeira(38, "Mandioqueira", 71.4, 89.1, 2.7, 10.6, 18971, 856),
                New Madeira(39, "Oiticica Amarela", 69.9, 82.5, 3.9, 10.6, 14719, 756),
                New Madeira(40, "Quarubarana", 37.8, 58.1, 2.6, 5.8, 9067, 544),
                New Madeira(41, "Sucupira", 95.2, 123.4, 3.4, 11.8, 21724, 1106),
                New Madeira(42, "Tatajuba", 79.5, 78.8, 3.9, 12.2, 19583, 940),
                New Madeira(44, "Pinho do Paraná", 40.9, 93.1, 1.6, 8.8, 15225, 580),
                New Madeira(45, "Pinus Caribea", 35.4, 64.8, 3.2, 7.8, 8431, 579),
                New Madeira(46, "Pinus Bahamensis", 32.6, 52.7, 2.4, 6.8, 7110, 537),
                New Madeira(47, "Pinus Hondurensis", 42.3, 50.3, 2.6, 7.8, 9868, 535),
                New Madeira(48, "Pinus Elliotti", 40.4, 66.0, 2.5, 7.4, 11889, 560),
                New Madeira(49, "Pinus Oocarpa", 43.6, 60.9, 2.5, 8.0, 10904, 538),
                New Madeira(50, "Pinus Taeda", 44.4, 82.8, 2.8, 7.7, 13304, 645),
                New Madeira(50, "Coníferas - C20", 20, 0, 0, 4, 3500, 400),
                New Madeira(50, "Coníferas - C25", 25, 0, 0, 5, 8500, 450),
                New Madeira(50, "Coníferas - C30", 30, 0, 0, 6, 14500, 500),
                New Madeira(50, "Dicotiledôneas C-20", 20, 0, 0, 4, 9500, 500),
                New Madeira(50, "Dicotiledôneas C-30", 30, 0, 0, 5, 14500, 650),
                New Madeira(50, "Dicotiledôneas C-40", 40, 0, 0, 6, 19500, 750),
                New Madeira(50, "Dicotiledôneas C-60", 60, 0, 0, 8, 24500, 800)
    }

    'get: precisa retornar algo (chama um valor)
    'set: digita o valot 
    'lista: coleção de algo - duas formas: 
    'primeiro: coleção de algo (clica no inserir valores)
    'segundo: vazio (sem passar valor nenhum)
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    Public Property ResistCompParalela() As Double
        Get
            Return _resistCompParalela
        End Get
        Set(value As Double)
            _resistCompParalela = value
        End Set
    End Property
    Public Property ResistTracaoParalela() As Double
        Get
            Return _resistTracaoParalela
        End Get
        Set(value As Double)
            _resistTracaoParalela = value
        End Set
    End Property
    Public Property ResistTracaoNormal() As Double
        Get
            Return _resistTracaoNormal
        End Get
        Set(value As Double)
            _resistTracaoNormal = value
        End Set
    End Property
    Public Property ResistAoCisalhamento() As Double
        Get
            Return _resistAoCisalhamento
        End Get
        Set(value As Double)
            _resistAoCisalhamento = value
        End Set
    End Property
    Public Property ModElasticidade() As Double
        Get
            Return _modElasticidade
        End Get
        Set(value As Double)
            _modElasticidade = value
        End Set
    End Property
    Public Property DensAparente() As Double
        Get
            Return _densAparente
        End Get
        Set(value As Double)
            _densAparente = value
        End Set
    End Property
    Public Property PropriedadesGeometricas() As PropriedadesGeometricas
        Get
            Return _propriedadesGeometricas
        End Get
        Set(value As PropriedadesGeometricas)
            _propriedadesGeometricas = value
        End Set
    End Property
    Public Property Tracao() As Tracao
        Get
            Return _tracao
        End Get
        Set(value As Tracao)
            _tracao = value
        End Set
    End Property
    Public Property Compressao() As Compressao
        Get
            Return _compressao
        End Get
        Set(value As Compressao)
            _compressao = value
        End Set
    End Property
    Public Property Cisalhamento() As Cisalhamento
        Get
            Return _cisalhamento
        End Get
        Set(value As Cisalhamento)
            _cisalhamento = value
        End Set
    End Property
    Public Property Flexao() As Flexao
        Get
            Return _flexao
        End Get
        Set(value As Flexao)
            _flexao = value
        End Set
    End Property
    Public Property ResistenciaCalculo() As ResistenciaCalculo
        Get
            Return _resistenciaCalculo
        End Get
        Set(value As ResistenciaCalculo)
            _resistenciaCalculo = value
        End Set
    End Property

End Class
