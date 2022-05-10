'orientação objeto para uma classe de madeira
'criando um objeto para definir propriedades e metódos.
'Objeto é definido por uma classe que descreve as variaveis.
'Classe onde espera as caracteristicas da madeira, que confrome a escolha que o usuario faz da madeira, manda as informações atraves dessa classe
Public Class Madeira
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
	Private _dadosIniciais As DadosIniciais = New DadosIniciais()
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

	Public Property DadosIniciais() As DadosIniciais
		Get
			Return _dadosIniciais
		End Get
		Set(value As DadosIniciais)
			_dadosIniciais = value
		End Set
	End Property

End Class
