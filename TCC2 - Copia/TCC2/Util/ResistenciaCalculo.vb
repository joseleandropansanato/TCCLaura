Public Class ResistenciaCalculo

	Private _resistCalculoTracaoParalela As Double = 0
	Private _resistCalculoTracaNormal As Double = 0
	Private _resistCalculoCompressaoParalela As Double = 0
	Private _resistCalculoCompressaoNormal As Double = 0
	Private _resistCalculoAoCisalhamento As Double = 0
	Private _moduloElasticidade As Double = 0
	Private _moduloElasticidadeTransversal As Double = 0
	Private _densidadeAparente As Double = 0
	Public Sub New()

    End Sub

	Public Property resistCalculoTracaoParalela() As Double
		Get
			Return _resistCalculoTracaoParalela
		End Get
		Set(value As Double)
			_resistCalculoTracaoParalela = value
		End Set
	End Property

	Public Property resistCalculoTracaNormal() As Double
		Get
			Return _resistCalculoTracaNormal
		End Get
		Set(value As Double)
			_resistCalculoTracaNormal = value
		End Set
	End Property

	Public Property resistCalculoCompressaoParalela() As Double
		Get
			Return _resistCalculoCompressaoParalela
		End Get
		Set(value As Double)
			_resistCalculoCompressaoParalela = value
		End Set
	End Property

	Public Property resistCalculoCompressaoNormal() As Double
		Get
			Return _resistCalculoCompressaoNormal
		End Get
		Set(value As Double)
			_resistCalculoCompressaoNormal = value
		End Set
	End Property

	Public Property resistCalculoAoCisalhamento() As Double
		Get
			Return _resistCalculoAoCisalhamento
		End Get
		Set(value As Double)
			_resistCalculoAoCisalhamento = value
		End Set
	End Property

	Public Property moduloElasticidade() As Double
		Get
			Return _moduloElasticidade
		End Get
		Set(value As Double)
			_moduloElasticidade = value
		End Set
	End Property

	Public Property moduloElasticidadeTransversal() As Double
		Get
			Return _moduloElasticidadeTransversal
		End Get
		Set(value As Double)
			_moduloElasticidadeTransversal = value
		End Set
	End Property

	Public Property densidadeAparente() As Double
		Get
			Return _densidadeAparente
		End Get
		Set(value As Double)
			_densidadeAparente = value
		End Set
	End Property

End Class
