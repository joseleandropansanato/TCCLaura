Public Class PropriedadesGeometricas

	'serve para receber valor e mandar quando pedir 
	Private _area As Double = 0
	Private _areaA1 As Double = 0
	Private _eixoXmi As Double = 0 'EIXO X MOMENTO DE INERCIA
	Private _eixoXmi1 As Double = 0 'EIXO X MOMENTO DE INERCIA PRINCIPAL EM X
	Private _eixoXrg As Double = 0 'EIXO X RAIO DE GIRAÇÃO
	Private _eixoXmr As Double = 0 'EIXO X 
	Private _eixoYmi As Double = 0 'EIXO Y MOMENTO DE INERCIA
	Private _eixoYrg As Double = 0 'EIXO Y RAIO DE GIRAÇÃO
	Private _eixoYmr As Double = 0 'EIXO Y MODULO DE RESISTENCIA
	Private _eixoXie As Double = 0 'EIXO X INERCIA EFETIVA
	Private _eixoYie As Double = 0 'EIXO Y INERCIA EFETIVA
	Private _eixoYmie As Double = 0 'EIXO Y MOMENTO DE INERCIA EFETIVA EM Y
	Private _eixoYmi1 As Double = 0 'EIXO Y MOMENTO DE INERCIA PRINCIPAL EM Y
	Private _coefBy As Double = 0 'COEFICIENTE DO ESPAÇADOR
	'Private 

	Enum Pilar
		Espaçador
		Chapa
	End Enum

	Public Sub New()

	End Sub
	Public Sub New(ByVal area As Double,
				   ByVal areaA1 As Double,
				   ByVal eixoXmi As Double,
				   ByVal eixoXmi1 As Double,
				   ByVal eixoXrg As Double,
				   ByVal eixoXmr As Double,
				   ByVal eixoYmi As Double,
				   ByVal eixoYrg As Double,
				   ByVal eixoYmr As Double,
				   ByVal eixoXie As Double,
				   ByVal eixoYie As Double,
				   ByVal eixoYmie As Double,
				   ByVal eixoYmi1 As Double,
				   ByVal coefBy As Double)

		_area = area
		_areaA1 = areaA1
		_eixoXmi = eixoXmi
		_eixoXmi1 = eixoXmi1
		_eixoXrg = eixoXrg
		_eixoXmr = eixoXmr
		_eixoYmi = eixoYmi
		_eixoYrg = eixoYrg
		_eixoYmr = eixoYmr
		_eixoXie = eixoXie
		_eixoYie = eixoYie
		_eixoYmie = eixoYmie
		_eixoYmi1 = eixoYmi1
		_coefBy = coefBy


	End Sub

	Public Property Area() As Double
		Get
			Return _area
		End Get
		Set(value As Double)
			_area = value
		End Set
	End Property

	Public Property AreaA1() As Double
		Get
			Return _areaA1
		End Get
		Set(value As Double)
			_areaA1 = value
		End Set
	End Property

	Public Property EixoXmi() As Double
		Get
			Return _eixoXmi
		End Get
		Set(value As Double)
			_eixoXmi = value
		End Set
	End Property

	Public Property EixoXmi1() As Double
		Get
			Return _eixoXmi1
		End Get
		Set(value As Double)
			_eixoXmi1 = value
		End Set
	End Property
	Public Property EixoXrg() As Double
		Get
			Return _eixoXrg
		End Get
		Set(value As Double)
			_eixoXrg = value
		End Set
	End Property

	Public Property EixoXmr() As Double
		Get
			Return _eixoXmr
		End Get
		Set(value As Double)
			_eixoXmr = value
		End Set
	End Property

	Public Property EixoYmi() As Double
		Get
			Return _eixoYmi
		End Get
		Set(value As Double)
			_eixoYmi = value
		End Set
	End Property

	Public Property EixoYmi1() As Double
		Get
			Return _eixoYmi1
		End Get
		Set(value As Double)
			_eixoYmi1 = value
		End Set
	End Property
	Public Property EixoYrg() As Double
		Get
			Return _eixoYrg
		End Get
		Set(value As Double)
			_eixoYrg = value
		End Set
	End Property

	Public Property EixoYmr() As Double
		Get
			Return _eixoYmr
		End Get
		Set(value As Double)
			_eixoYmr = value
		End Set
	End Property

	Public Property EixoXie() As Double
		Get
			Return _eixoXie
		End Get
		Set(value As Double)
			_eixoXie = value
		End Set
	End Property

	Public Property EixoYie() As Double
		Get
			Return _eixoYie
		End Get
		Set(value As Double)
			_eixoYie = value
		End Set
	End Property

	Public Property EixoYmie() As Double
		Get
			Return _eixoYmie
		End Get
		Set(value As Double)
			_eixoYmie = value
		End Set
	End Property

	Public Property CoefBy() As Double
		Get
			Return _coefBy
		End Get
		Set(value As Double)
			_coefBy = value
		End Set
	End Property

End Class
