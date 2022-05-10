Public Class Tracao

    Private _tensaoTracao As Double = 0
	Private _esbeltezTracaoX As Double = 0
	Private _esbeltezTracaoY As Double = 0
	Private _eixoXPecaComposta As Double = 0
	Private _eixoYPecaComposta As Double = 0

	Public Sub New()

	End Sub

	Public Sub New(ByVal tensaoTracao As Double,
				   ByVal esbeltezTracaoX As Double,
				   ByVal esbeltezTracaoY As Double,
				   ByVal eixoXPecaComposta As Double,
				   ByVal eixoYPecaComposta As Double)

		_tensaoTracao = tensaoTracao
		_esbeltezTracaoX = esbeltezTracaoX
		_esbeltezTracaoY = esbeltezTracaoY
		_eixoXPecaComposta = eixoXPecaComposta
		_eixoYPecaComposta = eixoYPecaComposta

	End Sub


	Public Property TensaoTracao() As Double
		Get
			Return _tensaoTracao
		End Get
		Set(value As Double)
			_tensaoTracao = value
		End Set
	End Property

	Public Property EsbeltezTracaoX() As Double
		Get
			Return _esbeltezTracaoX
		End Get
		Set(value As Double)
			_esbeltezTracaoX = value
		End Set
	End Property

	Public Property EsbeltezTracaoY() As Double
		Get
			Return _esbeltezTracaoY
		End Get
		Set(value As Double)
			_esbeltezTracaoY = value
		End Set
	End Property
	Public Property EixoXPecaComposta() As Double
		Get
			Return _eixoXPecaComposta
		End Get
		Set(value As Double)
			_eixoYPecaComposta = value
		End Set
	End Property
	Public Property EixoYPecaComposta() As Double
		Get
			Return _eixoYPecaComposta
		End Get
		Set(value As Double)
			_eixoYPecaComposta = value
		End Set
	End Property

End Class
