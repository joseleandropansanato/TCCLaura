Public Class Cisalhamento

	Private _tensaoCisalhanteX As Double = 0
	Private _tensaoCisalhanteY As Double = 0

	Public Sub New()
		'comentario
	End Sub

	Public Sub New(ByVal tensaoCisalhanteX As Double,
				   ByVal tensaoCisalhanteY As Double
				   )
		_tensaoCisalhanteX = tensaoCisalhanteX
		_tensaoCisalhanteY = tensaoCisalhanteY
	End Sub


	Public Property TensaoCisalhanteX() As Double
		Get
			Return _tensaoCisalhanteX
		End Get
		Set(value As Double)
			_tensaoCisalhanteX = value
		End Set
	End Property

	Public Property TensaoCisalhanteY() As Double
		Get
			Return _tensaoCisalhanteY
		End Get
		Set(value As Double)
			_tensaoCisalhanteY = value
		End Set
	End Property





End Class
