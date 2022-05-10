Public Class Flexao


	Private _tensaoFlexaoX As Double = 0
	Private _tensaoFlexaoY As Double = 0

	Public Sub New()

	End Sub

	Public Sub New(ByVal tensaoFlexaoX As Double,
				   ByVal tensaoFlexaoY As Double)


		_tensaoFlexaoX = tensaoFlexaoX
		_tensaoFlexaoY = tensaoFlexaoY


	End Sub


	Public Property tensaoFlexaoX() As Double
		Get
			Return _tensaoFlexaoX
		End Get
		Set(value As Double)
			_tensaoFlexaoX = value
		End Set
	End Property

	Public Property tensaoFlexaoY() As Double
		Get
			Return _tensaoFlexaoY
		End Get
		Set(value As Double)
			_tensaoFlexaoY = value
		End Set
	End Property

End Class
