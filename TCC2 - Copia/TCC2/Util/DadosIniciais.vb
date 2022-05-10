Public Class DadosIniciais
	'SEÇÃO T OU I
	Private _b1 As Double = 0 'bf2
	Private _b2 As Double = 0 'bw
	Private _b3 As Double = 0 'bf1
	Private _h1 As Double = 0 'hf2
	Private _h2 As Double = 0 'h
	Private _h3 As Double = 0 'hf1

	'caixão
	Private _b1c As Double = 0
	Private _b2c As Double = 0
	Private _b3c As Double = 0
	Private _b4c As Double = 0 'bf2-bw caixao
	Private _h1c As Double = 0
	Private _h2c As Double = 0
	Private _h3c As Double = 0
	Private _h4c As Double = 0 'hbf1 caixao
	Private _momentoEstaticoDeArea As Double = 0 'hf1



	Public Sub New()

	End Sub
	'SEÇÃO T
	Public Property b1() As Double
		Get
			Return _b1
		End Get
		Set(value As Double)
			_b1 = value
		End Set
	End Property

	Public Property b2() As Double
		Get
			Return _b2
		End Get
		Set(value As Double)
			_b2 = value
		End Set
	End Property
	Public Property b3() As Double
		Get
			Return _b3
		End Get
		Set(value As Double)
			_b3 = value
		End Set
	End Property
	Public Property h1() As Double
		Get
			Return _h1
		End Get
		Set(value As Double)
			_h1 = value
		End Set
	End Property

	Public Property h2() As Double
		Get
			Return _h2
		End Get
		Set(value As Double)
			_h2 = value
		End Set
	End Property
	Public Property h3() As Double
		Get
			Return _h3
		End Get
		Set(value As Double)
			_h3 = value
		End Set
	End Property

	'CAIXÃO
	Public Property b1c() As Double
		Get
			Return _b1c
		End Get
		Set(value As Double)
			_b1c = value
		End Set
	End Property

	Public Property b2c() As Double
		Get
			Return _b2c
		End Get
		Set(value As Double)
			_b2c = value
		End Set
	End Property
	Public Property b3c() As Double
		Get
			Return _b3c
		End Get
		Set(value As Double)
			_b3c = value
		End Set
	End Property
	Public Property b4c() As Double
		Get
			Return _b4c
		End Get
		Set(value As Double)
			_b4c = value
		End Set
	End Property
	Public Property h1c() As Double
		Get
			Return _h1c
		End Get
		Set(value As Double)
			_h1c = value
		End Set
	End Property

	Public Property h2c() As Double
		Get
			Return _h2c
		End Get
		Set(value As Double)
			_h2c = value
		End Set
	End Property
	Public Property h3c() As Double
		Get
			Return _h3c
		End Get
		Set(value As Double)
			_h3c = value
		End Set
	End Property
	Public Property h4c() As Double
		Get
			Return _h4c
		End Get
		Set(value As Double)
			_h4c = value
		End Set
	End Property
	Public Property momentoEstaticoDeArea() As Double
		Get
			Return _momentoEstaticodeArea
		End Get
		Set(value As Double)
			_momentoEstaticodeArea = value
		End Set
	End Property

End Class
