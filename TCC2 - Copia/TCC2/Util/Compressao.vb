Public Class Compressao
	Private _tensaoCompressao As Double = 0
	Private _tensaoCompressaoCurtaX As Double = 0
	Private _tensaoCompressaoCurtaY As Double = 0
	Private _tensaoCompressaoMedEsbeltaX As Double = 0
	Private _tensaoCompressaoMedEsbeltaY As Double = 0
	Private _tensaoCompressaoEsbeltaX As Double = 0
	Private _tensaoCompressaoEsbeltaY As Double = 0
	Private _tensaoMxd As Double = 0
	Private _tensaoMyd As Double = 0
	Private _esbeltezCompressaoX As Double = 0
	Private _esbeltezCompressaoY As Double = 0
	Private _forcaElasticaX As Double = 0
	Private _forcaElasticaY As Double = 0
	Private _lvinculado As Double = 0
	Private _ei As Double = 0 'excentricidade inicial
	Private _ea As Double = 0 'excentricidade acidental
	Private _edx As Double = 0 'excentricidade de projeto em x 
	Private _edy As Double = 0 'excentricidade de projeto em y
	Private _ec As Double = 0 'excentricidade suplementar de primeira ordem
	Private _e1ef As Double = 0 'excentricidade de primeira ordem
	Private _eig As Double = 0 'excentricidade de primeira ordem devido as cargas permanentes

	Public Sub New()

	End Sub

	Public Sub New(ByVal tensaoCompressao As Double,
				   ByVal tensaoCompressaoCurtaX As Double,
				   ByVal tensaoCompressaoCurtaY As Double,
				   ByVal tensaoCompressaoMedEsbeltaX As Double,
				   ByVal tensaoCompressaoMedEsbeltaY As Double,
				   ByVal tensaoCompressaoEsbeltaX As Double,
				   ByVal tensaoCompressaoEsbeltaY As Double,
				   ByVal tensaoMxd As Double,
				   ByVal tensaoMyd As Double,
				   ByVal esbeltezCompressaoX As Double,
				   ByVal esbeltezCompressaoY As Double,
				   ByVal forcaElasticaX As Double,
				   ByVal forcaElasticaY As Double,
				   ByVal lvinculado As Double,
				   ByVal ei As Double,
				   ByVal ea As Double,
				   ByVal edx As Double,
				   ByVal edy As Double,
				   ByVal e1ef As Double,
				   ByVal ec As Double,
				   ByVal eig As Double)

		_tensaoCompressao = tensaoCompressao
		_tensaoCompressaoCurtaX = tensaoCompressaoCurtaX
		_tensaoCompressaoCurtaY = tensaoCompressaoCurtaY
		_tensaoCompressaoMedEsbeltaX = tensaoCompressaoMedEsbeltaX
		_tensaoCompressaoMedEsbeltaY = tensaoCompressaoMedEsbeltaY
		_tensaoCompressaoEsbeltaX = tensaoCompressaoEsbeltaX
		_tensaoCompressaoEsbeltaY = tensaoCompressaoEsbeltaY
		_tensaoMxd = tensaoMxd
		_tensaoMyd = tensaoMyd
		_esbeltezCompressaoX = esbeltezCompressaoX
		_esbeltezCompressaoY = esbeltezCompressaoY
		_forcaElasticaX = forcaElasticaX
		_forcaElasticaY = forcaElasticaY
		_lvinculado = lvinculado
		_ei = ei
		_ea = ea
		_edX = edx
		_edy = edy
		_e1ef = e1ef
		_ec = ec
		_eig = eig
	End Sub


	Public Property TensaoCompressao() As Double
		Get
			Return _tensaoCompressao
		End Get
		Set(value As Double)
			_tensaoCompressao = value
		End Set
	End Property
	Public Property TensaoMxd() As Double
		Get
			Return _tensaoMxd
		End Get
		Set(value As Double)
			_tensaoMxd = value
		End Set
	End Property
	Public Property TensaoMyd() As Double
		Get
			Return _tensaoMyd
		End Get
		Set(value As Double)
			_tensaoMyd = value
		End Set
	End Property
	Public Property EsbeltezCompressaoX() As Double
		Get
			Return _esbeltezCompressaoX
		End Get
		Set(value As Double)
			_esbeltezCompressaoX = value
		End Set
	End Property

	Public Property EsbeltezCompressaoY() As Double
		Get
			Return _esbeltezCompressaoY
		End Get
		Set(value As Double)
			_esbeltezCompressaoY = value
		End Set
	End Property

	Public Property ForcaElasticaX() As Double
		Get
			Return _forcaElasticaX
		End Get
		Set(value As Double)
			_forcaElasticaX = value
		End Set
	End Property

	Public Property ForcaElasticaY() As Double
		Get
			Return _forcaElasticaY
		End Get
		Set(value As Double)
			_forcaElasticaY = value
		End Set
	End Property

	Public Property Lvinculado() As Double
		Get
			Return _lvinculado
		End Get
		Set(value As Double)
			_lvinculado = value
		End Set
	End Property

	Public Property Ei() As Double
		Get
			Return _ei
		End Get
		Set(value As Double)
			_ei = value
		End Set
	End Property
	Public Property Ea() As Double
		Get
			Return _ea
		End Get
		Set(value As Double)
			_ea = value
		End Set
	End Property

	Public Property Edx() As Double
		Get
			Return _edx
		End Get
		Set(value As Double)
			_edx = value
		End Set
	End Property
	Public Property Edy() As Double
		Get
			Return _edy
		End Get
		Set(value As Double)
			_edy = value
		End Set
	End Property
	Public Property E1ef() As Double
		Get
			Return _e1ef
		End Get
		Set(value As Double)
			_e1ef = value
		End Set
	End Property

	Public Property Ec() As Double
		Get
			Return _ec
		End Get
		Set(value As Double)
			_ec = value
		End Set
	End Property

	Public Property Eig() As Double
		Get
			Return _eig
		End Get
		Set(value As Double)
			_eig = value
		End Set
	End Property

	Public Property TensaoCompressaoCurtaX() As Double
		Get
			Return _tensaoCompressaoCurtaX
		End Get
		Set(value As Double)
			_tensaoCompressaoCurtaX = value
		End Set
	End Property
	Public Property TensaoCompressaoCurtaY() As Double
		Get
			Return _tensaoCompressaoCurtaY
		End Get
		Set(value As Double)
			_tensaoCompressaoCurtaY = value
		End Set
	End Property
	Public Property TensaoCompressaoMedEsbeltaX() As Double
		Get
			Return _tensaoCompressaoMedEsbeltaX
		End Get
		Set(value As Double)
			_tensaoCompressaoMedEsbeltaX = value
		End Set
	End Property
	Public Property TensaoCompressaoMedEsbeltaY() As Double
		Get
			Return _tensaoCompressaoMedEsbeltaY
		End Get
		Set(value As Double)
			_tensaoCompressaoMedEsbeltaY = value
		End Set
	End Property
	Public Property TensaoCompressaoEsbeltaX() As Double
		Get
			Return _tensaoCompressaoEsbeltaX
		End Get
		Set(value As Double)
			_tensaoCompressaoEsbeltaX = value
		End Set
	End Property
	Public Property TensaoCompressaoEsbeltaY() As Double
		Get
			Return _tensaoCompressaoEsbeltaY
		End Get
		Set(value As Double)
			_tensaoCompressaoEsbeltaY = value
		End Set
	End Property

End Class
