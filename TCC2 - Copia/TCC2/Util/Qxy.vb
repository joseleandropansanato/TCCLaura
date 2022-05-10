Class Qxy

    Dim _Qx As Double = 0
	Dim _Qy As Double = 0

	Public Sub New()

    End Sub

	Public Property Qx() As Double
		Get
			Return _Qx
		End Get
		Set(value As Double)
			_Qx = value
		End Set
	End Property

	Public Property Qy() As Double
		Get
			Return _Qy
		End Get
		Set(value As Double)
			_Qy = value
		End Set
	End Property


End Class
