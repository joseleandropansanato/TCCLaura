'Classe é usada para criar objetos
Public Class Kmod2

    Private _one As Double = 0
    Private _two As Double = 0
    Private _three As Double = 0
    Private _four As Double = 0

    Public Sub New()

    End Sub
    Public Sub New(
                ByVal one As Double,
                ByVal two As Double,
                ByVal three As Double,
                ByVal four As Double
            )
        _one = one
        _two = two
        _three = three
        _four = four

    End Sub
    Public Property One() As Double
        Get
            Return _one
        End Get
        Set(value As Double)
            _one = value
        End Set
    End Property
    Public Property Two() As Double
        Get
            Return _two
        End Get
        Set(value As Double)
            _two = value
        End Set
    End Property
    Public Property Three() As Double
        Get
            Return _three
        End Get
        Set(value As Double)
            _three = value
        End Set
    End Property

    Public Property Four() As Double
        Get
            Return _four
        End Get
        Set(value As Double)
            _four = value
        End Set
    End Property

End Class
