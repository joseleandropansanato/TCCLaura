'Classe é usada para criar objetos
Public Class Kmod3

    Private _se As Double = 0
    Private _s1 As Double = 0
    Private _s2 As Double = 0
    Private _s3 As Double = 0
    Private _s4 As Double = 0

    Public Sub New()

    End Sub
    Public Sub New(
                ByVal se As Double,
                ByVal s1 As Double,
                ByVal s2 As Double,
                ByVal s3 As Double
             )
        _se = se
        _s1 = s1
        _s2 = s2
        _s3 = s3

    End Sub
    Public Property Se() As Double
        Get
            Return _se
        End Get
        Set(value As Double)
            _se = value
        End Set
    End Property
    Public Property S1() As Double
        Get
            Return _s1
        End Get
        Set(value As Double)
            _s1 = value
        End Set
    End Property
    Public Property S2() As Double
        Get
            Return _s2
        End Get
        Set(value As Double)
            _s2 = value
        End Set
    End Property
    Public Property S3() As Double
        Get
            Return _s3
        End Get
        Set(value As Double)
            _s3 = value
        End Set
    End Property

End Class
