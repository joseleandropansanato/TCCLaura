'Classe é usada para criar objetos
Public Class Kmod1

    Private _permanente As Double = 0
    Private _longaDuracao As Double = 0
    Private _mediaDuracao As Double = 0
    Private _curtaDuracao As Double = 0
    Private _instantanea As Double = 0
    Public Shared fcok As Double = 0

    Public Sub New()

    End Sub
    Public Sub New(
            ByVal permanente As Double,
            ByVal longaDuracao As Double,
            ByVal mediaDuracao As Double,
            ByVal curtaDuracao As Double,
            ByVal instantanea As Double
        )
        _permanente = permanente
        _longaDuracao = longaDuracao
        _mediaDuracao = mediaDuracao
        _curtaDuracao = curtaDuracao
        _instantanea = instantanea

    End Sub
    Public Property Permanente() As Double
        Get
            Return _permanente
        End Get
        Set(value As Double)
            _permanente = value
        End Set
    End Property
    Public Property LongaDuracao() As Double
        Get
            Return _longaDuracao
        End Get
        Set(value As Double)
            _longaDuracao = value
        End Set
    End Property
    Public Property MediaDuracao() As Double
        Get
            Return _mediaDuracao
        End Get
        Set(value As Double)
            _mediaDuracao = value
        End Set
    End Property
    Public Property CurtaDuracao() As Double
        Get
            Return _curtaDuracao
        End Get
        Set(value As Double)
            _curtaDuracao = value
        End Set
    End Property
    Public Property Instantanea() As Double
        Get
            Return _instantanea
        End Get
        Set(value As Double)
            _instantanea = value
        End Set
    End Property

End Class
