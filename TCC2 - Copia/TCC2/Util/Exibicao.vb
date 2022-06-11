Public Class Exibicao

    Private _exibeTracao As Boolean = True
    Private _exibeCompressao As Boolean = True
    Private _exibeCisalhamento As Boolean = True
    Private _exibeFlexao As Boolean = True

    Public Sub New()

    End Sub

    Public Property ExibeTracao() As Boolean
        Get
            Return _exibeTracao
        End Get
        Set(value As Boolean)
            _exibeTracao = value
        End Set
    End Property

    Public Property ExibeCompressao() As Boolean
        Get
            Return _exibeCompressao
        End Get
        Set(value As Boolean)
            _exibeCompressao = value
        End Set
    End Property

    Public Property ExibeCisalhamento() As Boolean
        Get
            Return _exibeCisalhamento
        End Get
        Set(value As Boolean)
            _exibeCisalhamento = value
        End Set
    End Property

    Public Property ExibeFlexao() As Boolean
        Get
            Return _exibeFlexao
        End Get
        Set(value As Boolean)
            _exibeFlexao = value
        End Set
    End Property

End Class
